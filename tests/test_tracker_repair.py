"""Offline regression tests. All vacancy records below are synthetic fixtures."""
import copy
import datetime as dt
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import yaml
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from src import matching as m, scrapers as s, tracker as t, website as w, main as pipeline

ROOT=Path(__file__).resolve().parents[1]
PROFILE=yaml.safe_load((ROOT/'config/cv_profile.yaml').read_text(encoding='utf-8'))

def fixture(**kwargs):
    record=dict(title='Junior Sustainability Analyst',company='Synthetic Employer',city='Munich',location='Munich, Germany',url='https://example.org/job/123',description='GHG Protocol carbon accounting CSRD sustainability reporting data analysis',posted_date=m.today_local().isoformat(),relevance_score=8,match_type='Direct role',match_reason='Title: sustainability')
    record.update(kwargs)
    return record

class FakeResponse:
    def __init__(self,data=None,text='',url='https://example.org/jobs',status=200):
        self.data=data;self.text=text;self.url=url;self.status_code=status;self.cookies=type('Cookies',(),{'get_dict':lambda self:{}})()
    def json(self):return self.data
    def raise_for_status(self):
        if self.status_code>=400:raise s.requests.HTTPError(str(self.status_code))

class MatchingTests(unittest.TestCase):
    def decision(self,title,profile=PROFILE):return m.title_decision(title,profile['title_must_match'],profile['title_must_not_match'])
    def test_junior_esg_allowed(self):self.assertTrue(self.decision('Junior Sustainability Analyst')['keep'])
    def test_international_not_intern(self):self.assertTrue(self.decision('International ESG Reporting Analyst')['keep'])
    def test_partnerships_not_partner(self):self.assertTrue(self.decision('ESG Partnerships Analyst')['keep'])
    def test_senior_still_excluded(self):self.assertFalse(self.decision('Senior Sustainability Analyst')['keep'])
    def test_intern_still_excluded(self):self.assertFalse(self.decision('Sustainability Intern')['keep'])
    def test_junior_can_be_excluded_by_another_profile(self):
        profile=copy.deepcopy(PROFILE);profile['title_must_not_match'].append('junior')
        self.assertFalse(self.decision('Junior Sustainability Analyst',profile)['keep'])
    def test_unknown_location_not_confirmed(self):self.assertEqual(m.location_decision(fixture(location='Germany'))['status'],'unconfirmed')
    def test_outside_location_not_confirmed(self):self.assertEqual(m.location_decision(fixture(location='Berlin'))['status'],'outside')
    def test_preferred_area_confirmed(self):self.assertEqual(m.location_decision(fixture(location='Ottobrunn'))['status'],'confirmed')
    def test_age_boundary(self):
        today=dt.date(2026,10,22)
        self.assertTrue(m.date_decision(fixture(posted_date='2026-10-01'),today=today)['keep'])
        self.assertFalse(m.date_decision(fixture(posted_date='2026-09-30'),today=today)['keep'])
    def test_undated_policy_preserved(self):self.assertTrue(m.date_decision(fixture(posted_date=''))['keep'])
    def test_creation_is_not_publication(self):self.assertTrue(m.date_decision(fixture(posted_date='2020-01-01',_date_kind='creation_unverified'))['keep'])
    def test_related_evidence_required(self):
        self.assertFalse(m.role_decision(fixture(title='Consultant',description='We care about sustainability'),PROFILE)['keep'])
        self.assertTrue(m.role_decision(fixture(title='Consultant',description='CSRD sustainability reporting data quality'),PROFILE)['keep'])

class TrackerTests(unittest.TestCase):
    def setUp(self):self.temp=tempfile.TemporaryDirectory();self.path=Path(self.temp.name)/'tracker.xlsx'
    def tearDown(self):self.temp.cleanup()
    def update(self,jobs):return t.update_tracker(self.path,jobs,min_score=3,cv_profile=PROFILE)
    def rows(self):
        wb=load_workbook(self.path);rows=list(wb['Jobs'].values);wb.close();return rows
    def test_senior_not_added_with_esg_profile(self):self.assertEqual(self.update([fixture(title='Senior ESG Analyst')])['total_rows'],0)
    def test_junior_survives_storage(self):self.assertEqual(self.update([fixture()])['total_rows'],1)
    def test_existing_row_refreshes_score_title_location(self):
        self.update([fixture(relevance_score=4)])
        summary=self.update([fixture(title='Sustainability Data Analyst',relevance_score=9,location='Ottobrunn')])
        row=dict(zip(t.COLUMNS,self.rows()[1]))
        self.assertEqual(summary['already_tracked'],1)
        self.assertEqual(row['Relevance Score (1-10)'],9)
        self.assertEqual(row['Location'],'Ottobrunn')
        self.assertEqual(row['Job Title'],'Sustainability Data Analyst')
    def test_known_date_not_erased(self):
        self.update([fixture()]);self.update([fixture(posted_date='')])
        self.assertEqual(self.rows()[1][0],m.today_local().isoformat())
    def test_related_floor_seven(self):
        self.assertEqual(self.update([fixture(match_type='Related role',relevance_score=6)])['total_rows'],0)
        self.assertEqual(self.update([fixture(match_type='Related role',relevance_score=7)])['total_rows'],1)
    def test_old_rows_pruned(self):
        old=(m.today_local()-dt.timedelta(days=22)).isoformat()
        self.assertEqual(self.update([fixture(posted_date=old)])['total_rows'],0)

class AdapterTests(unittest.TestCase):
    def setUp(self):s.reset_headless_budget()
    def test_workday_listing_keeps_detail_route(self):
        page={'total':1,'jobPostings':[{'title':'Junior ESG Analyst','externalPath':'/job/Munich/ESG_R1','locationsText':'2 Locations','postedOn':'Posted Today'}]}
        with patch.object(s.requests,'post',side_effect=[FakeResponse(page),FakeResponse({'total':1,'jobPostings':[]})]):
            jobs=s.scrape_workday('Synthetic','https://example.wd3.myworkdayjobs.com/Careers')
        self.assertEqual(jobs[0]['_detail_api_url'],'https://example.wd3.myworkdayjobs.com/wday/cxs/example/Careers/job/Munich/ESG_R1')
    def test_workday_detail_enrichment(self):
        response=FakeResponse({'jobPostingInfo':{'title':'ESG Analyst','jobDescription':'<p>CSRD reporting</p>','location':'Berlin','additionalLocations':['Munich'],'startDate':'2020-01-01'}})
        with patch.object(s,'_safe_get',return_value=response):details,body=s.fetch_job_details('https://example.wd3.myworkdayjobs.com/wday/cxs/example/Careers/job/ESG_R1')
        self.assertEqual(details['description'],'CSRD reporting');self.assertIn('Munich',details['location'])
        self.assertNotIn('posted_date',details)
    def test_bad_workday_schema_is_not_a_job(self):
        with patch.object(s,'_safe_get',return_value=FakeResponse({'unexpected':[]})):
            self.assertEqual(s.fetch_job_details('https://example.wd3.myworkdayjobs.com/wday/cxs/example/Careers/job/R1'),({},''))
    def test_scoped_html_description(self):
        response=FakeResponse(text='<nav>carbon accounting</nav><main><h1>Sustainability Analyst</h1><div class="job-description">CSRD reporting</div></main><footer>GHG Protocol</footer>')
        with patch.object(s,'_safe_get',return_value=response):details,body=s.fetch_job_details('https://example.org/job/123')
        self.assertEqual(details['description'],'CSRD reporting');self.assertNotIn('GHG',body)
    def test_browser_budget_is_shared(self):
        s._headless_time_used=s._HEADLESS_BUDGET_SECONDS
        with patch.object(s,'_browse_public_board') as browse:
            jobs,stats=s.browse_public_board('Synthetic','https://example.org/jobs')
            browse.assert_not_called();self.assertEqual(stats['completion'],'browser_budget_exhausted')
    def test_actual_next_link_retains_pages(self):
        one=FakeResponse(text='<a href="/job/123">Sustainability Analyst</a><a rel="next" href="/jobs?page=2">Next</a>',url='https://example.org/jobs')
        two=FakeResponse(text='<a href="/job/124">ESG Analyst</a>',url='https://example.org/jobs?page=2')
        with patch.object(s,'browse_public_board',return_value=([],{'pages':0,'completion':'test'})):
            jobs=s.scrape_html_board('Synthetic','https://example.org/jobs',lambda url:one if '?' not in url else two)
        self.assertEqual(len(jobs),2)

class WebsiteTests(unittest.TestCase):
    def test_build_marker_and_no_empty_details(self):
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'tracker.xlsx';out=Path(tmp)/'index.html'
            t.update_tracker(path,[fixture()],cv_profile=PROFILE)
            with patch.dict(os.environ,{'GITHUB_RUN_ID':'12345','GITHUB_RUN_ATTEMPT':'2'}):w.generate(path,out)
            soup=BeautifulSoup(out.read_text(),'html.parser')
            self.assertEqual(soup.select_one('meta[name="tracker-build"]')['content'],'12345-2')
            self.assertEqual(len(soup.select('.job-row')),1)
            self.assertFalse(soup.select('.row-details'))
            self.assertTrue(soup.select_one('.job-row a.apply'))
    def test_saved_senior_hidden_and_junior_shown(self):
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'tracker.xlsx';out=Path(tmp)/'index.html'
            t.update_tracker(path,[fixture(),fixture(title='Senior ESG Analyst',url='https://example.org/job/999')])
            w.generate(path,out)
            soup=BeautifulSoup(out.read_text(),'html.parser')
            self.assertEqual(len(soup.select('.job-row')),1)
            self.assertIn('Junior Sustainability Analyst',soup.get_text())

class PipelineTests(unittest.TestCase):
    def test_direct_role_gets_description_before_scoring(self):
        with tempfile.TemporaryDirectory() as tmp:
            job=fixture(description='');job.pop('relevance_score')
            with patch.object(pipeline,'HEALTH_FILE',Path(tmp)/'health.json'),patch.object(pipeline,'scrape_company',return_value=[job]),patch.object(pipeline,'fetch_job_details',return_value=({'description':'CSRD sustainability reporting GHG Protocol','location':'Munich'},'CSRD sustainability reporting GHG Protocol')),patch.object(pipeline.time,'sleep'):
                jobs,counts,reports=pipeline.scrape_all_any_city([{'name':'Synthetic','city':'Munich','careers_url':'https://example.org/jobs'}],PROFILE)
            self.assertEqual(len(jobs),1);self.assertGreaterEqual(jobs[0]['relevance_score'],3)
            self.assertTrue(reports[0]['funnel_balanced'])
    def test_unusable_run_does_not_publish_fresh_summary(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);(root/'docs').mkdir()
            with patch.object(pipeline,'ROOT',root),patch.object(pipeline,'load_yaml',side_effect=[PROFILE,{'companies':[]},{'companies':[]}]),patch.object(pipeline,'scrape_all_any_city',return_value=([],{'source_statuses':{'EMPTY_UNVERIFIED':1}},[{'status':'EMPTY_UNVERIFIED','raw_candidates':0}])):
                with self.assertRaises(RuntimeError):pipeline.run()
            self.assertFalse((root/'docs/run_summary.json').exists())

    def test_direct_role_uses_scoped_body_fallback(self):
        with tempfile.TemporaryDirectory() as tmp:
            job=fixture(description='');job.pop('relevance_score')
            with patch.object(pipeline,'HEALTH_FILE',Path(tmp)/'health.json'),patch.object(pipeline,'scrape_company',return_value=[job]),patch.object(pipeline,'fetch_job_details',return_value=({'title':'Junior Sustainability Analyst'},'CSRD sustainability reporting GHG Protocol')),patch.object(pipeline.time,'sleep'):
                jobs,counts,reports=pipeline.scrape_all_any_city([{'name':'Synthetic','city':'Munich','careers_url':'https://example.org/jobs'}],PROFILE)
            self.assertEqual(len(jobs),1)
            self.assertIn('CSRD',jobs[0]['description'])
    def test_successful_run_writes_summary_and_page(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);(root/'docs').mkdir()
            health=root/'docs/scraper_health.json';tracker=root/'data/tracker.xlsx';out=root/'docs/index.html'
            with patch.object(pipeline,'ROOT',root),patch.object(pipeline,'TRACKER_FILE',tracker),patch.object(pipeline,'HEALTH_FILE',health),patch.object(pipeline,'load_yaml',side_effect=[PROFILE,{'companies':[{'name':'Synthetic','city':'Munich','careers_url':'https://example.org/jobs'}]},{'companies':[]}]),patch.object(pipeline,'scrape_company',return_value=[fixture()]),patch.object(pipeline,'generate_html',side_effect=lambda:w.generate(tracker,out)),patch.object(pipeline.time,'sleep'),patch.dict(os.environ,{'GITHUB_RUN_ID':'456','GITHUB_RUN_ATTEMPT':'1'}):
                pipeline.run()
            summary=json.loads((root/'docs/run_summary.json').read_text())
            self.assertEqual(summary['run_id'],'456')
            self.assertEqual(summary['tracker']['total_rows'],1)
            self.assertIn('456-1',out.read_text())
            self.assertIn('TRACKER_HEALTH_START',out.read_text())

if __name__=='__main__':unittest.main()
