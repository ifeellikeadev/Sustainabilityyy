"""Offline regression fixtures, not live vacancies or measured coverage."""
import copy,json,tempfile,time,unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import yaml
from src import matching as m,scrapers as s,main,tracker,website
from src.diagnostics import source_budget
P=yaml.safe_load((Path(__file__).parents[1]/'config/cv_profile.yaml').read_text())
class Repairs(unittest.TestCase):
 def setUp(self):s.reset_headless_budget()
 def test_lca_titles(self):
  for title in ['LCA & Environmental Data Specialist','Life-Cycle Assessment Consultant','Sustainability Consultant - Oekobilanzierung','Lebenszyklusanalyse Spezialist']:
   self.assertFalse(m.role_decision(dict(title=title,description='CSRD data quality'),P)['keep'],title)
 def test_incidental_lca_not_excluded(self):
  self.assertTrue(m.role_decision(dict(title='ESG Reporting Specialist',description='May collaborate with LCA team'),P)['keep'])
 def test_lca_no_scoring(self):
  self.assertFalse(any(x['term'].casefold() in ['lca','openlca','life cycle assessment'] for x in P['scoring_keywords']))
 def test_senior_excluded(self):self.assertFalse(m.role_decision(dict(title='Senior ESG Consultant'),P)['keep'])
 def test_junior_profile_preserved(self):self.assertTrue(m.role_decision(dict(title='Junior ESG Analyst'),P)['keep'])
 def test_intern_excluded(self):self.assertFalse(m.role_decision(dict(title='ESG Intern'),P)['keep'])
 def test_international_not_intern(self):self.assertTrue(m.role_decision(dict(title='International ESG Specialist'),P)['keep'])
 def test_hyphenated_responsibilities(self):
  j=dict(title='Consultant',description='GHG-Protocol and data-quality analysis')
  self.assertTrue(m.role_decision(j,P)['keep'])
 def test_missing_location_not_inferred(self):self.assertEqual(m.location_decision(dict(company='Munich firm',location=''))['status'],'unconfirmed')
 def test_outside_stays_outside(self):self.assertEqual(m.location_decision(dict(location='Berlin'))['status'],'outside')
 def test_preferred_location(self):self.assertEqual(m.location_decision(dict(location='Ottobrunn'))['status'],'confirmed')
 def test_unknown_date_kept(self):self.assertTrue(m.date_decision(dict(posted_date=''))['keep'])
 def test_old_date_excluded(self):self.assertFalse(m.date_decision(dict(posted_date='2020-01-01'))['keep'])
 def test_browser_reused(self):
  with patch.object(s,'_browse_public_board',return_value=([],dict(pages=1,completion='test'))) as f:
   s.browse_public_board('Fixture','https://example.org/jobs');s.browse_public_board('Fixture','https://example.org/jobs')
   self.assertEqual(f.call_count,1)
 def test_fair_browser_cap(self):
  s._BROWSER_BY_SOURCE['Fixture']=25
  with patch.object(s,'_browse_public_board') as f:
   self.assertEqual(s.browse_public_board('Fixture','https://example.org/jobs')[1]['completion'],'employer_browser_budget_exhausted');f.assert_not_called()
 def test_discovery_despite_five_cards(self):
  cards=[dict(title='Engineer',url='https://example.org/job/'+str(i)) for i in range(5)]
  extra=[dict(title='ESG Analyst',url='https://example.org/job/99')]
  with patch.object(s,'discover_real_careers_url',return_value=('https://jobs.ashbyhq.com/fixture','ashby','fixture')),patch.dict(s._PLATFORM_SCRAPERS,{'ashby':lambda *a:extra}):
   self.assertEqual(len(s._recover_thin_result('Fixture','https://example.org/careers',cards)),6)
 def test_observed_cross_host_search(self):
  response=SimpleNamespace(url='https://example.org/careers',text='<a href="https://careers.example.org/search">Search jobs</a>')
  with patch.object(s,'_safe_get',return_value=response):self.assertEqual(s.discover_real_careers_url(response.url)[0],'https://careers.example.org/search')
 def test_no_discovery_from_single_apply(self):
  response=SimpleNamespace(url='https://example.org/careers',text='<a href="https://jobs.lever.co/fixture/abc">Apply</a>')
  with patch.object(s,'_safe_get',return_value=response):self.assertIsNone(s.discover_real_careers_url(response.url)[0])
 def test_detail_explicit_label(self):
  response=SimpleNamespace(url='https://example.org/job/1',text='<main><h1>ESG Specialist</h1><span>Standort:</span><div>Ottobrunn</div><p>CSRD reporting</p></main>')
  with patch.object(s,'_safe_get',return_value=response):self.assertEqual(s.fetch_job_details(response.url)[0]['location'],'Ottobrunn')
 def test_workday_extra_location_descriptor(self):
  response=SimpleNamespace(json=lambda:dict(jobPostingInfo=dict(title='ESG Analyst',location='Berlin',additionalLocations=[dict(descriptor='Munich')],jobDescription='CSRD')))
  with patch.object(s,'_safe_get',return_value=response):self.assertEqual(s.fetch_job_details('https://fixture.wd3.myworkdayjobs.com/wday/cxs/fixture/Careers/job/1')[0]['location'],'Berlin; Munich')
 def test_direct_details_before_related(self):
  raw=[dict(title='Consultant',url='https://example.org/job/2',location='Munich',description='',posted_date=''),dict(title='ESG Specialist',url='https://example.org/job/1',location='',description='',posted_date='')]
  order=[]
  def detail(url):order.append(url);return dict(title='ESG Specialist',location='Munich',description='CSRD ESG data quality'),''
  with tempfile.TemporaryDirectory() as d,patch.object(main,'HEALTH_FILE',Path(d)/'health.json'),patch.object(main,'scrape_company',return_value=raw),patch.object(main,'fetch_job_details',side_effect=detail),patch.object(main.time,'sleep'):
   jobs,counts,rows=main.scrape_all_any_city([dict(name='Fixture',careers_url='https://example.org')],P)
   self.assertEqual(order[0],'https://example.org/job/1');self.assertTrue(rows[0]['funnel_balanced']);self.assertTrue(jobs)
 def test_saved_lca_removed(self):
  with tempfile.TemporaryDirectory() as d:
   path=Path(d)/'tracker.xlsx';wb=tracker._new_workbook();wb['Jobs'].append(['','Fixture','Munich','LCA Environmental Specialist',9,'Munich','https://example.org/job/1','Direct role','']);wb.save(path)
   self.assertEqual(tracker.load_or_create(path,P)['Jobs'].max_row,1)
 def test_listing_deadline_leaves_detail_time(self):
  with source_budget(.4) as budget:
   self.assertIsNone(budget.call(time.sleep,.15,max_seconds=.02,default=None,phase='listing'))
   budget.renew_detail_phase();self.assertTrue(budget.listing_expired);self.assertGreater(budget.remaining(),0);self.assertEqual(budget.call(lambda:42),42)
if __name__=='__main__':unittest.main()
