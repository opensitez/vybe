#!/usr/bin/env python3
"""Summarize paired frozen-control CPU rounds without selecting best repeats."""
import argparse,hashlib,json,statistics
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('directory',type=Path);p.add_argument('--rounds',type=int,default=101);a=p.parse_args();rows={};meta={}
for path in a.directory.glob('*/measurements.tsv'):
 for line in path.read_text().splitlines():
  f=line.split('\t')
  if f[0]=='PARITY':meta[(f[1],f[2])]={'bytes':int(f[3]),'accepted':f[4]=='true'}
  if f[0]=='TIME':rows.setdefault((f[1],f[2]),{}).setdefault(int(f[3]),{})[int(f[4])]=float(f[6])
result=[]
for key,rounds in sorted(rows.items()):
 assert len(rounds)==a.rounds and all(set(r)=={0,1,2}for r in rounds.values())
 ratio=statistics.median([r[2]/r[1]for r in rounds.values()]);result.append({'language':key[0],'label':key[1],**meta[key],'paired_ratio':ratio,'change_pct':(ratio-1)*100})
assert len(result)==90
report={'controls':result,'rounds':a.rounds,'improved':sum(x['paired_ratio']<1 for x in result),'largest_regression_pct':max(x['change_pct']for x in result),'best_gain_pct':min(x['change_pct']for x in result),'scope':'Paired same-executable thread CPU; exact capture/error/work equality checked before timing; isolated, not promoted.'}
plan=a.directory/'plan.json'
if plan.exists():
 data=json.loads(plan.read_text());root=Path(data['candidate']);assert all(hashlib.sha256((root/name).read_bytes()).hexdigest()==value for name,value in data['source_sha256'].items());report['source_unchanged']=True
report['measurements_sha256']={str(path.relative_to(a.directory)):hashlib.sha256(path.read_bytes()).hexdigest()for path in sorted(a.directory.glob('*/measurements.tsv'))}
(a.directory/'assessment.json').write_text(json.dumps(report,indent=2)+'\n');print(json.dumps({key:report[key]for key in ('improved','rounds','best_gain_pct','largest_regression_pct')},indent=2));print('Regressions:',[(x['language'],round(x['change_pct'],2))for x in result if x['paired_ratio']>=1])
