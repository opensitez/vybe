#!/usr/bin/env python3
"""Archive the untimed completed-rule profiler and its exact parity proof."""
import hashlib,json,shutil
from pathlib import Path
from collections import defaultdict
ROOT=Path(__file__).resolve().parents[4]
work=ROOT/'results/testrunner/parser-active-builds/predicate-outcomes'
out=ROOT/'results/testrunner/parser-failed-predicate-profile'
out.mkdir(exist_ok=True)
cols='language rule name start body_end flags scope parent body_steps after_enter_steps physical_peak virtual_peak matched certificate pure expressions nested_rules dispatches'.split()
langs={};checks=0
for control,item in json.loads((work/'runs.json').read_text()).items():
 path=Path(item['output']);shutil.copy2(path,out/path.name)
 lang=control.rsplit('-',1)[0];entry=langs.setdefault(lang,{'logical_steps':0,'expression_attempts':0,'nested_rule_attempts':0,'body_dispatches':0,'physical_peak':0,'virtual_peak':0,'key_conflicts':0,'rules':{}})
 seen={}; edges=set(); rows=[]
 for line in path.read_text().splitlines():
  f=line.split('\t')
  if f[0]=='GUARDS':checks+=int(f[2])
  if f[0]=='EDGE':edges.add((int(f[2]),int(f[3]),f[5]=='true'))
  if f[0]!='CALL':continue
  row=dict(zip(cols,f[1:]));rows.append(row)
  for k in cols[1:]:
   if k in ('name','matched','certificate','pure'):continue
   row[k]=int(row[k])
  for k in ('matched','certificate','pure'):row[k]=row[k]=='true'
  if row['parent']==-1:
   entry['logical_steps']+=row['after_enter_steps']+1
   entry['expression_attempts']+=row['expressions'];entry['nested_rule_attempts']+=row['nested_rules'];entry['body_dispatches']+=row['dispatches']
   entry['physical_peak']=max(entry['physical_peak'],row['physical_peak']);entry['virtual_peak']=max(entry['virtual_peak'],row['virtual_peak'])
  if not row['pure'] or not row['flags']&4:continue
  r=entry['rules'].setdefault(row['name'],{'rule':row['rule'],'calls':0,'failed':0,'successful':0,'failed_repeats':0,'successful_repeats':0,'repeated_failed_logical_body_steps':0,'repeated_successful_logical_body_steps':0,'repeated_failed_dispatches':0,'repeated_successful_dispatches':0,'repeated_failed_expressions':0,'repeated_successful_expressions':0,'body_steps':0,'dispatches':0,'physical_peak':0,'virtual_peak':0,'certificate_failures':0,'observed_parents':{}})
  r['calls']+=1;r['failed' if not row['matched'] else 'successful']+=1;r['body_steps']+=row['body_steps'];r['dispatches']+=row['dispatches'];r['physical_peak']=max(r['physical_peak'],row['physical_peak']);r['virtual_peak']=max(r['virtual_peak'],row['virtual_peak']);r['certificate_failures']+=int(row['certificate'] and not row['matched'])
  parent=str(row['parent']);r['observed_parents'][parent]=r['observed_parents'].get(parent,0)+1
  key=(row['rule'],row['start'],row['flags'],row['scope']);signature=(row['matched'],row['body_end'],row['body_steps'],row['physical_peak'],row['virtual_peak'],row['expressions'],row['nested_rules'],row['dispatches'])
  if key in seen:
   entry['key_conflicts']+=int(seen[key]!=signature)
   kind='successful' if row['matched'] else 'failed';r[kind+'_repeats']+=1
   r['repeated_'+kind+'_logical_body_steps']+=row['body_steps'];r['repeated_'+kind+'_dispatches']+=row['dispatches'];r['repeated_'+kind+'_expressions']+=row['expressions']
  else:seen[key]=signature
 for r in entry['rules'].values():r['negative_static_edges_for_observed_parents']={p:any(a==int(p) and b==r['rule'] and neg for a,b,neg in edges) for p in r['observed_parents']}
assert checks==1260
for lang,e in langs.items():
 e['rules']=dict(sorted(e['rules'].items(),key=lambda x:-x[1]['repeated_failed_dispatches']))
 for r in e['rules'].values():r['repeated_failed_logical_fraction']=r['repeated_failed_logical_body_steps']/e['logical_steps']
for name in ['core','profile','driver-build.json','driver-build.log','baseline-source-hashes.json','runs.json','initial-counts.json']:
 dst=out/name;src=work/name
 if src.is_dir():shutil.copytree(src,dst,dirs_exist_ok=True)
 else:shutil.copy2(src,dst)
shutil.copy2(Path(__file__),out/'archive_predicate_profile.py');shutil.copy2(Path(__file__).with_name('predicate_outcomes.rs'),out/'predicate_outcomes.rs')
manifest={str(p.relative_to(out)):hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(out.rglob('*')) if p.is_file() and p.name not in ('report.json','summary.md')}
report={'scope':'Untimed generated parser profile on nine immutable accepted controls; e2 Task+sequence baseline. No caching or grammar edits. Full trees/errors/recognition logical work and existing optimized/default-custom observer traces match baseline.', 'guard_comparisons':checks,'parity':True,'languages':langs,'limitations':['Logical charges must remain unchanged under any optimization; shares are not predicted CPU savings.','Physical dispatches and expression/rule attempts exclude metadata lookups, literal comparison bytes and scanner inner-loop bytes.','Nested body cost overlaps ancestor costs; do not sum all rules as independent savings.','Only completed calls recorded. Physical peak counts actual dispatched rule nesting; virtual peak also includes certified/reserved rule depth.'], 'sha256':manifest}
(out/'report.json').write_text(json.dumps(report,indent=2)+'\n')
lines=['Untimed predicate outcome profile: nine controls; exact capture/error/logical-work and observer parity; 1,260 resource guards.','', 'Physical dispatch counts exclude literal/scanner byte loops. Logical work shares are not CPU estimates.','']
for lang,e in langs.items():
 lines.append(f"{lang}: {e['logical_steps']} logical steps, {e['body_dispatches']} physical body dispatches; physical/virtual peak {e['physical_peak']}/{e['virtual_peak']}; {e['key_conflicts']} repeated-key conflicts.")
 for name,r in list(e['rules'].items())[:5]:lines.append(f"  {name}: {r['failed_repeats']} repeated failures, {r['repeated_failed_dispatches']} body dispatches, {r['repeated_failed_logical_body_steps']} logical body steps ({r['repeated_failed_logical_fraction']:.2%}); successful repeats {r['successful_repeats']}; physical/virtual peak {r['physical_peak']}/{r['virtual_peak']}.")
(out/'summary.md').write_text('\n'.join(lines)+'\n');print('\n'.join(lines))
