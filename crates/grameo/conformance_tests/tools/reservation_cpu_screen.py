#!/usr/bin/env python3
"""Link frozen reserve-policy candidate libraries to reusable reference drivers."""
import argparse,hashlib,json,os,subprocess
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--candidate',type=Path,required=True);p.add_argument('--generated-name',required=True);p.add_argument('--template',type=Path,required=True);p.add_argument('--output',type=Path,required=True);p.add_argument('--rounds',type=int,default=101);a=p.parse_args();cand=a.candidate.resolve();out=a.output.resolve();out.mkdir();meta=[json.loads(x)for x in(cand/'build.jsonl').read_text().splitlines()if x.startswith('{')]
def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def artifact(name):
 rows=[r for r in meta if r.get('reason')=='compiler-artifact'and r['target']['name']==name and r['profile']['opt_level']=='3'and any(x.endswith('.rlib')for x in r['filenames'])];return next(x for x in rows[-1]['filenames']if x.endswith('.rlib'))
libs={'after':artifact('vybe_parser'),'after_generated':artifact(a.generated_name)};plan=json.loads(a.template.read_text());plan['scope']='Bounded reservation candidate vs same frozen packed-mode grammar/control baseline';plan['candidate']=str(cand);plan['candidate_libraries']=libs;plan['rounds']=a.rounds;plan['source_sha256']={str(p.relative_to(cand)):sha(p)for d in ['core/src','codegen/src','generated_tests/src']for p in sorted((cand/d).rglob('*.rs'))}
for k,key in [('runtime','after'),('generated','after_generated')]:plan['binary_library_sha256'][k]=sha(libs[key])
for batch,row in plan['batches'].items():
 dst=out/batch;dst.mkdir();cmd=row['command'][:]
 for i,arg in enumerate(cmd):
  for key,value in libs.items():
   if arg.startswith(key+'='):cmd[i]=key+'='+value
 cmd[cmd.index('-o')+1]=str(dst/'cpu');cmd+=['-L','dependency='+str(Path(libs['after']).parent)];row['command']=cmd
 (out/'plan.json').write_text(json.dumps(plan,indent=2)+'\n')
 with(dst/'build.log').open('w')as f:subprocess.run(cmd,env=dict(os.environ,VYBE_CPU_DISPATCH=row['dispatch'],DEVELOPER_DIR='/Library/Developer/CommandLineTools'),stdout=f,stderr=f,check=True)
 with(dst/'measurements.tsv').open('w')as f:subprocess.run([str(dst/'cpu'),row['input_directory'],str(a.rounds)],stdout=f,check=True)
 row['binary_sha256']=sha(dst/'cpu');print(batch,flush=True)
(out/'plan.json').write_text(json.dumps(plan,indent=2)+'\n')
