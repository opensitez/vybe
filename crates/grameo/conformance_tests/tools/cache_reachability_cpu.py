#!/usr/bin/env python3
"""Four backend screen on immutable previously selected controls; no Pest rebuild."""
import argparse,json,os,statistics,subprocess,time,hashlib,shutil
from pathlib import Path

def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def closure(p):return {str(x.relative_to(p)):sha(x) for section in ['core/src','codegen/src'] for x in (p/section).rglob('*') if x.is_file()}
def main():
 ap=argparse.ArgumentParser();ap.add_argument('--output',type=Path,required=True);ap.add_argument('--rounds',type=int,default=101);a=ap.parse_args();repo=Path.cwd();controls=repo/'results/testrunner/parser-all-language-cpu-atomic-token-cache';prior=json.loads((controls/'report.json').read_text());out=a.output.resolve();out.mkdir(parents=True,exist_ok=True)
 candidate=repo/'results/testrunner/parser-active-builds/atomic-cache-reachability';dep=repo/'results/testrunner/parser-active-builds/atomic-cache-target/release/deps';runtime=dep/'libvybe_parser-03038302f5b45ab1.rlib';generated=dep/'libatomic_cache_reachability_generated-ae46f41084c7c91d.rlib';frozen=closure(candidate);source=Path(__file__).with_name('cache_reachability_cpu.rs').resolve()
 report={'scope':'Four backends in same executable: pest,old512,reachability,globalcache. Full captures/native errors/recognition logical steps before timing; identical frozen matched controls. Reused frozen independent pest reference libraries.','rounds':a.rounds,'source_start':frozen,'dependencies':{str(runtime):sha(runtime),str(generated):sha(generated)},'controls_report_sha256':sha(controls/'report.json'),'batches':{},'languages':{},'started_unix':time.time()}
 env=os.environ.copy();env.update(DEVELOPER_DIR='/Library/Developer/CommandLineTools',RUST_MIN_STACK='67108864')
 for batch in ['corpus-vb-shell-wast','corpus-dynamic','corpus-c-family','corpus-business','corpus-scripting']:
  d=out/batch;d.mkdir(exist_ok=True);oldcmd=prior['batches'][batch]['command'];dispatch=d/'dispatch.rs';s=(controls/batch/'dispatch.rs').read_text();import re
  s=re.sub(r'(&after_generated::all::(\w+)::Parser),',r'\1,&global_generated::all::\2::Parser,',s);dispatch.write_text(s);env['VYBE_CPU_DISPATCH']=str(dispatch)
  cmd=oldcmd.copy();cmd[cmd.index('-O')+1]=str(source);globalruntime=next(x[len('after='):] for x in cmd if x.startswith('after='));globalgen=next(x[len('after_generated='):] for x in cmd if x.startswith('after_generated='));cmd=[f'after={runtime}' if x.startswith('after=') else f'after_generated={generated}' if x.startswith('after_generated=') else x for x in cmd];cmd[cmd.index('-o')+1]=str(d/'cpu');cmd.extend(['--extern',f'global={globalruntime}','--extern',f'global_generated={globalgen}'])
  print('Compiling paired driver',batch,flush=True)
  with (d/'build.log').open('w') as f:subprocess.run(cmd,env=env,stdout=f,stderr=subprocess.STDOUT,check=True)
  print('Timing',batch,flush=True)
  with (d/'measurements.tsv').open('w') as f:subprocess.run([str(d/'cpu'),str(controls/batch),str(a.rounds)],env=env,stdout=f,stderr=subprocess.STDOUT,check=True)
  vals={};parity=set()
  for line in (d/'measurements.tsv').read_text().splitlines():
   v=line.split('\t')
   if v[0]=='PARITY':parity.add((v[1],v[2]))
   elif v[0]=='TIME':vals.setdefault((v[1],v[2],int(v[4])),{})[int(v[3])]=float(v[6])
  for lang in sorted({x[0] for x in vals}):
   report['languages'][lang]={'inputs':prior['languages'][lang]['inputs'],'workloads':{}}
   for sample in prior['languages'][lang]['inputs']:
    label=sample['label'];assert (lang,label) in parity;result={}
    for backend,name in enumerate(['pest','old512','reachability','globalcache']):
     xs=vals[(lang,label,backend)];assert len(xs)==a.rounds;values=list(xs.values());result[name]={'median_ns':statistics.median(values),'p95_ns':sorted(values)[int(.95*(len(values)-1))]}
    for backend,name in [(0,'pest'),(1,'old512'),(3,'globalcache')]:
     result['reachability']['median_ratio_to_'+name]=result['reachability']['median_ns']/result[name]['median_ns'];result['reachability']['paired_median_ratio_to_'+name]=statistics.median(vals[(lang,label,2)][i]/vals[(lang,label,backend)][i] for i in range(a.rounds))
    report['languages'][lang]['workloads'][label]=result
  report['batches'][batch]={'command':cmd,'binary_sha256':sha(d/'cpu'),'dispatch_sha256':sha(dispatch),'driver_sha256':sha(source),'reference_library_sha256':prior['batches'][batch]['reference_library_sha256']};(out/'report-progress.json').write_text(json.dumps(report,indent=2)+'\n')
 report['source_end']=closure(candidate);assert report['source_end']==frozen;report['complete']=True;report['finished_unix']=time.time();(out/'report.json').write_text(json.dumps(report,indent=2)+'\n');shutil.copy2(source,out/source.name);shutil.copy2(Path(__file__),out/Path(__file__).name)
if __name__=='__main__':main()
