#!/usr/bin/env python3
"""Same90 frozen controls with reusable Pest references and a source-frozen candidate."""
import argparse,json,os,statistics,subprocess,time,hashlib,shutil
from pathlib import Path
import full_corpus as gate

def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def main():
 p=argparse.ArgumentParser();p.add_argument('--controls',type=Path,required=True);p.add_argument('--candidate-root',type=Path,required=True);p.add_argument('--runtime',type=Path,required=True);p.add_argument('--generated',type=Path,required=True);p.add_argument('--output',type=Path,required=True);p.add_argument('--rounds',type=int,default=501);p.add_argument('--reuse-existing-binaries',action='store_true');a=p.parse_args()
 controls=a.controls.resolve();candidate=a.candidate_root.resolve();out=a.output.resolve();out.mkdir(parents=True,exist_ok=True);runtime=a.runtime.resolve();generated=a.generated.resolve();prior=json.loads((controls/'report.json').read_text());freeze=json.loads((candidate/'source-freeze.json').read_text());assert prior['complete']
 def verify():
  for name,digest in freeze.items():assert sha(candidate/name)==digest,name
 verify();source=Path(__file__).with_name('cache_cpu_compare.rs').resolve();env=os.environ.copy();env.update(DEVELOPER_DIR='/Library/Developer/CommandLineTools',RUST_MIN_STACK='67108864')
 report={'scope':'Same90 canonical frozen matched controls, pest/old512/new no-cache candidate; full captures/native errors/consumed/recognition logical steps before timing. Allocation/drop included; compile/preprocessing/IO/serialization/traversal excluded.','candidate_root':str(candidate),'candidate_source_freeze':freeze,'controls_report_sha256':sha(controls/'report.json'),'rounds':a.rounds,'dependencies':{str(runtime):sha(runtime),str(generated):sha(generated)},'languages':{},'batches':{},'started_unix':time.time()}
 for batch,langs in gate.BATCHES.items():
  d=out/batch;d.mkdir(exist_ok=True);dispatch=d/'dispatch.rs';shutil.copy2(controls/batch/'dispatch.rs',dispatch);env['VYBE_CPU_DISPATCH']=str(dispatch);cmd=prior['batches'][batch]['command'].copy();cmd[cmd.index('-O')+1]=str(source);cmd=[f'after={runtime}' if x.startswith('after=') else f'after_generated={generated}' if x.startswith('after_generated=') else x for x in cmd];cmd[cmd.index('-o')+1]=str(d/'cpu');cmd.extend(['-L',f'dependency={runtime.parent}'])
  for lang in langs:
   metadata=prior['languages'][lang];assert sha(candidate/f'languages/{lang}/src/grammar.pest')==metadata['grammar_sha256'];report['languages'][lang]={'inputs':metadata['inputs'],'grammar_sha256':metadata['grammar_sha256'],'audited_raw_corpus_sha256':metadata['audited_raw_corpus_sha256'],'workloads':{}}
  print('Compiling paired driver',batch,flush=True)
  if a.reuse_existing_binaries and (d/'cpu').exists():
   if (out/'interrupted-provenance.json').exists() and batch=='corpus-scripting':assert sha(d/'cpu')==json.loads((out/'interrupted-provenance.json').read_text())['binary_sha256']
   print('Reusing exact executable',batch,flush=True)
  else:
   with (d/'build.log').open('w') as f:subprocess.run(cmd,env=env,stdout=f,stderr=subprocess.STDOUT,check=True)
  print('Timing',batch,flush=True)
  with (d/'measurements.tsv').open('w') as f:subprocess.run([str(d/'cpu'),str(controls/batch),str(a.rounds)],env=env,stdout=f,stderr=subprocess.STDOUT,check=True)
  vals={};parity=set()
  for line in (d/'measurements.tsv').read_text().splitlines():
   v=line.split('\t')
   if v[0]=='PARITY':parity.add((v[1],v[2]))
   elif v[0]=='TIME':vals.setdefault((v[1],v[2],int(v[4])),{})[int(v[3])]=float(v[6])
  for lang in langs:
   for sample in report['languages'][lang]['inputs']:
    label=sample['label'];assert (lang,label) in parity;result={}
    for backend,name in enumerate(['pest','old512','candidate']):
     xs=vals[(lang,label,backend)];assert len(xs)==a.rounds;values=list(xs.values());result[name]={'median_ns':statistics.median(values),'p95_ns':sorted(values)[int(.95*(len(values)-1))]}
    for backend,name in [(0,'pest'),(1,'old512')]:
     result['candidate']['median_ratio_to_'+name]=result['candidate']['median_ns']/result[name]['median_ns'];result['candidate']['p95_ratio_to_'+name]=result['candidate']['p95_ns']/result[name]['p95_ns'];result['candidate']['paired_median_ratio_to_'+name]=statistics.median(vals[(lang,label,2)][i]/vals[(lang,label,backend)][i] for i in range(a.rounds))
    report['languages'][lang]['workloads'][label]=result
  report['batches'][batch]={'command':cmd,'binary_sha256':sha(d/'cpu'),'dispatch_sha256':sha(dispatch),'driver_sha256':sha(source),'reference_library_sha256':prior['batches'][batch]['reference_library_sha256']};(out/'report-progress.json').write_text(json.dumps(report,indent=2)+'\n')
 verify();assert sha(runtime)==report['dependencies'][str(runtime)] and sha(generated)==report['dependencies'][str(generated)];report['candidate_unchanged']=True;report['complete']=True;report['finished_unix']=time.time();(out/'report.json').write_text(json.dumps(report,indent=2)+'\n');shutil.copy2(source,out/source.name);shutil.copy2(Path(__file__),out/Path(__file__).name)
if __name__=='__main__':main()
