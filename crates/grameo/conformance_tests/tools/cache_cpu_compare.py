#!/usr/bin/env python3
"""Frozen old512/new cache/pest CPU on identical previously attested 90 controls."""
import argparse,hashlib,json,os,shutil,statistics,struct,subprocess,time
from pathlib import Path
import full_corpus as gate

def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def put(f,b):f.write(struct.pack('<Q',len(b)));f.write(b)
def main():
 p=argparse.ArgumentParser();p.add_argument('--controls',type=Path,required=True);p.add_argument('--audit',type=Path,required=True);p.add_argument('--target',type=Path,required=True);p.add_argument('--output',type=Path,required=True);p.add_argument('--rounds',type=int,default=501);a=p.parse_args()
 controls=a.controls.resolve();audit=a.audit.resolve();target=a.target.resolve();out=a.output.resolve();out.mkdir(parents=True,exist_ok=True)
 prior=json.loads((controls/'report.json').read_text());proof=json.loads((audit/'report.json').read_text());assert prior['complete'] and proof['immutable_bundle_gate_passed'] and proof['candidate_unchanged']
 crate=Path(proof['candidate_crate']);start=gate.implementation_hash(crate);assert start==proof['candidate_end_sha256']
 deps=prior['dependencies'];old=next(Path(x) for x in deps if Path(x).name.startswith('libvybe_parser-'));oldgen=next(Path(x) for x in deps if Path(x).name.startswith('libvybe_parser_generated_tests-'));pest=next(Path(x) for x in deps if Path(x).name.startswith('libpest-'));derive=next(Path(x) for x in deps if Path(x).name.startswith('libpest_derive-'))
 for x in [old,oldgen,pest,derive]:assert sha(x)==deps[str(x)]
 artifacts=[json.loads(x) for x in (target/'build.jsonl').read_text().splitlines() if x.startswith('{')]
 def artifact(name):
  x={Path(f) for r in artifacts if r.get('reason')=='compiler-artifact' and r['target']['name']==name and r['profile']['opt_level']=='3' for f in r['filenames'] if f.endswith('.rlib')};assert len(x)==1;return x.pop()
 new,newgen=artifact('vybe_parser'),artifact('vybe_parser_generated_tests')
 env=os.environ.copy();env.update(DEVELOPER_DIR='/Library/Developer/CommandLineTools',RUST_MIN_STACK='67108864',CARGO_MANIFEST_DIR=str(crate/'generated_tests'))
 report={'scope':'Same 90 previously attested matched controls, old512/new cache/pest full captures and native exact errors/recognition logical steps before timing. Allocation/drop included; grammar compilation, preprocessing, disk, traversal and assertions excluded.','controls_report_sha256':sha(controls/'report.json'),'audit_report_sha256':sha(audit/'report.json'),'source_sha256':start,'rounds':a.rounds,'dependencies':{str(x):sha(x) for x in [old,oldgen,new,newgen,pest,derive]},'timer':prior['timer'],'languages':{},'batches':{},'started_unix':time.time()}
 source=Path(__file__).with_name('cache_cpu_compare.rs').resolve();clock=controls/'clock.o';report['clock_object_sha256']=sha(clock)
 for batch,langs in gate.BATCHES.items():
  d=out/batch;d.mkdir(exist_ok=True);ref=d/'references.rs';ref.write_text((controls/batch/'references.rs').read_text().split('fn run_selected')[0].replace('mod ref_','pub mod ref_'))
  lib=d/('libfrozen_refs_'+batch.replace('-','_')+'.rlib');cmd=['rustc','--edition=2024','--crate-type=rlib','--crate-name','frozen_refs_'+batch.replace('-','_'),'-O',str(ref),'--extern',f'pest={pest}','--extern',f'pest_derive={derive}','-L',f'dependency={pest.parent}','-o',str(lib)]
  print('Compiling reusable references',batch,flush=True)
  with (d/'references-build.log').open('w') as f:subprocess.run(cmd,stdout=f,stderr=subprocess.STDOUT,env=env,check=True)
  calls=[]
  for lang in langs:
   metadata=prior['languages'][lang];assert metadata['grammar_sha256']==sha(crate.parents[1]/f'languages/{lang}/src/grammar.pest');samples=metadata['inputs'];entry='chunk' if lang=='lua' else 'program'
   with (d/f'{lang}.inputs').open('wb') as f:
    f.write(struct.pack('<Q',len(samples)))
    for i,s in enumerate(samples):
     matched=controls/batch/f'matched-{lang}-{i}.source';assert sha(matched)==s['matched_source_sha256'];data=matched.read_bytes();assert len(data)==s['bytes'];put(f,s['label'].encode());put(f,data);f.write(struct.pack('<Q',s['reference']=='accepted'))
   report['languages'][lang]={'grammar_sha256':metadata['grammar_sha256'],'audited_raw_corpus_sha256':metadata['audited_raw_corpus_sha256'],'inputs':samples,'workloads':{}}
   calls.append(f'run("{lang}",refs::ref_{lang}::Rule::{entry},refs::ref_{lang}::Reference::parse,&before_generated::all::{lang}::Parser,&after_generated::all::{lang}::Parser,"{entry}",args);')
  dispatch=d/'dispatch.rs';dispatch.write_text('fn run_selected(args:&[String]) {'+'\n'.join(calls)+'}\n');env['VYBE_CPU_DISPATCH']=str(dispatch);binary=d/'cpu'
  command=['rustc','--edition=2024','-O',str(source),'--extern',f'refs={lib}','--extern',f'pest={pest}','--extern',f'before={old}','--extern',f'before_generated={oldgen}','--extern',f'after={new}','--extern',f'after_generated={newgen}','-L',f'dependency={pest.parent}','-L',f'dependency={old.parent}','-L',f'dependency={new.parent}','-C',f'link-arg={clock}','-o',str(binary)]
  with (d/'build.log').open('w') as f:subprocess.run(command,stdout=f,stderr=subprocess.STDOUT,env=env,check=True)
  report['batches'][batch]={'reference_command':cmd,'reference_sha256':sha(ref),'reference_library_sha256':sha(lib),'command':command,'driver_sha256':sha(source),'binary_sha256':sha(binary)}
  print('Timing',batch,flush=True)
  with (d/'measurements.tsv').open('w') as f:subprocess.run([str(binary),str(d),str(a.rounds)],stdout=f,stderr=subprocess.STDOUT,env=env,check=True)
  measured={};parity=set()
  for line in (d/'measurements.tsv').read_text().splitlines():
   v=line.split('\t')
   if v[0]=='PARITY':parity.add((v[1],v[2]))
   elif v[0]=='TIME':measured.setdefault((v[1],v[2],int(v[4])),{})[int(v[3])]=float(v[6])
  for lang in langs:
   for sample in report['languages'][lang]['inputs']:
    label=sample['label'];assert (lang,label) in parity;result={}
    for backend,name in enumerate(['pest','old512','cache']):
     values=measured[(lang,label,backend)];assert len(values)==a.rounds;xs=list(values.values());result[name]={'median_ns':statistics.median(xs),'p95_ns':sorted(xs)[int(.95*(len(xs)-1))]}
    for name in ['pest','old512']:
     result['cache']['median_ratio_to_'+name]=result['cache']['median_ns']/result[name]['median_ns'];result['cache']['p95_ratio_to_'+name]=result['cache']['p95_ns']/result[name]['p95_ns'];backend=0 if name=='pest' else 1;result['cache']['paired_median_ratio_to_'+name]=statistics.median(measured[(lang,label,2)][r]/measured[(lang,label,backend)][r] for r in range(a.rounds))
    report['languages'][lang]['workloads'][label]=result
  (out/'report-progress.json').write_text(json.dumps(report,indent=2)+'\n')
 report['source_end_sha256']=gate.implementation_hash(crate);assert report['source_end_sha256']==start;report['finished_unix']=time.time();report['complete']=True;(out/'report.json').write_text(json.dumps(report,indent=2)+'\n');print('Completed all18 exact-controls cache CPU',flush=True)
if __name__=='__main__':main()
