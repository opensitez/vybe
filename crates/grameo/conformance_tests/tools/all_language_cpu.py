#!/usr/bin/env python3
"""Frozen all-language capture CPU baseline, with exact parity before timing."""
import argparse,csv,hashlib,json,os,shutil,statistics,struct,subprocess,time
from pathlib import Path
import full_corpus as gate
import native_differential as native

def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def put(f,data): f.write(struct.pack('<Q',len(data)));f.write(data)

def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--audit',required=True,type=Path)
    ap.add_argument('--target',required=True,type=Path)
    ap.add_argument('--baseline',required=True,type=Path)
    ap.add_argument('--output',required=True,type=Path)
    ap.add_argument('--rounds',type=int,default=501)
    args=ap.parse_args();audit=args.audit.resolve();target=args.target.resolve();baseline=args.baseline.resolve();out=args.output.resolve();out.mkdir(parents=True,exist_ok=True)
    proof=json.loads((audit/'report.json').read_text());assert proof['immutable_bundle_gate_passed'] and proof['candidate_unchanged']
    crate=Path(proof['candidate_crate']);frozen_hash=gate.implementation_hash(crate);assert frozen_hash==proof['candidate_end_sha256']
    artifacts=[json.loads(s) for s in (target/'build.jsonl').read_text().splitlines() if s.startswith('{')]
    def artifact(name):
        choices=[Path(f) for r in artifacts if r.get('reason')=='compiler-artifact' and r['target']['name']==name and r['profile']['opt_level']=='3' for f in r['filenames'] if f.endswith('.rlib')];assert len(set(choices))==1;return choices[0]
    parser,generated=artifact('vybe_parser'),artifact('vybe_parser_generated_tests');deps=baseline/'dependencies'
    pest=deps/'libpest-fa1172aed71c7dc5.rlib';derive=deps/'libpest_derive-dee07c710098f8f7.dylib'
    helpers={lang:next((baseline/'generated-sources').rglob(f'{lang}_preprocessor.rs')) for lang in ['python','fortran']}
    timer_source=Path(__file__).resolve().parents[4]/'results/testrunner/parser-diagnostic-batch512-experiment/cpu-harness/clock.c'
    shutil.copy2(timer_source,out/'clock.c')
    env=os.environ.copy();env.update(DEVELOPER_DIR='/Library/Developer/CommandLineTools',RUST_MIN_STACK='67108864',VYBE_AUDIT_PYTHON_HELPER=str(helpers['python']),VYBE_AUDIT_FORTRAN_HELPER=str(helpers['fortran']),CARGO_MANIFEST_DIR=str(crate/'generated_tests'))
    subprocess.run(['clang','-O2','-c',str(out/'clock.c'),'-o',str(out/'clock.o')],env=env,check=True)
    subprocess.run(['clang','-O2','-DTIMER_PROBE',str(out/'clock.c'),'-o',str(out/'clock-probe')],env=env,check=True)
    with (out/'clock-probe.log').open('w') as f:subprocess.run([str(out/'clock-probe')],stdout=f,env=env,check=True)
    report={'schema':1,'scope':'Representative canonical frozen inputs; full capture equality before batched in-memory parsing; allocation/drop included, preprocessing/compile/IO/traversal/serialization excluded from timing','audit_report_sha256':sha(audit/'report.json'),'source_sha256':frozen_hash,'rounds':args.rounds,'selection':'Three accepted byte-size quantiles10/50/90% and two rejected50/90%, canonical inputs16..65536bytes where available; fall back to all sizes if verdict pool empty. Python/Fortran raw and transformed canonical variants eligible. No synthetic mutations. Sampling is a performance workload selection, not a replacement for full conformance.','dependencies':{str(p):sha(p) for p in [parser,generated,pest,derive,*helpers.values()]},'timer':'CLOCK_THREAD_CPUTIME_ID, two calls per batch,250us adaptive target1..128 iterations, common maximum batch across backends, all six backend orders through one opaque dynamic callback timing callsite; one untimed self-warm parse immediately before each timed batch (steady-state CPU); timer cost not subtracted','languages':{},'batches':{},'started_unix':time.time()}
    for batch,languages in gate.BATCHES.items():
        batchdir=out/batch;batchdir.mkdir(exist_ok=True);modules=[];calls=[]
        for lang in languages:
            values=proof['languages'][lang];assert sha(audit/f'{lang}.tsv')==values['report_sha256'] and sha(audit/f'{lang}.bundle')==values['bundle_sha256']
            with (audit/f'{lang}.tsv').open() as f:rows=list(csv.DictReader(f,delimiter='\t'))
            raw=dict(native.bundled_sources(audit/f'{lang}.bundle',lang));chosen=[]
            for verdict,quantiles in [('accepted',[.1,.5,.9]),('rejected',[.5,.9])]:
                pool=[r for r in rows if r['reference']==verdict and 16<=int(r['bytes'])<=65536]
                if not pool:pool=[r for r in rows if r['reference']==verdict]
                assert pool,(lang,verdict);pool.sort(key=lambda r:(int(r['bytes']),r['label']))
                for q in quantiles:
                    row=pool[round(q*(len(pool)-1))]
                    if row not in chosen:chosen.append(row)
            samples=[]
            with (batchdir/f'{lang}.inputs').open('wb') as f:
                f.write(struct.pack('<Q',len(chosen)))
                for r in chosen:
                    label=gate.unescape(r['label']);transform=label.startswith('preprocessed:');original=label[len('preprocessed:'):] if transform else label;data=raw[original]
                    put(f,label.encode());put(f,data);f.write(struct.pack('<QQQQ',transform,int(r['bytes']),int(r['fingerprint'],16),r['reference']=='accepted'))
                    samples.append({'label':label,'bytes':int(r['bytes']),'reference':r['reference'],'raw_source_sha256':hashlib.sha256(data).hexdigest(),'matched_fingerprint':r['fingerprint']})
            gp=crate.parents[1]/f'languages/{lang}/src/grammar.pest';entry='chunk' if lang=='lua' else 'program'
            modules.append(f'mod ref_{lang} {{ #[derive(pest_derive::Parser)] #[grammar = {json.dumps(str(gp))}] pub struct Reference; }}')
            calls.append(f'run("{lang}",ref_{lang}::Rule::{entry},ref_{lang}::Reference::parse,&vybe_parser_generated_tests::all::{lang}::Parser,"{entry}",args);')
            report['languages'][lang]={'grammar_sha256':sha(gp),'audited_raw_corpus_sha256':values['raw_corpus_sha256'],'inputs':samples,'workloads':{}}
        reference=batchdir/'references.rs';reference.write_text('\n'.join(modules)+'\nfn run_selected(args:&[String]) {'+'\n'.join(calls)+'}\n');env['VYBE_CPU_REFERENCES']=str(reference)
        executable=batchdir/'cpu';source=Path(__file__).with_name('all_language_cpu.rs').resolve()
        command=['rustc','--edition=2024','-O',str(source),'--extern',f'pest={pest}','--extern',f'pest_derive={derive}','--extern',f'vybe_parser={parser}','--extern',f'vybe_parser_generated_tests={generated}','-L',f'dependency={deps}','-L',f'dependency={parser.parent}','-C',f'link-arg={out/"clock.o"}','-o',str(executable)]
        print('Compiling',batch,flush=True)
        with (batchdir/'build.log').open('w') as f:subprocess.run(command,stdout=f,stderr=subprocess.STDOUT,env=env,check=True)
        report['batches'][batch]={'command':command,'references_sha256':sha(reference),'binary_sha256':sha(executable),'driver_source_sha256':sha(source)}
        print('Timing',batch,flush=True)
        with (batchdir/'measurements.tsv').open('w') as f:subprocess.run([str(executable),str(batchdir),str(crate.parents[1]/'languages'),str(args.rounds)],stdout=f,stderr=subprocess.STDOUT,env=env,check=True)
        measured={};parity=set()
        for line in (batchdir/'measurements.tsv').read_text().splitlines():
            fields=line.split('\t')
            if fields[0]=='PARITY':parity.add((fields[1],fields[2]))
            elif fields[0]=='TIME':
                _,lang,label,accepted,round_index,backend,n,ns=fields;measured.setdefault((lang,label,int(backend)),[]).append(float(ns))
        for lang in languages:
            for index,sample in enumerate(report['languages'][lang]['inputs']):
                label=sample['label'];assert (lang,label) in parity;sample['matched_source_sha256']=sha(batchdir/f'matched-{lang}-{index}.source')
                result={}
                for backend,name in enumerate(['pest','generated','resolved']):
                    times=measured[(lang,label,backend)];assert len(times)==args.rounds;ordered=sorted(times);result[name]={'median_ns':statistics.median(times),'p95_ns':ordered[int(.95*(len(ordered)-1))]}
                for name in ['generated','resolved']:result[name]['median_ratio_to_pest']=result[name]['median_ns']/result['pest']['median_ns'];result[name]['p95_ratio_to_pest']=result[name]['p95_ns']/result['pest']['p95_ns']
                report['languages'][lang]['workloads'][label]=result
        (out/'report-progress.json').write_text(json.dumps(report,indent=2)+'\n')
    report['source_end_sha256']=gate.implementation_hash(crate);assert report['source_end_sha256']==frozen_hash;report['finished_unix']=time.time();report['complete']=True
    (out/'report.json').write_text(json.dumps(report,indent=2)+'\n')
    print('Completed all18 matched CPU baseline',flush=True)
if __name__=='__main__':main()
