#!/usr/bin/env python3
"""Untimed physical work in the largest frozen accepted CPU gaps."""
import argparse,hashlib,json,os,shutil,subprocess
from pathlib import Path

def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--cpu',type=Path,required=True);ap.add_argument('--grammars',type=Path,required=True);ap.add_argument('--output',type=Path,required=True);args=ap.parse_args()
    cpu,out,grammars=args.cpu.resolve(),args.output.resolve(),args.grammars.resolve();out.mkdir(exist_ok=True,parents=True);r=json.loads((cpu/'report.json').read_text());assert r['complete']
    dependencies=r['dependencies'];parser=next(Path(p) for p in dependencies if '/libvybe_parser-' in p);generated=next(Path(p) for p in dependencies if '/libvybe_parser_generated_tests-' in p)
    assert sha(parser)==dependencies[str(parser)] and sha(generated)==dependencies[str(generated)]
    src=Path(__file__).with_suffix('.rs').resolve();binary=out/'profile';command=['rustc','--edition=2024','-O',str(src),'--extern',f'vybe_parser={parser}','--extern',f'vybe_parser_generated_tests={generated}','-L',f'dependency={parser.parent}','-o',str(binary)];env=os.environ.copy();env['DEVELOPER_DIR']='/Library/Developer/CommandLineTools'
    with (out/'build.log').open('w') as f:subprocess.run(command,stdout=f,stderr=subprocess.STDOUT,env=env,check=True)
    build={'command':command,'source_sha256':sha(src),'binary_sha256':sha(binary),'parser_sha256':sha(parser),'generated_sha256':sha(generated)};(out/'build.json').write_text(json.dumps(build,indent=2)+'\n');shutil.copy2(src,out/'profile.rs')
    report={'scope':'Untimed generated-program physical observers; every Program metadata/certificate/flag method and original observer forwarded. Counters exclude tasks elided by bulk scanners/certificates; boolean metadata query counts are not CPU costs (constants may compile out). Plain/wrapped complete captures+consumed and recognition logical steps equal for every control. Rule attempts have no completion hook: repeated attempts are an upper bound mixing successes/failures, not cache hit counts. Same environment flags remain insufficient if stack/scope/depth/diagnostic state matter.','cpu_report_sha256':sha(cpu/'report.json'),'build':build,'forwarded_methods':['observe_rule_attempt','observe_expression','leading_bytes','atomic_expression','terminal_candidates','terminal_choice','fast_rule_failures','failure_terminals','atomic_tail_dispatch','tail_continuations','tail_dispatch','fast_trivia_failures','fast_predicate_failures','literal_choice_failure','trivia_failure_prefix','fast_ascii_repetitions','rule_id','rule','instruction','source_pratt','scope','whitespace_class','whitespace_prefix_class'],'controls':[],'rule_shapes':{}}
    for lang,batch in [('ruby','corpus-dynamic'),('fortran','corpus-business'),('washm','corpus-vb-shell-wast')]:
        assert sha(grammars/lang/'src/grammar.pest')==r['languages'][lang]['grammar_sha256']
        for index,sample in enumerate(r['languages'][lang]['inputs']):
            if sample['reference']!='accepted':continue
            p=cpu/batch/f'matched-{lang}-{index}.source';assert sha(p)==sample['matched_source_sha256'];log=out/f'{lang}-{index}.tsv'
            with log.open('w') as f:subprocess.run([str(binary),lang,str(p),str(grammars)],stdout=f,stderr=subprocess.STDOUT,env=env,check=True)
            v={'language':lang,'label':sample['label'],'bytes':sample['bytes'],'source_sha256':sample['matched_source_sha256'],'parity':True,'phases':{},'generated_median_ratio_to_pest':r['languages'][lang]['workloads'][sample['label']]['generated']['median_ratio_to_pest']}
            shapes=report['rule_shapes'].setdefault(lang,{})
            for line in log.read_text().splitlines():
                t=line.split('\t');kind=t[0]
                if kind=='CAPTURES':v['consumed']=int(t[1]);v['capture_nodes']=int(t[2])
                elif kind=='STEPS':v['recognition_logical_steps']=int(t[1])
                elif kind=='SOURCE':shapes.setdefault(t[1],{})['grammar_text']=bytes.fromhex(t[2]).decode()
                else:
                    phase=t[1];c=v['phases'].setdefault(phase,{'rules':[],'expressions':[],'opcodes':{},'incoming_rule_flags':{},'metadata':{}})
                    if kind=='TOTAL':c.update(rule_attempts=int(t[2]),physical_expression_evaluations=int(t[3]),unique_rule_offsets=int(t[4]),repeated_rule_offset_attempts=int(t[5]))
                    elif kind=='EXACT_FLAGS_REPEATS':c['repeated_exact_rule_offset_flags_attempts']=int(t[2])
                    elif kind=='RULE':c['rules'].append({'id':int(t[2]),'name':t[3],'attempts':int(t[4]),'repeated_offset_attempts':int(t[5])});shapes.setdefault(t[2],{})['name']=t[3]
                    elif kind=='RULEMETA':
                        c['rules'][-1].update(mode=t[3],expression_id=int(t[4]),atomic_expression_certificate=t[5]=='true',repeated_exact_flags_attempts=int(t[6]),direct_calls=t[7].split(',') if t[7] else [],own_expression_opcode_counts=json.loads(t[8]))
                        shapes.setdefault(t[2],{}).update({k:c['rules'][-1][k] for k in ['mode','expression_id','atomic_expression_certificate','direct_calls','own_expression_opcode_counts']})
                    elif kind=='EXPR':c['expressions'].append({'id':int(t[2]),'evaluations':int(t[3]),'instruction':'\t'.join(t[4:])})
                    elif kind=='OPCODE':c['opcodes'][t[2]]=int(t[3])
                    elif kind=='FLAGS':c['incoming_rule_flags'][t[2]]=int(t[3])
                    elif kind=='META':c['metadata'][t[2]]={'queries':int(t[3]),'available_or_enabled':int(t[4])}
            for c in v['phases'].values():c['rules'].sort(key=lambda x:x['attempts'],reverse=True);c['expressions'].sort(key=lambda x:x['evaluations'],reverse=True)
            report['controls'].append(v)
    assert len(report['controls'])==9;report['complete']=True;(out/'report.json').write_text(json.dumps(report,indent=2)+'\n');print('Profile complete: nine accepted controls, exact parity')
if __name__=='__main__':main()
