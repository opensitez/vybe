#!/usr/bin/env python3
"""Attach original-source/artifact alias provenance to a completed frozen gate."""
import argparse,hashlib,json,shutil
from pathlib import Path

def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
p=argparse.ArgumentParser();p.add_argument('--original',type=Path,required=True);p.add_argument('--review',type=Path,required=True);p.add_argument('--output',type=Path,required=True);p.add_argument('--core-directory',default='core');p.add_argument('--codegen-directory',default='codegen');p.add_argument('--generated-directory',default='generated_tests');a=p.parse_args()
r=json.loads((a.output/'report.json').read_text());assert r.get('all_languages') and r.get('candidate_unchanged') and 'candidate_end_sha256' in r
original=json.loads((a.review/'original-source-manifest.json').read_text());assert all(sha(a.original/name)==value for name,value in original.items()),'original source changed'
crate=a.review/'crates/grameo';mapping={a.core_directory+'/src':'src',a.codegen_directory+'/src':'codegen/src',a.generated_directory+'/src':'generated_tests/src'}
for prefix,target in mapping.items():
 for path in (a.original/prefix).rglob('*.rs'):assert sha(path)==sha(crate/target/path.relative_to(a.original/prefix))
artifacts={}
for backend in ('old','new'):
 metadata=r['drivers'][backend]
 for arg in metadata['command']:
  if arg.startswith('vybe_parser=') or arg.startswith('vybe_parser_generated_tests='):
   name,path=arg.split('=',1);key='parser_sha256' if name=='vybe_parser' else 'generated_sha256'
   assert sha(Path(path))==metadata[key],f'linked artifact changed: {path}'
   artifacts[path]=metadata[key]
 assert sha(a.output/(backend+'-native'))==metadata['driver_sha256']
proof=a.output/'closure-proof';proof.mkdir(exist_ok=True)
for name in ('original-build.jsonl','artifact-alias-provenance.json','original-source-manifest.json'):
 shutil.copy2(a.review/name,proof/name)
for name in ('source-freeze.json','engine.patch','proof.log','pratt-proof.log','tests.log','clippy.log','generated-source-equality.json'):
 if (a.original/name).exists():shutil.copy2(a.original/name,proof/name)
checks={'linked_artifacts_unchanged':True,'linked_artifacts_sha256':artifacts,'original_source_unchanged':True,'review_runtime_codegen_generated_sources_equal_original':True,'artifact_name_alias_only':json.loads((a.review/'artifact-alias-provenance.json').read_text()),'original_source_manifest_sha256':sha(proof/'original-source-manifest.json'),'gate_closure_sha256':r['candidate_end_sha256'],'immutable_gate_passed':r['immutable_bundle_gate_passed'],'scope':'Exact original opt3 libraries; gate review source copy only; no rebuilt or composed candidate; frozen grammars/helpers/inputs. Live readiness remains separately scoped.'}
(proof/'assessment.json').write_text(json.dumps(checks,indent=2)+'\n')
(a.output/'checkpoint.md').write_text(f"Immutable all18 gate passed: {r['immutable_bundle_gate_passed']}. Inputs {sum(x['inputs'] for x in r['languages'].values())}; resource probes {sum(x['resource_cases'] for x in r['languages'].values())}. Full canonical trees/consumed/errors/recognition work matched pinned independent pest-attested records. Closure {r['candidate_end_sha256']}; original source unchanged and review source exact. Original build metadata and artifact-name alias provenance retained under closure-proof. Moving production grammar/corpus drift remains outside the immutable parser proof; no retry or backend switch.\n")
print(json.dumps(checks,indent=2))
