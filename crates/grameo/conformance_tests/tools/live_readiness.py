#!/usr/bin/env python3
"""One live inventory scan separate from a frozen immutable parser certificate."""
import sys,json,time,hashlib
from pathlib import Path
import reconcile_corpus as rec,native_differential as native
out=Path(sys.argv[1]).resolve();report=json.loads((out/'report.json').read_text());repo=Path(sys.argv[2]).resolve();crate=Path(report['candidate_crate']);result={'checked_unix':time.time(),'scope':'Fresh live canonical inventory/grammar scan, no moving-source retries or replay; immutable full18 proof remains separate','languages':{},'candidate_engine_sha256':hashlib.sha256((crate/'src/engine.rs').read_bytes()).hexdigest()}
for lang,proof in report['languages'].items():
 live=rec.inventory(repo,lang);s=rec.corpus_digest(live);gp=repo/f'languages/{lang}/src/grammar.pest';frozen=crate.parents[1]/f'languages/{lang}/src/grammar.pest';gs=hashlib.sha256(gp.read_bytes()).hexdigest();fs=hashlib.sha256(frozen.read_bytes()).hexdigest();row={'live_raw_files':len(live),'live_raw_corpus_sha256':s,'audited_raw_files':proof['raw_files'],'audited_raw_corpus_sha256':proof['raw_corpus_sha256'],'corpus_matches':s==proof['raw_corpus_sha256'],'live_grammar_sha256':gs,'audited_grammar_sha256':fs,'grammar_matches':gs==fs}
 if not row['corpus_matches']:
  old=dict(native.bundled_sources(out/f'{lang}.bundle',lang));row.update(added=sorted(set(live)-set(old)),deleted=sorted(set(old)-set(live)),modified=sorted(k for k in set(old)&set(live) if old[k]!=live[k].read_bytes()))
 result['languages'][lang]=row
result['live_gate_passed']=all(x['corpus_matches'] and x['grammar_matches'] for x in result['languages'].values());(out/'live-readiness.json').write_text(json.dumps(result,indent=2)+'\n')
(out/'checkpoint.md').write_text('# Frozen candidate audit\n\nAll 18 languages and 145,608 immutable canonical inputs passed complete capture/consumed/error/recognition/logical-step equality against pest-attested OLD records; all 1,920 direct resource probes passed. Candidate dependency closure stayed unchanged: '+report['candidate_end_sha256']+'.\n\nRaw inputs: 129,267; transformed Python/Fortran inputs: 16,341. Accepted: 135,013; rejected: 10,595. These are syntax/capture gates, not all-language common AST/runtime adapter certification.\n\nFresh live canonical corpus and grammar hashes are in live-readiness.json. Drifted live sources are not certified by this immutable audit. Default backend remains pest. Exact same 90-control old512/cache/pest CPU comparison is running separately before global promotion.\n')
print({k:(v['corpus_matches'],v['grammar_matches']) for k,v in result['languages'].items() if not(v['corpus_matches'] and v['grammar_matches'])})
