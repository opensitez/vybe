# JavaScript frontend checkpoint — 2026-10-04

Production JavaScript continues to use pest. This checkpoint measures actual frontend changes separately from native capture-matcher improvements. The earlier frontend and atomic-matcher tables use the original JavaScript grammar. The final call-chain section compares explicit original/current grammar fixtures.

## Production changes

The frontend borrows the process-wide cached prelude and clones its statements directly into the final module body. The current pass reserves the extra strict directive slot exactly, scans the owned AST for `document` without cloning each statement, and reuses the original body vector when there are no top-level function declarations to hoist. The existing stable hoisting path handles modules that contain function declarations.

The frontend timing baseline is the immediately preceding pass: borrowed prelude cache, cloned document detection, unconditional hoisting vectors, and the missing strict directive capacity slot. Both baseline and current production variants use the same pest parser. The native variant uses the current frontend walk with the frozen native matcher. No production backend promotion, Pratt changes, or broader normalization rewrite is involved.

## Method and limits

Each round rotates previous/current/native order for ten warmups followed by 501 samples per variant. Timings include source indexing, parsing, the complete production AST walk, warm cached-prelude cloning/assembly, and AST destruction. They exclude filesystem reads, corpus comparisons, cold prelude initialization, WASM compilation, and execution. Allocation counts are recorded separately after caches are warm. The allocator counter is disabled during timing, with the same inexpensive allocator branch present for all variants.

Machine load was uncontrolled: other builds and the full source conformance gate ran concurrently. Large median/p95 drift is visible below. Three frontend rounds used the same release binary, with no rebuild between rounds. Geometric averages use every round's current/previous median ratio; no best run was selected. These results are directional rather than throughput or tail-latency certification.

## Complete frontend timing

Round 1, microseconds; each cell is median / p95:

| Workload | Previous pest frontend | Current pest frontend | Current native frontend |
|---|---:|---:|---:|
| tiny | 45.625 / 57.625 | 44.916 / 53.750 | 77.500 / 92.166 |
| declarations100 | 808.292 / 13539.792 | 780.833 / 15039.958 | 1921.875 / 20764.542 |
| expressions100 | 5340.417 / 44651.459 | 5502.375 / 47022.958 | 33478.250 / 80291.042 |
| comments100 | 955.167 / 10569.875 | 910.875 / 6358.375 | 1955.167 / 14278.458 |
| modern | 176.500 / 340.833 | 169.042 / 272.208 | 399.500 / 1610.625 |
| strict | 44.750 / 91.375 | 43.375 / 91.375 | 80.625 / 142.083 |
| for_loop | 88.084 / 183.708 | 85.125 / 176.084 | 177.041 / 365.292 |
| accessors | 71.417 / 152.542 | 69.375 / 167.125 | 138.292 / 735.125 |
| nested_document | 87.000 / 177.917 | 83.917 / 166.041 | 179.042 / 412.708 |
| document_strings | 52.791 / 100.833 | 51.666 / 83.375 | 95.333 / 162.125 |

All three rounds, geometric average of paired median ratios:

| Workload | Current/previous ratio | Time reduction | Native/current pest ratio |
|---|---:|---:|---:|
| tiny | 0.98297 | 1.70% | 1.79× |
| declarations100 | 0.96127 | 3.87% | 2.50× |
| expressions100 | 0.99580 | 0.42% | 3.75× |
| comments100 | 0.96125 | 3.87% | 2.13× |
| modern | 0.96334 | 3.67% | 2.42× |
| strict | 0.96952 | 3.05% | 1.87× |
| for_loop | 0.97522 | 2.48% | 2.11× |
| accessors | 0.97183 | 2.82% | 2.08× |
| nested_document | 0.96453 | 3.55% | 2.21× |
| document_strings | 0.97831 | 2.17% | 1.92× |

The initial expressions100 result was 3.0% slower despite fewer allocations, so two additional complete rounds were run specifically to investigate that unresolved regression. Retaining all results gives essentially flat expression timing (0.42% aggregate reduction), rather than a demonstrated expression throughput gain. The other workloads show 1.7–3.9% aggregate reductions. Native frontend parity with pest remains unmet.

| expressions100 round | Previous median / p95 | Current median / p95 | Native median / p95 | Current/previous median |
|---|---:|---:|---:|---:|
| 1 | 5340.417 / 44651.459 | 5502.375 / 47022.958 | 33478.250 / 80291.042 | 1.03033 |
| 2 | 3355.084 / 4731.375 | 3280.834 / 5730.542 | 9618.667 / 17249.958 | 0.97787 |
| 3 | 3324.750 / 4672.833 | 3258.500 / 4682.875 | 9669.125 / 14443.833 | 0.98007 |

Expression p95 worsens in every round. Its native median changes from 33.48 ms in the first round to 9.62/9.67 ms in subsequent rounds, showing substantial load drift. Other workloads also have mixed tails; no claim that all p95 values improve is justified.

## Frontend allocations

These are allocation-call / reallocation-call counts, not allocated bytes or peak live memory. Counts are identical across all three rounds.

| Workload | Previous pest frontend | Current pest frontend | Current native frontend |
|---|---:|---:|---:|
| tiny | 400 / 8 | 392 / 8 | 387 / 17 |
| declarations100 | 982 / 35 | 780 / 30 | 775 / 33 |
| expressions100 | 4382 / 37 | 3080 / 32 | 3075 / 36 |
| comments100 | 1288 / 36 | 986 / 31 | 981 / 32 |
| modern | 549 / 13 | 482 / 13 | 477 / 22 |
| strict | 391 / 11 | 386 / 10 | 381 / 18 |
| for_loop | 430 / 9 | 409 / 9 | 404 / 18 |
| accessors | 419 / 9 | 403 / 9 | 398 / 17 |
| nested_document | 407 / 9 | 397 / 9 | 392 / 18 |
| document_strings | 404 / 8 | 395 / 8 | 390 / 18 |

The 100-expression production frontend drops from 4,382 to 3,080 allocation calls (29.7%) and 37 to 32 reallocations. Declarations drop 20.6% and comments 23.4%. The strict module avoids one reallocation. These structural reductions hold even where timing gains remain unresolved.

The shared native capture node changed from 40 to 32 bytes on 64-bit targets. Typed pairs now share the original `Vec` allocation through `Rc<Vec<Node>>`, avoiding a full arena copy. That ownership preserves spare vector capacity until the pairs drop. Neither the node-size reduction nor the allocation-call counts establish a corresponding reduction in total peak memory.

## Native capture matcher timing

One frozen 501-sample rotating round. The before wrapper disables only `Program::atomic_expression`, retaining earlier terminal-choice, failure-prefix and tail-dispatch optimizations. Compact node storage applies to both native variants. This isolates compact atomic token execution rather than the complete historical memory change. Each cell is median / p95 in microseconds.

| Workload | Atomic tokens disabled | Current native capture | pest capture | Native time reduction | Current native/pest median |
|---|---:|---:|---:|---:|---:|
| tiny | 47.583 / 60.750 | 46.250 / 57.375 | 14.666 / 27.333 | 2.8% | 3.15× |
| declarations100 | 1816.375 / 7697.125 | 1720.917 / 6313.167 | 524.750 / 1463.542 | 5.3% | 3.28× |
| expressions100 | 19646.792 / 39305.625 | 18544.041 / 39454.500 | 2892.167 / 15241.667 | 5.6% | 6.41× |
| comments100 | 2268.666 / 11124.541 | 1720.375 / 9549.583 | 637.792 / 2956.333 | 24.2% | 2.70× |
| modern | 360.667 / 1077.583 | 357.375 / 1367.875 | 125.208 / 322.083 | 0.9% | 2.85× |

The compact token path cuts measured medians by 0.9–24.2% on these inputs, with the largest gain in comment-heavy source. Expression and modern-syntax p95 worsen. Current native capture medians remain 2.70–6.41 times pest in this round; native matcher parity is not achieved.

## Correctness scope

Before timing, the five capture workloads compare complete rule names, byte spans, nesting, capture counts and full consumption across reference, resolved, generated and atomic-disabled programs. The complete frontend compares all `Module` debug fields—including prelude, directives, spans and lowering results—on ten focused workloads: tiny, declarations, expressions, Unicode comments, modern class/arrow/optional-chain/template syntax, commented strict prologue, for loop, getter/setter accessors, nested document reference, and document only in strings/property names.

The frozen run also checked 64 deterministically spread repository JavaScript files no larger than 8 KiB plus 64 NUL mutations: 64 accepted programs have identical complete ASTs across previous/current/native frontends, and 64 inputs are rejected by all backends. Previous/current pest errors match exactly. Native syntax diagnostic wording is allowed to differ. The native walker uses the complete production walk with only the parser/Pair backend imports changed; the compiler's two accessor naming helpers are copied verbatim to avoid a compiler/runtime/plugin build.

This is a partial JavaScript frontend audit. It does not certify the full JS corpus, emitter/WASM output, runtime behavior, production native migration, or every language's common-AST path. The repository-wide source pass/fail gate is reported separately. Formatting, scoped diff checks, and strict Clippy pass for handwritten benchmark support; copied production walker code is excluded from benchmark-specific Clippy certification.

## Unshipped traversal-fusion experiment

A further experiment combined document detection with the existing postorder Promise normalization visitor. Detection ran after each rewrite. Resolve/reject retain their argument expressions, while supported array combinators retain their element expressions; only the Promise callee and array wrapper disappear, and neither can itself be `Ident(document)`. The rewrite introduces no document identifier. Exact before/after AST checks passed the original ten focused workloads, two additional Promise/document workloads, and the 128-source frontend gate.

The comparison used an explicit immediately-before-fusion frontend with all earlier allocation improvements still enabled. A single rotating 501-sample release round gave the following microseconds, median / p95:

| Workload | Before fusion | Fused candidate | Median time reduction |
|---|---:|---:|---:|
| tiny | 29.708 / 48.208 | 29.541 / 46.417 | 0.56% |
| declarations100 | 517.333 / 622.166 | 517.375 / 627.000 | -0.01% |
| expressions100 | 3321.833 / 4201.500 | 3327.375 / 4195.709 | -0.17% |
| comments100 | 797.875 / 938.417 | 797.833 / 922.375 | 0.01% |
| modern | 127.167 / 207.292 | 125.083 / 201.250 | 1.64% |
| strict | 33.959 / 66.417 | 33.666 / 60.834 | 0.86% |
| for_loop | 59.625 / 95.041 | 59.291 / 89.500 | 0.56% |
| accessors | 47.458 / 82.000 | 47.042 / 89.875 | 0.88% |
| promise_document | 101.916 / 137.958 | 101.833 / 143.292 | 0.08% |
| promise_without_document | 79.459 / 123.834 | 79.042 / 126.916 | 0.52% |
| nested_document | 67.625 / 116.250 | 67.042 / 108.666 | 0.86% |
| document_strings | 40.375 / 62.625 | 40.042 / 54.875 | 0.82% |

Large declarations, expressions and comment workloads are essentially flat; small cases mostly move less than 1%, with modern syntax at 1.6%. This does not demonstrate a meaningful throughput gain. The candidate was not shipped: production retains the separate Promise normalization and document-detection traversals described above. The historical timing/allocation results in this checkpoint still describe the retained production frontend.

The two Promise fixtures remain as extra semantic probes in the stock example. Experimental source snapshots and exact build arguments are saved under [javascript-fusion-experiment](../../../results/testrunner/lua-native-parser/javascript-fusion-experiment/), and [the complete fusion log](../../../results/testrunner/lua-native-parser/javascript-fusion-final.log) records all twelve comparisons. The temporary `--fusion-only` switch was removed after restoring production. [Restored verification](../../../results/testrunner/lua-native-parser/javascript-restored-verification.log) checks the stock generator/example and frozen core again.

## Validated call-chain factoring and parser profiling evidence

The validated call-chain common-prefix factoring is retained. Production grammar SHA-256: `60f845dcc2dd1e7207b7a2db0020165286296fa1a02fb4bc2e76b730c2ff4be6`. The saved original grammar is `d4fb47980f00d13a4a453de6ece65a9ff255fe3d5bd574a922a59771083bd5c4`. No decimal factoring or further grammar changes were made; continuing optimization targets the parser engine and codegen.

An actual-dispatch profiler found the unchanged expression workload repeatedly evaluating the silent `chain_lead = _{ NEWLINE* }` prefix five times at operator and argument boundaries. Common-prefix factoring quantified and reduced this cost; original/current grammars remain as explicit benchmark fixtures. It reduced expression rule attempts from 46,372 to 38,372, repeated rule/offset/mode attempts from 10,904 to 2,904, and expression evaluations from 223,496 to 196,296. Captures stayed at 9,002. Declaration rule attempts fell from 7,672 to 6,872.

The final paired run below used 501 samples per variant after the root's flat-sequence experiment had been rolled back. Earlier factoring timing used a different core and is not aggregated with this run. Units are microseconds, median / p95; both grammars use the same parser core within each paired binary.

| Workload | Original native captures | Candidate native captures | Original pest captures | Candidate pest captures |
|---|---:|---:|---:|---:|
| tiny | 43.584 / 74.708 | 40.208 / 69.458 | 14.791 / 33.875 | 14.000 / 32.083 |
| declarations100 | 1770.250 / 12228.750 | 1620.208 / 9957.625 | 516.416 / 2082.875 | 491.750 / 2578.125 |
| expressions100 | 19459.250 / 46788.167 | 16041.292 / 40429.875 | 3074.250 / 18715.667 | 2701.708 / 18115.250 |
| comments100 | 1688.834 / 11819.333 | 1533.959 / 11718.208 | 624.167 / 4683.209 | 600.125 / 4418.167 |
| modern | 351.542 / 890.042 | 318.042 / 926.041 | 124.625 / 223.250 | 113.875 / 230.833 |

| Workload | Original native frontend | Candidate native frontend | Original pest frontend |
|---|---:|---:|---:|
| tiny | 74.792 / 215.459 | 70.833 / 232.709 | 43.959 / 115.667 |
| declarations100 | 1858.541 / 9699.750 | 1674.750 / 11635.042 | 727.041 / 3885.584 |
| expressions100 | 20390.542 / 41835.167 | 17657.583 / 36316.708 | 4485.084 / 21238.625 |
| comments100 | 1969.333 / 15894.583 | 1800.917 / 14525.042 | 895.875 / 8156.000 |
| modern | 417.042 / 2505.208 | 380.917 / 1729.000 | 198.292 / 473.166 |

The expression current grammar's native captures were 17.6% faster and its complete native frontend was 13.4% faster, but it remained 5.94 times slower than current pest capture parsing. The production frontend remains pest-backed; pest capture parsing improved 12.1% on the same expression workload. Allocation/reallocation counts were unchanged. Heavy concurrent load caused long tails; several p95 values worsened despite lower medians. These timing tables do not measure the complete current pest frontend, so no complete production frontend percentage is inferred from its capture-only improvement.

Validation compared complete named capture trees, byte spans and nesting across original pest, candidate pest, original native and candidate native. All 576 chain-boundary probes passed, covering newlines, CR/CRLF, Unicode comments, optional/private/member access, bare calls/indexing, tagged templates, malformed suffixes, EOI and NUL. Fifteen focused inputs and 64 spread repository files also had identical complete common ASTs across all four original/current pest/native frontends including the runtime prelude; all 64 NUL-mutated files were rejected by all backends. This is finite fixture evidence, not a full equivalence proof or common-AST certification for every JavaScript input. The refreshed repository-wide JavaScript gate subsequently checked all 10,563 files: 10,543 accepted and 20 rejected, with zero capture/verdict/recognition/resource/unreadable discrepancies. This broader gate certifies parser agreement on the repository corpus, while common-AST validation remains scoped to the focused/spread frontend fixtures. [Repository gate evidence](../../../results/testrunner/parser-full-corpus-refresh/js.tsv) records the per-file results.

[The original dispatch profile](../../../results/testrunner/lua-native-parser/javascript-profile.log), [paired original/candidate profile](../../../results/testrunner/lua-native-parser/javascript-factored-profile-final.log), [final timing and parity log](../../../results/testrunner/lua-native-parser/javascript-factoring-saved-before.log), [timing summary](../../../results/testrunner/lua-native-parser/javascript-factoring-final-summary.json) and [artifact identity](../../../results/testrunner/lua-native-parser/javascript-factoring-artifact-proof.json) preserve the paired measurements. [Fresh four-frontend parity](../../../results/testrunner/lua-native-parser/javascript-factoring-final-parity.log) and [strict harness Clippy](../../../results/testrunner/lua-native-parser/javascript-factoring-clippy.log) also passed; the final profiler forwards all scanner metadata and its [build proof](../../../results/testrunner/lua-native-parser/javascript-final-profile-proof.json) identifies the linked release core. The `javascript_factoring` example uses saved grammar fixtures and can reproduce the original/current comparison without changing production grammar. The profiler suggests investigating a conservative fast success path for empty silent repetitions or nullable-aware failure certificates, with exact diagnostics and resource-limit replay. Engine and codegen implementation belongs to the parser track.

## Parser checkpoint layout

The retained parser optimization shares a checkpoint's capture-node cursor and builder journal cursor in one `usize`. Capture parsing always uses `NoBuilder`; semantic building and recognition never emit capture nodes. Input offset and grammar-stack undo cursors remain separate. This preserves journal values for semantic builders and node rollback for capture parsing, without changing grammar, span widths or the public builder API.

The isolated checkpoint-only comparison measured `Checkpoint` shrinking from 32 to 24 bytes and ordinary `Task` from 80 to 72 bytes. `AtomicTask` stayed at 48 bytes in this isolated comparison; the combined parser's separate atomic layout changes are outside this table. The expression workload retained task capacity 256 in both variants, reducing reserved ordinary-task arena storage from 20,480 to 18,432 bytes. Capture-node capacity and atomic-task capacities stayed equal. This is a task-arena measurement, not a total peak-memory claim.

Generated matcher, unchanged frozen grammar, 501 rotating samples per backend; microseconds, median / p95:

| Workload | Before checkpoint union | After checkpoint union | pest | Median reduction |
|---|---:|---:|---:|---:|
| tiny | 25.417 / 25.750 | 25.125 / 25.458 | 6.667 / 6.875 | 1.1% |
| declarations100 | 963.000 / 1383.625 | 953.334 / 1395.333 | 298.708 / 497.334 | 1.0% |
| expressions100 | 6182.708 / 8300.084 | 5968.292 / 7966.833 | 1553.708 / 2254.084 | 3.5% |
| comments100 | 993.750 / 1357.750 | 978.417 / 1300.625 | 374.708 / 557.833 | 1.5% |

An independent resolved-matcher comparison had 1.4–2.5% lower medians across the same four inputs. The results distinguish generated and resolved metadata access; their absolute medians come from separate runs under uncontrolled load. Declaration p95 worsened in the generated comparison, and native expression capture parsing remains about 3.84 times slower than pest in that run.

All 76 copied core integration tests passed, including builder rollback, Pratt, grammar-stack and resource-boundary tests. The isolated comparison also passed 42,000 exhaustive small-grammar mode/work/depth/diagnostic/capture probes, 50,400 shallow JavaScript resource probes, and exact complete before/after/pest capture equality on the four measured workloads. The shallow JavaScript depth sweep is supplementary; it does not exercise every deep JavaScript path. Fresh combined-production validation passed complete named capture and common-AST parity across all four original/current pest/native frontends: 576 boundary probes, 15 focused cases, and 64 spread accepted repository files plus 64 shared NUL rejections. Runtime prelude and UI directives were included. [Combined-production frontend parity](../../../results/testrunner/lua-native-parser/javascript-combined-production-parity.log) records the checks. The latest [dispatch profile](../../../results/testrunner/lua-native-parser/javascript-combined-final-profile.log) forwards the new terminal candidate and scanner hooks; [strict profiler Clippy](../../../results/testrunner/lua-native-parser/javascript-combined-profile-clippy.log) passed. Its [artifact proof](../../../results/testrunner/lua-native-parser/javascript-combined-profile-proof.json) uses current debug core/codegen artifacts for dispatch counts, not timing.

Evidence: [resolved timings](../../../results/testrunner/lua-native-parser/javascript-checkpoint-layout-experiment.log), [generated timings and capacities](../../../results/testrunner/lua-native-parser/javascript-checkpoint-generated-experiment.log), [76 core integration tests](../../../results/testrunner/lua-native-parser/javascript-checkpoint-core-tests.log), [minimal patch](../../../results/testrunner/lua-native-parser/javascript-parser-experiments/checkpoint-union.patch), [sizes/capacity summary](../../../results/testrunner/lua-native-parser/javascript-parser-experiments/checkpoint-summary.json). Source snapshots and hashes are saved in the same artifact directory. This checkpoint-only change was subsequently combined with the parser track's other validated changes.

## Rejected empty-repeat parser experiment

A parser-only shortcut for an empty silent repetition was tested in an isolated copy. It reused the child's certified first-byte failure, replayed diagnostics, and reserved the original work/depth costs, falling back to scheduled execution at tight budgets. The runtime-eligibility variant and a compiler-cached eligibility variant both passed scoped capture/resource checks but showed no useful speedup. The cached variant's 501-sample medians were 0.2–1.1% slower on all four workloads. Neither variant was promoted; source snapshots and [the cached comparison](../../../results/testrunner/lua-native-parser/javascript-empty-repeat-cached-experiment.log) remain as rejected evidence.

## Reproduction and artifacts

Use an isolated working directory to avoid root Cargo configuration and serialize builds with `-j 1`:

```sh
cd /private/tmp
DEVELOPER_DIR=/Library/Developer/CommandLineTools cargo run --release --manifest-path /Users/youness/www/html/vybe/crates/grameo/conformance_tests/Cargo.toml --offline -j 1 --features javascript-benchmark --example javascript_baseline -- 501
DEVELOPER_DIR=/Library/Developer/CommandLineTools cargo run --release --manifest-path /Users/youness/www/html/vybe/crates/grameo/conformance_tests/Cargo.toml --offline -j 1 --features javascript-benchmark --example javascript_ast_parity -- 501 --corpus=64
```

The earlier final frontend harness was compiled directly with `rustc -C opt-level=3` against finalized conformance release artifacts to avoid a shared Cargo lock. It links runtime core `libvybe_parser-ae33468390aed80f.rlib`, SHA-256 `fa2649bda0af99ba4c1becd99c75a6fa7bbd63534260907afb3b32b6fa3971be`. Its compiled artifact is newer than the frozen runtime sources. Exact arguments, source hashes, linked dependency paths and binary hash are saved with the logs:

- [Capture log](../../../results/testrunner/lua-native-parser/javascript-atomic-final.log)
- [Frontend round 1](../../../results/testrunner/lua-native-parser/javascript-frontend-final.log)
- [Frontend round 2](../../../results/testrunner/lua-native-parser/javascript-frontend-final-round2.log)
- [Frontend round 3](../../../results/testrunner/lua-native-parser/javascript-frontend-final-round3.log)
- [Frontend aggregate](../../../results/testrunner/lua-native-parser/javascript-frontend-final-summary.json)
- [Exact release arguments](../../../results/testrunner/lua-native-parser/javascript-frontend-final-arguments.json)
- [Frozen core and binary proof](../../../results/testrunner/lua-native-parser/javascript-frontend-frozen-core-proof.json)
- [Final strict Clippy log](../../../results/testrunner/lua-native-parser/javascript-final-clippy.log)

## Successful atomic-token memo experiment (unshipped)

An isolated narrow cache reused successful advancing `dec_digits` matches only in capture parsing. It keyed rule, offset and atomic/predicate/trivia flags, proved the token input-only, captureless, stackless and call-free, and retained the original execution whenever the remaining step/depth budget could not accommodate the complete recorded work. Failed, empty and semantic-builder matches were never cached. All 273,000 synthetic exact work/depth/capture/diagnostic probes passed; the four JS workload capture trees matched pest. Allocation/reallocation counts were unchanged.

The eight-slot 501-sample run reduced expression/declaration medians by 2.5%/3.0% but regressed the no-number comment control by 1.2%. A one-slot repeat reduced expression median by 14.9%, declarations by 0.1% and the comment control by 3.0%, with expression p95 around45ms. These shifts show machine-load confounding; neither run establishes a stable general gain. The experiment remains unshipped. Sources and proof: `results/testrunner/lua-native-parser/javascript-token-memo-experiment/`; raw logs: `javascript-token-memo-experiment.log` and `javascript-token-memo-one-slot.log`.

## Compact capture arena frontend check (isolated, unshipped)

A current production core snapshot was compared with the isolated compact arena under the identical generated JS grammar (`60f845dc…`). Only the capture arena and its adapters differed. All576 full capture/acceptance boundary probes,15 focused complete common AST cases,64 spread accepted corpus ASTs and64 rejected NUL mutations matched. The complete frontend includes production normalization, runtime prelude, directives and module destruction; it does not compile to WASM.

Allocator request volume includes the full requested size of reallocations. It measures allocation traffic, **not peak live memory, retained capacity or RSS**. Allocation/reallocation call counts stayed identical:

| Workload | Baseline bytes requested | Compact bytes requested | Reduction |
| --- | ---: | ---: | ---: |
| tiny |113,803|110,867|2.6%|
| declarations100 |635,541|537,373|15.4%|
| expressions100 |2,257,537|1,471,241|34.8%|
| comments100 |651,315|553,147|15.1%|

The501-sample rotating fullfrontend run gave these median/p95 microseconds:

| Workload | Baseline | Compact | Median change |
| --- | ---: | ---: | ---: |
| tiny |81.333 /222.084|80.042 /193.500|1.6% lower|
| declarations100 |1,946.667 /15,817.125|1,877.250 /12,890.250|3.6% lower|
| expressions100 |17,857.791 /35,379.125|18,155.000 /37,752.875|1.7% higher|
| comments100 |2,018.250 /8,800.708|2,056.416 /9,366.750|1.9% higher|

A no-rebuild paired501 repeat gave expressions29,240.916 /64,478.166 →26,862.833 /63,689.875 (8.1% lower median), comments2,047.292 /30,904.417 →2,272.334 /32,699.459 (11.0% higher median). Identical pest controls varied by up to7.7%; these frontend timings do not establish a stable speedup. The two comment regressions remain a concern.

Pre-parsed typed capture traversal has identical checksums. Baseline →compact median/p95 microseconds: tiny0.333/0.375 →0.500/0.542, declarations15.750/16.000 →22.750/23.250, expressions105.125/201.292 →148.334/285.333, comments16.708/18.333 →24.417/29.000. Compact-node decoding thus adds roughly40–50% to this traversal workload, although the43µs expression difference is small beside the loaded18ms frontend.

A separate snapshot added only `#[inline]` to `CompactNode::expand`. Samebinary originalcompact/inlinecompact expression traversal medians130.667/131.000µs were flat; the attribute did not resolve the traversal cost and was not promoted. Source, hashes, reproduction paths and scope are saved in `results/testrunner/lua-native-parser/javascript-compact-frontend-experiment/assessment.json`; raw logs are `javascript-compact-frontend.log`, `javascript-compact-frontend-repeat1.log`, `javascript-compact-traversal.log` and `javascript-compact-inline-traversal.log`.

### Cached typed node (isolated, unshipped)

A separate compact snapshot caches the decoded node already read by `Pairs::next`/`peek` in each `Pair`. Immutable arena ownership makes the cached fields stable. Complete capture/AST/corpus/NUL gates above pass, and three synthetic tests pass including cached clone/peek/span/string/rule/EOF operations in compact and wide arenas. `Pair` grows32→64bytes; allocation-call counts stay unchanged, but expression frontend request volume grows1,471,241→1,535,241bytes because temporary `Vec<Pair>` buffers also grow.

Paired501 traversal medians fall17–21%: expression152.416→126.542µs. Fullfrontend medians before→cached are tiny83.250→81.375, declarations1,791.291→1,757.708, expressions21,446.417→20,689.375, comments1,722.417→1,722.209µs. These0–3.5% frontend differences remain comparable to timing controls under heavy load; this experiment is unshipped. Proof and patch: `results/testrunner/lua-native-parser/javascript-compact-cached-pair-experiment/`.

### Shared typed context (isolated, unshipped)

Putting the borrowed source reference beside the arena in one `Rc` context reduces cached `Pair`64→48bytes and `Pairs`40→24bytes. Source lifetimes remain borrowed from the caller; the arena stays immutable. The same576/15/64/64 capture/AST/corpus/NUL checks and three compact/wide cached-handle tests pass; strict library Clippy passes.

Against the preceding64byte cached handle, paired501 traversal expression median improves35.750→33.958µs (5.0%). Fullfrontend medians are effectively flat: tiny49.250→49.500, declarations1,450.375→1,450.375, expressions11,996.958→11,958.542, comments1,489.334→1,492.667µs. Expression allocator request volume falls1,535,241→1,503,257bytes. A single context allocation grows16bytes per parse, while temporary handle buffers shrink. These are incremental variant comparisons; a matched production-wide versus final compact-context comparison is required to establish the production tradeoff. Evidence and patch: `results/testrunner/lua-native-parser/javascript-compact-cached-context-experiment/`.

### Final matched production-wide versus compact-context comparison (unshipped)

The final release executable compares the original production-wide core/original32byte `Pair` directly with compact storage plus48byte cached-context `Pair`, using identical generated grammar. It includes capture parsing, complete frontend, and traversal of pre-parsed typed captures. All576 capture/acceptance boundaries,15 focused complete ASTs,64 accepted corpus ASTs and64 rejected NUL mutations match. Exact core source hashes, release-library hashes, executable hash and reproduction paths are recorded in `results/testrunner/lua-native-parser/javascript-compact-final-experiment/`.

Allocator request bytes fall113,803→111,011 for tiny,635,541→537,389 declarations,2,257,537→1,503,257 expressions (33.4%),651,315→553,163 comments. Calls are unchanged. This is traffic rather than peak/retained memory. Typed `Pair` grows32→48bytes; `Pairs` shrinks40→24bytes.

Paired501 fullfrontend median/p95 microseconds:

| Workload | Production wide | Compact cached context |
| --- | ---: | ---: |
| tiny |82.250 /231.292|81.625 /272.083|
| declarations100 |1,810.208 /12,215.292|1,847.833 /13,658.583|
| expressions100 |21,639.208 /45,209.334|19,251.583 /45,502.459|
| comments100 |1,752.584 /14,909.958|1,762.625 /16,155.791|

Capture-only medians are43.708→43.000,1,514.708→1,498.167,16,673.875→15,990.625,1,507.333→1,508.542µs in the same order. Pre-parsed typed traversal medians fall0.458→0.416,21.959→18.875,135.834→124.083,21.625→16.500µs, resolving the original compact traversal regression. The expression frontend median is11.0% lower in this round, declarations2.1% higher, comments0.6% higher. Heavy tails and varying identical pest controls prevent a general speedup or improved-tail claim. The declarations regression remains unresolved; production is unchanged. Raw log: `javascript-compact-final-frontend.log`.

A targeted no-rebuild501 declaration repeat resolves the initial median-regression concern as unstable: capture median/p951,532.125/10,125.833 →1,506.041/9,636.292µs (1.7% lower median), complete frontend1,808.917/12,219.584 →1,795.334/13,483.042µs (0.8% lower median). Identical pest/frontend controls757.792→763.750µs differ0.8%. Across the two declaration rounds, the mean median ratio is approximately0.7% higher, effectively flat under those controls. Frontend p95 is worse in both rounds; this repeat does not establish improved tails. Typed traversal21.833→18.875µs remains lower. Separate raw log: `javascript-compact-final-declarations-repeat.log`. Production remains unchanged pending review.

### Combined root integration validation

After root integrated compact capture storage and the cached typed context while retaining ASCII batching, a fresh isolated copy of **all current parser/codegen sources** and the production JS grammar/walker/prelude/accessor helpers generated a new parser. The576 complete capture/acceptance boundaries,15 focused complete common AST cases,64 accepted corpus ASTs and64 rejected NUL mutations all match the frozen earlier native parser and pest. This run validates the exact combined root source; it performs no timing and does not promote the production JS frontend to native parsing or validate WASM compilation.

Frozen engine SHA256: `35c84bfff4a92bc0b8745ab353d9856f97d4a6605bc1ab7c665fd1152d958086`; compat: `d127f61461c726a4c20afc24bf62c20cee2706bb9cdcd27e9d718470a6c411b5`; grammar remains `60f845dcc2dd1e7207b7a2db0020165286296fa1a02fb4bc2e76b730c2ff4be6`. Full source/generated-output/binary hashes, reproduction command and harness are in `results/testrunner/lua-native-parser/javascript-root-combined-proof/`. Raw log: `javascript-root-combined-parity.log`. Earlier sections marked isolated/unshipped describe historical experiment snapshots; this note records the subsequently integrated combined root implementation.

### Batched thread-CPU capture benchmark of combined root

A separate release executable links the frozen pre-pass native parser, exact combined root parser and pest against the same grammar/input. Local SDK `_time.h` defines `CLOCK_THREAD_CPUTIME_ID=16`; a tiny C helper verifies42ns reported resolution,166ns minimum observed positive delta and approximately214ns CPU overhead per clock call. A21.1ms sleep advances thread CPU only11.7µs. Each sample batches64 tiny parses,8 declarations/comments or2 expression parses; two timer calls per batch are not subtracted. Capture construction **and destruction** are timed, with10warmup batches and501 recorded batches per backend/workload in rotated order. Input construction, grammar compilation, equivalence checks, result writes and AST normalization are excluded. CPU and wall values are saved separately for every batch.

CPU median/p95 microseconds per parse:

| Workload | Frozen native | Combined root | Pest |
| --- | ---: | ---: | ---: |
| tiny |43.265 /47.471|43.059 /47.102|12.755 /14.523|
| declarations100 |1,423.286 /1,499.599|1,417.031 /1,494.135|448.266 /521.255|
| expressions100 |8,220.438 /8,678.542|8,226.459 /8,699.188|2,106.584 /2,438.396|
| comments100 |1,432.849 /1,536.672|1,438.031 /1,526.281|571.505 /665.370|

The combined root changes CPU medians by−0.48%,−0.44%,+0.07%,+0.36% respectively. Paired current/old CPU ratio medians are1.0002,0.9925,0.9969,1.0006. This is effectively flat and **does not establish a parser CPU speedup for ASCII batching plus compact storage**. Current native still costs2.5–3.9times pest CPU on these inputs. Allocation/traversal improvements remain independent evidence. CPU tails are much tighter than wall tails: declaration current CPU p951,494µs vs wall3,788µs, expression CPU8,699µs vs wall21,115µs. Earlier wall gains should not be interpreted as equivalent CPU gains.

Full capture trees match all three backends before timing. Timer C source, Rust harness, exact source/generated/parser-library/binary hashes and detailed statistics are in `results/testrunner/lua-native-parser/javascript-cpu-bench-proof/`; raw batched samples are `javascript-cpu-bench-samples.csv`, output log `javascript-cpu-bench.log`.

### Generated guarded-literal atomic success (rejected, unshipped)

A frozen isolated candidate adds direct success for call-free literals followed by a negative ASCII boundary. Independent cost inspection agrees with root's focused resource tests: common keyword body11tasks plus one `FinishRule`; depth guard executes first, and insufficient remaining work uses the unchanged path. Failure probes and semantic builders retain ordinary execution.

The paired501 batched CPU comparison against combined production (`35c84bff…`), using unchanged JS/Lua grammars and exact capture parity with pest, rejects the ungated prototype. JS standard tiny/declaration/expression/comment CPU medians rise3.35%,2.59%,3.18%,2.50%; even a separately labeled keyword-heavy JS workload rises1.99%. Lua tiny/declaration/expression/comment changes are+1.72%,−0.05%,+1.94%,−0.81%. Paired CPU ratios likewise show JS regressions. The added probe overhead outweighs useful task elimination on these inputs. No production source or grammar changed. Source/generated/library/binary hashes, timer/source commands and statistics are archived in `results/testrunner/parser-atomic-success-experiment/cpu-proof/`; raw samples/log are `atomic-success-cpu-samples.csv` and `atomic-success-cpu.log`.

### Deferred diagnostics: Lua CPU validation (isolated, unshipped)

The immutable archived deferred-diagnostics candidate was copied into a standalone Lua-only harness and compared with frozen combined production (`35c84bff…`) and pest. Both native generators read the identical Lua grammar. Capture/result construction and destruction are measured with verified thread CPU and separate wall clocks, 501 paired batches in rotating backend order and 10 warmups. Compilation, input creation, parity checks and output are excluded. This measures parser capture work, not normalization or frontend startup.

| Lua workload | Production CPU median / p95 µs | Candidate CPU median / p95 µs | Pest CPU median / p95 µs | Candidate change |
|---|---:|---:|---:|---:|
| tiny | 23.86 / 26.42 | 15.09 / 16.67 | 9.32 / 11.17 | -36.74% |
| declarations100 | 822.90 / 871.90 | 640.19 / 665.95 | 436.34 / 518.10 | -22.20% |
| expressions100 | 4166.21 / 4370.94 | 3150.88 / 3271.58 | 1946.56 / 2285.04 | -24.37% |
| comments100 | 965.93 / 1025.42 | 753.14 / 795.46 | 526.95 / 599.49 | -22.03% |
| reject_early | 12.20 / 13.97 | 21.24 / 24.38 | 5.36 / 6.66 | +74.05% |
| reject_declarations100 | 857.92 / 906.04 | 1509.90 / 1576.85 | 445.85 / 544.10 | +76.00% |
| reject_expressions100 | 4232.52 / 4411.06 | 7427.21 / 7656.73 | 1948.65 / 2348.17 | +75.48% |

Accepted-input CPU medians fall 22–37%, and paired CPU ratios agree. Candidate/pest CPU median gaps are 1.43–1.62× on these accepted workloads. Rejected inputs cost 74–76% more CPU because the candidate first tries without diagnostic expectations and then replays the ordinary diagnostic parse. The early rejection is `local x = `; late rejections append NUL to the same 100-declaration/expression accepted inputs. This rejection cost is part of the result, not an excluded warmup.

All four accepted complete capture trees equal both production and pest. All three rejected inputs have identical complete native `ParseError` fields and are rejected by pest. This focused check does not replace the root resource-boundary or all-language gates. Production and the immutable candidate remain unchanged. Exact source/generated/library/binary hashes, commands and CPU/wall statistics are in `results/testrunner/parser-deferred-diagnostics-experiment/lua-cpu-proof/assessment.json`; `lua-cpu-samples.csv` preserves every batch and `lua-cpu.log` reports both clocks. Allocation counts and peak memory were not measured in this CPU run.

### Lazy expectation deduplication: CPU comparison (isolated, unshipped)

A separate parser-only candidate retains diagnostic collection and delays expectation deduplication, compacting every 32 pending records and before a final syntax error. It does not replay rejection. Independent inspection confirms that first-seen `Expected` variant/value order is preserved before rendering; writes cannot overwrite unread entries because the unique prefix never exceeds the read position. Farthest-offset advancement resets both the vector and prefix. Work/depth limits and other errors follow the original path. Pending redundancy is bounded to 31 records between compactions; retained vector capacity can exceed the original. Peak memory was not measured.

Against the same frozen production `35c84bff…` baseline, 501 paired rotating CPU batches give:

| Workload | Production CPU median / p95 µs | Candidate CPU median / p95 µs | Pest CPU median / p95 µs | Candidate change |
|---|---:|---:|---:|---:|
| js_tiny | 43.53 / 49.73 | 37.05 / 41.55 | 13.38 / 16.07 | -14.89% |
| js_declarations100 | 1447.72 / 1527.69 | 1257.26 / 1322.18 | 477.20 / 528.15 | -13.16% |
| js_expressions100 | 8524.54 / 8934.52 | 7884.83 / 8230.67 | 2222.33 / 2525.83 | -7.50% |
| js_comments100 | 1451.08 / 1531.39 | 1301.35 / 1362.78 | 581.15 / 653.42 | -10.32% |
| js_keywords100 | 5897.40 / 6625.54 | 4753.40 / 5319.23 | 1963.94 / 2188.75 | -19.40% |
| lua_tiny | 23.79 / 25.85 | 20.82 / 22.82 | 9.28 / 10.70 | -12.49% |
| lua_declarations100 | 841.85 / 873.58 | 695.63 / 725.19 | 432.24 / 507.24 | -17.37% |
| lua_expressions100 | 4207.40 / 4382.75 | 3544.65 / 3674.44 | 1940.35 / 2323.73 | -15.75% |
| lua_comments100 | 950.64 / 993.20 | 794.61 / 825.83 | 510.46 / 592.47 | -16.41% |
| js_reject_early | 17.80 / 19.80 | 17.76 / 19.60 | 5.88 / 7.46 | -0.18% |
| js_reject_declarations100 | 1100.09 / 1494.40 | 943.10 / 1294.57 | 371.25 / 508.31 | -14.27% |
| js_reject_expressions100 | 6312.88 / 8510.10 | 5854.40 / 7859.50 | 1777.71 / 2195.75 | -7.26% |
| lua_reject_early | 8.20 / 9.51 | 8.29 / 9.81 | 3.44 / 4.10 | +1.17% |
| lua_reject_declarations100 | 699.33 / 745.40 | 561.06 / 609.90 | 392.89 / 439.65 | -19.77% |
| lua_reject_expressions100 | 3516.23 / 3694.29 | 2879.83 / 3128.75 | 1797.00 / 1963.02 | -18.10% |

Standard accepted JS workloads improve 7.5–14.9% CPU and Lua improves 12.5–17.4%; the separately labeled keyword-heavy JS control improves 19.4%. Paired ratios agree. Early rejection is flat on JS (−0.18%) and slightly worse on Lua (+1.17%); late rejections improve 7.3–19.8%. These controls avoid the deferred candidate's approximately 75% rejected-input CPU increase. Separate benchmark rounds had different CPU frequency/cache/load, so absolute times from the deferred and lazy rounds should not be compared directly. CPU improvements also do not establish that every wall-time tail improves; raw CPU and wall measurements remain separate.

All nine accepted complete captures match production, candidate and pest exactly. Six rejected inputs have identical full native error fields and pest rejection. No grammar or production source changed. This focused proof does not replace resource-boundary/all-language validation. The measured source precedes a later Clippy fix that names the loop's initial bound; the only difference is preserved in `final-clippy-alias.diff`. Both forms evaluate the same initial range. The measured source, baseline snapshots, fresh generated programs, library/binary hashes, commands, raw samples and statistics are in `results/testrunner/parser-lazy-expectations-experiment/cpu-proof/`, with `cpu-samples.csv` and `cpu.log` beside it. Timing includes parser result construction/destruction and excludes compilation, proof and output; frontend/normalization are outside this run.

### General task and repetition frames: CPU comparison (isolated, unshipped)

This isolated candidate starts from the **final lazy-expectations revision**, including the Clippy initial-bound alias. Results below are additional gains beyond lazy expectations, not a comparison with production `35c84bff…`. General tasks shrink from 72 to 56 bytes on this 64-bit target. Repetition tasks carry indices into parser-owned active frames; a 64-byte frame holds repetition state and its iteration checkpoint once, reused across iterations. Frames are nested LIFO and removed at bounded completion, terminal child failure or skip failure. An inert initial checkpoint is overwritten by the original `Repeat` task, preserving semantic callback counts. There are no per-iteration frame allocations. The new vector retains its maximal active-frame capacity, so task size alone does not establish lower total allocation or peak memory.

`FinishRule` carries one `Option<usize>` selected by constant capture mode. Capture mode uses `NoBuilder`; semantic mode retains the exact rule ID. No sentinel is used, preserving custom semantic rule IDs such as `usize::MAX`. Existing logical task/work/depth boundaries and rollback are preserved.

| Workload | Final lazy CPU median / p95 µs | Task/frame CPU median / p95 µs | Pest CPU median / p95 µs | Incremental change |
|---|---:|---:|---:|---:|
| js_tiny | 37.25 / 41.25 | 34.88 / 39.15 | 13.16 / 15.53 | -6.36% |
| js_declarations100 | 1252.68 / 1316.63 | 1161.24 / 1205.86 | 476.46 / 529.01 | -7.30% |
| js_expressions100 | 7951.40 / 8361.44 | 7467.90 / 7800.25 | 2277.25 / 2561.75 | -6.08% |
| js_comments100 | 1278.18 / 1356.49 | 1188.35 / 1253.20 | 581.86 / 643.56 | -7.03% |
| js_keywords100 | 5054.52 / 5393.83 | 4656.29 / 4974.42 | 2122.65 / 2372.35 | -7.88% |
| lua_tiny | 20.60 / 23.32 | 19.48 / 22.34 | 9.38 / 11.21 | -5.44% |
| lua_declarations100 | 672.82 / 716.30 | 622.54 / 662.67 | 435.20 / 508.42 | -7.47% |
| lua_expressions100 | 3367.48 / 3623.19 | 3099.00 / 3381.27 | 1933.12 / 2262.25 | -7.97% |
| lua_comments100 | 771.72 / 825.71 | 710.70 / 774.47 | 500.60 / 582.90 | -7.91% |
| js_reject_early | 17.81 / 20.74 | 16.80 / 19.35 | 5.96 / 8.39 | -5.68% |
| js_reject_declarations100 | 1227.16 / 1312.71 | 1147.68 / 1239.25 | 465.28 / 531.43 | -6.48% |
| js_reject_expressions100 | 7647.58 / 8064.46 | 7203.88 / 7595.77 | 2172.83 / 2488.40 | -5.80% |
| lua_reject_early | 11.75 / 13.12 | 10.31 / 12.68 | 4.86 / 6.21 | -12.27% |
| lua_reject_declarations100 | 674.76 / 721.03 | 603.20 / 676.54 | 435.46 / 521.88 | -10.61% |
| lua_reject_expressions100 | 3331.44 / 3567.35 | 3053.56 / 3359.17 | 1917.81 / 2107.08 | -8.34% |

All 15 CPU medians improve and paired median ratios agree. Standard accepted JS improves 6.1–7.3%, Lua 5.4–8.0%, and the keyword-heavy JS control 7.9%. Rejected JS improves 5.7–6.5% and Lua 8.3–12.3%. Wall-time measurements remain separate: some medians and tails regress under heavy scheduling load, so this does not establish uniform latency gains. Candidate/pest CPU median gaps remain 2.04–3.28× for standard accepted JS and 1.42–2.08× for Lua.

The isolated core/codegen suite passes 87 tests, including layout and maximum semantic rule-ID checks. A fresh generated-parser release proof passes 92,736 comparisons of full captures/errors, recognition costs and complete semantic checkpoint/begin/finish/rollback traces plus committed history across nested/bounded/nullable/predicate/stack/atomic/compound repeats and tight work/depth limits. A separate optimized executable linked to the same release libraries passes 117 `usize::MAX` semantic rule-ID success/failure/resource comparisons. The benchmark's nine accepted trees equal both native backends and pest; six rejected inputs preserve complete native errors and pest rejection.

Exact frozen baseline/candidate source, patch, fresh generated programs, release library/binary hashes, commands and statistics are archived at `results/testrunner/parser-repeat-frames-experiment/cpu-proof/`. Raw `cpu-samples.csv` records all 22,545 batches (15 workloads × three backends × 501); `cpu.log` records both clocks. The core and handwritten harness Clippy checks are archived beside the proof. Production, grammar and previously audited candidates remain unchanged. Root's all-language gate is required before promotion. Allocation traffic and peak memory were not measured in this run.

### Wide-choice first-byte failure certificates (isolated, unshipped)

This candidate starts from the frozen final task/frame revision. It skips a certified whole-left subtree in wide ordered choices when the next byte proves failure, while reserving the complete original work/depth cost and preserving terminal expectation order. Tight guards and all semantic-builder paths retain ordinary dispatch. The refined version skips expectation decoding only when existing failure handling would discard every same-offset expectation. The `Program` hook defaults to `None`; generated programs use static borrowed metadata. Grammar and production source remain unchanged.

501 paired, rotated-order batches per backend use the verified thread CPU timer; wall time remains recorded separately. Compilation, grammar generation, input creation, proof and output are excluded.

| Workload | Final frame CPU median / p95 µs | Choice CPU median / p95 µs | Pest CPU median / p95 µs | Incremental change |
|---|---:|---:|---:|---:|
| js_tiny | 22.16 / 28.51 | 20.01 / 25.21 | 7.64 / 11.44 | -9.74% |
| js_declarations100 | 868.95 / 993.93 | 817.18 / 927.18 | 373.83 / 408.05 | -5.96% |
| js_expressions100 | 5956.96 / 6352.44 | 5702.65 / 6044.35 | 1820.69 / 1940.88 | -4.27% |
| js_comments100 | 797.64 / 984.66 | 743.68 / 908.51 | 416.80 / 479.12 | -6.77% |
| js_keywords100 | 2664.42 / 3134.38 | 2460.56 / 2933.27 | 1196.60 / 1469.19 | -7.65% |
| lua_tiny | 12.13 / 17.34 | 11.58 / 16.04 | 5.50 / 8.64 | -4.53% |
| lua_declarations100 | 461.92 / 556.83 | 479.64 / 555.51 | 354.93 / 387.02 | +3.84% |
| lua_expressions100 | 2146.08 / 2762.73 | 2148.23 / 2739.85 | 1496.85 / 1680.90 | +0.10% |
| lua_comments100 | 437.82 / 546.85 | 442.08 / 534.53 | 367.56 / 386.03 | +0.97% |
| js_reject_early | 10.27 / 13.51 | 9.80 / 12.72 | 3.42 / 5.38 | -4.58% |
| js_reject_declarations100 | 813.38 / 988.76 | 760.56 / 923.66 | 335.21 / 400.39 | -6.49% |
| js_reject_expressions100 | 6002.69 / 6472.58 | 5773.90 / 6268.65 | 1842.10 / 2076.08 | -3.81% |
| lua_reject_early | 8.20 / 10.45 | 7.33 / 9.13 | 3.92 / 5.03 | -10.63% |
| lua_reject_declarations100 | 518.96 / 578.33 | 519.92 / 578.29 | 371.82 / 433.56 | +0.18% |
| lua_reject_expressions100 | 2718.62 / 2895.92 | 2704.19 / 2858.67 | 1650.92 / 1947.96 | -0.53% |

All five accepted JS controls improve 4.3–9.7% CPU median, and rejected JS improves 3.8–6.5%. Lua tiny and early rejection improve, but accepted declarations regress 3.84% (paired median 3.50% worse); accepted expressions/comments and late rejections are approximately flat. This universal candidate remains unshipped. The additional per-choice metadata lookup can outweigh dispatch savings; narrower eligibility or a different lookup layout needs separate evidence. CPU p95 and wall-time tails are recorded without a uniform latency claim. Allocation traffic and peak memory were not measured.

Core/codegen tests (88), strict Clippy, 148,120 resolved capture/error resource comparisons and 169,050 freshly generated capture/error/recognition-cost/complete builder-trace comparisons pass. All nine benchmark accepted trees match both native backends and pest; six rejected inputs preserve complete native errors and pest rejection. Exact refined source/patch and focused proof are in `results/testrunner/parser-choice-prefix-experiment/refined-focused-proof/`; measured sources, generated programs, library/binary hashes, commands and statistics are in `cpu-proof/`, with all 22,545 raw batches in `cpu-samples.csv` and both clocks in `cpu.log`. The earlier unrefined focused proof remains preserved separately.

### Generated marked-choice dispatch (rejected, unshipped)

A separate candidate preserves the universal-lookup source/proof above. It appends `Instruction::CertifiedChoice` only for eligible generated choices, leaving ordinary `Choice` with its original zero-hook dispatch. The marked variant retains ordinary execution when the default `Program::choice_failure` hook returns `None`; atomic matching and terminal-choice shortcuts treat both variants identically. Resolved grammar instructions stay ordinary choices.

The same frozen final-frame baseline, unchanged JS/Lua grammar inputs, nine accepted and six rejected controls were tested in one release executable with 501 paired rotated batches and separate CPU/wall clocks. Absolute times between separate runs should not be compared: frequency/cache/scheduling conditions changed.

| Workload | Final frame CPU median / p95 µs | Marked-choice CPU median / p95 µs | Pest CPU median / p95 µs | Incremental change |
|---|---:|---:|---:|---:|
| js_tiny | 12.34 / 12.97 | 11.29 / 11.76 | 4.21 / 4.52 | -8.52% |
| js_declarations100 | 420.42 / 432.08 | 397.63 / 407.12 | 185.27 / 197.28 | -5.42% |
| js_expressions100 | 3827.79 / 5446.50 | 3702.23 / 5302.19 | 1286.15 / 1657.75 | -3.28% |
| js_comments100 | 442.57 / 454.64 | 412.47 / 423.25 | 233.99 / 248.38 | -6.80% |
| js_keywords100 | 1708.44 / 1754.62 | 1589.85 / 1631.31 | 766.60 / 814.21 | -6.94% |
| lua_tiny | 6.78 / 7.11 | 6.56 / 6.90 | 2.88 / 3.06 | -3.30% |
| lua_declarations100 | 213.35 / 220.19 | 229.35 / 235.82 | 186.22 / 193.29 | +7.50% |
| lua_expressions100 | 1122.19 / 1155.15 | 1190.65 / 1223.19 | 842.04 / 876.77 | +6.10% |
| lua_comments100 | 252.50 / 259.22 | 259.36 / 265.90 | 216.99 / 224.69 | +2.72% |
| js_reject_early | 5.98 / 6.26 | 5.74 / 5.98 | 1.94 / 2.06 | -3.95% |
| js_reject_declarations100 | 426.10 / 440.51 | 400.68 / 412.34 | 185.78 / 194.76 | -5.97% |
| js_reject_expressions100 | 2770.23 / 2991.71 | 2651.29 / 2871.62 | 929.23 / 1020.54 | -4.29% |
| lua_reject_early | 3.85 / 4.24 | 3.47 / 3.85 | 1.80 / 2.02 | -9.87% |
| lua_reject_declarations100 | 215.07 / 224.11 | 232.95 / 241.48 | 189.38 / 201.24 | +8.31% |
| lua_reject_expressions100 | 1132.06 / 1187.77 | 1205.71 / 1258.69 | 858.29 / 907.67 | +6.51% |

Removing the lookup on ordinary choices retains JS gains but does not fix Lua: accepted declarations/expressions and late rejections regress roughly 6–8%; comments regress 2.7%, with paired median ratios agreeing under much steadier controls. This candidate is rejected and remains isolated. The lookup on noneligible choices alone therefore does not explain the earlier regression; no grammar, production parser or frontend changed. Allocation traffic and peak memory were not measured.

All 88 core/codegen tests and strict Clippy pass. At 148,120 resource scenarios, ordinary instructions exactly match both marked instructions without a certificate and marked instructions with a certificate (296,240 comparisons). A fresh generated baseline/candidate proof passes 169,050 complete capture/error/recognition-cost/semantic callback-history comparisons. Benchmark accepted captures match both native backends and pest; rejected inputs retain complete native errors and pest rejection. Source/patch/focused proofs and measured source/generated programs/library and binary hashes/commands/raw timing are preserved separately at `results/testrunner/parser-certified-choice-experiment/focused-proof/` and `cpu-proof/`, with `cpu-samples.csv` and `cpu.log` beside them. A disk-full link failure was resolved by cleaning only disposable isolated debug/test build artifacts; the measured executable came from the successful retry.

### Structural atomic-token single-slot cache (isolated, unshipped)

This candidate starts from audited batch512 (`b0beb48a…`). Profiling found 1,800 successful nested leaf calls on JS expressions100, with 600 distinct rule/offset/flag keys and 1,200 one-slot hits (67%). All repeated calls were the unchanged digit token; eight slots offered no additional hits. Lua controls had no nested pure-leaf successes. Eligibility uses grammar structure, with no language or rule-name tests: silent, call-free unbounded ASCII class runs, including class followed by repetition of class or a case-sensitive separator plus class. Overlapping separators, folded separators, bounded loops, stack effects and rule calls retain ordinary dispatch.

Only certified generated call sites carry the new marker. Ordinary calls have no cache lookup; generated Lua has no marked calls and its opt-in is `None`. Resolved programs and custom programs retain default `None`. Recognition and semantic builder paths keep ordinary execution. The fixed single slot stores successful progress and the complete original body/finish cost; tight work/depth guards execute real tasks. Captures inside these silent atomic leaf calls are absent by construction, while outer captures stay unchanged.

Success is not diagnostic-free. On a terminating separator suffix the cache reports the ASCII builtin failure beyond the consumed end; otherwise it replays the builtin and separator literal failures at the end, in original order. Replay uses the original failure handler, preserving farthest and predicate/trivia suppression. Default custom-program observers retain every callback; explicit cache opt-in reports actual expression dispatches. A positive execution check reduced expression dispatches from 54 to 26 while retaining rule-entry observer counts.

| Workload | Round1 CPU median change | Round2 CPU median change | Median of two paired run ratios: change |
|---|---:|---:|---:|
| js_tiny | -5.94% | -5.64% | -6.05% |
| js_declarations100 | -5.42% | -6.81% | -6.03% |
| js_expressions100 | -6.23% | -8.34% | -6.66% |
| js_comments100 | -1.84% | -2.69% | -2.61% |
| js_keywords100 | -1.06% | -0.85% | -1.13% |
| lua_tiny | -1.10% | -0.73% | -1.06% |
| lua_declarations100 | +1.18% | +2.17% | +1.21% |
| lua_expressions100 | -1.53% | -2.02% | -1.65% |
| lua_comments100 | +0.68% | -2.90% | -0.73% |
| js_reject_early | -1.24% | -1.82% | -1.50% |
| js_reject_declarations100 | -5.37% | -5.49% | -5.71% |
| js_reject_expressions100 | -5.50% | -6.00% | -5.99% |
| lua_reject_early | -0.80% | -0.70% | -0.76% |
| lua_reject_declarations100 | +1.26% | +2.20% | +0.79% |
| lua_reject_expressions100 | -1.39% | -2.61% | -1.89% |

Both no-rebuild 501-sample rounds show approximately 5–8% CPU improvements on JS digit-heavy accepted/rejected controls. Lua declarations retain a small paired regression around 1.2%, so this version remains isolated while a separate opt-out specialization is tested. No best round was selected; all CPU/wall medians, p95, paired ratios and 45,090 batches are preserved. The first round’s current/pest CPU median gap is 1.28–2.02× for standard JS controls and 1.15–1.76× for Lua. Wall-time tails remain noisy; these data do not establish uniform latency improvement.

The enabled cache passes 87 core/codegen tests, strict Clippy, 763,488 fresh generated full capture/error/recognition-cost/semantic callback-history comparisons, 1,344 default-optout observer/capture/error comparisons and 220 guards spanning more than 512 expected records. The original prototype’s opt-in was initially emitted under an unrelated optional table; a structural eligibility unit test caught this, and the final proof explicitly asserts opt-in and cache execution. Initial fallback-only logs are retained but excluded from cache validation. Atomic task size stays 40 bytes and instruction size 24 bytes on the 64-bit target. Cache state is bounded and adds no cache heap allocation; total allocation traffic and peak memory were not measured.

Sources, exact frozen baseline/candidate revisions, fresh generated programs, release library/binary hashes, commands, calibrated thread timer, proofs and both raw runs are preserved at `results/testrunner/parser-atomic-token-cache-experiment/` and protected build artifacts at `results/testrunner/parser-active-builds/atomic-token-cache/`. Timing includes capture/result creation and destruction and excludes compilation, input creation, proof and output. No grammar, production parser or frontend changed.

### Atomic-token cache with no-eligibility specialization (isolated)

The frozen generated-only candidate specializes the atomic driver with a constant memo flag. Grammars without eligible structural tokens retain the false driver; the bounded cache fields remain in parser state. This removes the repeated Lua declaration regression seen in the initial candidate. Production is unchanged pending the full eighteen-language audit.

Two no-rebuild rounds each measured 501 paired, rotated batches for fifteen accepted/rejected controls against audited batch512 and pest. The table reports the median of the two run paired CPU median ratios.

| Input | Paired CPU change |
|---|---:|
| js_tiny | -6.95% |
| js_declarations100 | -5.96% |
| js_expressions100 | -7.08% |
| js_comments100 | -2.52% |
| js_keywords100 | -1.10% |
| lua_tiny | -1.47% |
| lua_declarations100 | -2.18% |
| lua_expressions100 | -2.68% |
| lua_comments100 | -1.84% |
| js_reject_early | -1.79% |
| js_reject_declarations100 | -6.00% |
| js_reject_expressions100 | -6.53% |
| lua_reject_early | -1.43% |
| lua_reject_declarations100 | -1.22% |
| lua_reject_expressions100 | -2.63% |

All fifteen paired medians improved in both rounds. Wall and CPU tails remain noisy; this does not establish a uniform latency improvement. Captures, consumed spans and native errors were exact before timing. The release proof passed 763,488 capture/error/work and semantic-builder callback/history comparisons, 1,344 default custom-program observer comparisons, and 220 expectation-threshold/resource cases. Core/codegen tests (87) and strict Clippy passed. Allocation traffic and peak memory were not measured.

Exact source hashes, patch, two raw rounds, CPU/wall medians and p95, generated parsers, binaries and library hashes are preserved in [the specialized assessment](../../../results/testrunner/parser-atomic-token-cache-specialized-experiment/assessment.json). Engine SHA `05e40e4642a54a8fcc06532628484428a90ee599db8696d62c7a8940b3152ca4`; unchanged audited batch512 baseline SHA `b0beb48a348198f11b10824d8bd258738688482c06da5864e5409c14d0acc2e6`. The full corpus audit is requested; no broader token cache is included.

### Successful atomic-body reuse profile (untimed)

A separate baseline512 snapshot observed completed atomic bodies on nine frozen Ruby, Fortran and shell controls. Instrumented and uninstrumented complete captures, consumed spans and recognition work were exact. This separates successful reuse from the earlier physical attempt counts.

| Token | Successful progress | Same-key successful repeats |
|---|---:|---:|
| Ruby identifier | 231 | 197 |
| Ruby keyword | 110 | 99 |
| Ruby constant | 33 | 26 |
| Fortran identifier | 162 | 125 |
| Fortran loop label | 297 | 264 |
| Fortran strict keyword | 48 | 36 |
| Shell name | 58 | 11 |
| Shell redirection descriptor | 15 | 10 |

The key includes rule, source offset and body environment flags. Some certified failures bypass body execution, so recorded failure completions are incomplete. Simulated cache reuse is an opportunity bound, not a timing result. General body costs exclude the outer FinishRule; nested atomic body costs include their FinishRule. [The outcome assessment](../../../results/testrunner/parser-atomic-token-outcomes-profile/assessment.json) preserves source hashes, raw events, matched inputs and these limits.


## Isolated guarded atomic fragments (unshipped)

The frozen general-entry prototype lowers eligible unchanged grammar bodies to a static ASCII class scanner, with an optional negative silent keyword predicate and optional single literal suffix. Ordered literals retain committed first-match behavior; insufficient work/depth budgets fall back to the VM before mutation. Successful scans replay terminating class and optional-suffix failures in the original order. Recognition, semantic builders and default custom programs retain their original paths.

The exact focused gate passed 2,136,220 capture/error/recognition/resource comparisons, also checking semantic-builder callbacks and committed history; 2,352 default observer comparisons, 85 tests and strict Clippy passed. One release round measured 501 rotated paired batches on each of 15 identical real controls, with baseline512, the frozen simple-class prototype, this guarded prototype and Pest in one executable. Thread CPU and wall time were recorded separately; compilation and normalization were excluded.

| Language | Accepted paired CPU change vs baseline512 (three controls) | Rejected change (two controls) |
|---|---|---|
| Ruby | −4.69%, −0.42%, −0.82% | +0.93%, −2.99% |
| Fortran | −3.97%, −5.12%, −5.23% | −5.50%, −5.14% |
| Shell | +0.44%, −0.08%, +0.97% | +0.28%, +0.70% |

This revision remains isolated: nested atomic calls still ignore fragment metadata, and the Ruby/shell results do not establish a broad improvement. CPU/wall tails are noisy; allocation request volume and peak memory were not measured. The archived four-backend summary is authoritative: the legacy human log's inherited `pairedratio` field compares the simple prototype rather than the guarded variant. [Frozen source, proof and raw timing](../../../results/testrunner/parser-atomic-fragment-guard-experiment/assessment.json), [complete four-backend summary](../../../results/testrunner/parser-atomic-fragment-guard-experiment/cpu-proof/summary.json).


A separate nested-entry prototype passed 2,638,860 exact capture/error/recognition/resource and builder comparisons, 3,024 default observer cases, 85 tests and strict Clippy. Its Option-bearing atomic rule-entry task nevertheless regressed all 15 paired CPU controls: accepted Ruby +5.39–9.05%, Fortran +3.92–7.21%, shell +9.43–11.22%; rejected controls +4.92–14.42%. It was rejected and remains unshipped. [Frozen source, exact proof and complete CPU evidence](../../../results/testrunner/parser-atomic-fragment-nested-experiment/assessment.json). The later sparse marked-entry variant passed the same exact gate and physically reduced nested expression dispatches from 54 to 6 in its focused example, but also regressed all 15 controls: accepted Ruby +3.80–6.45%, Fortran +1.35–4.68%, shell +9.42–10.92%; rejected +3.86–14.20%. It was rejected. A separate exact enum-layout probe confirms AtomicTask=40 bytes and Instruction=24 bytes before/after; unchanged size did not establish faster execution. [Sparse marked-entry source, proof and CPU archive](../../../results/testrunner/parser-atomic-fragment-nested-marked-experiment/assessment.json).


Untimed real-control profiling corrects the earlier missing-nested-path hypothesis: all targeted Ruby/Fortran identifiers and shell names use general entry; nested scans reach only 16 Ruby constants. Ruby identifier helpers yielded 231 hits, 110 keyword rejections, 144 head rejections and 30 EOI rejections. Fortran yielded 162/48/40/12 respectively. Each keyword rejection physically re-scanned the selected literal once in the VM. The head check already precedes guard matching and performs zero speculative guard comparisons on head/EOI failures. Shell names had 58 hits and zero fallbacks. These data justify testing generic failing-outcome certificates in the general-only candidate, while keeping the nested driver unchanged. [Frozen profile, raw events and limitations](../../../results/testrunner/parser-atomic-fragment-fallback-profile/assessment.json).


The general-only failing-outcome certificate then passed 3,311,080 exact capture/error/recognition/resource and builder comparisons, 3,360 default observer cases, 25 observed failed VM bypasses retaining rule-entry order, 85 tests and strict Clippy. Two 501-sample paired CPU rounds versus original512 aggregate to accepted Ruby −6.75%, −2.14%, −3.59%; Fortran −5.45%, −7.39%, −6.95%; shell −0.71%, −1.24%, −1.48%. Rejected controls aggregate to Ruby −4.15%, −6.50%; Fortran −6.91%, −7.23%; shell −0.91%, +0.02% (flat). Aggregation weights both run-level paired medians equally. This remains unshipped: incremental validation against the current Task48/sequence production engine is required. [Frozen source, proof, both raw rounds and complete timing summaries](../../../results/testrunner/parser-atomic-fragment-failing-experiment/assessment.json).


### Failing fragment composition with current Task48 + sequence parser

The original512 comparison does not establish the incremental benefit against the current parser. An isolated composition therefore preserves current Task48/sequence code and links the cached `ca064b511f1f6875` runtime and `5a63df64a26e989d` generated parser as the before side. Candidate engine SHA256 is `8cfabe550bf615564a760367cf3fcf9748c38179ec9df1c034471422ae8e0388`; codegen SHA256 is `1c82ff4c400a387c50d31afc2cadc701fd4cb080f1aa7f02e4d4a42896442bfd`.

All 3,311,080 focused capture/error/recognition/resource comparisons, semantic-builder histories and callbacks, 3,360 default custom observer comparisons, core/codegen tests, and strict Clippy pass. Two no-rebuild paired 501-sample CPU rounds give the following median of the two run-level paired median changes:

| Language | Accepted controls 0 / 1 / 2 | Rejected controls 3 / 4 |
| --- | --- | --- |
| Ruby | −3.37% / −1.15% / −1.06% | −1.94% / −4.38% |
| Fortran | −2.71% / −3.23% / −5.31% | −3.96% / −2.98% |
| shell | +3.01% / +2.99% / +3.14% | +4.99% / +4.43% |

This version is rejected and unshipped because the shell regression repeats. The general task loop normalizes a marked fragment entry before every task match; shell has only 58 successful fragment attempts and no scanner fallback on the three accepted controls. A separate isolated entry-helper experiment removes that global check while preserving the ordinary entry task representation. No production or grammar source changes occur here.

Evidence: [current-parser composition archive](../../../results/testrunner/parser-atomic-fragment-failing-e2-experiment/assessment.json), including frozen sources, exact commands, library hashes, both raw CPU/wall runs, per-workload medians/p95 and aggregate ratios. Allocation request volume and peak memory were not measured for this experiment. Exact real-control capture/error/Pest verdict checks pass before each timing workload; these are parser capture construction/destruction measurements, excluding frontend normalization.


The separate entry helper with ordinary `#[inline]` is also rejected. LLVM retained an outlined entry function for all three generated parsers, adding a function call and `Option<Task>` return on each ordinary entry. One paired 501-sample round regressed all 15 controls: Ruby accepted +1.76% / +3.74% / +4.87%, Fortran +5.35% / +2.92% / +1.17%, shell +10.12% / +10.69% / +9.20%, and shell rejected +12.48% / +12.42%. The same 3,311,080 focused comparisons, default observers, tests and Clippy pass. [Source, symbol inspection context and raw timing archive](../../../results/testrunner/parser-atomic-fragment-failing-entry-experiment/assessment.json). A separate forced-inline, constant-specialized entry refinement is being assessed; these results do not support promotion.


Forcing entry-helper inlining and specializing a constant fragment flag removes the outlined symbols but regresses all 15 controls by roughly 22–41% against the same current baseline. It is also rejected, with [source and CPU proof archived](../../../results/testrunner/parser-atomic-fragment-failing-entry-inline-experiment/assessment.json). Exact focused checks and Clippy remain green. A combined entry match arm is a separate experiment; no evidence here supports shipping either entry-helper variant.


The combined-only-entry match arm avoids the global normalization and helper call, but one paired 501-sample round still regresses all 15 controls by roughly 16–35%. Its [frozen source and evidence](../../../results/testrunner/parser-atomic-fragment-failing-entry-arm-experiment/assessment.json) are archived and it is unshipped. The full focused proof, core/codegen tests and Clippy pass. Further work must keep fragment state out of the ordinary dispatch path; changing only how the two entry cases share their body is insufficient.


### Pending fragment pointer, unchanged ordinary task layout

A separate candidate restores the general task declaration byte-for-byte to the e2 baseline. Eligible calls set one pending fragment reference and the immediately following ordinary rule entry consumes it before every early failure/depth exit. Recognition, builders and default custom programs do not enable it. Nested atomic dispatch remains unchanged. Candidate engine SHA256 `9c84c00db12a3cac23a7c4d54d4fc5d53d391d3d7facdea5ae584be18312be6c`; codegen stays `1c82ff4c400a387c50d31afc2cadc701fd4cb080f1aa7f02e4d4a42896442bfd`. Full focused comparisons, core/codegen tests and strict Clippy pass.

Two unchanged executable paired 501-sample CPU rounds against the cached current e2 runtime/generated baseline improve all 15 controls. The median of both run-level paired median changes is:

| Language | Accepted controls 0 / 1 / 2 | Rejected controls 3 / 4 |
| --- | --- | --- |
| Ruby | −7.76% / −5.78% / −5.73% | −6.57% / −9.21% |
| Fortran | −7.69% / −8.15% / −8.98% | −8.03% / −7.97% |
| shell | −1.52% / −2.02% / −1.63% | −1.69% / −1.24% |

The candidate remains isolated pending composition with the newer packed-environment parser and a broad corpus audit. This comparison establishes incremental gains over e2, not the newer packed environment. The [frozen archive](../../../results/testrunner/parser-atomic-fragment-failing-pending-experiment/assessment.json) includes all source/grammar/library hashes, commands, both raw CPU/wall rounds, median/p95 summaries and run-level aggregate. No allocation or peak-memory claim is made. Native capture/error and Pest acceptance/full accepted capture controls pass before each timing workload. Symbols show the compiler inlined the scanner in this executable; only the pending-state layout/dispatch is prescribed by this experiment.


### Packed-environment composition and no-eligible-language screen

The pending-pointer scanner composed with current packed-environment baseline engine `5c4d6bc40cff0ce8a23b1e677c5d0afe8a43e303ba8f56ca3d74453c668828da` passes the original 3,311,080 focused comparisons plus 195,200 additional non-tail-dispatch capture/error/resource comparisons, with exact recognition and semantic-builder observer histories. Core/codegen tests and strict Clippy pass. Candidate engine is `12caeb889430c3229748dd9fb189933ec1d201b27e8f83a0622c5c5577cf1a9c`.

Two paired 501-sample rounds against this exact newer baseline give run-level aggregate changes:

| Language | Accepted controls 0 / 1 / 2 | Rejected controls 3 / 4 |
| --- | --- | --- |
| Ruby | −7.05% / −4.40% / −4.78% | −5.34% / −7.91% |
| Fortran | −13.08% / −13.54% / −14.54% | −13.22% / −13.48% |
| shell | −3.98% / −3.33% / −2.37% | −1.76% / −2.60% |

[Packed composition source/proof/timing archive](../../../results/testrunner/parser-atomic-fragment-failing-packed-experiment/assessment.json). This isolated version is **not ready to ship**: an all-18-language, 90-real-control CPU screen found no-eligible-language overhead. The full generated parser package uses the preserved `capture-task-review/languages` grammar inventory, with no grammar edits. All 90 full native/Pest capture and rejection controls pass.

The first 101-sample screen improves many eligible languages: Go roughly 22–27%, Dart 13–14%, C# 7–9%, Lua 5–8%, Pascal 9–10%, and Ruby/Fortran 7–10%. However, all five JavaScript controls regress 1.07–1.54%. A no-rebuild 501-sample repeat across 60 controls confirms JavaScript accepted regressions 5.17% / 5.21% / 4.76%, rejected 1.67% / 1.59%; PHP reaches +1.01%, and some PowerShell controls reach +1.24%. A first COBOL +2.03% outlier disappears on repeat. Generated JavaScript, PHP and COBOL have no fragment markers or enabling override. These measurements justify further work on ordinary instruction/state overhead before any immutable full-corpus audit or promotion.

[Full 101-screen and 501-repeat evidence](../../../results/testrunner/parser-atomic-fragment-packed-cpu-screen/assessment-repeat.json) preserves exact plans, all six reference libraries/dependency sets, raw CPU rows, parity checks and binary/library hashes. CPU percentages are per-control paired medians; absolute values and tails vary under machine load. No peak-memory or allocation claim is made. A separate opcode-order experiment appends the new instruction variant after all existing variants while leaving scanner/State code unchanged; its result will be reported independently.


Moving only the fragment instruction to the end of the enum fails to remove no-eligible-language overhead: the 101-sample screen gives JavaScript +1.82–3.19% and accepted PHP +2.37–3.41%. It is [archived and unshipped](../../../results/testrunner/parser-atomic-fragment-failing-opcode-end-experiment/assessment.json). A separate constant-specialized driver uses a unit slot when fragments are disabled. Its focused 3,311,080 + 195,200 comparisons pass, along with default observers, tests and Clippy. A cfg-test compares the disabled State against a mirror of the original declaration, asserting equal size, alignment and every original field offset. CPU screening remains pending; structural simplification alone does not establish a speedup.


The constant-specialized driver largely restores no-eligible controls, but its 501 repeat still gives JavaScript +0.24–1.15% and two PHP rejection controls +1.85% / +1.88%. [Frozen unit-slot variant](../../../results/testrunner/parser-atomic-fragment-failing-specialized-experiment/assessment.json) remains unshipped. Its all-18 generated build took 3m20s wall under load; compile CPU was not measured for that run.

The next isolated variant restores the `Instruction` declaration byte-for-byte as well. Enabled capture drivers query a dense expression fragment table only in ordinary Call handling; disabled drivers eliminate that query. Fragment specifications are stored once per rule, with no grammar changes. Focused 3,311,080 + 195,200 comparisons, default observer checks, core/codegen tests and strict Clippy pass. The all-18 generated build records 132.82s wall, 56.25s user CPU and 1.65s system CPU with dependency artifacts cached; these are not a cold workspace build or a controlled comparison to the earlier wall-only builds. Auxiliary BSD time memory statistics were unavailable after a sandbox kernel-query denial, while Cargo completed successfully.

The [opcode-free metadata snapshot](../../../results/testrunner/parser-atomic-fragment-failing-metadata-experiment/assessment.json) is engine `958ea9bc663ede347cdc470bf40dfc85f55eacc4b98cdd200cff68a58bc76de5`, codegen `5ec760ac77f34a6be3d6d8f9945a7fc80a362a34aa48bf60331c2c360b7a81a3`. Its first 90-control, 101-sample paired screen removes systematic JavaScript overhead (−1.27% to +0.37%), retaining Go roughly 17–22%, Dart 10–12%, Lua 4–7%, Ruby 4–9% and Pascal 7% gains. Some other controls show small regressions, including Kotlin accepted +1.11–1.86%. The subsequent 75-control no-rebuild 501 repeat confirms Kotlin all five +1.53–2.36%, while JavaScript is −0.15–0.76%. No full-corpus gate or production promotion has occurred for these scanner variants.


The [general-owner compiler refinement](../../../results/testrunner/parser-atomic-fragment-failing-general-owner-experiment/assessment.json) removes fragment metadata from call expressions owned exclusively by declared Atomic rules whose entire body uses the compact atomic driver. Ownership is conservative: shared expressions union their owners; Normal, Silent, CompoundAtomic, uncertified Atomic and stackful bodies retain the ordinary eligibility. There are no grammar changes. Instruction/Task declarations remain unchanged. Runtime files are exactly the Rustfmt output of the preceding metadata variant; only codegen eligibility changes.

[Actual Kotlin profile](../../../results/testrunner/parser-atomic-fragment-general-eligibility-profile/profile.log) shows 38,029 general Call lookups across five real controls and zero usable fragment results: the eligible digit rule is called only inside certified atomic number bodies. The refinement disables that unreachable scanner path; the corresponding profile makes zero metadata lookups. Focused 3,311,080 capture/error/resource/builder comparisons, 195,200 non-tail comparisons, 3,360 default custom observer checks, core/codegen tests and strict Clippy pass. New compiler tests retain Normal, CompoundAtomic and stackful callers while eliminating nested-only certified Atomic callers. Its cached all-18 generated build records 77.51s wall and 49.52s user CPU, with no controlled cold-build comparison implied.

The 90-control/101-round screen and 75-control/501-round repeat retain substantial accepted gains on Ruby, Fortran, Go, Dart, Lua and Pascal; Kotlin accepted controls become roughly 0.3–0.4% lower CPU rather than 1.5–2.4% higher. JavaScript remains near flat. However, independent per-backend calibration exposed a benchmark confound: PHP rejection control 4 reports roughly +11% with baseline batch size 2/candidate size 1, and roughly −7% when those sizes reverse. Both logs are retained. These conflicting results do not establish a parser regression or gain. A new paired 501-round screen uses an identical calibrated batch count for all three backends on each control, with unchanged frozen libraries. The scanner remains isolated and unshipped pending that evidence and the immutable all-language corpus gate.

The [identical-batch 90-control/501-round screen](../../../results/testrunner/parser-atomic-fragment-general-owner-equal-batch-cpu-screen/assessment.json) resolves the large rejection outliers and retains accepted gains: Ruby 3.5–6.9%, Fortran 7.3–9.5%, Go 19.5–21.3%, Dart 10.6–11.4%, Lua 6.2–6.8% and Pascal 6.6–7.1% lower CPU. Kotlin accepted controls are 0.3–0.6% lower, and COBOL is approximately flat. JavaScript is 0.4–0.9% higher in this run, while PowerShell is 2.3–4.6% higher across all five controls. The latter prevents promotion of this variant. PowerShell's readonly profile finds 17,858 general metadata lookups but only 74 eligible results across five controls; it supports replacing the expression-sized pointer table with a small generated match on the target rule ID already carried by ordinary Call instructions. That next experiment is isolated; these results remain unchanged and unshipped.


The [sparse target-rule lookup](../../../results/testrunner/parser-atomic-fragment-failing-rule-match-experiment/assessment.json) passes the same focused correctness suite but does not resolve PowerShell's overhead: identical-batch 90-control/501-round CPU results remain +2.7–4.8% there, while other eligible gains persist. It stays unshipped. The following isolated body-entry design removes the pending pointer and all Call-side lookup/set operations completely. It selects a scanner inside the existing certified Atomic body entry branch, before compact atomic execution. Both const-specialized State variants now have the original size/alignment and all 17 field offsets. Instruction and Task declarations remain unchanged. Its focused 3,311,080 + 195,200 + 3,360 comparisons, tests and strict Clippy pass; broad equal-batch CPU screening is pending.


The [body-entry screen](../../../results/testrunner/parser-atomic-fragment-body-entry-equal-batch-cpu-screen/assessment.json) completes all 90 controls with 501 rotated paired CPU samples and identical per-control batch sizes. Accepted gains are Ruby 3.2–7.7%, Fortran 6.9–9.2%, Go 19.2–19.8%, Dart 11.4–12.1%, Lua 6.1–6.3%, and Java 6.6–7.3%. JavaScript improves 0.6–1.0% on accepted controls in this run. PowerShell's remaining +0.4–1.6% and one PHP rejection +2.4% remain material tradeoffs; no uniform improvement is claimed. This is the best isolated entry architecture, [preserved with source/proof/generated hashes](../../../results/testrunner/parser-atomic-fragment-failing-body-entry-experiment/assessment.json), without production promotion or a new full-corpus audit. Its cached package build uses 112.44s wall/56.49s user CPU. Guarded identifier fragments already contain first-byte literal buckets; no further bucket experiment is needed.


An [untimed structural/outcome inventory](../../../results/testrunner/parser-atomic-literal-boundary-profile/report.json) identifies 558 call-free Atomic literal-plus-negative-ASCII-boundary shapes across the frozen grammars. Most rejection attempts already use first-byte failure certificates, so rule frequency alone overstates the remaining opportunity. For example, audited Fortran controls have 1,405 relevant calls but only 2,179 physical body dispatches; Ruby has 1,929 relevant calls and only 134 dispatches. These counts are from the earlier e2 profile, not a current CPU estimate.

A separate [literal-boundary prototype](../../../results/testrunner/parser-atomic-fragment-literal-boundary-experiment/assessment.json) extends the body-entry matcher without grammar, Task, Instruction or State layout changes. It handles successful keyword tokens, continuation-boundary rejection and literal mismatch, with exact ordered literal expectations, folded ASCII handling, grouped expression fuel and tight-budget fallback. Focused release proof passes 6,048,760 exact capture/error/resource/recognition comparisons, 331,840 non-tail builder/callback comparisons and 5,712 default opt-out observer comparisons. Examples reduce expression dispatches from 7/6/2/18/7 to zero for ordinary success, boundary rejection, mismatch, grouped success and case-folded success respectively. Core/codegen tests and strict Clippy pass. The next CPU screen compares this addition against frozen body-entry, rather than attributing cumulative gains to the new shape. It remains unshipped.


The [incremental literal-boundary screen](../../../results/testrunner/parser-atomic-fragment-literal-boundary-incremental-cpu-screen/assessment.json) finishes all 90 controls with 501 paired samples, identical batch counts and full capture/error/recognition/work parity. Relative to frozen body-entry, accepted CPU falls 3.7–4.5% on Fortran, 3.1–3.7% on Java, 2.2–3.8% on Lua, 0.6–1.5% on JavaScript and 0.6–1.9% on shell. Some PHP/COBOL controls are slightly slower: PHP accepted +0.5–1.1%, one rejection +1.3%; COBOL accepted −0.8% to +0.5%, one rejection +1.2%. These are incremental results, not the cumulative gain against production or a uniform speedup. Generated Rust grows from 6,638,726 to 7,384,330 bytes across the all-18 package; the cached release build takes 4m13s wall, with CPU unmeasured and no controlled compile-time comparison. The best matcher remains isolated. Composition with the latest bounded32 allocation baseline starts as a separate candidate/proof, preserving both frozen measured inputs.


The [bounded32 composition](../../../results/testrunner/parser-atomic-fragment-literal-boundary-bounded32-experiment/assessment.json) passes the expanded focused proof, core/codegen tests and strict Clippy against a fresh snapshot of current production. Its [90-control/501-round identical-batch screen](../../../results/testrunner/parser-atomic-fragment-literal-boundary-bounded32-cpu-screen/assessment.json) measures cumulative direct matcher gains against production **including** bounded allocation: accepted Go 20–23%, Dart 13–14%, Fortran 10–12%, Java about 8%, Lua 7–8%, C# 5–6%, Ruby 4–7% and Pascal 7–8% lower CPU. Some controls regress modestly. A [20-control no-rebuild repeat](../../../results/testrunner/parser-atomic-fragment-literal-boundary-bounded32-cpu-screen/no-rebuild-regression-repeat/assessment.json) retains accepted C +1.5–2.2%, Kotlin +0.3–0.9%, JavaScript +0.8–1.3% and COBOL +0.2–0.7%. Both runs remain archived; no claim is made that every language improves.

[Readonly lookup profiling and ARM64 inspection](../../../results/testrunner/parser-atomic-fragment-literal-boundary-bounded32-cpu-screen/no-hit-profile/abi-review.json) preserve captures/errors and recognition/work results. C makes 33–584 sparse fragment queries per control with zero usable metadata; COBOL makes zero queries. Kotlin and JavaScript exercise keyword fragments. The C getter returns an optional borrowed pointer in `x0`, without a hidden aggregate result, and its caller inlines the lookup. The scanner is also inlined, enlarging the hot driver; external static metadata may limit constant propagation. This is a possible cause for tradeoffs, not a measured causal attribution. The candidate remains frozen and unshipped pending the full immutable corpus gate. No grammar changed.


A fused atomic-body classification experiment replaces two eligibility queries with one generated lookup. Exact focused proof passes, but its three-variant return makes LLVM stop inlining the VM executor on eligible grammars; the first broad screen regresses roughly 14–35%. ARM64 inspection shows register returns, rather than a hidden aggregate result. Forcing executor inlining removes that large regression but still leaves mostly 1–5% overhead. A separate annotation-only control on production trivia/bitmap also mostly regresses 1–3%, so this is not a production-wide inlining improvement. Both experiments remain [rejected and archived](../../../results/testrunner/parser-atomic-fragment-body-classification-inline-experiment/assessment.json), with the [independent annotation control](../../../results/testrunner/parser-production-execute-inline-experiment/assessment.json) kept separate.

The [scalar-pair classification refinement](../../../results/testrunner/parser-atomic-fragment-body-pair-experiment/assessment.json) preserves the expanded exact proof and avoids the large regression, but its 90-control/101-round screen is mixed: most changes are small, while Lua accepted CPU rises 0.2–2.6% and rejected CPU about 4%. It remains unshipped; the proven boundary matcher is retained without classification changes.

The [boundary matcher composed with production trivia/bitmap](../../../results/testrunner/parser-atomic-fragment-literal-boundary-trivia-experiment/assessment.json) passes 6,048,760 exact main comparisons, 331,840 non-tail resource/builder comparisons and 5,712 default opt-out observer comparisons, plus core/codegen tests and strict Clippy. Its [90-control/101-round identical-batch CPU screen](../../../results/testrunner/parser-boundary-trivia-cumulative-cpu-screen/assessment.json) measures against frozen production engine `41dfd88f`, explicitly before the later compact trivia refinement: accepted Go 21–24%, Dart 11–14%, Fortran 11–14%, Pascal 10–11%, Lua 9–10%, Ruby 6–10%, Java 7–8% and C# 5–7% lower CPU. JavaScript is 0.7–1.2% lower and C 2.2–2.5% lower in this run. PHP has two accepted controls 0.4–0.8% higher; no universal gain is claimed. Source, generated artifacts, frozen baselines and proof are archived. The matcher remains isolated pending composition onto the latest production and the full immutable corpus gate; grammars and production source are unchanged by these experiments.


The [latest composition](../../../results/testrunner/parser-atomic-fragment-literal-boundary-compact-prefix-experiment/assessment.json) keeps the proven matcher and composes it onto the root's frozen **unpromoted compact-task/borrowed-expectation-prefix baseline**, preserving generated compact dispatch and default custom wide dispatch. Focused release proof again passes 6,048,760 main comparisons, 331,840 non-tail resource/builder histories and 5,712 default observer comparisons; core/codegen tests and strict Clippy pass. Enabled/disabled fragment specializations preserve wide and compact State sizes and alignment. Task and Instruction declarations and grammar bytes remain unchanged.

Two [90-control/501-round identical-batch CPU runs](../../../results/testrunner/parser-boundary-compact-prefix-cumulative-cpu-screen/assessment.json), including a [full no-rebuild repeat](../../../results/testrunner/parser-boundary-compact-prefix-cumulative-cpu-screen/all90-no-rebuild-repeat/assessment.json), retain incremental accepted gains against that exact frozen baseline: Go about 22–24%, Dart 14–15%, Fortran 12–14%, Lua 10–12%, Pascal 7–8%, Java 6–7%, Ruby 4–8%, C# 5–6% and shell 4–5% lower CPU. JavaScript stays effectively flat. Tradeoffs remain: VB roughly 0.5–0.8% higher, WAST up to 1.3% higher, Python rejection roughly 2–3% higher, and C moves from slightly lower to 0.2–0.7% higher on repeat. Python accepted controls improve modestly. A separate 20-control no-rebuild repeat also preserves these small tradeoffs. All runs compare complete native capture/error and recognition/work results outside timing, with accepted Pest capture parity. This does not establish universal speedup or production Pest parity. The immutable all-language gate is prepared from the frozen source closure; no production promotion has occurred.


The root subsequently integrated that exact compact/borrowed-prefix baseline after its immutable 145,608-record and 1,920-resource audit passed. The frozen boundary matcher's own immutable gate then started, without rebuilding or altering its qualification snapshot.

A separate [scalar-byte eligibility refinement](../../../results/testrunner/parser-atomic-fragment-body-class-byte-experiment/assessment.json) preserves the expanded exact proof, tests and strict Clippy but regresses eligible languages about 20–40% in a matched 90-control/101-round incremental screen. PHP/COBOL/WAST, which do not enable fragments, stay approximately flat. [ARM64 evidence](../../../results/testrunner/parser-body-class-byte-incremental-cpu-screen/arm64-review/assessment.json) shows a tiny byte-return getter but newly outlined capture VM execution with a 0x130 stack frame, plus an outlined scanner; those standalone functions are absent from the proven matcher's executable. Changing the return representation alone does not prevent the compiler's inlining decision from changing. This variant is rejected and unshipped; the qualified matcher and its corpus gate remain unchanged.


The frozen boundary scanner subsequently passes its immutable 145,608-record plus 1,920-resource audit and is integrated by the root. A separate [borrowed existing-eight-alternative terminal descriptor experiment](../../../results/testrunner/parser-terminal-choice-borrowed-experiment/assessment.json) confirms surviving shell getters return 24-byte aggregates through ARM64's hidden x8 slot. The optional borrowed hook removes that copy, while retaining original owned tables/getters and default custom query behavior. Expanded main proof, tests and Clippy pass; an independent proof adds 493,506 resource/builder cases, 496,392 complete default physical traces and 9,475,369 identical owned queries with zero default borrowed calls. Seven real descriptors match owned/CompiledGrammar contents across valid IDs; generated getters both return None for invalid IDs. CompiledGrammar's existing invalid-index panic is preserved, not changed.

The 90-control/101-sample screen and same-binary 501 repeat show mostly tiny mixed changes: shell accepted repeat −0.40%, −0.56%, +0.19%; JavaScript approximately flat to slightly lower. Some PHP controls remain slightly higher. Removing the ABI copy does not establish a material broad runtime gain, so this standalone variant remains unshipped. The integrated direct scanner and its proof are unchanged.

### Scalar tail fuel-bound experiment (unshipped)

Against the integrated `f69` scanner and `b614` compiler, a generated maximum tail-class weight permits one checked bound over the whole remaining input. When it proves enough fuel, the existing scalar tail scan omits its per-byte overflow and budget checks; otherwise it retains the old loop. No grammar, task, state, or frontend behavior changes. Exact comparisons passed 6,386,312 cases, 97 tests, strict Clippy, and an all18 emitted-property check covering 19 class fragments.

The initial 90-control, 101-pair common-batch screen passed all capture/error/recognition-work checks. 69 controls had lower paired thread CPU, with a median change of −0.506% across controls; paired wall time changed −0.541%. Accepted JavaScript controls changed −0.07%, −1.20%, and −0.08%; PHP accepted controls changed +0.32%, +0.44%, and +0.46%. These small mixed differences warrant repeat qualification before promotion. Ambient one-minute load ranged from 31.7 to 80.2 on 10 logical CPUs; thread CPU removes descheduling but does not isolate contention, cache, or thermal effects. See [compact assessment](../../../results/testrunner/parser-atomic-fragment-tail-fuel-bound-cpu-screen/summary.json) and [source/proof assessment](../../../results/testrunner/parser-atomic-fragment-tail-fuel-bound-experiment/assessment.json). Temporary driver executables were removed and raw timings compressed under the parser-results retention budget.

The focused 501-pair repeat of Lua, Ruby, Fortran, and PHP did not confirm the initial broad gain: 10/20 controls improved, median paired CPU changed −0.027%, and paired wall time −0.085%. Lua accepted controls changed −0.35%, −0.40%, +0.45%; Ruby −2.19%, +0.21%, +0.24%; Fortran +0.25%, −0.93%, +0.69%; PHP +0.29%, −0.25%, +0.24%. The earlier PHP rejection regression decreased from +1.40% to +0.45%, and the other rejection changed −2.08%, so the small regressions cannot be confidently attributed to runtime changes rather than variation. Ambient one-minute load ranged 8.24–30.15. This experiment remains unshipped; the conditional concurrent-throughput test was omitted because stable broad benefit was not established. All repeat executables were removed after [compact assessment](../../../results/testrunner/parser-atomic-fragment-tail-fuel-bound-cpu-screen/focused501-repeat/summary.json), with hashes retained and timings compressed.

### Atomic repeat SIMD experiments (unshipped)

An [actual bulk-repeat profile](../../../results/testrunner/parser-active-builds/atomic-repeat-chunk-profile/report.json) instruments the existing `AtomicTask::Repeat` class scan, including predicate, skipping, discarded, and rejected paths. All 90 controls preserve exact captures/errors and recognition work. Runs of at least 32 bytes account for 92% of scanned PHP/Python bytes, 74% of JavaScript bytes, and 72% of Fortran bytes; Ruby has no such runs and 10,032 zero-length scan attempts. Long runs mostly come from comments and strings. These are scanned-byte shares, not CPU shares or expected speedups.

Three isolated NEON variants preserve the existing class membership/count result and work accounting. Each passes 12,544,252 kernel comparisons over observed classes, arbitrary masks, all byte values, and vector boundaries. The first also passes the 6,386,312-case resource/builder/default-observer proof. Its initial timing compared a precompiled baseline grammar with a client-compiled candidate grammar; those runs are retained as unqualified historical evidence. Corrected timing compiles both native grammar modules inside the same executable, with identical input, batch counts, rotated backend order, thread CPU and wall timers.

The matched first variant improves only 39/90 controls, with median paired CPU +0.137%; PHP accepted rises about 2.1% and Ruby 1.5–1.7%. A second, shorter initial probe still raises PHP 0.6–1.0% and Ruby 3.7–4.8%. The final variant keeps scalar membership checks at entry and invokes NEON only after 16 actually matching bytes, with at least 32 bytes remaining. Its 25-control/101-pair screen improves accepted Python 1.6–2.1% and Fortran 0.8–1.8%, but raises PHP 1.3–1.9% and Ruby 1.8–2.2%; JavaScript changes −0.62% to +0.26%. Ambient one-minute load is 9.15–10.53 on 10 logical CPUs. No broad speedup is established. All three remain [rejected and archived](../../../results/testrunner/parser-atomic-repeat-byteclass-neon-experiment/assessment.json), with [final matched results](../../../results/testrunner/parser-atomic-repeat-byteclass-neon-v3-cpu-screen/summary.json). Full resource proof was not repeated for the two later rejected variants, and no immutable corpus gate or promotion followed. Temporary private libraries/executables totaling 16.2 MB were removed after recording their hashes; sources, compact assessments, proof logs, and compressed timings remain.

### Terminal-choice ASCII dispatch experiments (unshipped)

Own ARM64 assembly confirms that the terminal-choice ASCII builtin matcher loads a function pointer and branches to it. A direct inline builtin match removes that call but retains a local class jump table. A second variant replaces class dispatch with a 336-byte immutable bit table and supported-builtin mask; its success path has no indirect branch or call. Both affect only this terminal-choice callsite and preserve the other ASCII predicate users, grammar, Program, Task, and State declarations.

Each variant passes all 21 builtin values across all 256 bytes, including unsupported `None` results, eight direct unit tests, strict Clippy, and exact captures/errors/recognition work for 20 frozen PHP/Python/Ruby/Fortran controls. Each uses 101 rotated, identical-count CPU/wall pairs with both native generated modules compiled inside the same executable. The direct match's median paired CPU changes −0.244%, with accepted Fortran +0.3–0.7%. The table's median changes +0.478%: accepted PHP +0.49–1.19%, Ruby −0.89% to +2.28%, Python −1.08% to +0.81%, and Fortran −0.37% to +0.97%. Table ambient one-minute load ranges 23.26–38.14; these measurements reflect the busy-server setting and do not establish a consistent gain. Both remain [rejected and archived](../../../results/testrunner/parser-terminal-choice-ascii-experiment/assessment.json); no 501 repeat or full integration gate followed. Owned compiled outputs totaling 11.8 MB were removed, while sources, assembly, proof logs, hashes, and compressed timings remain.


The isolated compact-cache eligibility bitmap screen retained the exact selected rules and runtime. A trimmed constant-u64 getter passed all-valid and invalid-ID properties, 900 capture/resource plus 900 recognition/work checks, and strict codegen Clippy. In matched 20 controls over 101 rounds CPU timing it reduced the median paired CPU time by2.13%; all 20 controls were lower in that run. The independent baseline was corrected-hash H, not production; the main cache composition owns broader qualification. Source and attribution are preserved in `results/testrunner/parser-failed-rule-cache-eligibility-bitmap-trim-experiment`.

A separate sparse lazy-depth refinement experiment from compact-bookkeeping P preserved normal Entry32/Frame40 storage and bounded its separate certificates to 128. It passed 17 unit tests, strict Clippy,20,440real capture/resource and matching recognition/work checks, 1,170,288synthetic virtual-depth/fullbuilder resource cases, and 12,168default-optout physical observer cases. Despite nonvacuous certified cache unlocks, its20 controls over 101 rounds screen was flat overall (median paired CPU−0.01%;10/20lower), with Ruby median+0.68%. One Ruby rejection improved10.72%, but other Ruby controls slowed0.68–1.81%. It remains unshipped; no501repeat or full-corpus audit was requested for this mixed result. Exact sources, compressed timing and assessment are in `results/testrunner/parser-failed-rule-cache-sparse-depth-experiment`.


The separately frozen unconditional global high-water alternative passed the same resource/observer proof plus 18 unit tests. Its active ARM update was four instructions versus nine for a nonempty sparse frame, but it writes even before any refinement. In the same 20 controls over 101 rounds screen versus P it regressed median paired CPU by3.97% (Ruby+4.02%,Fortran+6.35%,Python+3.92%); PHP improved3.66%. This variant is unshipped. Source/proof/compressed timing are preserved in `parser-failed-rule-cache-global-depth-experiment`; a distinct lazy-activation plus min-of-valid-certificates followup is being evaluated.


The combined lazy-activation/min-certificate variant passed 19 unit tests, strict Clippy and the same complete focused resource/observer matrices. It also remains unshipped: matched 101-round CPU medians were Ruby+5.05%,Fortran+6.82%,Python+3.97%, while PHP improved3.49%. This measured combined mechanism is not evidence for min updates alone. All three completed depth-variant runtime/proof executables and libraries were removed with size/SHA manifests; their exact source, proof logs and compressed timing remain. The separate cost-derived depth-bound audit may eliminate tracking overhead entirely.


A compiler-only cache-eligibility worklist replaces per-rule HashSets with one dense stamp buffer and reusable DFS stack, and repeated impurity scans with reverse edges. All 18 frozen grammars emit byte-identical Rust under default and experimental options. Both 31-pair host-opt0 runs improve every language; the final median CPU reductions are18.83% for emission and13.52% for grammar compile plus emission (JS10.62%,Ruby10.43%,Fortran10.45%,PHP21.77%,Python15.82%). Both variants link the same opt3 grammar parser/resolve dependency, so this is not an all-dependencies-opt0 build measurement. Four unit tests,72 reference comparisons, formatting and strict normal/test Clippy pass. This change affects generation time; parser runtime and generated Rust remain unchanged. Exact sources and evidence are in `results/testrunner/parser-failed-rule-cache-eligibility-worklist-experiment`; production integration belongs to the parent task.

### Bloom cache admission (2026-10-05, rejected)

An isolated 128-byte Bloom admission filter skipped failure-cache frames on first encounters. Exact completed keys were still required for replay; collisions could only add ordinary evaluation/bookkeeping. Fifteen units, strict Clippy, both task-layout proofs (206,712 exact parser/resource/builder cases and 32,448 default observer cases total) passed.

The 33-control ×101 paired thread-CPU screen improved only one control. Against the completion-tag cache, language median changes were Ruby +7.42%, PHP +6.06%, Python +5.23%, and Fortran +0.22%. Against production 18e, previous regression controls remained slower: COBOL +2.13%, PowerShell +3.29%, WAST +4.23%. This policy is rejected and unshipped; no production or grammar edits occurred. Exact sources, common batch counts, separate CPU/wall timings, ambient load, commands and compressed records are retained in `results/testrunner/parser-failed-rule-cache-bloom-admission-experiment/summary.md` and `screen-assessment.json`. Completed private compiled outputs were removed after assessment.

### Actual production generation comparison (2026-10-05)

The optimized eligibility worklist remains faster than its private cache compiler parent175f. A separate comparison against actual production c492 measures the net feature-generation cost: 18 frozen grammars ×31 paired alternating rounds per default/experimental emit and compile+emit mode, both opt0 emitter libraries sharing the same supplied opt3 CGU1 core. Default median CPU increased4.56% for emit and2.97% for compile+emit; experimental medians increased4.61% and3.34%. These results do not support a production compile-speed improvement claim.

All18 outputs, for both options, were identical after removing only the added failed-rule eligibility getter and mapping cached capture entry names to production names. Unnormalized output differs as expected; runtime cache correctness remains a separate qualification. Exact source/artifact provenance, separate CPU/wall timing, load, commands and compressed records are in `results/testrunner/parser-failed-rule-cache-worklist-production-comparison/{summary.md,assessment.json,measurements.tsv.gz}`. Production source was untouched and completed private compiled outputs were removed.
