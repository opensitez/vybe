# C Frontend Optimization Gate

## Current Integration Status (2026-10-16)

The fresh root-debug Cargo gate passes **345 tests, zero failures, 19 ignored,
zero filtered**:
`cargo test --offline -p vybe_language_c --lib --jobs 1 -- --test-threads=4`.
Executable: `target/debug/deps/vybe_language_c-4acf8386f0c944fe`.
The unchanged C direct-call tests now pass: a real completed source edge reduces
the fixture caller from 51 to 34 bytecode bytes, while mutable/shadowed callbacks
remain indirect. Three public Go/PHP shared-emitter gates also pass, including
metadata-free validated WASM reload/execution and a wrong-result trap control.

The complete saved-Doom emission gate passes on that same root cohort:
5,143 units, 2,329 chunks, **24,568,366 bytecode bytes and 4,718 actual source
direct-call sites**. The 71,966,332-byte metadata-free standard-WASM artifact
independently validates and has zero dynamic-equality calls/function references.
Artifact: `/private/tmp/nyx-doom-direct-calls-20261016.wasm`, SHA256
`4ca51e84d3bb81703ab3a836addbf81fcb4bdd4e11344881ec308548319d4bbc`.
Exact commands, identities and boundaries:
`/private/tmp/nyx-c-direct-calls-root-handoff-20261016.md`.

Native committed units and source-Pratt are the production default. Direct
initializer, body/control, primary, syntax/binary, scalar declaration/parameter,
function, call, index/member and value-return constructors reach the ordered C
consumer. The Doom gate exercises 1,060 value returns, 1,402 functions, 5,682
calls, 1,587 indices, 1,058 members and 5,911 source-Pratt handoffs. Ordinary and
switch-case returns share one consumer; complete contextual AST/span and
validated WASM/noEq/VM controls pass. Context-dependent forms retain captures
and typed lowering; this is not full capture-free construction or streaming.

The shared publication inventory has 35 exact-source safety gates. Every current
expression variant is explicitly audited, including hidden object-method,
default, bound and case-condition slots. Unknown containers and binding-policy
changes still reject optimization. Existing completed-entry/capture/arity/
receiver/alias guards remain in force; no language-name exception was added.
Computed pointer-member addresses preserve stride, backing and single evaluation.
Declared W64 reads no longer produce the nine dynamic-equality references
previously reproduced by the strict width fixture; float reads stay floating.

Instrumented full-Doom frontend time was 323.60s, with 63.76s inclusive matching;
shared emission took 4.88s, writer 33.60s, validation/audit 7.24s. These are
diagnostic observations under shared load, not a controlled speedup. Most of the
frontend total remains outside matching. The private generated compact/
stack-free unit-offset candidate passes eight focused gates and saved-source
capture parity, but is not activated or included in these root results.

CLI `vx` was not rebuilt in this cohort. Fresh gameplay/FPS, standalone Doom
linkage/execution, the reported room freeze and full parser/emitter overlap
remain unproven. A reproduced nested-label bug in `F_CastTicker` has a private
six-test CFG candidate, not a deployed fix. The owned typed-unit preparation
transaction retains original storage and its semantic barriers; its 28
exact-source gates do not prove parser-to-emitter overlap.

### Earlier Cohorts

The latest source-linked debug cohort, `/private/tmp/nyx-c-switch-unit-pipeline-tests`,
passes **309 tests, zero failures, 19 ignored**. SHA-256:
`501f509d3abbc3f28e3d77654ea2076222a26022d39f16c34ae439a01ce8c2ce`.
Log: `/private/tmp/nyx-c-switch-pipeline-20261016.log`.
Direct `if`, `while`, `do-while`, `for` and `switch` statement shells now reach
the ordered C consumer. Loop condition/body
evaluation order, scope effects and constant-loop rewrites retain the reference
policy; eight complete PEG/reference/Native loop AST cases pass. The producer
tests cover anchor validation, rollback, reentrancy and committed-unit lifetime.
The small execution gate now explicitly drains `take_owned_units`/`next_unit`,
checks source indices and ownership exhaustion, compares ordinary/owned standard
WASM bytes, validates them, audits dynamic equality and executes both chunks.
The final fresh link includes loop execution for condition effects, continue and
the do-while-zero rewrite, plus switch fallthrough/default execution. Eight
switch AST cases preserve signed/unsigned/wide promotion, nested labels/scopes
and braceless forms. The producer does not collect or convert case values.

The shared reference and atomic preparation accumulators now separate fact
collection/sealing from ordered typed-unit consumption. Ten combined source-linked
split-boundary parity gates pass against the original traversal, including
closures/defaults, scope redirects, folding, late atomic invalidation and
source-drop isolation. Log: `/private/tmp/nyx-staged-preparation-source-20261016.log`. This
shared-source increment is not included in the C cohort's older linked SDK;
full compiler/C integration remains a separate gate.

Actual CLI export also validates as standard WASM and passes the independent
dynamic-equality audit: `/private/tmp/nyx-doom-pipeline-cli.wasm`, 70,667,705 bytes,
SHA-256 `8c296f41af193fdf3b1f967f1c6f2555cd4b38a2222570107595d348adc3ca1e`.
The reader loads 2,333 chunks, but execution fails to link `env.__vybe_autoderef`.
Compiler initializer ownership has nine source-linked gates; the writer import
fix passes four root-debug Cargo integration gates, including validated
metadata-free callable reload/execution and a wrong-result trap control. Explicit
host imports retain binding/mutability; local function initializers do not invent
host imports. This old artifact is not gameplay proof and predates loop integration;
the combined CLI rebuild/re-export/reload remains pending.

Parser-to-emitter streaming remains incomplete: C needs a prepared provider and
its translation-unit semantic facts, while shared preparation still analyzes
whole bodies and emits hoisted functions before unit delivery. Staged semantic
fact accumulation is the next shared boundary, not a claim of existing overlap.

The integrated production default is **Native committed-unit matching**, not PEG.
`frontend::selected_backend` selects Native when `VYBE_C_FRONTEND` is unset or
`native`; explicit `peg` and unknown selections retain PEG. `owned`, `mixed`,
and contextual `pratt` remain opt-in. Native and PEG retain the same non-worker
lowering policy; worker lowering is enabled only by the owned/mixed opt-ins.
The latest source-linked debug library cohort passes **290 tests, zero failures,
19 ignored**: `/private/tmp/nyx-c-pipeline-verified-tests`, SHA-256
`cf40b47d84f804e267d503154d719ee1aff4dcdf185ecd347d581e8c4c3e8fda`.
Log: `/private/tmp/nyx-c-pipeline-params-library-20261016.log`.
This includes Native choice-head guards, direct binary consumers, exact linear
capture-shape checks and named scalar parameter handoffs. Parameters reuse the
reference constructor and ordered type/access policy; aliases requiring
pointer/array/callback handling, explicit complex declarators, unnamed parameters
and runtime-bound arrays retain their tested fallback paths.
Native matching now hands supported positional literal initializer ASTs directly
to both initializer consumers. Certified complete regions omit descendant
captures; contextual/designated/effectful forms retain fallback. The explicit
Native path and the explicit source-Pratt lane consume matcher-issued operand handles and reduction receipts,
instead of rediscovering their precedence from capture children. Contextual
typing and unsupported nodes still use the C walker. This is not a complete
capture-free frontend or streaming compilation. Native now matches each
`frontend_unit` at an absolute source offset and merges direct initializer and
Pratt receipts across units. The C walker consumes these units after its existing
global predeclaration barrier. No speculative unit reaches emission.

The independent grammar gate passes two focused tests plus complete saved-Doom
capture/span parity: 1,808,719 source bytes, 4,370 committed units and 1,399,744
non-wrapper captures identical to the whole-program parse. The slow gate is
`streaming_unit_tests::committed_units_preserve_complete_saved_translation_unit`
with `VYBE_C_FRONTEND_PREPROCESSED`. This proves syntax boundary preservation,
not typed AST or emitted WASM parity. The integrated `cargo check --offline
-p vybe_language_c --tests --jobs 1` now passes after fixing the two source-pass
test adapters' parser re-exports. The fresh linked execution gate passes.

Direct block-child sequences and scope-independent control statements now feed
the C consumer as well as initializers. `walk_block` retains scope entry/exit;
owned and contextual statements share the existing source-span stamping and
debug-timing path. Typed returns and declarations remain contextual. The new
`direct_body_and_control_statements_preserve_contextual_scopes_and_spans` gate
compares fourteen PEG/Native complete modules against the old statement walk
and proves actual body/statement handoffs.

Primary literals/unresolved names, simple numeric prefixes and comma sequences
also reach the ordered C consumer as grammar-built AST. Name normalization is
shared with capture lowering; numeric prefixes use the existing `apply_prefix`
path. Contextual operators, calls, lvalues and typed returns retain fallback.
The Body/Primary/Syntax consumer oracles pass 56 complete PEG/Native comparison
pairs, including scope and source-span policy.

The source-linked debug executable
`/private/tmp/nyx-c-pipeline-declarations-tests` reports **271 passed, zero failed,
18 ignored**. Native scalar declaration ASTs now reach the ordered consumer;
explicit pointers, arrays, function declarators and contextual aliases preserve
capture fallback. User-defined and lexically bound calls retain their binding
instead of being rewritten to libc adapters. Log:
`/private/tmp/nyx-c-pipeline-declarations-library-20261016.log`.

The subsequent emission gate in `/private/tmp/nyx-c-pipeline-emission-tests`
compares ordinary compilation with `begin_emission_owned`/`emit_next` for all
unit-boundary and call-binding fixtures. Metadata-free standard WASM bytes are
identical, independently validate, and contain no dynamic equality calls; the
fixtures execute successfully. This exercises the shared incremental emitter
after complete AST preparation, not parser-to-emitter streaming.

The full-module gate is
`frontend::tests::committed_units_emit_complete_valid_standard_wasm_without_dynamic_equality`.
Set `VYBE_C_FRONTEND_PREPROCESSED` to a complete saved multi-unit source and run
the ignored test explicitly. `VYBE_C_FRONTEND_WASM_ARTIFACT` optionally retains
its metadata-free WASM. It requires actual committed-unit/Pratt matching rather
than whole-program recapture, emits through the owned incremental shared path,
validates the complete module and checks bytecode plus WASM equality call sites.
Reported phase durations include test instrumentation and are not frontend
performance acceptance. It does not execute the game or certify the CLI bridge.
The BinaryRegion cohort `/private/tmp/nyx-c-pipeline-emission-tests`, SHA-256
`b9b0198b9fcc6d14e69101e73adc57529a9155e79a365a1df11f38dd40f7e708`,
passes this full gate on the saved Doom source (1,808,719 bytes, SHA-256
`a3cb210bde65a45bb83426dfaff8f983bb73b4c97a6d3a81aa994df86e6f271a`).
It emits 5,130 source units / 2,329 chunks / 75,688,109 WASM bytes, validates the
complete artifact and passes the equality-call audit. Artifact:
`/private/tmp/nyx-doom-pipeline-20261016.wasm`, SHA-256
`622b275601db1c9673d03c8c156da61e0274ded9ef83bd5a75b3e234aaac5401`.
Log: `/private/tmp/nyx-doom-pipeline-wasm-20261016.log`.

Direct binary-region consumption is now wired to grammar-issued Pratt
reductions. Each original unary operand is normalized in source order under the
current C scope, followed by the same integral-value policy; the original
binary-region rewrite runs once. Nested contextual regions and effectful
operands retain capture lowering. The new consumer oracle, all-operator-pair
gate pass in that cohort, including 5,744 complete-Doom direct binary regions,
12,342 direct operands, 2,504 direct declarations and 4,720 direct block bodies.
The broad run reports 281 passed, three counter assertions failed, 19 ignored;
the failed cases' complete AST comparisons passed. Counter updates and a linear
exact-capture test helper are source-fixed but await a refreshed SDK link.
That same pinned BinaryRegion executable also passes complete saved-Doom
PEG/Native AST equality: 5,130 statements and 166,036,219 identical AST/debug
bytes, including spans/directives. Log:
`/private/tmp/nyx-c-pipeline-binary-doom-ast-20261016.log`.
The 785.13-second duration includes two parses and formatting under shared load;
it is not a compile-time benchmark. Named scalar parameter receipts/consumer are
the subsequent 290-test cohort above: focused old/PEG/Native AST equality and
ordinary/incremental validated WASM plus VM execution pass, but the earlier full
Doom artifact/AST results do not certify that later addition.

The earlier `/private/tmp/nyx-c-pipeline-tests-current` executable passes complete saved-Doom PEG/Native AST parity for the
new committed-unit and Direct consumers: 1,808,719 source bytes, 5,131 statements
and 166,118,766 identical AST/debug bytes. The 345.91-second correctness run
includes two parses and full AST formatting/comparison; it is not a compile-time
benchmark. Log: `/private/tmp/nyx-c-pipeline-doom-ast-20261016.log`.

An earlier generated-table library gate passes **225 tests, zero failures,
11 ignored**. A separately linked production-library integration gate passes
**nine tests, zero failures, four ignored**, including a counter proving actual
public Native initializer handoffs. Six executable fixtures run under all five
backends with identical metadata-free standard WASM bytes and independent
validation. The new fixture covers nested tables, records, compound literals,
strings, numeric widths and contextual/effectful fallback.

Performance acceptance is not established. Three-sample table-source frontend CPU
medians are PEG 0.853s, Native/direct 1.529s, Owned 1.197s, Mixed 1.070s and
Pratt 1.726s. Complete compilation of `m_random.c` measures PEG 21.59ms versus
Native/direct 27.35ms CPU. These comparisons do not isolate direct construction
from native matching; an old-Native/direct-disabled control and repairs are
required. Do not claim a compilation speedup from compatibility passes.

A later compile-only whole-Doom CLI cohort does improve: 178.96 -> 123.34
seconds wall, 167.48 -> 119.61 seconds user CPU, both compiling 2,332 chunks.
That single before/after comparison contains multiple compiler/SDK changes and
does not isolate Pratt or Direct. Exact binary hashes and failed-build caveats
are recorded in [the current Doom report](../../../documentation/runtime/doom_debug_current_run.md).
It does not supersede the isolated table-source regressions or establish a
repeatable frontend-only gain.

## Initializer Predeclaration Ownership

The scalar-array predeclaration proof now leases a supported direct initializer
instead of consuming its first owned AST. Source-order definition selection and
byte packing read that immutable value; dropping the lease restores it for the
ordered declaration walk. Unsupported/contextual initializers keep normal
lowering. Reparse generation checks prevent restoration into a replacement cache.
This removes a duplicate recipe replay, not the contextual lowering pipeline.

`initializer_predeclaration_preserves_owned_ast_without_replay` compares complete
PEG/Native results against a test-only consume-first reference and counts owned
handoffs. It includes tentative definitions, enum extents, numeric prefixes,
typedef widths and designated/contextual fallbacks. The ignored
`measure_initializer_predeclaration_reuse` first requires complete AST parity,
then measures six alternating reference/candidate thread-CPU and wall pairs.
Set `VYBE_C_FRONTEND_PREPROCESSED` to benchmark a real saved translation unit;
otherwise it uses twelve 1,024-element signed tables.

The root-debug library executable `vybe_language_c-4acf8386f0c944fe` passes
**242 tests, zero failures, 17 ignored**. The focused ownership test passes all
twelve PEG/Native comparisons, including complete AST equality and fewer owned
handoffs. This cohort also covers the direct builder's compact frames, sparse
rollback anchors, numeric admission and standard-WASM execution fixtures.
The alternating performance benchmark remains unrun; do not count this candidate
as a measured speedup yet. Subsequent default-Pratt activation and broader
direct-AST changes require their own fresh gate.

## Predeclaration Signature Reuse Candidate (2026-10-16)

Top-level predeclaration now passes its already decoded name and parameters to
function registration instead of walking the same declarator a second time.
This changes contextual AST construction work, not grammar acceptance, emission,
or runtime behavior. A test-only reference path retains the former second walk.
`predeclaration_reuses_decoded_signature_without_ast_changes` compares complete
results for PEG and Native, including callback declarations, variadic signatures,
array/VLA parameters, records and rejected input, and counts decoder invocations.
The ignored `measure_predeclaration_signature_reuse` compares six alternating
full-frontend CPU/wall pairs on a 300-prototype input after complete AST parity.

The root-linked debug C library gate passes **214 tests, zero failures, nine
ignored**. The signature oracle passes all 18 PEG/Native comparisons, preserving
complete AST/debug results (including VLA side effects and rejected syntax).
Decoder counts fall, for example, from 24 to 20 for a callback-taking function
and from 23 to 17 for a typedef/prototype/callback combination.

Six alternating full-frontend pairs on 300 prototypes pass AST parity first.
Reference/candidate CPU milliseconds are 288.569/250.635, 279.116/249.569,
277.950/247.145, 268.462/236.938, 272.203/232.390 and 259.741/231.639.
Every pair reduces CPU, about 11--15%. Early wall samples overlapped external
build load and another short correctness run. This is a prototype-heavy
micro-workload, not a Doom compile-time percentage or production streaming claim.
The source-linked executable is `/private/tmp/nyx-c-predeclaration-tests`.
The full 1,808,719-byte preprocessed Doom translation unit also passes the
old-decoding/new-decoding complete AST comparison: 166,118,766 identical debug
bytes, including spans and directives. The two parses and comparison took
249.29 seconds under shared load; this is a correctness gate, not a speed claim.
The later integration executable is `/private/tmp/nyx-c-frontend-integration-tests`.
Its broader interim gate reports 215 passed, two failed, ten ignored: both
failures require regeneration of the optimized parser for the new Pratt operator
capture anchors. They remain failures, not passes established by the older gate.

## Completed-Node Rewrite Boundary

`Walker::rewrite_binary_result` now bypasses the sixteen binary-only passes when
construction already produced a non-Binary root. Their existing non-Binary arms
are identities; pointer result-type retention also only admits Binary +/-.
The ordered chain is unchanged for Binary roots. A test-only reference switch
runs the former full chain and compares complete modules under PEG, Native and
Pratt (39 comparisons). Literal returns remove 16 identity calls; a char-index
expression removes 32 of 48, and a local-array/index example removes 80 of 96
under PEG. Not every expression saves work; wide arithmetic examples can retain
all calls. The current library gate passes **228 tests, zero failures, 13 ignored**.

This is a work-count reduction, not yet a verified CPU speedup. Paired expanded
`sha1.c` measurements were noisy and did not establish an improvement. The first
per-pass-guard candidate was replaced with one entry guard to avoid repeated
discriminant checks on still-Binary expressions. `measure_completed_binary_rewrites`
keeps complete AST equality ahead of six alternating frontend CPU/wall pairs.

Dated measurements and pending-candidate notes below describe their original
cohorts. In particular, the older PEG/Pratt benchmark does not compare Native
against Pratt and does not describe today's default policy.

## October 16 Source Pratt Receipts And Direct Composition

Source Pratt receipt matching is now enabled in normal Native matching. The
declaration-free builder shares one match with Direct; original contextual C
lowering and Native's completion-worker policy remain unchanged. Set
`VYBE_C_SOURCE_PRATT_RECEIPTS=0` for the old Native diagnostic control. The
explicit `pratt` lane still selects capture-Pratt unless receipts are requested
with `VYBE_C_SOURCE_PRATT_RECEIPTS=1`. Activation follows the user's feature-
integration priority, not a claim that the earlier regressing cohort is faster.

The fresh pre-activation root library baseline passes 242 tests, zero failures,
17 ignored (`target/debug/deps/vybe_language_c-4acf8386f0c944fe`). The default-
activation standalone root-debug cohort passes **243 tests, zero failures,
17 ignored**, including the new capture-shape/resource-error oracle and Direct's
VM/byte-identical validated standard-WASM tests. Executable:
`/private/tmp/c-native-pratt-default-tests`, SHA-256
`9cebd94f63a20626290591efc88d1826bdc7c2aeea08468a52c28fdceb19147d`.
The full library run took 55.14 seconds; this is a test duration, not a frontend
speed comparison. The saved-Doom default-Native/PEG full-AST gate also passes:
1,808,719 source bytes, 5,131 statements and 166,118,766 identical AST/debug bytes.
Its 219.79-second duration is a correctness-run duration, not a speed comparison.
This cohort precedes the committed-unit producer below.

### Committed Top-Level Unit Producer

`parse_top_level_units(source, options)` now returns the production walker's
top-level `Vec<Pair>` from separate absolute-offset capture arenas. Native uses
`build_capture_program_at` with declaration-free Source Pratt and Direct in each
match. It merges Source receipts and Direct initializer ownership across units,
filters `frontend_trivia`/EOI, accepts semicolon-only progress, and releases no
items to the walker before the entire sequence succeeds. Main retains the full
global-predeclaration barrier before the ordered AST walk. This is not streaming
emission or permission to publish speculative functions.

Finite work budgets/nondefault depth and non-Native modes keep whole-program
matching. A failed later unit clears prior handoffs and falls back to the
whole-program path for original diagnostics. Failed reparses do not keep stale
Direct ASTs. Unit captures and operand handles retain the original full input
and absolute offsets, not rebased source slices.

Producer source compiles and its narrow executable gate passes 3,144 complete
capture/text/span/error comparisons across empty/trivia/semicolon inputs,
UTF-8 comments, declarations/functions, late syntax errors and finite budgets.
It also proves two Direct initializer ASTs and two Source Pratt receipts survive
cross-unit merging. `/private/tmp/c-pratt-unit-producer-gate.rs` includes the real
production frontend, Grameo and AST rlibs; it substitutes only a decimal-integer
literal constructor for the selected initializer fixtures and has no contextual
walker. Consequently it is not full-module AST or VM/WASM evidence.

Permanent tests: `top_level_units_preserve_absolute_captures_and_merge_pratt_receipts`,
`top_level_units_merge_direct_ownership_and_clear_late_failures`, and
`top_level_units_keep_whole_program_errors_and_finite_budgets`. Fresh whole-
library execution of the latest correctness fixes is pending main's fresh
source-linked cohort. The committed-unit/Direct Body/Primary/Syntax full-Doom
AST gate passes: 1,808,719 source bytes, 5,131 statements and 166,118,766 identical
AST/debug bytes (345.91 seconds, correctness duration only, before the binding
and array-parameter repairs below). No Cargo, release or alternate target was
used by this agent. Native-only four-lane timing remains pending too.

`committed_units_are_consumed_by_the_typed_walker_with_both_handoffs` exercises
`crate::parse_optimized`, not just the producer. With Native Pratt explicitly
enabled it requires complete PEG AST equality, multiple `frontend_unit` matches,
no `program` recapture, actual Source Pratt/Direct handoffs and no owned-
declaration policy. Public tests now cover forward record typedefs, later global
pointer definitions, variadic callback variables, static typed arrays, VLA/array
typedef boundaries and late rejections. The execution gate requires PEG/Native
metadata-free WASM byte equality, independent validation, zero VM exit status,
chunk noEq and emitted WASM abstract-equality call/reference audits. These new
typed/execution tests now have fresh source-linked production execution evidence
below; the new counter-strengthened full-Doom gate still needs its next cohort.

The unit-boundary fixture set and execution/noEq audit are shared by the public
target and `frontend::tests::committed_units_preserve_standard_wasm_execution_and_noeq`
through `tests/support/frontend_unit_compatibility.rs`. Both wrappers now cover
seven unit-boundary programs and eight user-call binding programs: forward
record typedef/global pointer, variadic callback, scalar/pointer/row arrays,
while/continue/break/goto controls, user functions and callback shadowing. The shared helper
compares complete typed ASTs before compilation, then standard WASM bytes,
validation, VM exit and chunk/WASM noEq. The public test target passes a fresh
metadata-only typecheck. Fresh production source-linked execution passes all
15 programs under PEG and default Native using
`/private/tmp/libvybe_language_c_binding_array_fix.rlib` and
`/private/tmp/c-user-binding-array-wasm-gate`. Its old footer counts only seven
base programs, but the shared runner chains all eight binding programs too.
There is one audit implementation, not two diverging copies.

The original Node/read fixture first exposed unconditional libc adapter dispatch
overriding a user function. C call normalization now resolves macros first,
then actual user definitions (including later definitions) and active object
bindings ahead of every library adapter. Extern-only prototypes remain adapters.
Local bindings are tracked at declaration time in block/for/switch scopes, not
with the whole-function local-name overapproximation. Callback return inference
and the fread/fseek result exceptions honor that same binding priority.
The unchanged Word-table fixture then exposed array-parameter metadata counting
the decay pointer as an element pointer. Parameter adjustment and subscripting
now keep explicit pointer depth, decay and row bounds distinct; scalar arrays,
pointer arrays and pointer-to-row execution controls all pass. A pre-array-repair
library snapshot passed 263 tests with this sole remaining fixture failure;
fresh whole-library verification of both repairs is pending main's current build.

### Native Choice Heads

The production grammar now explicitly selects `top_level_item`, `statement`,
`type_specifier_strict` and `type_specifier` for generic choice heads. Native,
Source Pratt and Direct opt into the capability; the composed builder requires
both participants to grant it. Engine guards still retain observed alternatives,
predicate/trivia behavior, finite budgets, diagnostics and original PEG fallback.
This is feature integration, not a measured frontend speedup.

The new source-linked producer gate passes 3,144 capture/span/error/budget
comparisons with generated choice metadata and actual Source Pratt/Direct unit
composition, plus cross-unit handoff/late-failure checks. Its decimal literal
and statement-constructor stand-ins remain explicit: this is not typed C/WASM
proof. All current C test source passes a metadata-only rustc check. Permanent
tests require selected-root metadata, composed capability AND semantics, complete
typed module/error parity, and enabled/disabled capture and resource-budget
parity. Both permanent choice tests pass in
`/private/tmp/nyx-c-native-choice-heads-tests` (0.13 seconds, test duration only).
The strengthened declaration-composition test passes with Direct disabled and
enabled (0.30 seconds), retaining the legacy-handoff requirement in its disabled
control and requiring actual Direct declarations in the combined cohort. This
source-linked executable uses current C source/private tables, root rlibs and
the saved root runtime archive of the same crate hash; no runtime rebuild or
alternate Cargo graph was used. The complete library cohort passes **278 tests,
zero failures, 19 ignored** in 146.23 seconds (test duration, not a speed
comparison). Binary SHA-256:
`ff9ea4a54d46fe35c06f6c5d958089e73cc0c56cb90eaca7d46d02b7fd01aa0e`.
Log: `/private/tmp/nyx-c-native-choice-heads-library-20261016.log`.
Only private generated tables were written; the queued root build owns root
table regeneration. Full-Doom choice-enabled AST/WASM/noEq remains pending.

### Next Emission Boundary

The production Bundle owned lane already calls
`compile_with_imports_owned` -> `begin_emission_owned_prepared` -> `finish`;
`finish` pumps `emit_next`. An explicit caller loop alone therefore does not
create parser/emitter overlap. C still commits a complete normalized `Module`
after its capture/global-predeclaration barrier, helper closure discovery and
module-wide storage/ABI normalization.

The next shared ownership hook is an owned producer cursor inside
`PreparedEmission`: retain the original source indices and function-first order,
release each owned statement only after successful emission, and leave borrowed
modules unchanged. Cache the strict-prologue decision before draining functions.
Retire address-keyed source-body receipts before dropping their owning ASTs;
`source_atomic_free_bodies` is a concrete negative-proof cache whose identities
must not survive allocator reuse. Preserve poisoned-session errors, abort, one
final relocation and the prohibition on executing/exporting pending chunks.

Actual parser/emitter overlap needs a further generic exact-parser provider,
not a C profile switch or runtime ABI change. Its committed inventory must cover
directives/imports/canon, global/type/function declarations, address/atomic and
rest-arity facts, plus the normalized helper closure. Current frontend preparation
returns a full Module and optional scheduling hints, not an owned typed-unit
producer. Until those global facts and C normalization are committed, parsed
capture units are not safe emission units. Shared compiler/runtime files were
inspected only by this agent; no streaming-completion claim is made.

`owned.rs::SourcePrattBuilder` uses the existing Grameo Source Pratt matcher,
transactional unary capture tokens and ordered reduction callbacks. Committed
recipes retain operands and reductions, not context-bound expressions. The
existing `construct_contextual_binary` walker hook resolves those operands and
uses the same typed unary lowering and binary combination. It no longer builds
an operator-token stream, repeats precedence parsing or walks immediate captures
again for successful receipt handoffs. Single-infix receipts combine directly.
All downstream C type/pointer/floating rewrites remain walker-owned.

The authorized grammar change makes each silent `c_*_op` a positive guard plus
the original atomic `binary_op`, retaining captures and longest-choice behavior.
Admission checks the generated Program shape. Finite step limits/nondefault
depth retain reference matching; failed Source Pratt matches recapture for exact
reference diagnostics. Declaration entry probes are excluded, and declaration
initializers retain their original semantic operator callbacks.

`frontend.rs::WithDirect` runs the primary builder and Newton's Direct builder
in one native match. It journals begin-time observer selections (Direct changes
literal observation while a rule is open), paired semantic cursors and hook
frames. Reductions, capture observations/tokens and committed hooks are forwarded.
Direct's certified initializer capture boundary is forwarded too: Grameo denies
omission when the primary observes any omitted unary capture. No second source
parse or live C scope mutation is introduced by composition.

Verified in root-debug standalone rustc cohorts, without Cargo, release or a
different Cargo target:

- 65 frontend tests pass, seven ignored, including full typed-module AST/error
  parity, 324 operator-pair modules, reentrant bindings, failed-RHS/budget journal
  restoration, same-match Direct/Pratt publication and invalidation, and Direct's
  existing full-module VM/standard-WASM gate.
- Root optimized parser and the earlier private generated copy have identical
  SHA-256 `a724c3b284e964e4592778b5a4e4030b1c7b653fcd9bcaf146baebdc6f5110e7`.
- Full frontend parsing plus AST destruction, eight rotating/reversing rounds
  per backend, complete PEG AST equality required before timing. Raw real files,
  not expanded projects; CPU and wall are separate. Upper medians in milliseconds:

| Input / Lane | CPU | Wall |
| --- | ---: | ---: |
| r_draw.c / NativeCapture (Direct disabled) | 179.621 | 182.238 |
| r_draw.c / NativeDirect | 195.509 | 199.383 |
| r_draw.c / CapturePratt+Direct | 193.979 | 195.892 |
| r_draw.c / SourcePratt+Direct | 212.860 | 215.428 |
| tables.c / NativeCapture (Direct disabled) | 914.872 | 928.438 |
| tables.c / NativeDirect | 1322.142 | 1338.191 |
| tables.c / CapturePratt+Direct | 1358.979 | 1367.341 |
| tables.c / SourcePratt+Direct | 1329.772 | 1338.248 |

Source receipts lose to both Native controls on r_draw.c and do not establish
a win over NativeDirect on tables.c. All Direct lanes lose to NativeCapture on
tables.c. External workspace load was variable; this agent ran no other work
during the timed batches. `r_draw.c` has zero initializer handoffs; `tables.c`
has three Direct initializer handoffs and five Source Pratt handoffs in its
composed lane. These are not whole-Doom or whole-compilation speed measurements.

Test/timing executable SHA-256:
`ba5c4d28697cc8e92e5ee19f58fb3dfba8f84f76d5f5734f51083e0775bfbe1a`.
Logs: `/private/tmp/c-source-pratt-controlled-frontend-tests.log`,
`/private/tmp/c-source-pratt-r-draw-four-lanes.log`, and
`/private/tmp/c-source-pratt-tables-four-lanes.log`.
This cohort precedes main's subsequent one-entry rewrite-chain refinement and
the receipt consumer's final wrong-rule rejection guard. A current production
rlib has separately compiled. Its public integration suite with receipt matching
explicitly enabled reports **8 passed, 1 failed, 4 ignored**. The failure occurs
at the NativeDirect/PEG comparison for
`int a[]={0x1e-2,0xbe+2}; int main(void){return a[0]!=28;}` before the Pratt
comparison: the Direct initializer boundary scanner admits hex `e` followed by
a sign as though it were a decimal exponent, despite grammar arithmetic there.
Newton/main own the Direct fix. Five-backend VM/byte-identical standard-WASM,
registered subprocess selection, expanded-source order and the other public
gates pass; this does **not** make the overall compatibility gate pass.
Log: `/private/tmp/c-source-pratt-production-public-gate.log`.
Production rlib SHA-256
`8d9aaa1d94307cea15fa39255e46d2d0c91bfaeb41ac69136e454a6a65d921c4`;
public test SHA-256
`6408d509a1f64954bf0660659e1fc09e63f1f17f29aace870ddf913b1985df98`.

The saved-Doom five-lane complete AST gate subsequently passed: PEG,
NativeCapture, NativeDirect, CapturePratt+Direct and SourcePratt+Direct agree for
all 1,808,719 source bytes and 5,131 statements. The composed lane consumes 9,178
source receipts and 279 Direct initializers. Its 790.28-second total is a
correctness run under shared load, not a benchmark. Log:
`/private/tmp/c-source-pratt-doom-full-ast.log`.

### Declaration-Free Native Integration Candidate

The next implementation replaces Begin/Unary/Reduce receipt events and their
post-commit frame replay with a shared immutable reduction graph assembled at
the actual matcher callbacks. Compact Grameo `EditArena` checkpoints restore
the builder stack; side-vector rollback restores graph nodes/roots. Only unary
captures that are actual Pratt operands become graph atoms. Single-infix
consumption remains allocation-free. Existing capture anchors are retained.

A subsequent stack reduction stops retaining completed binary values: the
enclosing unary capture supplies the actual operand handle. Nested unary
wrappers no longer need observation frames, and no completed root values remain
to make observer queries scan earlier expressions. Unary capture observation is
now limited to active Pratt frames, allowing Direct's existing certified literal
initializer elision outside those frames. No operator captures are suppressed.
The same-match composition test explicitly requires an empty initializer
descendant list. This latest capture-elision composition still needs execution
in the fresh full-frontend cohort; the allocation table below predates it.

Native can now compose a declaration-free SourcePrattBuilder with Direct in one
transactional match. It consumes receipts through the existing contextual unary
callback; a missing receipt returns to the original Native typed fold. Native
does not acquire owned declaration construction or completion-worker policy.
This is now Native's default path, with a diagnostic opt-out. Fresh activation
correctness and timing gates are being run; no speedup is implied by activation.

Actual saved-Doom global-allocation attribution (calls include reallocations;
bytes are cumulative requested bytes, not live/RSS measurements):

| Builder | Calls | Requested Bytes |
| --- | ---: | ---: |
| Original event journal + declaration hooks | 384,221 | 286,144,788 |
| Compact graph + declaration hooks | 305,085 | 292,036,912 |
| Compact graph, no declaration hooks | 194,278 | 215,136,008 |

The original journal's commit alone allocated 79,246 times / 19,413,952 bytes;
the latest declaration-free graph commit allocates twice / 671,792 bytes.
All three retain 1,399,746 captures. These are matcher-only allocation counts,
not full-frontend speed evidence. The temporary attribution harness includes
the real builder/generated parsers and AST/Grameo root rlibs, with unrelated
walker entry points stubbed. It must not be used to claim full-module parity.

The graph passes 324 operator-pair schedules with exact unary handle text/spans,
and failed-RHS/resource-budget rollback in both builder modes. New Native-only
whole-module AST/rejection tests, all 324 full-AST operator cases, and
`measure_native_pratt_full_frontend` are ready. The latter compares NativeCapture,
NativeDirect, NativePrattCapture and NativePrattDirect in rotating/reversing
order after complete PEG AST parity. Current source typechecks, but root-linked
execution is blocked by the missing matching Webcore rlib in the compiler's
forced plugin dependencies. No Cargo or alternate target was started.

Pending: latest graph/Native whole-frontend execution, corrected Direct/public
parity, full-frontend timing and default activation backed by that evidence.
The separate identifier-only work below is historical, not this Pratt deliverable.

## October 16 Identifier-Probe Builder Bypass

`frontend.rs` now extends Pratt's existing `ident_name` entry bypass to Owned
and Mixed. This atomic identifier/keyword-guard rule cannot produce expression
or declaration handoffs. Matching uses the existing `NativeBuilder` and the
same options; same-source cache invalidation still precedes matching, including
on failure. It no longer constructs semantic builders or publishes empty cache
entries for these probes. Native/PEG defaults, grammar, walker, normalization,
AST construction and contextual lowering are unchanged. `owned.rs` is untouched.

Verified with a standalone, opt-level-zero `rustc --test` harness including the
current frontend sources and the existing root-debug generated parsers and
Grameo/AST/libc rlibs. No Cargo, release or alternate target was used. The harness
deliberately stubs walker entry points; it cannot prove full-module AST or VM
compatibility. The initial attempted C-archive link failed on its dependency
chain, before switching to this matcher-only harness.

- New parity test passes: ten accepted/rejected identifier inputs, 97 step
  limits and three depth limits; complete capture rule/text/spans and Debug
  errors agree with both former semantic builders (5,820 comparisons).
- New cache test passes for Owned/Mixed/Pratt on successful and failed probes:
  same-source entries disappear, unrelated entries survive, no empty publication.
- Corrected matcher-only suite: 19 passed; owned construction suite: 26 passed.
  Logs: `/private/tmp/c-identifier-probe-{matching,owned}-tests.log`.
  An earlier filter accidentally included the walker-dependent operator-pair
  test: 19 passed, one deliberate stub panic. That test was excluded from the
  corrected run and remains pending, not a compatibility pass.

Before/after debug microbenchmark: six samples per backend per batch, 60,000
identifier matches per sample, including capture/cache teardown. Batches ran
before/after/after/before; backend order reverses each round. No other harness
was running during the repeat pair. External workspace load was uncontrolled;
the first after batch overlapped a short correctness run. Median milliseconds:

| Batch | Native CPU / Wall | Owned CPU / Wall | Mixed CPU / Wall |
| --- | ---: | ---: | ---: |
| Before | 209.584 / 211.729 | 304.602 / 308.045 | 677.153 / 684.926 |
| After | 211.617 / 213.730 | 238.373 / 240.489 | 253.835 / 255.975 |
| After repeat | 235.444 / 239.644 | 263.414 / 268.367 | 302.806 / 307.030 |
| Before repeat | 304.144 / 532.608 | 459.738 / 840.415 | 985.830 / 1807.987 |

The unchanged Native control shows substantial load drift. Median per-round
CPU ratios to Native are Owned 1.464/1.147/1.131/1.523 and Mixed
3.247/1.221/1.283/3.230 in batch order. This supports reduced probe overhead,
not a whole-frontend or whole-compilation speedup, and changes no backend default.
Harness and raw logs: `/private/tmp/c-identifier-probe-harness.rs` and
`/private/tmp/c-identifier-probe-{before,after}{,-repeat}.log`.

Executable SHA-256 (rechecked after root artifacts were reported disappearing):
before `de7699b06e36410be782a60ed5dc09f00095d9ecde368db3f7213a1bfcf1bc44`;
after `330a43a362c148fa3028b1fb1baaef5b113092c2f30a0be4a727cc85751406c6`.
Linked `libgrameo-de2ba92b1eca0540.rlib` SHA-256
`684bc27a0bba4f10b00e05d65df12b780c38991faa529108e0136731f40198c6` and
`libvybe_ast-fe038211cd101cf9.rlib` SHA-256
`ad46801a52e53311c83d2839929d4c9438710e41e9f657ddc7cd28bcaa084ff5`
were unchanged and still present at that recheck. The before executable was
compiled before the patch; the after includes the two new tests.

Pending: current linked whole-module AST/spans/directives/context parity,
operator-pair walker gate, VM/native-oracle/WASM gates, representative complete
frontend CPU/wall and allocation census. Do not infer those from this harness.

## One-Command Correctness Gate

When the shared build queue is free, run the existing root debug test targets:

```sh
python3 tools/test_c_frontend_gate.py --output /private/tmp/c-frontend-gate-001
```

This is a thin sequence of `cargo test --offline --jobs 1` invocations using
the workspace manifest and its existing `target/debug` tree: nine default
suites, each with its own invocation, log, exit code, and receipt entry. It does not use
release mode, another build tree, or overlapping Cargo jobs. Cargo recompiles
current dependencies instead of guessing which prebuilt executable is fresh.
Use `--dry-run` to print commands without starting builds or tests, and a new
output directory for each real run. Commands, partial C frontend source hashes,
per-suite logs, exit codes, complete Rust summaries (including ignored/filtered
counts), and wall times are saved in `receipt.json`. The partial snapshot omits
normalizers/libc/shared files and is not binary provenance. Results are incomplete
until every requested suite finishes; zero-test and ignored-only runs cannot
succeed. Ignored tests and omitted optional tiers are not counted as passes.
`--suite-timeout SECONDS` is a positive integer, default 1800, covering each
suite's lock wait, build, and tests. Timeout is recorded separately from assertion
failure and makes the overall receipt fail; later suites continue. Each Cargo
invocation starts a new POSIX session/process group. Cleanup signals only that
runner-owned group (TERM, up to five seconds for the leader, then KILL including
remaining descendants), then reaps Cargo. It never targets unrelated builds.
Non-POSIX runs fail closed rather than kill only Cargo and orphan its children;
`--dry-run` remains portable. A wrapper already running old code does not gain
timeout protection from an on-disk update.

| Existing Target | Evidence |
| --- | --- |
| C lib tests | Grammar/capture/span/error-budget, native/PEG/Pratt AST and lowering unit regressions; not exclusively parser tests |
| `frontend_optimization` | Complete bounded AST/spans/directives/rejection parity; five-backend VM execution, byte-identical metadata-free WASM and validator checks; Native is a parser backend, not a native C oracle |
| Four ABI/storage suites | Separate sequential Cargo invocations for `scalar_pointer_carriers`, `pointer_array_places`, `scalar_index_store_capture`, and `global_record_field_owner`; native/VM/parity/validator/noEq checks according to each existing target. A failure in one cannot hide the others. |
| `stdio_descriptors` | Standard-stream and file-expression descriptor transport, single evaluation, Native/PEG AST parity, VM execution, metadata-free WASM validation and abstractEq/Ne call audit; deterministic terminal-capability selection and absence checks |
| `doom_patterns` | Native C oracle, Vybe VM assertions, chunk-level noEq and metadata-free WASM validation using the existing matrix |
| `record_places` | Existing source-backed subsystem/native oracle checks, VM assertions, chunk noEq and emitted WASM Call/ReturnCall/RefFunc abstractEq/Ne audit where its shared validation helper is used; coverage varies by test |

These suites are run once with the integrated default (the wrapper clears an
inherited frontend override). Explicit backend parity lives in the existing
frontend tests; this is not a claim that every record fixture ran every backend.
Wall times include Rust builds, C preparation/compilation and execution, not VM-only timings
or parser speed measurements. Failing gates remain failures; later targets
still run so one regression does not hide others.

Optional tiers reuse existing infrastructure:

```sh
# Slow emission/audit only: full Doom module, no GUI or gameplay execution.
python3 tools/test_c_frontend_gate.py --output /private/tmp/c-frontend-gate-slow \
  --slow --preprocessed /private/tmp/nyx-doom-preprocessed.c

# Separate CLI exit-code check, not parser parity; repeat with native/pratt.
VYBE_C_FRONTEND=peg target/debug/testrunner run \
  tests/c/c_pointer_to_array_multidim --vybex target/debug/vx --cold -j 1 --json
```

`--slow` invokes only
`doom_module_contracts::complete_doom_has_no_dynamic_equality_and_validates_as_wasm`:
all-chunk noEq plus complete metadata-free WASM validation, **not full gameplay**.
`--preprocessed` runs the existing ignored full-input Native/PEG AST parity test,
not a benchmark or whole-input Pratt proof. Testrunner checks exit status only;
it supplies neither AST parity, native C oracle, WASM validation nor noEq proof.
Its files must carry the normal `vybe-test:` discovery header. No results for
this new orchestration command are claimed until a real receipt is recorded.

## October 15 Bounded Pratt Handoff Candidate

`frontend/owned.rs::construct_contextual_binary` now looks ahead through four
immediate captures. Exactly three children are validated as unary/operator/unary
before invoking either typed C callback. Their owned handles are then consumed
directly, removing token/span buffer construction and a second immediate-child
capture traversal from the existing single-infix fast path. Longer regions feed
the retained prefix and untouched iterator into the same token validator and
Grameo Pratt engine. No capture suppression, contextual classifier, grammar,
walker rewrite, default backend, or streaming behavior changes.

The new `contextual_pratt_prefix_lookahead_preserves_complete_operand_spans`
unit covers singleton, one-infix, inline and spilled lengths, comments, nested
calls/sizeof, exact operand text/span/order, and rejection of a non-expression
boundary without a callback. Existing all-operator single-infix, reentrant,
complete AST, budget/error and five-backend execution gates remain required.
Rustfmt syntax parsing passes; this candidate has not been built or executed.
No CPU/allocation improvement or production activation is claimed.

The historical 29.7% median CPU regression below compares PEG against native
matching plus contextual Pratt, not Pratt against native matching alone. Native
matching, operator-token validation, duplicate child traversals and Grameo's
long-region reduction stacks are possible costs; the existing measurement does
not attribute percentages to them. Shared-load variation was substantial. A
current-source before/after Pratt cohort and a same-cohort native-only control
are required to isolate this change; full-source AST parity must precede timing.
Whole-source captures and downstream typed lowering/normalization remain in all
these modes. This patch does not eliminate the dominant historical walk phase.

`frontend::tests::measure_contextual_pratt_frontend` now prepares that three-way
control: complete Native and ContextualPratt ASTs must each equal PEG before
timing. Each three-round cycle rotates PEG/Native/Pratt through all run positions;
the next cycle reverses order. Set `VYBE_C_FRONTEND_SAMPLES=6` for both cycles
(the default remains three). Measurements include module destruction, report
per-backend CPU/wall samples and ranges, and report per-round candidate/baseline
ratios plus medians of those paired ratios for Native/PEG, Pratt/Native and
Pratt/PEG. Construction profiling is outside timed rounds. This test change is
syntax-checked only; shared libtest execution and measurements remain pending.

## Enabled Behavior

At the original PEG-default implementation stage, its successful singleton handoff
skips the unused fallback capture traversal, as do other successful handoffs.
This production change does not change operand evaluation or C rewrites;
its standalone performance effect has not been isolated.

The retained PEG fallback collects decoded `(BinOp, precedence)` tuples
in `walk_binary` for its sole `fold_binary` consumer, removing one temporary
`String` allocation per binary operator. Operand lowering order, source spans,
existing folds, and downstream typed C rewrites remain unchanged. The new
`decoded_binary_operators_preserve_precedence_order_spans_and_folds` unit
covers 324 ordered operator pairs, singleton spans, equality folds, and wide
shift folds, but has not yet been executed against a fresh build. Fresh
frontend parity and performance gates are pending; no speedup is claimed,
and the historical evidence below does not verify this change.

`VYBE_C_FRONTEND=pratt` now selects the existing certified native capture
matcher followed by Grameo Pratt binary-AST construction at the ordered C
walker scope boundary. Unary operands are lowered exactly once through the
existing typed C context hook. The existing integer, pointer, logical and
floating-point rewrite pipeline still runs afterward.

The next opt-in composition also journals raw scalar declarators during the
same native match, using the mixed backend's existing deferred declaration
target. Pratt consumes them only at the existing ordered walker boundary;
unsupported shapes retain contextual lowering. It does not construct mixed
binary recipes alongside Pratt, and identifier-only probes keep the native
no-observer path. This source change and the borrowed backing-classification
change have not yet passed fresh execution gates. Exact complete-AST parity
and alternating paired CPU measurements are pending; the negative historical
benchmark below remains the measured baseline, not evidence for this change.

### October 14 Composition Verification

The root debug no-run build completed successfully. Direct gates passed:
capture/span/budget parity (1 test), contextual Pratt units (5 tests), the
324-operator-pair parity gate (1 test), and frontend integration (8 tests).
Integration includes identical metadata-free standard WASM bytes, validation,
and VM execution across all five backends on its five bounded fixtures.

The saved preprocessed Doom input passed complete AST parity: 1,808,719 bytes,
5,073 module statements, 23,287 parser calls, 11,655 contextual Pratt handoffs,
and zero declaration handoffs. This is AST parity, not a full-game execution
gate. Whole-source captures and contextual lowering remain.

Three alternating paired samples completed successfully (seconds):

| Pair | PEG CPU | Pratt CPU | PEG Wall | Pratt Wall |
| --- | ---: | ---: | ---: | ---: |
| 0 | 170.705 | 176.387 | 253.543 | 262.032 |
| 1 | 86.734 | 122.509 | 87.241 | 124.824 |
| 2 | 94.474 | 110.435 | 94.763 | 111.463 |
| Median | 94.474 | 122.509 | 94.763 | 124.824 |

The composed route used **29.7% more median CPU** and **31.7% more median
wall time** than PEG. Each paired CPU result was slower. There is no speedup
claim or Pratt default promotion; PEG was the default in that cohort. Native-only matching was
not sampled by this benchmark, so this does not isolate the marginal cost of
composition. Concurrent builds and substantial sample variation qualify the
measurement; Rust build/link time is excluded.

The C library compile began at 22:34:36 UTC, with walker SHA-256
`c7316defc9c534eaf932a28775cde9193a261b86a9c96c4fe64473c6616ac79a`
and frontend SHA-256
`81e441ebbb1be23b6caa4b509db066f33b1a32f9cbef6b3520082e9240dbc71d`.
Later source guard fixes are not covered by these results. The test executable
was removed by a concurrent target clean while its already-mapped benchmark
process continued to successful termination.

Input: `/private/tmp/nyx-doom-preprocessed.c`, SHA-256
`a3cb210bde65a45bb83426dfaff8f983bb73b4c97a6d3a81aa994df86e6f271a`.
Logs: `/private/tmp/nyx-pratt-{capture-parity,unit-parity,operator-parity,frontend-parity}.log`
and `/private/tmp/nyx-pratt-doom-paired-benchmark.log`.

The remaining prepared-change notes below describe individual earlier changes;
the results above verify the composition cohort, not isolated performance wins
for any one allocation reduction.

At that stage `native`, `owned` and `mixed` were explicit selections. Mixed recipes
store a committed reduction schedule rather than reconstructing precedence
on every handoff; singleton regions avoid a child-vector allocation.

The contextual-Pratt opt-in now prepares up to seven infix tokens (four
operands) in a stack buffer instead of allocating its token `Vec`. Larger
regions spill to a heap vector without changing token spans/order, complete
validation before C callbacks, or the existing Grameo Pratt reduction engine.
Singletons already bypass construction in `walk_binary`; this does not add a
second singleton optimization. Instrumented construction profiles now report
inline versus spilled Pratt regions. Uninstrumented parsing records neither
counter. Pair/iterator clones share Grameo's capture arena and are not deep
capture-tree copies.

The opt-in also reduces a completely validated two-operand/one-infix region
directly through the same ordered unary callbacks and `combine` operation.
This avoids Grameo's temporary `values` and `pending` vectors for that shape;
longer regions still use the existing Pratt engine. Instrumented profiles
count these regions separately. The new
`contextual_pratt_single_infix_keeps_complete_peg_ast` unit covers all eighteen
operators, comments/spans, typedef/sizeof context, pointer decay, side effects
and nested expressions; the reentrancy/order unit also includes this path.
These source changes are prepared, not freshly compiled or executed. The
allocation reduction is structural, not a measured whole-Doom speedup.

The three new `contextual_pratt_inline_` units cover token/spill boundaries,
complete PEG AST parity, nested handoffs and callback order. They are prepared
but uncompiled/unrun against the fresh cohort. Whole-Doom complete AST parity,
five-backend VM/metadata-free WASM parity, native C oracle gates and alternating CPU/wall
measurements remain pending. No allocation census or speedup is claimed from
historical results. Grameo's reduction stacks and whole-source captures remain;
this is not capture-free parsing or streaming.

PHP's current builder uses `supports_pratt_at`, `accepts_pratt`, transactional
`reduce`/rollback and `suppress_pratt_captures` for accepted owned regions.
C's contextual mode does not implement that capture-suppression protocol:
its complete unary captures still feed the scope-sensitive typed hook. This
local scratch-allocation change does not replace or implement those hooks.

`expected_execution_and_standard_wasm_match` compares PEG/native-matcher/owned/
mixed/Pratt output and VM exits; "native" there is a frontend backend, not a
native C oracle. Independent native C fixtures remain a promotion requirement.
The whole-Doom timing gate reports PEG versus current Pratt. To attribute an
effect specifically to inline token storage requires before/after Pratt
cohorts with otherwise identical sources/dependencies; a PEG comparison alone
does not isolate this allocation change.

This is not capture-free parsing or streaming. Whole-source PEG captures,
predeclaration and module normalization remain. No per-function streaming
compilation/commit path is implemented. Unobserved-failure-prefix capability
is declared by builders, but source-matched C code generation emitted no
nullable-prefix certificates; no C speedup from that capability is claimed.

## Evidence (2026-10-11)

The preprocessor hook dump `/private/tmp/nyx-doom-preprocessed.c` contains
1,808,719 bytes, SHA-256
`a3cb210bde65a45bb83426dfaff8f983bb73b4c97a6d3a81aa994df86e6f271a`.
The whole prepared AST, including spans and directives, matched PEG: 4,986
top-level statements, 11,719 contextual Pratt handoffs, 26,560 operands.

Current-source frontend unit tests: 37 passed, 2 ignored. Integration tests:
7 passed, 4 ignored, including complete AST/rejection parity and byte-identical
standard Wasm with zero execution exits for five fixtures. All 324 ordered
binary-operator pairs also match the real PEG walker AST.

Initial three alternating whole-frontend debug CPU samples (including module
drop) had medians PEG 55.987 s and Pratt 51.343 s. Ranges were 55.742-69.683 s
and 36.817-80.062 s respectively. These noisy samples do not authorize a
production-default change. A second complete-AST gate also passed; its CPU
medians were PEG 74.520 s and Pratt 63.728 s, with ranges 66.099-77.761 s
and 61.119-84.522 s. In both runs the final paired Pratt sample was slower.
The repeat log is `/private/tmp/c-doom-contextual-pratt-repeat.log`.

The production fallback-traversal guard was subsequently verified against
37 frontend tests and 2 compiler-capability tests; its test binary SHA-256 is
`04be39a28969dd4f6fc79ae1fabac6c5689a0201ae0f96887ab90ee70698dc02`.
Logs are `/private/tmp/c-contextual-pratt-guard-unit.log` and
`/private/tmp/c-contextual-pratt-guard-capability.log`.

Instrumented old-CLI wall phases were matcher
31.115 s and walk 207.224 s; those are not CPU measurements or evidence for
the latest root binary.

Standalone `rustc` linked current C sources against the existing coherent
shared dependency snapshot, not a new Cargo target directory. The timed
binary SHA-256 is
`e1f31ba126d5410f528df4d1acdbf992a9129697d46cbbb5302bd4a0d71062d8`.
Its source hashes are `/private/tmp/c-contextual-pratt-current-source-hashes.txt`;
other agents' subsequent walker edits are not represented by that binary.
Logs: `/private/tmp/c-doom-contextual-pratt-frontends.log`,
`/private/tmp/c-contextual-pratt-coherent-unit.log`, and
`/private/tmp/c-contextual-pratt-integration-all.log`.

## Reproduce The Whole-Source Gate

### Current Promotion Gates (2026-10-13, Pending Execution)

`expanded_real_sources_keep_include_and_macro_order` now compares the complete
contextual-Pratt AST with PEG for the expanded calculator, Doom function-pointer
probe, and multi-glob fixture. These are full debug representations, including
spans and directives, not function outlines or matcher-only comparisons.

`environment_selected_production_dispatch_matches_peg_in_isolated_processes`
starts separate test subprocesses with `VYBE_C_FRONTEND=peg` and `pratt`; it never
changes the parent process environment. Each child verifies the registered parse
function, accepted/rejected source parity on a fresh thread, and complete
project-prepared AST parity for the same three expanded sources. Production
project loads retain their registered parser without an override. Construction
counters prove successful Pratt dispatch actually reaches the handoff; PEG
records zero Pratt handoffs. Instrumented timing/source-dump options are removed
from child environments.

These new tests are source-prepared only: no Cargo or fresh test execution was
performed while the parent's GUI and focused compiler verification were active.
The production default is unchanged. Capture-free parsing and streaming
compilation remain absent. Earlier October 11 results do not certify these tests
or the latest contextual walker. Full-Doom CLI AST parity, current backend
execution/noEq/metadata-free WASM gates, and repeatable representative CPU
measurements remain required before any default switch.

With an agreed shared debug build, run the ignored unit test
`frontend::tests::measure_contextual_pratt_frontend` using
`VYBE_C_FRONTEND_PREPROCESSED` set to the preprocessor hook dump and
`VYBE_C_FRONTEND_SAMPLES=3` or greater. Leave `VYBE_C_DEBUG_TIMINGS` unset;
collect instrumented phase traces separately. The test compares the complete
prepared AST before alternating thread-CPU and wall samples.
