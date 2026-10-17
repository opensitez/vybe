# Python registered owned hybrid frontend

`parse` and language registration now use the owned hybrid frontend.
`parse_with_direct_ast` exposes that same frontend; the older
`parse_with_direct_literals` name is a compatibility alias. `parse_with_peg`
is the explicit reference frontend. Ordinary Grameo parsing remains PEG.
This converts Python's registered path to hybrid construction, but does not
claim complete direct ownership of every Python production.

The current registered backend uses immutable expression recipes instead of eagerly
building common AST during speculative matching. The explicit
`parse_with_eager_ast` entry preserves the earlier eager backend for comparison.

| Production path | Current construction | Contextual lowering |
| --- | --- | --- |
| Bitwise, shift, additive and multiplicative expressions | Real Pratt island over full `unary` atoms; immutable Atom/Binary recipes preserve precedence and byte spans | Original operator helpers construct common AST left-to-right at live-scope consumption |
| Calls, member/index chains, powers, prefixes, comprehensions, generators and f-strings | Original unary operand captures are retained and reused without rematching | Original walker lowers each operand in the current scope |
| Lists, tuples, assignments, returns and leaf statements | Captures remain; nested arithmetic can use recipes | Existing statement/container normalization preserves order and spans |
| Functions, classes, parameters, defaults and decorators | Eligible nested arithmetic uses recipes | Original declaration headers, bodies, scopes and descriptor handling remain |
| Imports, directives and module policy | Existing complete directive record preserved | Original library mounting and alias policy |
| Final normalization | Existing semantic applicability facts avoid irrelevant walks | Required passes keep their established order |

The recipe cache requires original source identity, rule and absolute byte span.
Repeated consumption resolves the immutable recipe against original operand
captures in the then-current scope. Matching errors retry the PEG reference;
contextual lowering errors propagate unchanged. Ordinary grammar bodies and
rule IDs remain unchanged; the optimized `bitor_expr` annotation now binds full
`unary` operands. Whole-file commitment still prevents streaming.

The following implementation notes describe the earlier eager comparison
backend where applicable.

Grammar performance directives bind additive to a context-neutral scalar atom with left-associative `+`, `-`
and higher-precedence `*`, `/`, `//`, `%`, `@`. The PEG grammar remains the
fallback certificate. A lexical hint selects regions; grammar matching and the
balanced owned-value check are authoritative. The hint for selection, begin and
finish is cached over immutable source; speculative rollback cannot invalidate
it. Rejected hybrid input retries the reference frontend to preserve exact
PEG diagnostic ordering.

Transactional arena compositions use immutable frame IDs and preserve rollback.
Exact-span expression wrappers preserve their completed child and remove only
the immutable frame with one undo entry instead of pop/pop/push. Changed-span
wrappers retain their original contextual handling. First reads move common AST values; repeated reads rebuild certified lexical
leaves and composition events. Keys include the anchor rule and range, and reads
verify original source identity. Python's postfix normalizer resets nested
container spans; the ownership bridge preserves scalar versus comma-tuple
boundaries and restores the same emitted spans.

The compulsory bytes/memoryview walk records deepcopy, logging, asyncio, Await, future-callback
and membership triggers from lowered AST nodes after previous rewrites. There is
no additional whole-tree scan. Enum helpers retain their walk when the semantic
enum registry is populated or membership calls occur; membership can also
recognize ordinary value objects. Nonliteral bytes/memoryview replacements, deepcopy candidates and enum replacements
conservatively retain subsequent walks because new descendants may be unvisited.
Otherwise absent triggers skip the respective walk in the hybrid frontend; the
explicit PEG oracle retains all walks. Statement-root updates remain ordered.
`VYBE_PYTHON_NORMALIZATION_STATS=1` reports both late and enum candidate/retained
walk counts. On numbers, the hybrid skips all 124 late, 31 enum and 31 deepcopy candidates.

Module recognition now indexes only the immutable static surface; dynamic
namespace declarations and live import/alias maps retain their existing lookup
order. Scope mutation cannot make the static index stale.

Production generation explicitly enables native ordered-choice and rule-head
dispatch through `compile_with_options(optimizations: true)` and
`emit_with_api(HYBRID_CAPTURES)`: 62 choice certificates and
46 rule heads target matching work. The grammar and Pratt bindings are unchanged;
the shared emitter and ordinary PEG defaults remain unchanged. Real sequence,
finish callbacks and rollback frames remain intact. Function metadata readers
now share immutable body snapshots; the forwarding adapter copies a snapshot
before editing and emitting an independent declaration.

No shared Grameo API blocks remaining coverage. Full ownership needs a
language-local deferred representation and ordered normalization for functions,
classes, comprehensions, calls, imports and general control flow. Speculative
construction must not use a default walker before contextual facts exist.

Validation and CPU measurements are recorded in
[direct_ast_evidence.json](direct_ast_evidence.json). The assignment and
normalization frontier and wrapper checkpoint passed 61 source tests, 25 exact-expected compiler/VM
cases and 13 selected real fixtures through both paths, including all four
formerly timed-out exception fixtures, deepcopy, inspect, pydoc, doctest, nested
qualification and JSON hooks.
The combined corpus passes 11,501 accepted and 305 rejected files with zero
AST, span, directive or error differences. Its 11,812 paths include six encoding
skips; the bounded stdlib subset excludes 42 recorded PEG crashes/timeouts.
Two initial assignment differences exposed the existing silent-augmentation
recovery policy inside string text; those statements now use that same adapter.
Existing generated HashSet ordering is narrowly canonicalized for comparison
and independently reproduced with PEG versus PEG.

The matching follow-up uses isolated frontend O2, matching the Python dev package
setting, and fixed historical compiler/runtime/AST pins. It measures neither the
whole current debug CLI nor release execution. Rule-head fusion improves numeric
hybrid/PEG ratios from 1.094 to 1.074 initially and 1.075 to 1.071 in reversed
executable order; closure ratios improve modestly in both orders. Absolute CPU
moves with the PEG control, and the pydoc gain does not repeat, so there is no
blanket raw CPU claim. Numeric hybrid remains about 7% slower than PEG. The body
snapshot ownership change is not presented as an independently measured speedup.
Currentframe filtering, a broader borrowed function resolver, a membership
shortcut and a larger lexical-hint cache were removed after CPU checks.

Native numeric sampling identifies `module_namespace_path` in 174 of 1,640
active-worker samples (10.6%). It previously scanned the static export table and
cloned ordinary identifier strings before returning no namespace. The static
module-name index now initializes once, and negative recognition borrows names.
Runtime registration, import and alias queries remain live. Ordinary `str`/`repr`
checks likewise borrow their operand until a bytearray rewrite needs ownership.

The latest O2 before/after measurements use the same frozen dependencies, ten
alternating samples of 100 parses per mode, and repeat in reverse executable order.
Registered numbers CPU falls from 3.151ms to 2.574ms initially (18.3%), and from
3.870ms to 3.433ms in reverse order (11.3%). PEG also improves; the latest
hybrid/PEG ratio remains 1.0541. The initial closure gain does not reproduce in
reverse order, so this round claims the numeric improvement only. Atom-only
Pratt activation and its unused extra depth counter were removed; their CPU
change was near neutral. Two lexical allocation hints failed CPU gates and
were discarded. Grammar acceptance, predicates, trivia, work/depth limits and
rollback remain separately tested against the captured PEG matcher.

Matched real-file measurements distinguish three effects. On the numbers
fixture, package optimization level2 reduces registered frontend CPU from
32.246ms to4.168ms per parse compared with level0, about7.74x. At the same
level2, hybrid remains9.2% slower than PEG (3.816ms); on the closure fixture
the gap is2.7%. Hook allocation wins are not claimed as CPU speedups. Earlier
same-profile measurements improved from1.162 to1.099 after skipping semantic-
inapplicable late walks. Earlier measurements include lexical-hint memoization
and keep applicability collection out of the explicit PEG reference.

A capture-only diagnostic measures about0.944ms of matching and0.024ms of
indentation preprocessing, leaving most frontend CPU in initialization,
construction and normalization. Separate warm closure diagnostics measured
2.176ms frontend,22.814ms compiler and11.240ms VM wall time; these are phase
diagnostics, not matched speedup evidence. The four exception-isinstance fixtures that previously timed out now use
linear generated boolean guards: each StrictEq predicate occurs once, in its
established evaluation order. Their generated AST intentionally changes.
Twelve simple predicates normalize to58 nodes instead of14,332; the full
32-candidate guard contains222 nodes. Six new exception/type/multiple-inheritance
execution cases, ten existing cases and all four real fixtures pass. The
exception_isinstance frontend now takes about2.45ms CPU per parse; its former
reference run exceeded15s. This is documented separately in
[exception_predicate_evidence.json](exception_predicate_evidence.json).

Rebuild vybex to load the registered frontend and Python package profile
setting; an old executable keeps its previously linked parser. The compact
`examples/frontend_cpu.rs` command in the evidence record reproduces complete
frontend comparisons without including compiler or VM work.

The 2026-10-09 grammar-directive migration emits byte-identical optimized parser
bytes to the former Rust binding configuration. Ordinary `compile` ignores the
directives and also emits unchanged bytes. The migrated build passed 61 source
tests, 25 expected-output execution cases and the combined corpus gate
(11,501 accepted, 305 rejected, zero AST/span/directive/error differences; six
encoding skips). Configuration migration itself has no CPU gain claim. Prefix
unary ownership and position-aware observation remain frozen experiments. Constructor
applicability facts are separately gated and registered as described below.

The constructor frontier now records three exact semantic shapes during the
existing syntax walk and skips the subsequent constructor walk when none occurs.
Syntax rewrites only fold literals, clone/move previously visited boolean
operands, or create a slice helper with leaf defaults, so candidate facts remain
complete. This removes all 31 constructor walks on the numbers fixture; three
compulsory final expression walks remain. Its independent gates passed 62 source
tests, 25 expected-output execution cases and the full combined corpus with zero
differences. Two real constructor fixtures and a types.new_class field-population
case pass both frontends; the historical runtime's dynamic attached-method case
retains its identical existing error. Private frontend microbenchmarks show small
normalized closure/pydoc benefits in both executable orders; numeric CPU is
inconclusive. These opt2/debuginfo0/codegen16 controls do not measure the full
DEBUG CLI or Aster. See `constructor_frontier_evidence.json`.

A subsequent same-binary diagnostic showed percent-level ratio variation, larger
than some constructor observations. Treat those small CPU changes as directional
measurements rather than a statistically established gain; removing certified
absent walks and exact semantic parity are the concrete retained results.

The additive Pratt island now matches an appended `pratt_scalar_atom` rule
(numbers, names, True/False/None), bypassing five ordinary matching wrappers for
already-selected arithmetic. Ordinary PEG productions do not call that atom and
existing user-rule IDs are stable. Scoped postfix normalization remains in the
existing arithmetic lowering; boundary callbacks retain exact spans. Independent
gates passed 62 source tests, 25 expected-output VM cases and 11,501 accepted/305
rejected corpus files with zero differences. Small private normalized numbers
and pydoc changes point positive in both executable orders, while closure is
mixed; noise prevents a decisive CPU claim. See `scalar_pratt_evidence.json`.
Bare atom-only Pratt remains disabled, and prefix/position-API/edit experiments
remain outside production.

The registered bitwise/shift region now owns scalar operand ASTs across six
precedence levels through multiplicative arithmetic. Ordinary PEG productions
remain unchanged, with three unused operator rules appended. Existing user-rule IDs remain
stable; the synthetic EOI enum tail moves three positions. Both frontends
share the original ordered Counter/dict/set/enum/literal/protocol lowering in
the current walker scope. Gates passed 64 source tests, 33 expected-output VM
cases, and the combined 11,501 accepted/305 rejected corpus with zero differences.
Numeric hybrid/PEG CPU ratios improve directionally in both private measurement
orders but remain 3.5–4.1% above PEG; closure/pydoc are mixed. No broad CPU or
whole DEBUG compilation gain is claimed. Existing frozen-runtime Counter and
int(IntFlag) failures are documented in `bitwise_pratt_evidence.json`.

A later paired bitwise/additive eligibility cache passed 65 source tests,
4,126,613 original/new classification pairs, the full combined corpus, and
33 expected-output VM cases. CPU signs flipped between executable orders, so
the experiment is not activated and has no speedup claim. Its compact results
are in `paired_selection_experiment_evidence.json`; retained production remains
the independently gated bitwise/shift conversion above.

The registered scalar subscript hook now transfers one complete owned index
operand to the existing scoped postfix index lowering. It preserves the
original `__getitem__`/`__index__` order and child expression spans; slices,
multiple items and unsupported operands retain their captures. The grammar is
unchanged. Gates passed 66 source tests, 44 expected-output VM cases, and the
11,501 accepted/305 rejected corpus with zero differences. Three indexed
fixtures remove 20 capture nodes per owned subscript (three of four subscripts
owned). Actual indexed CPU measurements are mixed between executable orders;
no speedup is claimed. Zero-subscript controls are excluded from attribution.
See `scalar_subscript_evidence.json`. Chains combining calls or indexes retain their enclosing captures; pure member
chains are now owned as described below. The complete argument-list conversion remains private: its corpus
and 40 VM cases pass, but pydoc regresses in both measurement orders.

The registered postfix Pratt island now constructs complete scalar/name member
chains as common `Member` AST nodes. The grammar preserves trailing trivia,
matching the original postfix spans. The raw lexical base enters the exact
shared ordered member hook in the current scope; the complete chain is then
desugared once. Existing PEG productions remain unchanged, with one unused
operator rule appended. Existing user-rule IDs stay stable; synthetic EOI shifts
one slot. Calls, indexes and grouped bases keep their current adapter.

Final gates passed 68 source tests, 47 expected-output VM cases and the combined
11,501 accepted/305 rejected corpus with zero AST/span/directive/error differences.
The same cached member replay matches after declared/ordinary/declared scope
changes, and deleted-name reads occur once. The numbers fixture owns all 40
member operations, removes 234 capture nodes and expands owned returns from one
to four. The property fixture owns all six member operations, removes 60 capture
nodes and expands owned returns from one to two. Its registered frontend CPU
falls directionally 1.18%/1.42% in both private executable orders. Numeric raw CPU
is mixed; its hybrid/PEG ratio drops from roughly 4–5% overhead to roughly 2%,
without a numeric raw gain claim. Pydoc owns zero members and is excluded from
feature attribution. These measurements do not represent whole DEBUG
compilation or Aster. See `member_pratt_evidence.json`. The out-of-line helper
control was rejected after increasing actual CPU; the retained helper is inlined.

Zero-argument calls in this same real postfix Pratt region remain frozen. The
latest reusable-box and threaded-chain candidates pass 71 source tests, 53
expected-output VM cases and the full combined corpus with zero differences.
Actual eligible call-chain CPU is mixed or regresses, so they are not activated;
see `empty_call_pratt_experiment_evidence.json`. Full argument-list ownership
also remains private because its pydoc CPU regression persists. General call
arguments, function/class/control-flow scaffolding and comprehensions/generators
remain contextual; complete frontend conversion is not claimed.

A private module-facade applicability frontier passes the compatibility corpus
and skips certified-absent walks, but its initial CPU candidate regresses on
numbers. It is not registered. The separately gated latched fact-collection variant
also passes 69 source tests, 47 expected-output VM cases, three real concurrency
fixtures and the full combined corpus with zero differences, but closure/pydoc
CPU regresses and numeric signs change by executable order. Neither candidate
is retained. See `facade_frontier_experiment_evidence.json`.


A frozen simple-assert statement hook passes 71 source tests, including 1,524
budget/depth rollback cases, 53 expected-output VM cases, three real assertion
fixtures and the full combined corpus with zero differences. It transfers one
complete owned condition, normalizes in the live scope, then shares the original
If/not/AssertionError lowering. Message-bearing and unsupported conditions keep
their adapter. Paired CPU fails retention: the tiny assertion fixture worsens relative to its
PEG control, and docstring/caught assertion signs change by executable order.
Assertions remain outside registered ownership. See `assert_ownership_experiment_evidence.json`.


The isolated six-rule expression-anchor lookup guard passes 68 source tests,
47 expected-output VM cases and full combined corpus parity, but CPU is mixed
and it is not retained. Inspection of actual callers corrected the initial
hot-path estimate: the existing walker already collapses transparent wrappers
and most bridge calls already use real anchor rules. See
`anchor_guard_experiment_evidence.json`.


The retained identifier opaque hook now enables its existing certified atomic
fragment in the hybrid matcher. It preserves the outer identifier capture and
finish callback; keyword predicates and character classes need no descendant
owned hooks. PEG/default grammar is unchanged. Independent current-matcher
controls pass 68 source tests, 47 expected-output VM cases and full corpus parity;
the identifier change passes 69 source tests (including 2,794 work/depth rollback
checks), 47 VM cases and the same 11,501 accepted/305 rejected corpus with zero
differences. Scope/import/member/keyword-prefix/continuation/deleted-name cases
remain exact.

Six tighter private pairs show median registered frontend CPU 5.67% lower on
numbers, 3.19% lower on the property fixture and 2.51% lower on pydoc. Numeric
PEG-normalized deltas improve in all six pairs; its median hybrid/PEG ratio moves
from 1.031 to 0.984. Closure remains inconclusive: after six additional
closure-only pairs, the combined twelve-pair registered median is -0.22% and
PEG-normalized median -0.12% (seven improving/five regressing pairs). Registered
range is -5.13% to +16.87%, normalized range -6.00% to +4.28%. This is not a universal frontend or whole
DEBUG/Aster gain claim. Parsing sources match live; the historical emitter
snapshot is required only for coherent frozen compiler/runtime execution gates.
See `identifier_opaque_evidence.json`. The private assert/facade/cache experiments
remain outside production.


The whole scalar-index postfix conversion remains held outside live source.
The broader private call/index corpus subsequently exposed an extra deleted-base
read normalization in the shared index-chain lowering. The prepared patch is
corrected base ordering and added deleted-base cache regressions now pass
70 source tests, 54 historical expected-output VM cases and the full
11,501 accepted/305 rejected corpus with zero differences. Its refreshed
guarded patch remains held during the source freeze. CPU results below describe
the earlier candidate; corrected paired CPU has not been claimed. The prepared three-file
`scalar_index_postfix_ready.patch` adds a postfix grammar operator and common
Index AST/replay construction; source-order scoped index/member lowering remains
at the existing normalization hooks. Calls, slices and complex/grouped index
operands still fall back. Gates pass 70 source tests, 54 expected-output VM cases,
11,501 accepted/305 rejected corpus cases with zero differences, 2,540 additional
work/depth rollback cases, and repeated cached indexes under changing scopes.
Twelve paired private frontend samples show class-index registered CPU median
9.51% lower (all twelve improve), linecache 3.69% lower (ten improve), and list
index 0.97% lower (seven improve, weaker). Zero-index numeric raw timing remains
inconclusive; no index gain is attributed to zero-index controls. These are not
whole DEBUG compiler/Aster results. See
`scalar_index_postfix_experiment_evidence.json` for scope and historical compiler
limitations. The private completion experiment remains separate and unretained.

Private full-call and complex-index Pratt operators now use the existing full
operand grammar with explicit transactional stack contracts and contextual
normalization hooks. The first corpus found one deleted-base AST difference
(11,501 accepted/305 rejected); the private correction passes 71 source tests,
including changing deleted-base scopes and 5,610 malformed work/depth rollback
cases, against the checked semantic-operator core. Corrected candidate and shared-helper control each pass the full corpus
(11,501 accepted/305 rejected, zero differences) and 54 historical VM cases.
Six paired CPU samples reject broader call/index activation: numeric registered
CPU +12.37%, pydoc +7.28%, linecache +6.15%; closure is neutral. An isolated
rollback-safe boundary-shape hook guard is being tested privately to remove
empty speculative operator frames. Neither candidate is activated in production.

The boundary guard passes 72 source tests and 54 historical VM cases but remains
private, without a new full-corpus claim. Six three-way paired samples still
reject activation versus the corrected scalar control: numeric +7.53%, pydoc
+13.21%, linecache +10.51%; closure is neutral and getitem improves 5.78%.
It improves property/list versus unguarded (~5.9%/5.4%) but does not solve the
overall regression. Partial operand capture handoff is being investigated to
avoid replaying entire outer calls after contextual/incomplete operands; that
cost diagnosis is provisional. Terminal results are in
`semantic_call_postfix_experiment_evidence.json`.

Mixed contextual operand recipes now preserve the original child capture anchor
and reuse owned fragment caches during live-scope normalization, avoiding parent
rematching for the measured numeric/pydoc/linecache calls. Explicit argument-frame
ownership and spread-container lowering fixes pass 11,501 accepted/305 rejected
corpus cases with zero differences, 74 source tests and 61 historical VM cases.
The V5 construction variant folds single-argument metadata in place, removes
exact immutable call wrappers and avoids a temporary operation list for plain
identifier callees. Its six paired CPU samples remain mixed: getitem improves
2.28% registered (all six samples), while numeric is +0.11% registered/+2.51%
PEG-normalized and other changes are small. This remains private; neither the
corrected scalar whole-index patch nor broader calls/complex indexes is activated.
The V6 punctuation-boundary observer guard passes 75 source tests, all 61
historical VM cases and the same full corpus with zero differences. It journals
only changes to the actual opening-parenthesis/bracket hint after grammar trivia;
all grammar acceptance, contextual capture retention and rollback remain intact.
Six three-way paired samples show numeric -5.04% raw/-2.10% PEG-normalized versus
V5, but versus scalar control numeric is -1.99% raw/+1.40% normalized and pydoc
-3.28% raw/+2.80% normalized. Several other medians remain mixed, with substantial
raw dispersion; no broad gain or activation is claimed. These are private
opt2/debuginfo0 frontend measurements, not current DEBUG CLI/Aster results.

An untimed actual-dispatch census over the retained frozen matcher separates
executed ordinary/atomic tasks from native-head/prefix event replay and Pratt
FIRST misses. Seven real test files execute 54,963–117,795 hybrid tasks, 11.9–22.8%
fewer than PEG in the same instrumented pin. This is task-count evidence, not
CPU evidence. COMMENT is a larger remaining matching target: 1,682–1,768 atomic
NEWLINE checks exactly match comment characters plus one end check per comment.
Atomic Eval represents 15–23% of all executed tasks. The existing quoted scanner
requires a closing literal, so it cannot certify the current unterminated
prefix-plus-newline-negated COMMENT run. A generic exact scanner proof is being
evaluated by the shared engine owner; no scanner or grammar change is activated.
See `matcher_dispatch_census_evidence.json` for exact counts, pins and limitations.

The private declaration integration now registers six generated/schema-validated
AST bindings for full-rich function/class headers and contextual parameter/base/
body handoff, using the existing transactional arena and original live scopes.
It also collects construction facts without an extra walk: the numeric fixture
delivers 42 function headers/one class and skips 82 of 84 optional nested-class/
nested-function rebuilds. Parameters/bases/bodies still use captures, mandatory
normalization scans remain, and whole-file commit still prevents streaming.
The candidate passes 74 source tests (960 rollback/budget/depth cases), 11,501
accepted/305 rejected canonical corpus cases with zero differences, and 61
historical expected-output VM cases. It remains private: twelve paired frontend
CPU runs regress 0.3–2.2% median registered CPU; PEG-normalized medians also
regress, with wide dispersion. This is isolated opt2/debuginfo0 frontend evidence,
not current DEBUG CLI/Aster compatibility or speed. Ordinary parser output with
bindings disabled remains byte-identical. Frozen evidence and sources are in
`results/testrunner/grameo-python-2026-10-10/owned-declarations/evidence.json`.
A private V2 stores immutable source spans instead of copying header text
through rollback frames. Its 74 source/61 historical VM/full canonical corpus
gates also pass. Six paired runs under recorded external build load regress
7.3–12.7% raw median CPU and 5.3–13.6% PEG-normalized median CPU. It is not
activated. The current core already keeps the hot undo journal at one byte;
typed edit/composition side-vector layout is being audited before another variant.

Private declaration V3 uses the existing separate edit/composition arena types
with compact field inverses and one cursor; only a popped completed frame is
boxed. Edit records shrink 176→104 bytes, but the same 74 source/61 historical
VM/full canonical corpus gates plus six CPU pairs show no runtime gain
(+4.8–17.6% raw, +6.5–18.7% normalized medians under recorded external load).
It remains private. Actual attempted declaration-hook counts nearly equal
completions, disproving failed alternatives as a dominant cost; no position-aware
API or keyword guard was introduced. The six-entry binding scan executes
1,333–4,277 times per fixture; this is a census, not CPU attribution.
See persistent `owned-declarations/v3-evidence.json` and
`owned-declarations/declaration-attempt-evidence.json` under the dated results path.


Private declaration dispatch/normalization attribution (2026-10-10): failed function/class alternatives remain negligible per the actual attempt census. New V4 uses a const mask derived from generated binding registration; V5 restricts irrelevant declaration handler dispatch and observes header operands only while a transactional header is pending. Function/class observation stays unconditional. Both have 75 passing source tests, 61 expected-output execution checks against pinned workspace compilerccc/runtimeacf/ASTfe, and 11,501 accepted/305 rejected canonical AST/span/directive/error comparisons with zero differences. A captured-header normalization control has 72 source tests and the same execution/full parity gates. No live source/core/API/Cargo/registration change.

New same-closure diagnostic CPU after conservative JS-build overlap exclusion favors current V3 and the normalization control, but the new mask/dispatch guards are mixed (numeric dispatch/state +4.28%), so those are not retained. Root correctness-load terminal time remains unconfirmed, and the normalization control includes ordered helper refactoring; no isolated fact/guard cause or whole DEBUG CLI/Aster speedup claim. Exact source guards, controls, timestamps, dispersion and results: `results/testrunner/grameo-python-2026-10-10/owned-declarations/dispatch-facts-evidence.json` and its README. Broader rich declaration header ownership remains private; original contextual parameters/bases/bodies and mandatory scans remain.


Private deferred-expression recipe prototype (2026-10-10): `results/testrunner/grameo-python-2026-10-10/deferred-expression-recipes/`. Matching owns immutable byte-span Atom/Binary recipes for the complete bitwise/shift/add/multiply Pratt family over full rich unary operands. Common AST is constructed and Python operators lowered only at committed live-scope consumption, left-to-right from original captured operand pairs. Calls, members, indexes, powers, comprehensions, await, defaults/decorators/directives and generators retain existing contextual walker semantics. No new operand rematching, shared API/Cargo/live frontend source changes. Explicit eager same-grammar control preserved; ordinary no-optimization generated parser byte-identical.

72 source tests pass (including12,288 depth/work rollback cases and repeated ownership restoration),61 expected-output VM cases pass against pinned workspace closure, and both recipe/eager controls have11,501 accepted+305 rejected full AST/span/directive/error parity with zero differences. Both-order12pair diagnostic favors recipes: retained ratios0.649–0.820, same-rich-eager0.841–0.929, ownPEG0.804–0.839 across4real tests+2complete stdlib files. Cargo CLI compilation ran concurrently under root authorization; this is private opt2/debuginfo0 frontend evidence, not quiet actual DEBUG CLI/Aster. Remains private pending clean repeat/integration; no whole3-second goal claim. Exact hashes/timestamps/dispersion in recipe evidence.json/cpu/cpu.json.


Retained current-core recipe integration (2026-10-10): the registered frontend now uses immutable arithmetic/bitwise recipes over rich unary captures; `parse_with_eager_ast` preserves the eager comparison backend. Actual combined production source passes72 source tests,61 expected-output pinned VM cases and11,501 accepted/305 rejected canonical AST/span/directive/error comparisons with zero differences. Concurrent field-lowering/emitter edits were preserved and audited by source hashes. Ordinary no-optimization parser output remains byte-identical.

The refreshed six-pair-per-order busy-server repeat shows4–14% less frontend CPU versus same-rich eager in both orders across six fixtures. Five improve5–13% versus retained scalar control; closure1–3% is inconclusive. All raw/normalized dispersion and samples are retained, including potential initial correctness-job overlap. These frozen-source opt2/debuginfo0 measurements predate concurrent walker/emitter changes and do not establish current DEBUG CLI/Aster or whole-application speed. Historical18–35% figures are not current retained gains. Evidence: `results/testrunner/grameo-python-2026-10-10/deferred-expression-recipes/current-core/evidence.json`. Whole-file commitment, contextual declarations/statements and rich unary lowering remain; explicit-stack long-chain recipe resolution is the next private refinement.


Checked shared call-prefix integration (2026-10-11): the opted-in RecipeAst path reuses the expression prefix between generator and ordinary call arguments, removing exponential nested-call speculation while preserving the existing PEG grammar. Default compilation and builders without `supports_shared_prefix` are unchanged. Current Program geometry, SourcePratt replacements, observers, scopes, atomic trivia, grammar-stack rollback, finite limits and live capture handles are gated; unsafe handle relocation falls back to replay. The original rich argument productions remain intact.

Live production source tests pass 76/76; repository core/codegen contract suites pass 16/16. Current-source full frontend parity covers 11,537 accepts + 308 identical rejects (11,845 sources), plus 28 focused fixtures; all canonical PEG/direct Modules, spans, directives, scopes and errors match. Both-order development-profile frontend CPU including AST drop gains about 16% on pydoc, 13–20% on inspect, 36–39% on ordinary calls and 15–16% on member calls. Nested depths 8/12/16 improve 11–14x / 86–89x / 761–1143x. Numbers regresses 1–6% (67–230 microseconds); linecache is order-sensitive. Active builds/load average 47–66 are recorded; these are not whole CLI/runtime timings. Exact guards, source/build provenance and all samples are in `results/testrunner/grameo-python-2026-10-10/deferred-expression-recipes/borrowed-call-signature-validation/operand-handles/shared-call-prefix-current/`.
