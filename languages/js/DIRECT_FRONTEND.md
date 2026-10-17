# JavaScript direct frontend coverage

Registered `parse` and source-only `parse_eval` now use the journaled owned
frontend followed by the existing contextual and module normalization pipeline.
`parse_direct_bodies` exposes that same route. Explicit `parse_peg` and
`parse_source_peg` entries preserve golden references. Shared Grameo and
unmodified languages remain PEG by default. This is an activated hybrid frontend;
full conversion of every syntax variant remains unfinished.

The JS grammar explicitly selects `nullable_failure_heads`. Grameo codegen can
then certify a failed rule whose silent nullable prefix consumes nothing. The
selected recipe builder grants permission to skip that unobserved failed path;
ordinary PEG parsing and builders without that permission retain the original
execution. Certificates preserve ordered expectations, logical work and depth
limits, resolved trivia scopes, and every observed rule hook. Unicode, predicates,
stack operations and unknown or potentially consuming paths use normal matching.
The current JS certificates cover failed call chains and class declaration heads.
Full frontend measurements include construction and AST destruction; these are
not measurements of the CLI or VM. Evidence is recorded in
`results/testrunner/grameo-js-2026-10-07/hybrid/js_nullable_failure_prefix_20261010`.

| Production path | Owned coverage in body mode | Contextual fallback / remaining work |
| --- | --- | --- |
| Scalars and arithmetic | ASCII identifiers, ordinary numbers/quoted strings; this/super/boolean/null/undefined leaves; additive/multiplicative/exponentiation over owned or deferred contextual unary atoms; direct -, +, !, ~, typeof and void over complete owned operands | BigInt, globalThis, await/delete/increments, grouping, optional/new/template/regex/object/function atoms still lower from their original captures; incomplete all-owned trailing-comment islands retain PEG fallback |
| Expressions | Exact expression/initializer wrappers; comparison, equality, shifts, in/instanceof, logical/nullish and bitwise chains over certified scalar/arithmetic/call operands | Contextual typeof equality normalization, assignments, conditional/comma, constructors, optional chains, templates, regex, objects, named/generator functions and arrows |
| Ordinary chains | Complete ordinary member/index/call chains over owned primaries; empty calls, owned positional arguments, trailing comma, nested chains, canonical `.length` lowering; exact enclosing arithmetic operand spans | Optional chains, spread arguments, constructors, tagged templates, unsupported children, and special Array/Intl/concat/regex/Symbol-index call lowering |
| Arrays | Complete ordinary and spread elements over owned expressions, exact holes/elisions, nested arrays, trailing commas | Any unsupported element keeps contextual construction; key/value and hole marker policies remain exact |
| Function expressions | Anonymous sync/async nongenerator functions with owned identifier/rest/default parameters and complete bodies | Named function IIFE binding normalization, generators/yield discovery, destructuring/prologues and unsupported bodies remain contextual |
| Statement lists | Complete function bodies and nested block statements, without lexical body preselection | Any list containing an unowned statement retains captures and uses contextual construction |
| Lexical declarations | `var`/`let`/`const`, multiple identifier declarators, owned initializers, omitted initializer where allowed | Destructuring, unsupported initializers; missing const initializer preserves the contextual diagnostic |
| Function declarations | Sync/async named non-generator declarations, identifier/rest parameters with owned defaults, complete owned bodies | Destructuring/prologues, unsupported defaults/bodies, explicit generators and every positive literal yield spelling retain scope-aware generator discovery |
| Control flow | If/else, while, do/while, throw, implicit/labeled break and continue | For/for-in/for-of/for-await, switch, labels |
| Exception scopes | Complete try/catch/finally with optional identifier catch binding | Destructuring catch and any incomplete branch; block-body consumers share the owned cache |
| Classes | Owned method body statements and parameter initializer islands can be reused by the existing class walker | Class/member/receiver/protocol normalization, decorators, computed names, private validation and all class declaration construction |
| Resource and module declarations | Owned descendants can be reused | Using/await-using lowering, imports/exports and declaration policy construction |
| Directives and context | Original source spans, strict string statements and module directive record preserved | Source-prologue classification, prelude insertion, hoisting, static TDZ, optional-chain counters, using/generator/class/model normalization retain their existing order |

A region frame certifies the actual grammar-produced owned child kinds; it does
not replace grammar matching. Unsupported children publish gaps, so parents do
not guess that a partial body is complete. Reductions move existing common AST
children and journal their original routing and metadata. Rollback restores
operands, names, default parameters, declaration order and block spans.

An incomplete parent can still consume a complete parameter/declarator/catch
region. Those carriers keep the initializer's original lexical/anchor metadata
or the block metadata. On cache handoff, unconsumed headers publish the existing
owned initializer/block under its captured anchor, keeping fallback captures
usable without a second parse. This is necessary even when the enclosing body
will use the contextual walker.

For fully ready owned roots, first cache use transfers the statement/body/tree and reads compact
lexical/model facts collected at the existing reductions. It performs no replay
recipe allocation or AST traversal. Only an actual repeat read recaptures that
certified rule with `grameo::engine::parse_program_at`, retaining the full original
source and absolute offsets, then constructs it with a fresh contextual walker.
Source allocation identity remains mandatory; scalar repeat reads still decode
only their certified lexical ranges. Repeated-read tests include nonzero offsets,
Unicode prefixes, exact spans and facts. Registered parsing never recursively
calls its own frontend for fallback.

Immutable per-rule frame tokens are allocated before the arena's initial commit
and reused on its journaled stack, avoiding a full Value slot at every rule entry.
Rollback cannot invalidate those IDs. Statement caches transfer their existing
arena Box rather than copying the whole statement into a HashMap slot.

Newly owned calls certify their enclosing arithmetic wrapper's exact span and
rule even when the old scalar/arithmetic preselection was inactive. This keeps
trailing trivia on comparison operands and supplies the correct rule for repeat
recapture. Capture suppression requires that exact certified anchor; a complete
child alone cannot stand in for an enclosing wrapper with a different span.

Primary construction is enabled for eligible identifier callees. Arithmetic
already owns its scalar leaves; primaries inside unsupported member, unary and
container paths keep their captures. A permissive prefix filter only decides
whether to attempt ownership, while grammar-produced children certify every
completed call. An immutable ignored-call token keeps nesting and rollback exact.

A complete conversion still needs direct reducers and exact inverse recipes for
the remaining expression variants, destructuring/for/using, full callable and
class policies. Grameo's current builder receives rule IDs and source spans,
not a borrowed completed capture view; migration of those contextual walkers
requires explicit language reductions or a shared contextual capture handoff.
Removing the remaining captures does not itself eliminate the required ordered
module normalizations or prove a complete compilation/runtime speedup.

The class-member normalization registry currently borrows its inputs. Shared
owned method/constructor/base-call helpers are available, but class/member
callbacks still need to use them while preserving modifiers, defaults, receiver
and protocol policies.

Ordinary property carriers retain source spans until the enclosing chain owns
them, then allocate names once. Member carriers reuse the property slot through
a small reversible edit. Array item carriers similarly reuse expression slots;
spread and elision routes retain exact inverses, while complete array roots move
the element buffer. Anonymous function expressions move complete parameter and
body vectors into a common FunctionExpr, preserving the default inner statement
span and the original outer expression span.

Before contextual arithmetic recipes, source-only/frontend evidence was 49 tests and two complete 10,586-file
gates (10,565 accepted, 21 identical rejects) with zero AST/span/directive/error
differences. The public API gate includes prelude insertion; its isolated
frontend wrapper does not prove linked compiler/runtime execution. Full hybrid
CPU remains about 1.51–1.56 times PEG on the five selected WordPress JavaScript
files in the latest matched function-expression run. Conversion is active and
compatibility-gated, but the intended frontend speedup is still unfinished.

Exact single-child identity wrappers now use the shared arena's reversible
`remove_value_below_top` operation instead of popping and re-pushing the owned
child. Matching frame identity, child kind, source bounds and trivia are checked
before removal. Exact no-op span/anchor edits are skipped. The matched routing
change measured -3.2%/-3.0% CPU on hooks/plugins, -1.5% on image, approximately
neutral on style, and +0.5% on media; this is a local reduction, not the intended
whole frontend speedup. A test-only retained journal verifies three successful
parse rollback cycles; normal parsing still commits normally.

Frozen sources, complete gates, CPU comparisons and coverage profiles are in
[chain/array/function evidence](../../results/testrunner/grameo-js-2026-10-07/hybrid/chain_arrays_functions_20261008/README.md).

The retained shared matcher task layout reduces wide/compact/stack-free hybrid
continuations from 56/48/40 to 48/40/32 bytes without narrowing capture or semantic
cursors. All 14 core unit, 126 integration and 49 frontend tests pass; complete corpus
parity remains unchanged. Matched 120-pair CPU improves the no-op hybrid by
1.5–4% and four full frontend inputs by 0.7–3.1%, with hooks neutral. Full hybrid
still ranges about 1.44–1.60 times PEG across these runs. The linked CLI smoke evidence
preceded this layout and does not validate the newly linked layout yet.
[Shared task layout evidence](../../results/testrunner/grameo-js-2026-10-07/hybrid/shared_task_layout_20261008/README.md)
includes the exact retained sources and negative getter/atomic-guard probes.

Later [compact carrier](../../results/testrunner/grameo-js-2026-10-07/hybrid/compact_carrier_probe_20261008/README.md)
and [broader arithmetic operand selection](../../results/testrunner/grameo-js-2026-10-07/hybrid/chained_pratt_probe_20261008/README.md)
probes passed complete compatibility gates but produced mixed CPU results, so
neither is retained. Diagnostic phase attribution and sampling locate most of
the remaining cost in matching/builder work rather than duplicate module passes
or cache handoff. The intended major frontend speedup is not complete.

JavaScript compilation explicitly enables grammar optimization directives with
`compile_with_options`; `emit_with_api` reads their selected native features.
The shared matcher fuses certified impossible left branches at 37 native ordered
choice sites, preserving complete observer order and logical limits. Ordinary
code generation and PEG reference parsing retain their existing dispatch. The
new path passes the complete compatibility gates; 120 matched alternating pairs
show 1.5–3.8% less isolated parser CPU on the five selected inputs. This does not
prove complete `target/debug/vybex` WordPress execution under the user's 3-second
target, and linked runtime validation of this newly generated route remains
separate. [Native choice evidence](../../results/testrunner/grameo-js-2026-10-07/hybrid/native_choice_fusion_20261009/README.md)
records the retained implementation, exact scope and remaining performance gap.

The certified left-failure proof also accepts sequences whose left child is
already certified. The right child never runs on that byte miss; original
sequence work, depth, observer order and rollback remain exact. This increases
JavaScript eligible choice sites from 37 to 40. Two incremental 120-pair runs
show 0.4–2.3% and 0.9–2.7% lower CPU on the five frozen fixtures. These isolated
results do not establish the complete debug WordPress target.

The explicit native generator also emits bounded left-sequence rule-head recipes:
33 JavaScript rule heads (30 with sequences). The matcher preserves real
FinishRule/FinishExpr/Sequence frames, observes every original expression/rule,
and charges each original logical Eval/Enter transition before executing it.
Wrapper hooks, captures, spans, scope, suffix/trivia matching and rollback remain
active. Two incremental 120-pair runs show 0.7–4.0% and 0.2–2.7% less CPU on the
five frozen fixtures; the smallest differences are within control variation.
PEG defaults remain unchanged. PHP's 79 recipes have representative matcher
gates; full registered PHP AST/CPU gates remain a separate requirement.

Pratt metadata now lives in `src/grammar.pest` through `@grameo pratt` and
`@grameo operator` line comments. `@grameo optimize` selects native choices and
rule heads. Both AST constructor annotations and performance directives require
explicit, independent compile options; ordinary compilation ignores them, even
when malformed. The migrated generated JS parser is byte-identical to the
previous explicit Rust Pratt/native bindings. This change centralizes grammar
configuration; it adds no AST ownership coverage or measured runtime speedup.

## Retained contextual arithmetic recipes (2026-10-09)

Body mode now selects the arithmetic Pratt binding over complete grammar unary
atoms. A binary recipe moves ready operands and retains source/capture references
for contextual operands. On first use it temporarily removes its cache root,
indexes the retained unary captures once, and lowers each contextual atom through
the caller's live JsWalker scope. Optional-chain normalization receives the
certified operand span before normalization; typeof equality keeps its original
contextual fold. Actual repeat reads recapture only the certified arithmetic
rule. Statement/body ownership still requires complete ready children, so this
is partial expression conversion rather than complete statement/class conversion.

`operator_trivia=op_nl` executes JavaScript's real explicit newline rule between
an infix operator and its next atom. Accepted programs use no whole-source
syntax replay. Rejected syntax obtains the ordinary PEG diagnostic ordering.
Ordinary shared PEG parsing and explicit JS PEG references remain unchanged.

The exact retained bounded v2 passes 52 source tests, 24,576 small-budget rollback
cases, and the complete 10,586-file gate (10,565 accepted, 21 identical rejects,
zero AST/span/directive/error differences). Thirteen private coherent
compiler/VM fixtures agree with PEG, including captured object operands, local
shadowing and operator newlines. A linked DEBUG CLI gate remains the parent
integration step.

Two 60-pair probe orders give normalized hybrid/previous-hybrid CPU ratios:
hooks 0.980/0.972, jquery.form 0.975/0.985, image 0.991/0.984, draggable
1.002/1.002. The last result shows no gain. These are small local improvements;
full hybrid remains roughly 1.4–1.5 times PEG on these probes. Later comment-proof
variants remain private because their CPU signs vary by probe.

The prototype census found 25 contextual binary roots / 40 operators and 34
captured atoms on draggable, three roots on jquery.form, one on image, and none
on hooks. Captured atoms still use contextual lowering; root/operator counts do
not establish a complete frontend conversion or complete request speedup.

### Whole-expression Pratt matching (2026-10-09)

The registered hybrid now explicitly binds `logical_expr` over unary atoms, with the existing flat logical/bitwise and comparison/equality/shift levels followed by arithmetic levels. Ordinary PEG productions and default Grameo compilation remain unchanged. Captured contextual atoms lower in live scope; repeated use recaptures the certified owner rule. `typeof` strict-equality normalization, source spans and rollback remain exact.

This retained extension passes54source tests,36,864additional rollback budgets,10,586corpus AST/span/directive/error comparisons and21compiler/VM fixtures. Four production-file probes show3.5–6.4% isolated frontend CPU reduction in both orders; hybrid remains approximately1.40–1.55times PEG. Diagnostic rule queries fall9.2–10.2%, begin callbacks14.5–15.5%, and retained captures14.5–17.1%. Contextual and unsupported regions still retain captures; body-cache coverage is zero on these probes. These are isolated dev-setting results, not a complete DEBUG request speed claim. Exact evidence: `results/testrunner/grameo-js-2026-10-07/hybrid/js_whole_expression_pratt_20261009`.
