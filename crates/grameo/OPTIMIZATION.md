# Explicit grammar optimization directives

Ordinary `grameo::compile` treats performance annotations as comments, including
malformed annotations. Enable them explicitly when compiling an optimized
frontend:

```rust
let grammar = grameo::compile_with_options(source, grameo::CompileOptions {
    optimizations: true,
    ..Default::default()
})?;
let parser = grameo_codegen::emit_with_api(
    &grammar, grameo_codegen::ParserApi::HYBRID_CAPTURES,
)?;
```

Only actual line comments carry directives. Text inside grammar literals or block
comments is ordinary grammar text. Enabled malformed directives produce located
grammar diagnostics.

```pest
// @grameo optimize native_choices native_rule_heads
// @grameo pratt sum atom=atom trailing=consume
// @grameo operator sum plus precedence=10 fixity=infix_left
// @grameo operator sum power precedence=20 fixity=infix_right
sum = { product ~ (plus ~ product)* }
product = { atom ~ (power ~ product)? }
atom = @{ ASCII_DIGIT+ }
plus = @{ "+" }
power = @{ "**" }
```

The grammar author supplies and certifies each Pratt table; precedence is never
inferred from PEG alternative order. `trailing=consume` and `trailing=preserve`
select the existing island trivia policy explicitly. Operator fixities are
`prefix`, `postfix`, `infix_left` and `infix_right`; precedence must fit `u16`.
Operator declarations retain their source order within each target. Existing
`bind_pratt` validation rejects unknown, nullable, ambiguous and stack-dependent
operator contracts. Captures continue to use the original PEG structure;
cooperating direct builders select Pratt through their existing hooks.

`native_choices` selects certified first-byte failure dispatch for ordered
choices. `native_rule_heads` selects bounded Sequence-to-Call scheduling recipes.
They can be selected independently. Both retain original observers, semantic
hooks, captures, spans, logical work/depth limits, grammar stack and rollback.
Their ordinary PEG path remains available. `emit_with_native_hybrid_choices`
continues to explicitly select both features for callers using manual bindings.

AST constructor comments remain independently controlled by
`CompileOptions::ast_bindings`. Enable both fields when a frontend uses both AST
bindings and performance directives. See [AST bindings](AST.md) for constructor
schemas and target normalization hooks.

The PHP and JavaScript migration uses exactly their prior Pratt tables and
produces byte-identical generated parsers. Centralizing that configuration adds
no AST ownership coverage or runtime speedup by itself. The complete debug
WordPress performance target requires separate end-to-end evidence.

## Explicit operator separators

An opted-in Pratt declaration can bind an explicit separator after each infix
operator, for source grammars whose newline/comment rule is outside the lexical
operator rule:

```text
// @grameo pratt additive atom=unary trailing=consume operator_trivia=op_nl
```

The separator executes through the ordinary rule engine, including its captures,
scopes, builder callbacks and grammar-stack journal. It may be nullable. A failed
separator restores the infix checkpoint together with reductions and operand
state. Prefix/postfix operators keep their existing behavior. Omitted properties
and existing `bind_pratt` calls retain `None`; normal compilation ignores the
annotation. The equivalent explicit API is
`set_pratt_operator_trivia(target, separator)` after `bind_pratt`.

JavaScript uses this binding with deferred captured-atom recipes. Its complete
AST/span/directive/error gates pass; the measured gain is small and remains
separate from complete DEBUG request performance.

## Explicit ordered-choice token heads

```pest
// @grameo choice_heads statement primary
```

`choice_heads` names normal or silent rules whose root is an ordered choice,
with at most 64 alternatives. Compile with `CompileOptions.optimizations = true`
and generate the hybrid capture API. Default compilation treats the annotation
as a comment and emits the same PEG program as before. Unknown targets, duplicate
targets, unsupported root modes and source-Pratt target conflicts are diagnosed.

The generator proves mandatory ASCII token heads and emits a byte-to-branch mask.
Unknown or unsupported heads always retain their original branch. Matching visits
admitted original branches in their original order; it preserves syntax, capture
identity and successful hook delivery. This is failure admission, not a grammar
rewrite. No helper productions become public parse roots.

Semantic builders additionally opt in through `Builder::supports_choice_heads()`
(default `false`). That permission requires pure selection/checkpoint queries.
Observed prefix rules disable admission; successful callbacks still execute.
Finite work budgets, diagnostic matching, predicates, trivia and unsupported
execution modes use the original PEG path. A depth reservation protects exact
maximum-depth behavior.

Generated immutable programs implement `Program::choice_heads_candidate` and
`Program::choice_heads`, together with `choice_heads_metadata_is_current()` and
`match_observers_are_noop()`. The metadata permission defaults to `false`; custom
wrappers must not forward it unless their instructions, literals, rule modes,
scopes and source-Pratt bindings exactly match the certified program. Observer
wrappers must keep the no-op permission false. Forwarding metadata alone never
enables this shortcut.

The JS frontend explicitly selects `statement` and `primary` and admits it in its
direct-AST builder. The checked DEBUG frontend comparison and complete corpus
evidence are in `proposals/grameo-js-choice-heads-20261011`; CLI/Wasm validation
remains a separate integration gate.


## Shared capture-preserving wrapper prefixes

`// @grameo shared_wrapper_prefix postfix branch=call fallback=read` permits
sharing an ordered-choice prefix when `call = prefix ~ suffix` and `read =
prefix`. Compilation must explicitly enable optimizations, and the semantic
builder must permit this exact wrapper pair. Unannotated grammars and default
builders retain the ordinary PEG path. Captures, rollback, depth limits and
observed hooks retain their existing contracts; finite work budgets use the
ordinary matcher.

The parser reuses per-parse wrapper frames rather than allocating a frame for
each attempt. Generated and resolved immutable programs certify structural
geometry through `Program::shared_wrapper_prefix_metadata_is_current()`.
Custom forwarding programs default to `false` and revalidate the geometry.
This permission does not remove semantic checks: SourcePratt bindings on either
wrapper or a direct call in its prefix still disable sharing, including bindings
registered after the structural certificate was created.

## Direct token compilation and filtered choices

`// @grameo optimize native_tokens` explicitly requests generated Rust recognizers for certified atomic token bodies. Ordinary PEG compilation ignores the comment; the existing public `OptimizationSelection` fields are unchanged. An embedding can inspect the separate `CompiledGrammar::native_token_compilation()` getter. The code generator accepts only acyclic, stack-free, capture-suppressing closures built from ASCII literals/classes, ordered choices, sequences, predicates, and non-nullable repetitions. UTF-8 `ANY` advances exactly one scalar; non-ASCII literals/ranges and other unsupported Unicode operations keep the ordinary matcher. Stateful stack operations, recursive rules, and capture-enabling mode changes also retain ordinary matching. Existing exact quote/delimited certificates keep their specialized scanners.

Direct recognizers retain PEG choice priority and greedy repetition. They run only on the unlimited, non-diagnostic lane with no matcher observers; failures replay ordinary matching. The rule's own capture and semantic begin/finish hooks remain. A semantic builder must explicitly permit opaque token bodies before inner hooks can be omitted. Finite work limits, checked rule-depth limits, diagnostic replay, absolute source offsets, and normal fallback still use their existing contracts. Forwarding `Program` implementations default to an uncertified native-token geometry even if they forward a scanner method.

Existing `// @grameo choice_heads <rule>...` selection can also use positive lookahead to prove a necessary lexical head. For example, `&"$" ~ variable` no longer becomes an unknown first-byte set. Negative and empty lookahead retain the unknown fallback. Candidate filtering checks each omitted alternative's reached semantic rules separately: observed alternatives remain on the original PEG path, while unrelated inert alternatives can be omitted. A possibly successful first alternative retains the original short path. Builders still opt in with `supports_choice_heads`; normal PEG/default builders do not enable this policy.

These are parser-throughput mechanisms, not a native machine-code JIT or a claim about whole-program compilation time. Language AST bindings and compiler-as-service paths require their own complete AST, source-intent, emission, and workload measurements before activation.

### Certified negative token results

```pest
// @grameo optimize native_token_failures
identifier = @{ ASCII_ALPHA ~ ASCII_ALPHANUMERIC* }
```

This separate opt-in implies `native_tokens`. Without
`CompileOptions.optimizations`, the annotation remains a comment.
`CompiledGrammar::native_token_failures()` reports the selected capability; the
existing public `OptimizationSelection` struct is unchanged.

Generated `Program::native_token_result` returns `Some(Some(end))` for a proved
match, `Some(None)` for a proved failure, and `None` when the matcher must retain
PEG execution. Only fully certified atomic rule closures can claim failure.
Unknown rules, insufficient rule depth and non-scalar-aligned starting offsets
return `None`; unsupported or recursive rules keep PEG. The default method adapts
legacy `native_token` successes and never turns its `None` into a failure proof.

The engine consumes negative certificates only on the existing unlimited-work,
non-diagnostic lane with no matcher observers. A failed token leaves the cursor
at entry; ordinary rule, choice, predicate and builder rollback still run.
Finite-work parsing, diagnostic replay, observable programs and non-opaque
semantic hooks retain their ordinary paths. Opaque token permission has the same
meaning as for positive native matches; no directive grants permission to omit
unregistered scope or semantic effects. This is an optional parser algorithm,
not a promise of whole-compiler or application speed.
