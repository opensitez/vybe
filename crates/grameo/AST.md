# Grammar-driven AST construction

## Source loading and request code storage

Optimized grammars can declare storage policy separately from include execution
and name-resolution scope:

```pest
// @grameo code_space request
// @grameo source_unit load target=path selector=require_once loader=source phase=execute repeat=once missing=error scope=caller
```

`code_space` accepts `unit` or `request`. Without the directive the policy is
`unit`. It is interpreted only when grammar optimizations are enabled, such as
`compile_with_options(source, CompileOptions { optimizations: true,
..Default::default() })`. Ordinary PEG compilation ignores these comments,
including malformed optimization comments. Optimized compilation rejects invalid
or duplicate `code_space` declarations.

The compiled grammar and generated parser expose the policy through
`grameo::program::Program::code_space()`, returning
`grameo::source_units::CodeSpace`. Forwarding implementations keep the default
`Unit` unless they explicitly forward this contract.

`Request` tells an embedding that it may retain validated loaded code within one
live VM request. It does not load a file, execute an include during parsing,
change namespaces, or change caller scope, repetition, and missing-file behavior.
The embedding must consume the property and implement its storage policy; a
directive alone does not eliminate compilation or linking. A resident-code
consumer must preserve repeated plain-include execution, `_once` behavior,
source/dependency validity, caller context, and invalidation on request reset,
reload, or VM replacement. It must retain a working fallback for units it cannot
reuse. The grammar contract does not authorize reuse across requests or mutation
of installed code with stale derived caches.

## AST bindings

AST bindings are comments in an ordinary PEG grammar. Other PEG consumers ignore
them; Grameo validates them and its code generator emits static construction
tables and a `Parser::build_bound_ast::<Target>(rule, input)` entry point.

```pest
// @grameo ast program = Block(body: return_statement*)
program = { return_statement* ~ EOI }
// @grameo ast return_statement = Return(value: integer?, span: rule)
return_statement = { "return" ~ integer? ~ ";" }
// @grameo ast integer = Integer(text: source, span: rule)
integer = @{ ASCII_DIGIT+ }
WHITESPACE = _{ " " | "\n" }
```

Fields select constructed child rules, not capture-tree nodes:

| Selector | Value passed to the target |
| --- | --- |
| `integer` | One owned child |
| `integer?` | Optional owned child |
| `statement*` | A vector of owned children in source order |
| `children`, `children?`, `children*` | All immediate semantic children, with the declared cardinality |
| `expression[0]` | The first child constructed by that rule |
| `expression[1]` | The second child constructed by that rule |
| `source` | Borrowed source text for this rule |
| `rule` | This rule's half-open byte span |

`source`, `rule` and `children` are reserved selector names. Selected rules need their own
bindings. Transparent, unbound rules do not produce values. A bound rule must
consume every semantic child exactly once; ambiguous selections and cardinality
errors are explicit build failures, not silent AST corruption. Occurrence indices
are zero-based. Lists may be empty.

## Target integration

Implement `grameo::ast::builder::Target` with:

- `Value`: the target's node enum, such as expression/statement/body/module values.
- `Context` and `prepare`: immutable per-input data, such as one source line index.
- `construct`: a generated numeric constructor ID, span, and owned field arguments
  produce a node plus target-selected inverse metadata.
- `split`: recover the exact original children, flattened in field order, using
  that inverse metadata. This is required for speculative PEG rollback.
- Optional `validate`: reject invalid literals or other semantic construction
  inputs with a located `BuildError` before child ownership is transferred.

Generated `AstConstructor` enum variants identify constructors without runtime
string dispatch. Arguments follow the declaration's field order. Each argument is
`One`, `Optional`, `Many`, `Text`, or `Span`. `split` returns only child node values,
not text/span arguments. Constructors and inverses must not panic or perform
external effects. Lowering that changes operands must retain sufficient inverse
data; do not clone entire ASTs just to implement rollback.

Use `grameo_codegen::generate_with_ast_schema(source, schema)` in build scripts to
check constructor names, field order/names, optional/list cardinalities, and child
value types against target-provided `ConstructorSchema` entries. This provides
target type validation without adding an AST/compiler dependency to Grameo. Plain
`generate` still emits bindings, but only validates grammar-side names and selectors.
Schema generation assigns constructor IDs by schema position, even when a grammar
uses only some constructors. A shared target can therefore serve multiple grammars.

The generated bound AST entry uses semantic matching without constructing a
capture tree. Ordinary `parse`/`parse_pairs` are unchanged. Custom hybrid bindings
can use `AST_BINDINGS` with `AstBuilder` and the existing capture-builder entry;
successfully constructed regions prune their child captures. The generic binding
adapter currently uses PEG matching; custom builders retain the existing Pratt
island API.

## Applying this to a compiler

A class constructor can consume already built members directly; a function can
consume its parameters/body; a loop can consume its condition/body. Targets can
emit common AST directives and declaration metadata in those constructors, and
perform reversible local lowering to primitive operations. Name resolution that
requires later declarations or runtime values still needs a semantic stage.
The parser does not replace that stage with guesses or language-name switches.

`../vybe_ast_bindings` supplies the reusable `CommonAst` target and `SCHEMA` for
Vybe frontends. Generate with that schema, then call:

```rust,ignore
let module = Parser
    .build_bound_ast::<vybe_ast_bindings::CommonAst>(Rule::program, source)?
    .into_module();
```

The generated integration test in `generated_tests/tests/common_ast.rs` constructs
real Vybe modules, namespaces, classes, methods, functions, while loops and returns directly.
The initial shared constructors support parameterless functions, method-only
classes and integer expressions. Module metadata defaults to an anonymous module
with no source-language policy; frontends must supply their own module identity
and directives. Parameters, fields, modifiers, primitive lowering and slot
assignment still require expanded target bindings or semantic processing.
That fixture is not a complete source-language frontend. Production PHP migration
to these bindings is still pending; its existing experimental direct-expression
builder remains available.

## Constructing normalized classes

The shared target also accepts `NormalMethod`, `NormalStaticMethod`, `NormalField`,
`NormalStaticField`, and `NormalClass`. Methods move their bodies into
`NormalMethod` when their rule finishes; fields move their initializers into
`NormalField`. The class reduction routes these completed members straight into
the existing compiler-consumed `NormalClass` buckets. It does not create a raw
`ClassDecl` and run a separate class-member normalization pass.

```pest
// @grameo ast class_decl = NormalClass(name: identifier, members: member*, span: rule)
class_decl = { "class" ~ identifier ~ "{" ~ member* ~ "}" }
// @grameo ast member = NormalMember(value: children)
member = _{ method | field }
// @grameo ast method = NormalMethod(name: identifier, body: block, span: rule)
method = { "function" ~ identifier ~ "(" ~ ")" ~ block }
// @grameo ast field = NormalField(name: identifier, init: integer?, span: rule)
field = { "var" ~ identifier ~ ("=" ~ integer)? ~ ";" }
```

Use the `Name`, `Block`, and expression bindings above for child rules. Parsing a
class root returns `Value::NormalClass(Box<NormalClass>)`; this is a declaration
value, not a `Module` statement. The shared baseline uses public members, unchanged
canonical names, parameterless methods and default class policy. Language targets
must supply richer metadata and protocol roles explicitly rather than infer them
from a language name. The common AST also exposes `from_method_stmt_owned` for
incremental migration of existing function-shaped members without cloning bodies.

Undo retains only member routing order and moves original values back out of the
normalized buckets. Inheritance, trait composition, resolved storage/slot layout,
and scope-dependent calls remain deferred. Ordinary whole-file builds commit
only after successful parse completion.

For a caller-selected stable unit, `engine::build_capture_program_at(program,
entry, full_source, start_offset, options, builder)` runs hooks and opted-in Pratt
islands at an absolute UTF-8 boundary. `CompiledGrammar::build_captures_at` offers
the default-options convenience entry. Generated Parser types implement Program
and can use the engine entry directly. Captures, diagnostics and hook spans refer
to the original full input; SOI/EOI retain full-input semantics. Success commits
that entry once; a later failed unit restores its own initial checkpoint without
discarding prior successful units.

This is a unit-building primitive, not an automatic streaming frontend. The
caller must select a rule that cannot be rolled back by an enclosing alternative,
establish declaration and scope visibility, and handle forward references before
publishing a compilation task or external event. Speculative finish hooks remain
transactional and cannot publish external side effects. Frontends needing
whole-file validation can retain the original whole-file entry.

## Indexed contextual receipts

A custom hybrid `Builder` can request a finalized capture token for selected
observed rules with `observes_capture(rule)` and receive it in
`try_finish_capture`. Both hooks are optional: the default requests no tokens
and delegates to the existing finish policy. Journal tokens with semantic values
so failed speculation restores them. After success, resolve each token with
`root.resolve_capture(token)` rather than walking the subtree to find its rule.

Resolution validates the input identity, capture index, rule, span and anchor
bounds. Receipts must stay attached to that successful build; they are not
persistent or cross-parse cache keys. Keep any captures the consumer still needs,
and retain the original contextual path when a token is absent or invalid.
PHP uses this for declaration discovery while preserving its namespace,
by-reference, scope and directive processing. This removes discovery traversals;
it does not eliminate all contextual normalization.


## Failed-prefix hook permission

The opt-in `nullable_failure_heads` matcher optimization may skip a proven failed
prefix only when Program observers are explicitly inert and every reached Builder
hook is either unobserved or explicitly covered by
`Builder::failure_prefix_hooks_are_inert(rule)`. The latter returns `false` by
default, including through borrowed builders. Grant it only when skipping hooks
on the certified failed prefix cannot suppress a checked error, change semantic
state, or cause an externally visible difference. A finish-only capture receipt
hook can meet this contract when the certified failed prefix never finishes that
rule. Successful matches still execute their hooks. Capture observation alone
never grants permission; existing builders and PEG grammars need no changes.


## Owned prefix and postfix composition

A hybrid Builder can move an operand into its target AST with
`MetadataArena::unary_with_metadata(child, make, split)`. `make` returns the parent
and binding-selected rollback metadata; `split` must recover the exact original
operand from them. This supports prefix/postfix reductions that replace raw spans
or contextual flags without cloning a subtree or storing undo metadata inside the
committed AST. Like other owned arena compositions, the operand is unavailable
until rollback restores it; commit discards inverse metadata while preserving the
owned parent. Constructors and inverses must not panic or mutate external state.
This API is opt-in and does not change matching, PEG captures, or existing hooks.

## Checked shared expression prefixes

An optimized grammar can avoid parsing the same expression twice in a choice:

```pest
// @grameo shared_prefix call_arg branch=generator_arg prefix=expression
call_arg = { keyword_arg | generator_arg | expression }
generator_arg = { expression ~ generator_clause+ }
```

The directive is a comment unless compilation enables optimizations. A hybrid
Builder must also explicitly grant `supports_shared_prefix(generator_arg)`;
the default is `false`, including through borrowed builders. Ordinary PEG parsing
and builders without this grant keep their original execution and capture shape.

The compiler checks the exact named prefix and fallback, a normal unobserved
branch wrapper, and matching scopes. The engine rechecks the current Program
geometry and declines forwarded Pratt replacements and active match observers.
The original suffix sequence keeps its trivia and atomicity behavior. A failed
suffix retains the successful prefix only when removing the wrapper cannot
invalidate capture handles; otherwise it rolls back and replays the ordinary
fallback. Finite work budgets also retain the ordinary path. Depth limits, errors,
spans, grammar stack state and capture structure remain equivalent.

Grant this hook only when the builder's transactional prefix semantics permit
reuse across the failed branch and fallback. It does not permit arbitrary prefix
factoring, wrapper lifting, or changes to scope/directive processing. Python's
registered RecipeAst path uses it for generator-versus-ordinary call arguments;
its PEG control and rich argument productions remain intact.

### Capture-preserving shared wrapper prefixes

`// @grameo shared_wrapper_prefix target branch=call_rule fallback=read_rule` selects a direct `call_rule | read_rule` pair within the target's ordered-choice spine. The branch body must be `prefix ~ suffix`, and the fallback body must structurally equal that prefix. All three rules must be normal rules in the same scope; designated whitespace/comment wrappers are rejected. Unsupported shapes produce a source-spanned G015 diagnostic when optimizations are enabled. Ordinary PEG compilation ignores the directive.

A builder must explicitly grant `supports_shared_wrapper_prefix` or its source-aware variant. Both wrappers must be unobserved. The matcher preserves original suffix/trivia sequencing, rule depth, descendant capture indices and handles; suffix failure retains the prefix and retags only its inert wrapper. Changed live IR, wrapper scopes/modes, designated trivia, conflicting Pratt bodies, Program observers, predicates/trivia contexts and finite work budgets retain ordinary matching. This mechanism is language-independent; adding framework support does not activate a language grammar.

### Checked generator versus parenthesized grouping

```pest
// @grameo shared_group_prefix primary branch=generator fallback=expression_list prefix=expression
primary = { "(" ~ generator ~ ")" | "(" ~ expression_list? ~ ")" | other }
generator = { expression ~ generator_clause+ }
expression_list = { expression ~ ("," ~ expression)* ~ ","? }
```

This explicit optimization requires an additional pure Builder grant,
`supports_shared_group_prefix(generator, expression_list)`, whose default is
false. Original productions stay unchanged. Default PEG compilation ignores the
directive; ordinary captures and unregistered builders use original matching.
The checked shape shares only the first successful expression. A failed generator
suffix restores its prefix checkpoint and retags the inert wrapper for the list
alternative; comma tails and closing delimiters still run the original expressions.
Empty groups and failed first expressions use the original fallback.

Both wrappers must be normal, unobserved (including capture hooks), and in the
same scope. The matcher checks current Program geometry, modes, delimiters and
list tails; conflicting wrapper/target Pratt overrides, active observers,
predicates/trivia contexts and finite work budgets retain ordinary matching.
Rule depth, grammar-stack effects, capture indices, issued child handles, errors,
spans and transactional semantic state are preserved. Wrapper captures finish
before delimiter trivia, as on PEG. Successful first-expression semantics must
permit reuse rather than rollback/replay; this is part of the explicit grant.
Bound source-intent adapters decline by default. Python registers the grant on its
RecipeAst route without changing PEG bodies or scope/directive lowering.

### Owned capture boundaries

```pest
// @grameo owned_capture_boundary owned_expression
owned_expression = { primary ~ suffix* }
```

This directive is interpreted only when `CompileOptions.optimizations` is
enabled. Ordinary PEG compilation ignores it, including malformed directives.
The selected rule must be normal, non-trivia and have no Pratt override. The
compiler records its original body, scope and complete descendant rule closure,
including scoped trivia and Pratt entries reached through child rules.

An AST builder must explicitly grant `supports_owned_capture_boundary(rule)`
and return a pure exact end prediction from `owned_capture_end`. It must own the
complete original AST at this boundary; omitted descendants cannot subsequently
be recovered from captures. Semantic begin/finish hooks still run and must
remain transactional. The original root capture is retained, including its
capture token, while descendant captures are suppressed before allocation.
Any descendant capture observer prevents this optimization.

The matcher still executes the original grammar body with unchanged rule-depth
limits. If the predicted end differs from the actual successful match, it rolls
back to the post-begin checkpoint and replays the body with captures, without
duplicating the root hooks. Checked construction failures roll back normally.
Finite work budgets, diagnostic matching, predicates, trivia, atomic/silent
rules and non-no-op Program observers keep ordinary matching. Program wrappers
and builders grant no permission by default. A language must verify full AST,
scope, directive, error and DEBUG CPU parity before enabling its producer.


### Checked intra-rule list and slice prefixes

```pest
// @grameo shared_inner_prefix list_inner prefix=expression kind=list
// @grameo shared_inner_prefix subscript_item prefix=expression kind=index
list_inner = { expression ~ clause+ | (expression ~ ("," ~ expression)*)? ~ ","? }
subscript_item = { expression? ~ ":" ~ expression? ~ (":" ~ expression?)? | expression }
```

These directives require `CompileOptions.optimizations` and a pure explicit
`Builder::supports_shared_inner_prefix(target)` grant, which defaults to false.
Original PEG productions remain authoritative. Unsupported shapes are rejected
only in optimized compilation. Default PEG parsing, unregistered builders,
Program observers, predicates/trivia and finite work budgets retain original
execution.

The matcher revalidates the target, choice, calls, repetitions, delimiters,
normal rule mode, scope, trivia role and target Pratt policy against the current
Program. The identical prefix call can itself use an admitted Pratt island.
A successful nonempty expression is retained at its original capture address
and semantic checkpoint. List-comprehension suffix failure reuses that expression
with the original comma and optional trailing-comma tails. Slice colon failure
retains it as the ordinary index expression; successful slices continue through
the original optional upper bound and step. Empty or failed prefixes run the
original alternatives. No capture wrapper is removed, retagged or relocated.
Rule depth, grammar-stack state, issued handles, spans and diagnostics remain
unchanged.

The grant certifies deterministic prefix hooks and permits retaining successful
prefix state instead of rolling it back and replaying it. It cannot publish
speculative external effects. Unbound source-unit adapters delegate the owner's
grant; bound collectors decline. Python registers this on RecipeAst while keeping
its PEG bodies, contextual AST lowering, scopes and directives unchanged.
