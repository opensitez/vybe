# Vybe parser: independent grammar frontend

First implementation milestone of [`grammarplan.md`](../../grammarplan.md): compile existing pest-compatible grammar **definitions** into a located, reference-resolved arena IR. This does not yet parse guest language source, generate Rust parsers, construct the common AST, or implement editor recovery.

The package is its own Cargo workspace, with no dependencies. Ordinary tests do not build pest, language plugins, the compiler, or the VM. The parent workspace excludes this directory; existing language parsers remain unchanged.

From the repository root:

```sh
cargo test --manifest-path crates/vybe_parser/Cargo.toml --offline
cargo run --manifest-path crates/vybe_parser/Cargo.toml --offline --bin grammarcheck -- languages/*/src/grammar.pest
cargo tree --manifest-path crates/vybe_parser/Cargo.toml --offline
```

`grammarcheck` reads one or more UTF-8 grammar files, reports rules/expression counts and syntax-plus-resolution time, and exits nonzero on errors. Its timing is a development aid, not a benchmark or a speed claim relative to pest. Parent `.cargo/config.toml` overrides may produce unused-patch warnings; those crates are not dependencies and are not compiled.

## Implemented

- Native Rust grammar lexer, comments, decoded string/character escapes and byte locations.
- All five rule modes; ordered choice, sequence, groups, predicates, repeats and bounded repeats.
- Pratt parsing of **grammar expressions**, preserving postfix/predicate/sequence/choice precedence. Source-language Pratt parsing comes later.
- Structured `PUSH`, `PUSH_LITERAL`, and indexed `PEEK`; ordinary stack builtin references.
- Dense rule IDs and builtin resolution, forward references, duplicate/undefined-rule diagnostics.
- Topologically ordered expression arena; bounded repeats do not expand into copies.
- Grammar nesting limit and iterative flat chains; deterministic independent parse contexts.
- Repository corpus test reading all 18 grammar files without importing language crates.

## API

`parse` / `parse_with_options` return `GrammarSyntax` with rules, source ranges, and expression IDs. `compile` adds name resolution and returns `CompiledGrammar`; `syntax()`, `rule_id()` and `reference()` expose the located IR and resolved symbols. Grammar parsing reports the first lexical/syntactic error; resolution collects duplicates and unresolved references with related locations.

All locations are half-open UTF-8 byte ranges in the input grammar. Diagnostic rendering reports one-based lines and Unicode-scalar columns; this is not yet an LSP position conversion API.

## Deliberate remaining work

This milestone covers the constructs and builtin names used by the repository, not the entire pest extension/Unicode-property catalog. Tags currently produce an explicit unsupported-feature diagnostic. Parsing a stack operation does not implement its source-matching behavior. Compilation does not yet validate left recursion, progress/nullable repetitions, or certify that a grammar is safe to execute.

Next: specify execution/mode/stack rollback contracts, add conservative grammar analysis, and implement the reference recognition engine with conformance fixtures. The generated performance backend, source-language Pratt islands, direct AST builders and LSP sessions follow the plan. Keep those stages independently testable here.

Implementation inputs are the documented pest syntax and Vybe's grammars. No pest or Tree-sitter source is used; no JavaScript/Node grammar tooling is required.
