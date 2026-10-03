# Vybe grammar compiler and parser

Independent implementation of [grammarplan.md](../../grammarplan.md). Existing language defaults remain unchanged. Production uses Rust and pest-compatible grammar text; it has no pest dependency and requires no JS/Node tooling.

The core compiles grammar definitions, resolves dense IDs, and analyzes nullable/effect/recursion/progress facts. An iterative engine matches source with ordered choices, all rule modes, implicit trivia, Unicode XID, transactional grammar stacks, work/depth guards, and compact borrowed captures. `parse_source` builds captures; `recognize` omits them.

`codegen` emits typed Rust rule enums and static programs without runtime grammar loading or resolution. It currently shares the iterative engine rather than emitting specialized matching functions. `engine::build_program` instead executes transactional semantic hooks with no capture arena. A generated fixture builds actual `vybe_ast` return statements while matching and checks rollback; full language AST bindings are still pending.

The independent source-expression Pratt module has typed transactional builders, explicit operator precedence/associativity, groups and resource limits. A generated token grammar feeds it into common-AST expression nodes with source spans. Automatic island lowering and call/index/ternary integration remain pending. Certified ASCII whitespace has a bounded byte-set fast path; uncertain/effectful trivia retains the full engine.

`CompiledGrammar::bind_pratt` also binds existing atom/operator rule names directly to a source-expression island. No token vector is required. Bindings include explicit trailing trivia behavior and preserve ordered operator retries after failed operands. `Builder::supports_pratt` opts in and `reduce` builds semantic nodes; capture parsing retains the original grammar. Authors must certify the bindings match their expression language. Generated Rust preserves the binding tables. Tests cover native common-AST construction, rollback, and an isolated Lua profile; production frontends have not switched.

Checked builder hooks (`try_begin`, `try_finish`, `try_reduce`) report `BuildFailure` for semantic construction errors and restore the initial checkpoint. Errors are fatal for that build, including inside speculative syntax; bindings must defer contextual validation where appropriate. Existing infallible hooks remain supported. Atomic repetitions of a single ASCII builtin use a byte scanner with identical work-budget accounting and diagnostics; other repetitions retain the full engine.

`source::Index` converts byte offsets and zero-based UTF-8/UTF-16/scalar positions, with explicit scalar/surrogate/CRLF errors and sparse checkpoints for long lines. This is the coordinate foundation used by the basic editor session.

The package is its own Cargo workspace. Default tests build only the core and `unicode-ident`; `codegen` is an optional workspace member. Neither path builds the compiler, language plugins or VM. Optional conformance/generated checks have separate manifests and build directories.

From the repository root:

```sh
# Fast core loop
cargo test --manifest-path crates/vybe_parser/Cargo.toml --offline
cargo run --manifest-path crates/vybe_parser/Cargo.toml --offline --bin grammarcheck -- languages/*/src/grammar.pest
# Core plus Rust generation correctness
cargo test --manifest-path crates/vybe_parser/Cargo.toml --workspace --offline
# Black-box acceptance/capture comparisons with test-only reference parser
cargo test --manifest-path crates/vybe_parser/conformance_tests/Cargo.toml --offline
# Generated matcher parity and direct-common-AST fixture
cargo test --manifest-path crates/vybe_parser/generated_tests/Cargo.toml --offline
# Compile generated Rust for every repository grammar
cargo test --manifest-path crates/vybe_parser/generated_tests/Cargo.toml --offline --features all-grammars
# Local performance baseline (no timing assertions in tests)
cargo run --manifest-path crates/vybe_parser/conformance_tests/Cargo.toml --offline --release --example baseline
```

`grammarcheck` reports grammar errors and analysis warnings with byte spans rendered as one-based Unicode-scalar line/columns. Its timing includes syntax, resolution and analysis. CLI columns use Unicode scalars; `source::Index` provides LSP UTF-16 conversion separately. Parent Cargo configuration may print unused-patch warnings; those crates are not built.

`parse` returns located `GrammarSyntax`; `compile` resolves it and rejects proven analysis errors. `resolve` is the lower-level API retaining analysis errors for inspection. Uncertain findings are warnings backed by runtime guards. Ordinary corpus tests compile all 18 repository grammars without importing the language crates.

Remaining: broader Unicode-property/extension support (tags currently report an unsupported-feature error), specialized generated matching and lexical dispatch, automatic Pratt-island lowering, complete common-AST bindings, declarative module loading, richer recovery/REPL/LSP sessions, language migration and end-to-end performance gates. The current iterative matcher remains several times slower than the reference on the local Lua baseline; replacing syntax tooling alone has not achieved the performance objective.

No pest or Tree-sitter source was inspected or reused. Only documented syntax and repository grammar files are implementation inputs.

Format packages explicitly with `cargo fmt -p vybe_parser -p vybe_parser_codegen --manifest-path crates/vybe_parser/Cargo.toml`. For optional packages, select `-p conformance_tests` or `-p vybe_parser_generated_tests`. Avoid `--all` there: Cargo follows the AST path dependency into its outer workspace.

`compat` provides an optional owned typed-pair view for walker migration; generated `Parser::parse_pairs` exposes rule enums (including EOI), child iteration, text, spans and cheap handle cloning. The compatibility position method follows the one-based scalar/LF convention; indexed LSP positions use `source::Index` instead. Broader walker API compatibility still needs inventory and migration checks.

`arena::Arena<T>` journals semantic node/value handles without requiring AST nodes to implement Clone. `editor::Session` applies revision-checked edits and produces immutable source/index/capture reports with document/revision-scoped IDs and complete-document diagnostics. It is currently a full-reparse correctness baseline; incremental reuse/recovery remains pending.

`arena::OwnedArena<T>` moves owned AST children into parents and uses explicit composition inverses to recover them on rollback. It returns the owned root without cloning or a final AST walk. The generated native-Pratt fixture uses this with real `vybe_ast::Expression` nodes; inverse functions and exclusive child ownership are binding contracts.

Generate Rust directly with `cargo run --manifest-path crates/vybe_parser/Cargo.toml -p vybe_parser_codegen --bin grammargen --offline -- input.grammar output.rs`. Build scripts can call `vybe_parser_codegen::generate` and write the module into OUT_DIR. Unchanged CLI output preserves its timestamp.

`modules::compile` links ordinary grammar sources using explicit Rust import/export metadata. Rules retain module-local trivia; local names, exported entry aliases and cross-file diagnostics are preserved. `vybe_parser_codegen::generate_modules` emits the same scoped program. Declarative manifest loading/reexports/package resolution are still pending. Existing standalone grammars need no changes.
