# Grameo codegen

Generate Rust parsers from PEG grammar files. Generated parsers use the `grameo` runtime and do not compile grammars at runtime.

```sh
cargo install grameo_codegen
grammargen input.pest output.rs
```

For build scripts, add `grameo_codegen = "0.1"` to `[build-dependencies]`, call `grameo_codegen::generate(&source)`, and write the returned Rust module to `OUT_DIR`. Add `grameo = "0.1"` to `[dependencies]` and include the generated module in your code. Its `Parser` exposes `parse` and `recognize` methods accepting the generated `Rule` enum.

`// @grameo ast` annotations also generate `AST_BINDINGS`, an `AstConstructor` enum and `build_bound_ast::<Target>`. Use `generate_with_ast_schema` to validate target constructor signatures before generation.

Opted-in hybrid frontends can select `// @grameo choice_heads rule_name ...`
for conservative token-head admission at ordered-choice roots. Enable
`CompileOptions.optimizations` and the cooperating builder permission; ordinary
PEG generation treats it as a comment. See [optimization directives](../OPTIMIZATION.md)
for generated-program, observer, scope and budget contracts.

Dual-licensed under GPL-3.0-only or a commercial license; contact the [Opensitez Vybe Team](https://opensitez.com/). Grammar definitions and generated parser code may be used under any license. See [LICENSE.md](LICENSE.md).
