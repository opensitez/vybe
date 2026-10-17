# Grameo

Compile pest-compatible PEG grammars, parse input into capture trees, or validate input without captures. Implemented in Rust.

Add it with `cargo add grameo`.

## Grammar basics

Rules use `name = { expression }`. Choices are ordered: the first successful alternative wins.

| Syntax | Meaning |
| --- | --- |
| `"text"`, `'a'..'z'` | Literal or character range |
| `a ~ b`, `a \| b` | Sequence or ordered choice |
| `a?`, `a*`, `a+` | Optional, zero or more, one or more |
| `&a`, `!a` | Positive or negative lookahead |
| `name = _{ ... }` | Silent rule: no capture |
| `name = @{ ... }` | Atomic rule: no implicit whitespace inside |
| `SOI`, `EOI` | Start and end of input |

A silent `WHITESPACE` rule defines implicit spacing between expressions. `ASCII_DIGIT` matches a digit. Use `SOI` and `EOI` to require a complete input match.

## Example

```rust
fn main() {
    let grammar = grameo::compile(r#"
        WHITESPACE = _{ " " | "\t" | "\n" }
        number = @{ ASCII_DIGIT+ }
        program = { SOI ~ number ~ ("," ~ number)* ~ EOI }
    "#).expect("valid grammar");

    let tree = grammar.parse_source("program", "12, 34").expect("valid input");
    for root in tree.pairs() {
        for pair in root.into_inner() {
            if pair.rule_name() == "number" {
                println!("{}", pair.as_str()); // Prints 12, then 34.
            }
        }
    }

    assert!(grammar.recognize("program", "12, 34").is_ok());
    assert!(grammar.recognize("program", "12, nope").is_err());
}
```

`parse_source` returns captures; `recognize` validates without allocating a capture tree. Grammar compilation and input parsing return errors when they fail. Input parsing has no default work cutoff; callers can set an explicit limit with `MatchOptions`.

For direct AST construction, implement `grameo::builder::Builder` and call `grammar.build_ast(...)` or a generated parser's `build_ast(...)`. Builders support transactional rollback and optional Pratt expression bindings.

Grammar comments such as `// @grameo ast return_statement = Return(value: expression?, span: rule)` can generate owned AST reductions through a target's constructors. See [AST bindings](AST.md) for selectors, schema validation and rollback requirements.

Generate a Rust parser with static grammar tables:

```sh
cargo run --manifest-path crates/grameo/Cargo.toml -p grameo_codegen --bin grammargen -- input.pest output.rs
```

Performance directives can select native matching and Pratt bindings through
explicit compile options. See [grammar optimization directives](OPTIMIZATION.md).

## License

The package is dual-licensed under **GPL-3.0-only or a commercial license**. Contact the [Opensitez Vybe Team](https://opensitez.com/) for commercial licensing.

Your grammars and generated parser code may be used, modified and distributed for any purpose, including proprietary and commercial use, under a license of your choice. The parser runtime itself remains subject to the package license. See [LICENSE.md](LICENSE.md).
