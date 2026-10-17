//! Capture and diagnostic fixtures for certifying Python short-string bodies.
//! The alternate grammar only names silent bodies under the original atomic rule.
use grameo::{CompiledGrammar, ParseTree, Span};

const LEGACY: &str = r####"string_literal = @{
    string_prefix? ~ (triple_double_string | triple_single_string
        | "'" ~ ("\\'" | "\\\\" | (!("'") ~ !NEWLINE ~ ANY))* ~ "'"
        | "\"" ~ ("\\\"" | "\\\\" | (!("\"" ) ~ !NEWLINE ~ ANY))* ~ "\"")
}
triple_double_string = _{ "\"\"\"" ~ (!"\"\"\"" ~ ANY)* ~ "\"\"\"" }
triple_single_string = _{ "'''" ~ (!("'''") ~ ANY)* ~ "'''" }
"####;

fn grammars() -> (CompiledGrammar, CompiledGrammar) {
    let source = include_str!("../src/grammar.pest");
    let start = source.find("string_literal = @{").unwrap();
    let prefix = source[start..].find("\nstring_prefix =").unwrap() + start;
    let end = source[prefix + 1..]
        .find('\n')
        .unwrap_or(source.len() - prefix - 1)
        + prefix
        + 1;
    let prefix_rule = &source[prefix..end];
    let newline = source
        .lines()
        .find(|line| line.starts_with("NEWLINE ="))
        .unwrap();
    let root = "root = { SOI ~ string_literal ~ EOI }\n";
    (
        grameo::compile(&format!("{root}{newline}\n{LEGACY}{prefix_rule}")).unwrap(),
        grameo::compile(&format!(
            "{root}{newline}\n{}{prefix_rule}",
            &source[start..prefix]
        ))
        .unwrap(),
    )
}

#[test]
fn production_short_bodies_are_certified_with_named_newline() {
    let grammar = grameo::compile(include_str!("../src/grammar.pest")).unwrap();
    for name in ["single_quoted_string", "double_quoted_string"] {
        let rule = grammar.rule_id(name).unwrap();
        let fragment =
            grameo::quoted::certify(&grammar, rule).expect("short body must use scanner");
        assert!(fragment.forbidden_newline);
        assert!(fragment.newline_crlf_only);
        assert!(!fragment.any_escape);
    }
}

#[test]
fn python_named_newline_preserves_bare_carriage_return() {
    let (old, new) = grammars();
    for delimiter in ['\'', '"'] {
        for grammar in [&old, &new] {
            assert!(
                grammar
                    .parse_source("root", &format!("{delimiter}a\rb{delimiter}"))
                    .is_ok()
            );
            assert!(
                grammar
                    .parse_source("root", &format!("{delimiter}a\nb{delimiter}"))
                    .is_err()
            );
            assert!(
                grammar
                    .parse_source("root", &format!("{delimiter}a\r\nb{delimiter}"))
                    .is_err()
            );
        }
    }
}

fn captures(tree: &ParseTree<'_, '_>) -> Vec<(String, Span, usize)> {
    fn visit(pair: grameo::Pair<'_, '_, '_>, depth: usize, out: &mut Vec<(String, Span, usize)>) {
        out.push((pair.rule_name().into(), pair.as_span(), depth));
        for child in pair.into_inner() {
            visit(child, depth + 1, out);
        }
    }
    let mut out = Vec::new();
    for pair in tree.pairs() {
        visit(pair, 0, &mut out);
    }
    out
}

fn assert_parity(old: &CompiledGrammar, new: &CompiledGrammar, input: &str) {
    match (
        old.parse_source("root", input),
        new.parse_source("root", input),
    ) {
        (Ok(a), Ok(b)) => {
            assert_eq!(a.consumed(), b.consumed(), "{input:?}");
            assert_eq!(captures(&a), captures(&b), "{input:?}");
        }
        (Err(a), Err(b)) => assert_eq!(a, b, "{input:?}"),
        _ => panic!("string recognition changed for {input:?}"),
    }
}

#[test]
fn silent_short_bodies_preserve_all_prefixes_captures_and_errors() {
    let (old, new) = grammars();
    let bodies = [
        "''",
        "\"\"",
        "'plain'",
        "\"世界🙂\"",
        "'e\u{301}'",
        "'é🙂'",
        "'\\'quoted\\''",
        "\"\\\"quoted\\\"\"",
        "'a\\\\b'",
        "'a\\qb'",
        "'a\nline'",
        "'a\rline'",
        "'a\r\nline'",
        "'a\\\nline'",
        "'a\\\rline'",
        "'unterminated",
        "'trailing\\",
        "'a'junk",
        "'''multi\nline'''",
        "\"\"\"世界🙂\r\ntriple\"\"\"",
        "'''unterminated",
        "\"\"\"unterminated",
    ];
    for prefix in [
        "", "r", "R", "b", "B", "u", "U", "rb", "Rb", "rB", "RB", "br", "bR", "Br", "BR",
    ] {
        for body in bodies {
            assert_parity(&old, &new, &format!("{prefix}{body}"));
        }
    }
}

#[test]
fn quoted_delimiter_escape_runs_and_unicode_keep_exact_boundaries() {
    let (old, new) = grammars();
    let atoms = ["", "a", "'", "\"", "\\", "\n", "\r", "\r\n", "é", "🙂"];
    for delimiter in ['\'', '"'] {
        for left in atoms {
            for middle in atoms {
                for right in atoms {
                    assert_parity(
                        &old,
                        &new,
                        &format!("{delimiter}{left}{middle}{right}{delimiter}"),
                    );
                }
            }
        }
        for count in 0..12 {
            assert_parity(
                &old,
                &new,
                &format!("{delimiter}{}{}{delimiter}", "\\".repeat(count), delimiter),
            );
        }
        assert_parity(
            &old,
            &new,
            &format!("{delimiter}{}{delimiter}", "hello 世界 🙂 ".repeat(20_000)),
        );
    }
}
