//! Black-box regressions for full-corpus heredoc acceptance and tail rotation.
use pest::Parser as _;
type Shape = Vec<(String, usize, usize, usize)>;
#[derive(pest_derive::Parser)]
#[grammar = "../../../languages/ruby/src/grammar.pest"]
struct Reference;
fn shape<R: pest::RuleType>(pairs: pest::iterators::Pairs<'_, R>) -> Shape {
    fn visit<R: pest::RuleType>(pair: pest::iterators::Pair<'_, R>, depth: usize, out: &mut Shape) {
        out.push((
            format!("{:?}", pair.as_rule()),
            pair.as_span().start(),
            pair.as_span().end(),
            depth,
        ));
        for child in pair.into_inner() {
            visit(child, depth + 1, out);
        }
    }
    let mut result = Vec::new();
    for pair in pairs {
        visit(pair, 0, &mut result);
    }
    result
}
fn native_shape(tree: &grameo::ParseTree<'_, '_>) -> Shape {
    fn visit(pair: grameo::Pair<'_, '_, '_>, depth: usize, out: &mut Shape) {
        let span = pair.as_span();
        out.push((pair.rule_name().to_owned(), span.start, span.end, depth));
        for child in pair.into_inner() {
            visit(child, depth + 1, out);
        }
    }
    let mut result = Vec::new();
    for pair in tree.pairs() {
        visit(pair, 0, &mut result);
    }
    result
}
fn compare<R: pest::RuleType>(
    grammar: &grameo::CompiledGrammar,
    entry: &str,
    source: &str,
    expected: Result<Shape, pest::error::Error<R>>,
) {
    let capture = grameo::engine::parse_program(
        grammar,
        entry,
        source,
        grameo::MatchOptions::default(),
    );
    let recognition = grameo::engine::recognize_program(
        grammar,
        entry,
        source,
        grameo::MatchOptions::default(),
    );
    match (expected, capture, recognition) {
        (Ok(expected), Ok(capture), Ok(recognition)) => {
            assert_eq!(native_shape(&capture), expected, "source={source:?}");
            assert_eq!(capture.consumed(), recognition.consumed);
            assert_eq!(capture.consumed(), expected[0].2);
        }
        (Err(_), Err(capture), Err(recognition)) => {
            assert_eq!(capture.kind, grameo::ParseErrorKind::Syntax);
            assert_eq!(capture, recognition);
        }
        (expected, capture, recognition) => panic!(
            "source={source:?}; reference={expected:?}; capture={capture:?}; recognition={recognition:?}"
        ),
    }
}
#[test]
fn heredoc_reference_branch() {
    let grammar =
        grameo::compile(include_str!("../../../../languages/ruby/src/grammar.pest")).unwrap();
    for source in [
        include_str!("../../../../tests/ruby/string_methods/str_squiggly_heredoc.rb"),
        "x = <<~X\nX\n",
        "x = <<~X\n a\nX\n",
        "<<~X\nX\n",
    ] {
        let expected = Reference::parse(Rule::program, source).map(shape);
        assert!(
            expected
                .as_ref()
                .unwrap()
                .iter()
                .any(|node| node.0 == "heredoc")
        );
        compare(&grammar, "program", source, expected);
    }
}
macro_rules! probe {
    ($name:ident, $grammar:literal, [$($source:literal),+]) => {
        mod $name {
            use super::*;
            #[derive(pest_derive::Parser)]
            #[grammar_inline = $grammar]
            struct Probe;
            #[test]
            fn tail_repetition() {
                let grammar = grameo::compile($grammar).unwrap();
                for source in [$($source),+] {
                    compare(&grammar, "main", source, Probe::parse(Rule::main, source).map(shape));
                }
            }
        }
    };
}
probe!(
    literal_repeat,
    r#"main = @{ "a" ~ ("b" ~ "c"*)* ~ "b" ~ EOI }"#,
    ["ab", "abb", "abcb", "abcbb", "abc"]
);
probe!(
    newline_repeat,
    r#"main = @{ "a" ~ ("\n" ~ (!"\n" ~ ANY)*)* ~ "\n" ~ EOI }"#,
    ["a\n", "a\nX\n", "a\nX\n\n", "a\nX"]
);
probe!(
    call_newline_repeat,
    r#"main = @{ "a" ~ (nl ~ (!nl ~ ANY)*)* ~ nl ~ EOI } nl = { "\n" }"#,
    ["a\n", "a\nX\n", "a\nX\n\n", "a\nX"]
);
probe!(
    newline_choice_repeat,
    r#"main = @{ "a" ~ (nl ~ (!nl ~ ANY)*)* ~ nl ~ EOI } nl = { "\n" | "\r\n" }"#,
    ["a\n", "a\nX\n", "a\nX\n\n"]
);
probe!(
    full_tail_repeat,
    r#"main = @{ "a" ~ nl ~ (!nl ~ ANY)* ~ (nl ~ (!nl ~ ANY)*)* ~ nl } nl = { "\n" | "\r\n" }"#,
    ["a\n", "a\nX\n", "a\nX\n\n"]
);
probe!(
    full_choice_tail_repeat,
    r#"main = @{ "a" ~ nl ~ (!nl ~ ANY)* ~ (nl ~ (!nl ~ ANY)*)* ~ nl | "b" ~ nl ~ (!nl ~ ANY)* ~ (nl ~ (!nl ~ ANY)*)* ~ nl } nl = { "\n" | "\r\n" }"#,
    ["a\n", "a\nX\n", "a\nX\n\n"]
);
probe!(
    only_scan_repeat,
    r#"main = @{ (!nl ~ ANY)* } nl = { "\n" | "\r\n" }"#,
    ["", "X", "X\n", "\n"]
);
probe!(
    full_tail_eoi,
    r#"main = @{ "a" ~ nl ~ (!nl ~ ANY)* ~ (nl ~ (!nl ~ ANY)*)* ~ nl ~ EOI } nl = { "\n" | "\r\n" }"#,
    ["a\n", "a\nX\n", "a\nX\n\n"]
);
probe!(
    full_tail_nonatomic,
    r#"main = { "a" ~ nl ~ (!nl ~ ANY)* ~ (nl ~ (!nl ~ ANY)*)* ~ nl ~ EOI } nl = { "\n" | "\r\n" }"#,
    ["a\n", "a\nX\n", "a\nX\n\n"]
);
probe!(
    literal_tail_atomic,
    r#"main = @{ "a" ~ ("b" ~ "c"*)* ~ "b" }"#,
    ["ab", "abb", "abcb", "abcbb", "abc"]
);
probe!(
    literal_tail_normal,
    r#"main = { "a" ~ ("b" ~ "c"*)* ~ "b" }"#,
    ["ab", "abb", "abcb", "abcbb", "abc"]
);
probe!(
    single_tail_atomic,
    r#"main = @{ "a" ~ "b"* ~ "b" }"#,
    ["ab", "abb", "abbb", "a"]
);
probe!(
    call_newline_tail,
    r#"main = @{ "a" ~ (nl ~ (!nl ~ ANY)*)* ~ nl } nl = { "\n" | "\r\n" }"#,
    ["a\n", "a\nX\n", "a\nX\n\n"]
);
probe!(
    named_nullable_tail,
    r#"main = { "a" ~ (p ~ rest)* ~ p } p = { "b" } rest = { "c"* }"#,
    ["ab", "abb", "abcb", "abc"]
);
probe!(
    grouped_eoi_tail,
    r#"main = { "a" ~ ((p ~ rest)* ~ p) ~ EOI } p = { "b" } rest = { "c"* }"#,
    ["ab", "abb", "abcb", "abc"]
);
probe!(
    extra_groups_tail,
    r#"main = { "a" ~ ((("b") ~ "c"*)*) ~ ("b") }"#,
    ["ab", "abb", "abcb", "abc"]
);
probe!(
    nonnullable_tail,
    r#"main = { "a" ~ (p ~ "c")* ~ p } p = { "b" }"#,
    ["ab", "abcb", "abcbc", "abb"]
);
probe!(
    stackful_reference_tail,
    r#"main = { value ~ "X" ~ POP_ALL ~ EOI } value = { "a" ~ (p ~ "c"*)* ~ p } p = { PUSH("b") }"#,
    ["abXb", "abbXbb", "abcbXbb", "abX", "abcbXb"]
);
probe!(
    bounded_tail,
    r#"main = { "a" ~ ("b" ~ "c"*){0,3} ~ "b" }"#,
    ["ab", "abb", "abbb", "abbbb", "abbbbb"]
);
probe!(
    plus_tail,
    r#"main = { "a" ~ ("b" ~ "c"*)+ ~ "b" }"#,
    ["ab", "abb", "abbb", "abcb", "abcbb"]
);
probe!(
    group_tail_no_eoi,
    r#"main = { "a" ~ ((p ~ rest)* ~ p) } p = { "b" } rest = { "c"* }"#,
    ["ab", "abb", "abcb", "abc"]
);
probe!(
    deeper_group_eoi_tail,
    r#"main = { (("a" ~ ((((p ~ rest)* ~ p))))) ~ EOI } p = { "b" } rest = { "c"* }"#,
    ["ab", "abb", "abcb", "abc"]
);
probe!(
    deeper_group_tail_no_eoi,
    r#"main = { (("a" ~ ((((p ~ rest)* ~ p))))) } p = { "b" } rest = { "c"* }"#,
    ["ab", "abb", "abcb", "abc"]
);
probe!(
    grouped_prefix_suffix_tail,
    r#"main = { "a" ~ (((p)) ~ "c"*)* ~ ((p)) } p = { "b" }"#,
    ["ab", "abb", "abcb", "abc"]
);
