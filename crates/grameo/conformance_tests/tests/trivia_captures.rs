//! Black-box capture semantics for designated implicit trivia rules.
use pest::Parser as _;
type Shape = Vec<(String, usize, usize, usize)>;

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
    let mut out = Vec::new();
    for pair in pairs {
        visit(pair, 0, &mut out);
    }
    out
}
fn native_shape(tree: &grameo::ParseTree<'_, '_>) -> Shape {
    fn visit(pair: grameo::Pair<'_, '_, '_>, depth: usize, out: &mut Shape) {
        let span = pair.as_span();
        out.push((pair.rule_name().to_owned(), span.start, span.end, depth));
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

macro_rules! case {
    ($name:ident, $grammar:literal) => {
        mod $name {
            use super::*;
            #[derive(pest_derive::Parser)]
            #[grammar_inline = $grammar]
            struct Reference;
            #[test]
            fn implicit_trivia_capture_semantics() {
                let reference = shape(Reference::parse(Rule::main, "a b").unwrap());
                let grammar = grameo::compile($grammar).unwrap();
                let capture = grameo::engine::parse_program(
                    &grammar,
                    "main",
                    "a b",
                    grameo::MatchOptions::default(),
                )
                .unwrap();
                assert_eq!(native_shape(&capture), reference, "{}", stringify!($name));
                let recognized = grameo::engine::recognize_program(
                    &grammar,
                    "main",
                    "a b",
                    grameo::MatchOptions::default(),
                )
                .unwrap();
                assert_eq!(capture.consumed(), recognized.consumed);
            }
        }
    };
}
case!(
    silent_normal_child,
    r#"main = { "a" ~ "b" ~ EOI } WHITESPACE = _{ child } child = { inner } inner = { " " }"#
);
case!(
    normal_normal_child,
    r#"main = { "a" ~ "b" ~ EOI } WHITESPACE = { child } child = { inner } inner = { " " }"#
);
case!(
    silent_nonatomic_child,
    r#"main = { "a" ~ "b" ~ EOI } WHITESPACE = _{ child } child = !{ inner } inner = { " " }"#
);
case!(
    normal_nonatomic_child,
    r#"main = { "a" ~ "b" ~ EOI } WHITESPACE = { child } child = !{ inner } inner = { " " }"#
);
case!(
    silent_compound_child,
    r#"main = { "a" ~ "b" ~ EOI } WHITESPACE = _{ child } child = ${ inner } inner = { " " }"#
);
case!(
    normal_compound_child,
    r#"main = { "a" ~ "b" ~ EOI } WHITESPACE = { child } child = ${ inner } inner = { " " }"#
);
case!(
    atomic_trivia,
    r#"main = { "a" ~ "b" ~ EOI } WHITESPACE = @{ child } child = { inner } inner = { " " }"#
);
case!(
    compound_trivia,
    r#"main = { "a" ~ "b" ~ EOI } WHITESPACE = ${ child } child = { inner } inner = { " " }"#
);
case!(
    nonatomic_trivia,
    r#"main = { "a" ~ "b" ~ EOI } WHITESPACE = !{ child } child = { inner } inner = { " " }"#
);
case!(
    silent_comment_child,
    r#"main = { "a" ~ "b" ~ EOI } COMMENT = _{ child } child = { inner } inner = { " " }"#
);
case!(
    normal_comment_child,
    r#"main = { "a" ~ "b" ~ EOI } COMMENT = { child } child = { inner } inner = { " " }"#
);
case!(
    silent_comment_nonatomic_child,
    r#"main = { "a" ~ "b" ~ EOI } COMMENT = _{ child } child = !{ inner } inner = { " " }"#
);
case!(
    normal_comment_compound_child,
    r#"main = { "a" ~ "b" ~ EOI } COMMENT = { child } child = ${ inner } inner = { " " }"#
);
case!(
    explicit_normal_trivia,
    r#"main = ${ "a" ~ WHITESPACE ~ "b" ~ EOI } WHITESPACE = { child } child = { inner } inner = { " " }"#
);
case!(
    explicit_silent_trivia,
    r#"main = ${ "a" ~ WHITESPACE ~ "b" ~ EOI } WHITESPACE = _{ child } child = { inner } inner = { " " }"#
);
case!(
    explicit_nonatomic_trivia,
    r#"main = ${ "a" ~ WHITESPACE ~ "b" ~ EOI } WHITESPACE = !{ child } child = { inner } inner = { " " }"#
);
case!(
    explicit_compound_trivia,
    r#"main = ${ "a" ~ WHITESPACE ~ "b" ~ EOI } WHITESPACE = ${ child } child = { inner } inner = { " " }"#
);
case!(
    explicit_normal_comment,
    r#"main = ${ "a" ~ COMMENT ~ "b" ~ EOI } COMMENT = { child } child = { inner } inner = { " " }"#
);
