use vybe_parser::{
    MatchOptions, ParseErrorKind,
    engine::parse_program,
    program::{Instruction, Program, RuleSpec},
};

#[derive(Debug)]
struct Unoptimized<'g>(&'g vybe_parser::CompiledGrammar);
impl Program for Unoptimized<'_> {
    fn rule_id(&self, name: &str) -> Option<usize> {
        self.0.rule_id(name)
    }
    fn rule(&self, id: usize) -> RuleSpec<'_> {
        self.0.rule(id)
    }
    fn instruction(&self, id: usize) -> Instruction<'_> {
        self.0.instruction(id)
    }
}

#[test]
fn single_ascii_trivia_certificate_matches_original_rules_and_errors() {
    let source = r##"WHITESPACE = _{ " " | '\t'..'\r' | ^"a" }
COMMENT = _{ "#" ~ (!NEWLINE ~ ANY)* }
root = { SOI ~ "x" ~ "y" ~ EOI }"##;
    let grammar = vybe_parser::compile(source).unwrap();
    let class = grammar.whitespace_class().unwrap();
    for byte in 0u8..=127 {
        let text = char::from(byte).to_string();
        assert_eq!(
            class.contains(byte),
            grammar.recognize("WHITESPACE", &text).is_ok()
        );
    }
    for input in [
        "xy",
        "x y",
        "xaaaAAA\t\ry",
        "x #comment\ny",
        "x é y",
        "x\ny!",
        "x #bad",
    ] {
        let optimized = grammar.parse_source("root", input);
        let unoptimized = Unoptimized(&grammar);
        let original = parse_program(&unoptimized, "root", input, MatchOptions::default());
        match (optimized, original) {
            (Ok(a), Ok(b)) => {
                assert_eq!(a.consumed(), b.consumed());
                assert_eq!(a.capture_count(), b.capture_count());
            }
            (Err(a), Err(b)) => assert_eq!(a, b),
            (a, b) => panic!("{input:?}: {a:?} vs {b:?}"),
        }
    }
}

#[test]
fn captures_stack_unicode_nullable_and_multibyte_trivia_stay_on_full_engine() {
    for source in [
        r#"WHITESPACE = { " " }"#,
        r#"WHITESPACE = _{ child } child = { " " }"#,
        r#"WHITESPACE = _{ PUSH(" ") }"#,
        r#"WHITESPACE = _{ "é" }"#,
        r#"WHITESPACE = _{ NEWLINE }"#,
        r#"WHITESPACE = _{ " "? }"#,
    ] {
        let syntax = vybe_parser::parse(source).unwrap();
        assert!(
            vybe_parser::resolve(syntax)
                .unwrap()
                .whitespace_class()
                .is_none()
        );
    }
}

#[test]
fn fast_trivia_obeys_work_and_depth_guards() {
    let grammar = vybe_parser::compile(r#"WHITESPACE = _{ " " } root = { "x" ~ "y" }"#).unwrap();
    assert_eq!(
        grammar
            .parse_source_with_options(
                "root",
                "x y",
                MatchOptions {
                    max_rule_depth: 1,
                    ..MatchOptions::default()
                }
            )
            .unwrap_err()
            .kind,
        ParseErrorKind::DepthLimit
    );
    let input = format!("x{}y", " ".repeat(100_000));
    assert_eq!(
        grammar
            .parse_source_with_options(
                "root",
                &input,
                MatchOptions {
                    max_steps: 100,
                    ..MatchOptions::default()
                }
            )
            .unwrap_err()
            .kind,
        ParseErrorKind::WorkLimit
    );
}

#[test]
fn leading_terminal_trivia_can_scan_before_comment_fallback_without_reordering_it() {
    let grammar = vybe_parser::compile(
        r#"WHITESPACE = _{ " " | COMMENT | "*" }
COMMENT = ${ "*" ~ "!" }
root = { SOI ~ "x" ~ "y" ~ EOI }"#,
    )
    .unwrap();
    assert!(grammar.whitespace_class().is_none());
    let prefix = grammar.whitespace_prefix_class().unwrap();
    assert!(prefix.contains(b' '));
    assert!(
        !prefix.contains(b'*'),
        "terminal after comment must retain lower priority"
    );
    let unoptimized = Unoptimized(&grammar);
    for input in ["xy", "x y", "x *! y", "x * y", "x *!*! y", "x *!z"] {
        match (
            grammar.parse_source("root", input),
            parse_program(&unoptimized, "root", input, MatchOptions::default()),
        ) {
            (Ok(a), Ok(b)) => {
                assert_eq!(a.consumed(), b.consumed());
                assert_eq!(a.capture_count(), b.capture_count());
                let flatten = |tree: &vybe_parser::ParseTree<'_, '_>| {
                    tree.pairs()
                        .flat_map(|pair| pair.into_inner())
                        .map(|pair| (pair.rule_name().to_owned(), pair.as_span()))
                        .collect::<Vec<_>>()
                };
                assert_eq!(flatten(&a), flatten(&b));
            }
            (Err(a), Err(b)) => assert_eq!(a, b),
            (a, b) => panic!("{input:?}: {a:?} vs {b:?}"),
        }
    }
}
