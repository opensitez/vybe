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
fn trivia_failure_certificates_preserve_full_errors_steps_and_depth_limits() {
    use vybe_parser::engine::recognize_program;
    #[derive(Debug)]
    struct Baseline<'g>(&'g vybe_parser::CompiledGrammar);
    impl Program for Baseline<'_> {
        fn fast_ascii_repetitions(&self) -> bool {
            true
        }
        fn rule_id(&self, name: &str) -> Option<usize> {
            self.0.rule_id(name)
        }
        fn rule(&self, id: usize) -> RuleSpec<'_> {
            self.0.rule(id)
        }
        fn instruction(&self, id: usize) -> Instruction<'_> {
            self.0.instruction(id)
        }
        fn scope(&self, id: usize) -> vybe_parser::program::Scope {
            self.0.scope(id)
        }
    }
    for source in [
        r##"WHITESPACE = _{ " " | COMMENT } COMMENT = _{ "#" ~ (!NEWLINE ~ ANY)* } root = { SOI ~ "a" ~ "b" ~ EOI }"##,
        r##"COMMENT = _{ nested | ^"é" ~ ANY | ('中'..'文') ~ ANY } nested = ${ ("/*" | "//") ~ (!NEWLINE ~ ANY)* } root = { SOI ~ "a" ~ "b" ~ EOI }"##,
        r##"COMMENT = _{ "!" ~ POP | "#" ~ PUSH("x") } root = { SOI ~ "a" ~ "b" ~ EOI }"##,
        r##"COMMENT = _{ ASCII_DIGIT{2,3} ~ "x" } root = { SOI ~ "a" ~ "b" ~ EOI }"##,
        r##"COMMENT = _{ PUSH("#") ~ "x" } root = { SOI ~ "a" ~ "b" ~ EOI }"##,
    ] {
        let grammar = vybe_parser::compile(source).unwrap();
        let comment = grammar.rule_id("COMMENT").unwrap();
        assert!(grammar.trivia_failure_prefix(comment).is_some());
        let baseline = Baseline(&grammar);
        for input in [
            "", "ab", "a b", "a#hi\nb", "a!b", "a/*hi\nb", "aéb", "a中b", "a12xb", "a#xb", "a13b",
        ] {
            for depth in 0..=8 {
                for budget in [0, 1, 2, 4, 8, 16, 24, 31, 32, 40, 64, 100, 200, 500] {
                    let options = MatchOptions {
                        max_steps: budget,
                        max_rule_depth: depth,
                    };
                    let a = recognize_program(&grammar, "root", input, options);
                    let b = recognize_program(&baseline, "root", input, options);
                    assert_eq!(a, b, "{source}\n{input:?}; {options:?}");
                    let captures = |tree: vybe_parser::ParseTree<'_, '_>| {
                        (
                            tree.consumed(),
                            tree.capture_count(),
                            tree.pairs()
                                .map(|p| (p.as_rule(), p.as_span()))
                                .collect::<Vec<_>>(),
                        )
                    };
                    assert_eq!(
                        parse_program(&grammar, "root", input, options).map(captures),
                        parse_program(&baseline, "root", input, options).map(captures),
                        "captures {source}\n{input:?}; {options:?}"
                    );
                }
            }
        }
    }
    for source in [
        r##"COMMENT = _{ POP ~ "#" } root = { "a" ~ "b" }"##,
        r##"COMMENT = _{ !"a" ~ "#" } root = { "a" ~ "b" }"##,
        r##"COMMENT = _{ "#"? ~ "x" } root = { "a" ~ "b" }"##,
    ] {
        let grammar = vybe_parser::compile(source).unwrap();
        assert!(
            grammar
                .trivia_failure_prefix(grammar.rule_id("COMMENT").unwrap())
                .is_none()
        );
    }
}

#[test]
fn atomic_ascii_repetitions_preserve_captures_diagnostics_and_every_budget_boundary() {
    use vybe_parser::engine::recognize_program;
    #[derive(Debug)]
    struct RepeatBaseline<'g>(&'g vybe_parser::CompiledGrammar);
    impl Program for RepeatBaseline<'_> {
        fn rule_id(&self, name: &str) -> Option<usize> {
            self.0.rule_id(name)
        }
        fn rule(&self, id: usize) -> RuleSpec<'_> {
            self.0.rule(id)
        }
        fn instruction(&self, id: usize) -> Instruction<'_> {
            self.0.instruction(id)
        }
        fn scope(&self, id: usize) -> vybe_parser::program::Scope {
            self.0.scope(id)
        }
    }
    let source = r#"
root = { SOI ~ (number ~ "x" | number) ~ EOI }
number = @{ ASCII_DIGIT+ }
bounded = ${ ASCII_HEX_DIGIT{2,4} }
normal = { ASCII_DIGIT+ }
WHITESPACE = _{ " " }
probe = { &number ~ number ~ EOI }
"#;
    let grammar = vybe_parser::compile(source).unwrap();
    let unoptimized = RepeatBaseline(&grammar);
    for entry in ["root", "number", "bounded", "normal", "probe"] {
        for input in [
            "", "1", "1234", "12345", "1 2", "12x", "12é", "é", "abc", "01AF",
        ] {
            for budget in 0..100 {
                let options = MatchOptions {
                    max_steps: budget,
                    ..MatchOptions::default()
                };
                assert_eq!(
                    recognize_program(&grammar, entry, input, options),
                    recognize_program(&unoptimized, entry, input, options),
                    "{entry} {input:?} budget={budget}"
                );
                let optimized = parse_program(&grammar, entry, input, options);
                let ordinary = parse_program(&unoptimized, entry, input, options);
                match (optimized, ordinary) {
                    (Ok(a), Ok(b)) => {
                        assert_eq!(a.consumed(), b.consumed());
                        assert_eq!(a.capture_count(), b.capture_count());
                        let pairs = |tree: &vybe_parser::ParseTree<'_, '_>| {
                            tree.pairs()
                                .map(|p| (p.as_rule(), p.as_span()))
                                .collect::<Vec<_>>()
                        };
                        assert_eq!(pairs(&a), pairs(&b));
                    }
                    (Err(a), Err(b)) => assert_eq!(a, b, "{entry} {input:?} budget={budget}"),
                    (a, b) => panic!("{entry} {input:?} budget={budget}: {a:?} / {b:?}"),
                }
            }
        }
    }
    // Every byte is checked for each supported builtin, including bytes that
    // begin a multi-byte scalar. Unicode characters must never be split.
    for builtin in [
        "ASCII",
        "ASCII_DIGIT",
        "ASCII_NONZERO_DIGIT",
        "ASCII_BIN_DIGIT",
        "ASCII_OCT_DIGIT",
        "ASCII_HEX_DIGIT",
        "ASCII_ALPHA_LOWER",
        "ASCII_ALPHA_UPPER",
        "ASCII_ALPHA",
        "ASCII_ALPHANUMERIC",
    ] {
        let grammar = vybe_parser::compile(&format!("root = @{{ {builtin}{{2,4}} }}")).unwrap();
        for ch in (0..=127).map(char::from).chain(['é', '😀']) {
            let input = ch.to_string().repeat(3);
            assert_eq!(
                recognize_program(&grammar, "root", &input, MatchOptions::default()),
                recognize_program(
                    &RepeatBaseline(&grammar),
                    "root",
                    &input,
                    MatchOptions::default()
                ),
                "{builtin} {input:?}"
            );
        }
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
