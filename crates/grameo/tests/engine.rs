use grameo::{MatchOptions, ParseErrorKind, compile};

fn accepts(grammar: &str, entry: &str, source: &str) -> bool {
    compile(grammar)
        .unwrap()
        .parse_source(entry, source)
        .is_ok()
}

#[test]
fn ordered_choices_are_not_longest_match_and_failed_sequences_rollback() {
    let grammar = compile(
        r#"main = { ("a" | "ab") ~ EOI } rollback = { (word ~ "x" | word) ~ EOI } word = { "a" }"#,
    )
    .unwrap();
    assert!(grammar.parse_source("main", "a").is_ok());
    assert!(grammar.parse_source("main", "ab").is_err());
    let tree = grammar.parse_source("rollback", "a").unwrap();
    let root = tree.pairs().next().unwrap();
    let children: Vec<_> = root
        .into_inner()
        .map(|pair| (pair.rule_name(), pair.as_str()))
        .collect();
    assert_eq!(children, [("word", "a"), ("EOI", "")]);
    assert_eq!(tree.capture_count(), 3);
}

#[test]
fn whitespace_is_inserted_only_at_boundaries_and_newlines_can_remain_visible() {
    let grammar = r#"WHITESPACE = _{ " " } main = { SOI ~ "a" ~ "b" ~ EOI } single = { "a" }"#;
    assert!(accepts(grammar, "main", " a  b "));
    assert!(!accepts(grammar, "main", "a\nb"));
    assert!(!accepts(grammar, "single", " a"));
    let compiled = compile(grammar).unwrap();
    assert_eq!(compiled.parse_source("single", "ab").unwrap().consumed(), 1);
}

#[test]
fn predicates_restore_both_pops_and_pushes_and_never_emit_captures() {
    let grammar = compile(r#"main = @{ PUSH("a") ~ &POP ~ POP ~ EOI } push = @{ &PUSH("a") ~ "a" ~ !DROP ~ EOI } negative = { !leaf ~ ANY ~ EOI } leaf = { "a" }"#).unwrap();
    assert!(grammar.parse_source("main", "aa").is_ok());
    assert!(grammar.parse_source("push", "a").is_ok());
    let tree = grammar.parse_source("negative", "b").unwrap();
    assert_eq!(tree.capture_count(), 2);
}

#[test]
fn failed_alternatives_restore_popped_values_not_just_stack_length() {
    let grammar = compile(r#"main = @{ PUSH("a") ~ (POP ~ "x" | POP) ~ EOI }"#).unwrap();
    assert!(grammar.parse_source("main", "aa").is_ok());
    assert!(grammar.parse_source("main", "aax").is_ok());
}

#[test]
fn repeat_attempt_restores_skipping_captures_and_stack_on_failure() {
    let grammar = compile(r#"WHITESPACE = { " " } main = { word* } word = { "a" } stack = @{ PUSH("a") ~ (POP ~ "x")* ~ POP ~ EOI }"#).unwrap();
    let tree = grammar.parse_source("main", "a a ").unwrap();
    assert_eq!(tree.consumed(), 3);
    assert_eq!(
        tree.pairs()
            .next()
            .unwrap()
            .into_inner()
            .map(|p| (p.rule_name(), p.as_str()))
            .collect::<Vec<_>>(),
        [("word", "a"), ("WHITESPACE", " "), ("word", "a")]
    );
    assert!(grammar.parse_source("stack", "aa").is_ok());
}

#[test]
fn bounded_repetition_is_greedy_and_does_not_expand_or_backtrack_counts() {
    assert!(accepts(r#"a = { "x"{2,3} ~ EOI }"#, "a", "xxx"));
    assert!(!accepts(r#"a = { "x"{2,3} ~ EOI }"#, "a", "x"));
    assert!(!accepts(r#"a = { "x"{2,3} ~ "x" ~ EOI }"#, "a", "xxx"));
    assert!(accepts(r#"a = { ""{3} ~ EOI }"#, "a", ""));
}

#[test]
fn scalar_matching_ascii_case_fold_and_unicode_xid_are_distinct() {
    assert!(accepts(r#"a = { ANY ~ EOI }"#, "a", "😀"));
    assert!(accepts(r#"a = { ^"if" ~ EOI }"#, "a", "IF"));
    assert!(!accepts(r#"a = { ^"é" ~ EOI }"#, "a", "É"));
    assert!(accepts(
        r#"a = { XID_START ~ XID_CONTINUE* ~ EOI }"#,
        "a",
        "α_2"
    ));
    assert!(!accepts(
        r#"a = { XID_START ~ XID_CONTINUE* ~ EOI }"#,
        "a",
        "2a"
    ));
    let grammar = compile(r#"a = { 'α'..'ω' ~ EOI }"#).unwrap();
    let tree = grammar.parse_source("a", "β").unwrap();
    assert_eq!(tree.pairs().next().unwrap().as_span().end, 2);
}

#[test]
fn recognition_and_capture_modes_match_and_fail_identically() {
    let grammar =
        compile(r#"a = { SOI ~ (word ~ ",")* ~ word? ~ EOI } word = @{ ASCII_ALPHA+ }"#).unwrap();
    for source in ["", "a", "a,b", "a,", "a,,b", "é", "abc,def"] {
        let tree = grammar.parse_source("a", source);
        let recognition = grammar.recognize("a", source);
        match (tree, recognition) {
            (Ok(tree), Ok(recognition)) => assert_eq!(tree.consumed(), recognition.consumed),
            (Err(tree), Err(recognition)) => assert_eq!(tree, recognition),
            _ => panic!("capture/recognition disagree: {source:?}"),
        }
    }
}

#[test]
fn source_nesting_is_iterative_and_resource_errors_are_not_syntax_failures() {
    let grammar = compile(r#"a = { "(" ~ a ~ ")" | "x" }"#).unwrap();
    let source = format!("{}x{}", "(".repeat(600), ")".repeat(600));
    assert_eq!(
        grammar.recognize("a", &source).unwrap().consumed,
        source.len()
    );
    let options = MatchOptions {
        max_rule_depth: 20,
        ..MatchOptions::default()
    };
    assert_eq!(
        grammar
            .recognize_with_options("a", &source, options)
            .unwrap_err()
            .kind,
        ParseErrorKind::DepthLimit
    );
    let options = MatchOptions {
        max_steps: 10,
        ..MatchOptions::default()
    };
    assert_eq!(
        grammar
            .recognize_with_options("a", &source, options)
            .unwrap_err()
            .kind,
        ParseErrorKind::WorkLimit
    );
    assert_eq!(
        grammar.recognize("missing", "").unwrap_err().kind,
        ParseErrorKind::UnknownEntry
    );
    // Low-level resolution preserves unsafe IR so runtime guards remain tested
    // even though the normal compile API rejects this definite error.
    let nonprogress = grameo::resolve(grameo::parse(r#"a = { ""* }"#).unwrap()).unwrap();
    assert_eq!(
        nonprogress.recognize("a", "").unwrap_err().kind,
        ParseErrorKind::NonProgress
    );
    let stack = compile("a = { POP }").unwrap();
    assert_eq!(
        stack.recognize("a", "").unwrap_err().kind,
        ParseErrorKind::InvalidStack
    );
}

#[test]
fn pair_clones_and_root_and_child_iterators_do_not_copy_nodes() {
    let grammar = compile(r#"a = _{ word ~ word ~ EOI } word = { "a" }"#).unwrap();
    let tree = grammar.parse_source("a", "aa").unwrap();
    assert_eq!(tree.capture_count(), 3);
    let pairs = tree.pairs();
    assert_eq!(pairs.clone().count(), 3);
    assert_eq!(
        pairs.map(|pair| pair.rule_name()).collect::<Vec<_>>(),
        ["word", "word", "EOI"]
    );
}

#[test]
fn stack_slices_match_bottom_up_and_full_stack_matches_top_down() {
    let grammar = compile(r#"a = @{ PUSH("a") ~ PUSH("b") ~ PEEK[0..] ~ PEEK_ALL ~ POP_ALL ~ !DROP ~ EOI } b = @{ PUSH_LITERAL(")") ~ POP ~ EOI }"#).unwrap();
    assert!(grammar.recognize("a", "ababbaba").is_ok());
    assert!(grammar.recognize("b", ")").is_ok());
}

#[test]
fn repetition_segment_trivia_is_retained_without_leaking_failed_children() {
    let grammar = compile(
        r#"
WHITESPACE = { " " }
word = { "a" }
star = { word* }
plus = { word+ }
minimum = { word{2,} }
bounded = { word{,2} }
rollback = { (bounded ~ "!" | "a") ~ EOI }
root = { bounded ~ EOI }
"#,
    )
    .unwrap();
    let plus = grammar.parse_source("plus", "a ").unwrap();
    assert_eq!(plus.consumed(), 2);
    let children = plus
        .pairs()
        .next()
        .unwrap()
        .into_inner()
        .map(|pair| (pair.rule_name(), pair.as_str()))
        .collect::<Vec<_>>();
    assert_eq!(children, [("word", "a"), ("WHITESPACE", " ")]);
    assert_eq!(grammar.parse_source("star", "a ").unwrap().consumed(), 1);
    assert_eq!(grammar.parse_source("plus", "aa ").unwrap().consumed(), 2);
    assert_eq!(
        grammar.parse_source("minimum", "aa ").unwrap().consumed(),
        3
    );
    let bounded = grammar.parse_source("bounded", " a").unwrap();
    assert_eq!(bounded.consumed(), 2);
    let children = bounded
        .pairs()
        .next()
        .unwrap()
        .into_inner()
        .map(|pair| (pair.rule_name(), pair.as_str()))
        .collect::<Vec<_>>();
    assert_eq!(children, [("WHITESPACE", " "), ("word", "a")]);
    let restored = grammar.parse_source("rollback", "a").unwrap();
    assert_eq!(restored.capture_count(), 2); // root and EOI only
    struct Journal(Vec<(usize, grameo::Span)>);
    impl grameo::builder::Builder for Journal {
        fn checkpoint(&self) -> usize {
            self.0.len()
        }
        fn rollback(&mut self, mark: usize) {
            self.0.truncate(mark);
        }
        fn begin(&mut self, _: usize, _: usize) {}
        fn finish(&mut self, rule: usize, span: grameo::Span, _: &str) {
            self.0.push((rule, span));
        }
    }
    let initial = (usize::MAX, grameo::Span { start: 42, end: 43 });
    let mut built = Journal(vec![initial]);
    let complete = grameo::engine::build_program(
        &grammar,
        "root",
        " a",
        MatchOptions::default(),
        &mut built,
    )
    .unwrap();
    assert_eq!(complete.consumed, 2);
    assert_eq!(
        &built.0[1..],
        &[
            (
                grammar.rule_id("word").unwrap(),
                grameo::Span { start: 1, end: 2 }
            ),
            (
                grammar.rule_id("bounded").unwrap(),
                grameo::Span { start: 0, end: 2 }
            ),
            (
                grammar.rule_id("root").unwrap(),
                grameo::Span { start: 0, end: 2 }
            ),
        ]
    );
    let mut rejected = Journal(vec![initial]);
    assert!(
        grameo::engine::build_program(
            &grammar,
            "root",
            " a!",
            MatchOptions::default(),
            &mut rejected
        )
        .is_err()
    );
    assert_eq!(rejected.0, [initial]);
    for max_steps in 0..complete.steps {
        let mut builder = Journal(vec![initial]);
        assert!(
            grameo::engine::build_program(
                &grammar,
                "root",
                " a",
                MatchOptions {
                    max_steps,
                    ..MatchOptions::default()
                },
                &mut builder
            )
            .is_err()
        );
        assert_eq!(builder.0, [initial], "budget {max_steps}");
    }
    for source in ["", "a ", " a", " x", "aa "] {
        let steps = grammar.recognize("bounded", source).unwrap().steps;
        for max_steps in 0..steps {
            let options = MatchOptions {
                max_steps,
                ..MatchOptions::default()
            };
            let captured = grammar
                .parse_source_with_options("bounded", source, options)
                .unwrap_err();
            let recognized = grammar
                .recognize_with_options("bounded", source, options)
                .unwrap_err();
            assert_eq!(captured.kind, ParseErrorKind::WorkLimit);
            assert_eq!(captured, recognized);
        }
    }
}
