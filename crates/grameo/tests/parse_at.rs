use grameo::{compile, engine::parse_program_at, MatchOptions, ParseErrorKind, Span};

#[test]
fn captures_keep_original_input_and_absolute_nested_spans() {
    let grammar = compile(r#"entry = { item ~ EOI } item = { "a" ~ "b" }"#).unwrap();
    let source = "λ:ab";
    let tree = parse_program_at(&grammar, "entry", source, 3, MatchOptions::default()).unwrap();
    assert_eq!(tree.source().as_ptr(), source.as_ptr());
    assert_eq!(tree.consumed(), 5);
    let root = tree.pairs().next().unwrap();
    assert_eq!(root.as_span(), Span { start: 3, end: 5 });
    let children: Vec<_> = root.into_inner().map(|p| (p.rule_name(), p.as_span(), p.as_str())).collect();
    assert_eq!(children, [("item", Span { start: 3, end: 5 }, "ab"), ("EOI", Span { start: 5, end: 5 }, "")]);
}

#[test]
fn full_input_boundaries_do_not_become_fragment_boundaries() {
    let grammar = compile(r#"soi = { SOI ~ "a" } eoi = { "a" ~ EOI } prefix = { "a" } empty = { EOI }"#).unwrap();
    let options = MatchOptions::default();
    assert!(parse_program_at(&grammar, "soi", "xa", 1, options).is_err());
    assert!(parse_program_at(&grammar, "eoi", "xab", 1, options).is_err());
    assert_eq!(parse_program_at(&grammar, "prefix", "xab", 1, options).unwrap().consumed(), 2);
    assert_eq!(parse_program_at(&grammar, "empty", "λ", 2, options).unwrap().consumed(), 2);
}

#[test]
fn rejects_invalid_utf8_and_out_of_bounds_starts() {
    let grammar = compile("entry = { ANY }").unwrap();
    for start in [1, 3, usize::MAX] {
        let error = parse_program_at(&grammar, "entry", "λ", start, MatchOptions::default()).unwrap_err();
        assert_eq!(error.kind, ParseErrorKind::Syntax);
        assert_eq!(error.offset, start);
        assert!(error.message.contains("UTF-8 boundary"));
    }
}

#[test]
fn stack_predicates_trivia_and_work_limits_use_absolute_offsets() {
    let grammar = compile(r#"WHITESPACE = _{ " " } entry = { PUSH("a") ~ &POP ~ POP ~ "b" ~ EOI }"#).unwrap();
    for source in ["aab", "a a b", "aac", "a"] {
        let full = format!("λ:{source}");
        for max_steps in 0..120 {
            let options = MatchOptions { max_steps, ..MatchOptions::default() };
            let baseline = grammar.parse_source_with_options("entry", source, options);
            let at = parse_program_at(&grammar, "entry", &full, 3, options);
            match (baseline, at) {
                (Ok(a), Ok(b)) => { assert_eq!(a.consumed() + 3, b.consumed()); assert_eq!(a.capture_count(), b.capture_count()); }
                (Err(a), Err(b)) => { assert_eq!(a.kind, b.kind); assert_eq!(a.offset + 3, b.offset); assert_eq!(a.expected, b.expected); assert_eq!(a.message, b.message); }
                _ => panic!("offset parse differs for {source:?}, budget={max_steps}"),
            }
        }
    }
}

#[test]
fn default_entry_errors_and_depth_limits_remain_located() {
    let grammar = compile(r#"entry = { nested } nested = { "a" }"#).unwrap();
    let unknown = parse_program_at(&grammar, "missing", "xa", 1, MatchOptions::default()).unwrap_err();
    assert_eq!(unknown.kind, ParseErrorKind::UnknownEntry);
    assert_eq!(unknown.offset, 1);
    let depth = parse_program_at(&grammar, "entry", "xa", 1, MatchOptions { max_rule_depth: 1, ..MatchOptions::default() }).unwrap_err();
    assert_eq!(depth.kind, ParseErrorKind::DepthLimit);
    assert_eq!(depth.offset, 1);
}
