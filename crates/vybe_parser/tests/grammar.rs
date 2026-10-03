use vybe_parser::{
    Builtin, ExprKind, ParseOptions, Reference, RuleMode, compile, parse, parse_with_options,
};

fn root(source: &str) -> (vybe_parser::GrammarSyntax, usize) {
    let syntax = parse(source).unwrap();
    let expression = syntax.rules[0].expression;
    (syntax, expression)
}

#[test]
fn grammar_precedence_is_postfix_then_predicate_then_sequence_then_choice() {
    let (syntax, id) = root(r#"main = { &"a"+ ~ "b" | "c" }"#);
    let ExprKind::Choice { left, right } = syntax.expressions[id].kind else {
        panic!("expected choice")
    };
    assert!(
        matches!(syntax.expressions[right].kind, ExprKind::Literal { ref text, .. } if text == "c")
    );
    let ExprKind::Sequence { left, right } = syntax.expressions[left].kind else {
        panic!("expected sequence")
    };
    assert!(
        matches!(syntax.expressions[right].kind, ExprKind::Literal { ref text, .. } if text == "b")
    );
    let ExprKind::Predicate {
        expression,
        positive: true,
    } = syntax.expressions[left].kind
    else {
        panic!("expected predicate")
    };
    assert!(matches!(
        syntax.expressions[expression].kind,
        ExprKind::Repeat {
            min: 1,
            max: None,
            ..
        }
    ));
}

#[test]
fn grouping_changes_precedence_and_preserves_original_ranges() {
    let source = r#"main = { ("a" | "b") ~ "c" }"#;
    let (syntax, id) = root(source);
    let ExprKind::Sequence { left, .. } = syntax.expressions[id].kind else {
        panic!()
    };
    let group = &syntax.expressions[left];
    assert_eq!(&source[group.span.start..group.span.end], r#"("a" | "b")"#);
    let ExprKind::Group(child) = group.kind else {
        panic!()
    };
    assert!(matches!(
        syntax.expressions[child].kind,
        ExprKind::Choice { .. }
    ));
    assert_eq!(
        &source[syntax.rules[0].span.start..syntax.rules[0].span.end],
        source
    );
}

#[test]
fn all_rule_modes_and_ascii_case_insensitive_literal() {
    let syntax = parse(r#"a = { "a" } b = _{ a } c = @{ a } d = ${ a } e = !{ ^"IF" }"#).unwrap();
    let modes: Vec<_> = syntax.rules.iter().map(|rule| rule.mode).collect();
    assert_eq!(
        modes,
        [
            RuleMode::Normal,
            RuleMode::Silent,
            RuleMode::Atomic,
            RuleMode::CompoundAtomic,
            RuleMode::NonAtomic
        ]
    );
    assert!(
        matches!(syntax.expressions[syntax.rules[4].expression].kind, ExprKind::Literal { insensitive: true, ref text } if text == "IF")
    );
}

#[test]
fn decoded_literals_and_unicode_character_ranges() {
    let (syntax, id) = root(r#"main = { "\n\r\t\\\"\'\x41\u{1f600}" ~ '\u{03b1}'..'ω' }"#);
    let ExprKind::Sequence { left, right } = syntax.expressions[id].kind else {
        panic!()
    };
    assert_eq!(
        syntax.expressions[left].kind,
        ExprKind::Literal {
            text: "\n\r\t\\\"'A😀".into(),
            insensitive: false
        }
    );
    assert_eq!(
        syntax.expressions[right].kind,
        ExprKind::Range {
            start: 'α',
            end: 'ω'
        }
    );
}

#[test]
fn every_repeat_form_keeps_compact_bounds() {
    for (suffix, min, max) in [
        ("*", 0, None),
        ("+", 1, None),
        ("?", 0, Some(1)),
        ("{4}", 4, Some(4)),
        ("{2,5}", 2, Some(5)),
        ("{,5}", 0, Some(5)),
        ("{2,}", 2, None),
        ("{1000000000}", 1_000_000_000, Some(1_000_000_000)),
    ] {
        let (syntax, id) = root(&format!("main = {{ ANY{suffix} }}"));
        assert_eq!(syntax.expressions.len(), 2);
        assert!(
            matches!(syntax.expressions[id].kind, ExprKind::Repeat { min: actual_min, max: actual_max, .. } if actual_min == min && actual_max == max)
        );
    }
}

#[test]
fn stack_operations_are_structured_and_resolved() {
    let grammar = compile(r##"main = { PUSH("#"*) ~ PUSH_LITERAL(")") ~ PEEK[-2..] ~ PEEK[0..1] ~ POP ~ PEEK_ALL ~ POP_ALL ~ DROP }"##).unwrap();
    let expressions = &grammar.syntax().expressions;
    assert!(
        expressions
            .iter()
            .any(|e| matches!(e.kind, ExprKind::Push(_)))
    );
    assert!(
        expressions
            .iter()
            .any(|e| matches!(e.kind, ExprKind::PushLiteral(ref value) if value == ")"))
    );
    assert!(expressions.iter().any(|e| e.kind
        == ExprKind::PeekSlice {
            start: -2,
            end: None
        }));
    assert!(expressions.iter().any(|e| e.kind
        == ExprKind::PeekSlice {
            start: 0,
            end: Some(1)
        }));
    for (id, expr) in expressions.iter().enumerate() {
        if matches!(expr.kind, ExprKind::Reference(ref name) if name == "POP") {
            assert_eq!(
                grammar.reference(id),
                Some(Reference::Builtin(Builtin::Pop))
            );
        }
    }
}

#[test]
fn forward_references_and_user_newline_rule_take_precedence() {
    let grammar =
        compile(r#"main = { word ~ NEWLINE ~ EOI } word = { ASCII_ALPHA+ } NEWLINE = { "\n" }"#)
            .unwrap();
    for (id, expr) in grammar.syntax().expressions.iter().enumerate() {
        if let ExprKind::Reference(name) = &expr.kind {
            let expected = match name.as_str() {
                "word" => Reference::Rule(1),
                "NEWLINE" => Reference::Rule(2),
                "EOI" => Reference::Builtin(Builtin::Eoi),
                "ASCII_ALPHA" => Reference::Builtin(Builtin::AsciiAlpha),
                _ => panic!(),
            };
            assert_eq!(grammar.reference(id), Some(expected));
        }
    }
    assert_eq!(grammar.rule_id("word"), Some(1));
    assert_eq!(grammar.rule_id("missing"), None);
}

#[test]
fn duplicate_and_undefined_rules_have_exact_primary_and_related_ranges() {
    let source = "a = { missing }\na = { another }";
    let errors = compile(source).unwrap_err();
    assert_eq!(errors.len(), 3);
    let duplicate = &errors[0];
    assert_eq!(duplicate.code, "G009");
    assert_eq!(&source[duplicate.span.start..duplicate.span.end], "a");
    assert_eq!(duplicate.related[0].0.start, 0);
    for (error, name) in errors[1..].iter().zip(["missing", "another"]) {
        assert_eq!(error.code, "G010");
        assert_eq!(&source[error.span.start..error.span.end], name);
    }
}

#[test]
fn comments_are_trivia_only_outside_literals() {
    let grammar = compile(
        "/* outer /* inner */ end */\na = { \"//not-comment\" ~ \"/*literal*/\" } // tail\r\n",
    )
    .unwrap();
    assert_eq!(grammar.syntax().rules.len(), 1);
    assert_eq!(grammar.syntax().expressions.len(), 3);
}

#[test]
fn malformed_grammar_reports_errors_without_panicking() {
    for (source, code) in [
        ("", "G005"),
        ("/*", "G002"),
        ("a = { \"unfinished", "G002"),
        (r#"a = { "\q" }"#, "G003"),
        (r#"a = { "\u{d800}" }"#, "G003"),
        (r#"a = { "\u{110000}" }"#, "G003"),
        (r#"a = { "\x80" }"#, "G003"),
        ("a = { 'ab'..'z' }", "G003"),
        ("a = { 'z'..'a' }", "G007"),
        ("a = { ANY{5,2} }", "G007"),
        ("a = { ANY{,} }", "G007"),
        ("a = { ANY{4294967296} }", "G004"),
        ("a = { (ANY }", "G005"),
        ("a = { ^ANY }", "G005"),
        ("a = { PUSH_LITERAL(ANY) }", "G005"),
        ("a = { #name = ANY }", "G008"),
        ("a = { PEEK[2147483648..] }", "G004"),
        ("a = { ANY ~ }", "G005"),
        ("a = { ANY } garbage", "G005"),
    ] {
        let error = parse(source).unwrap_err();
        assert_eq!(error.code, code, "{source}: {error}");
        assert!(error.span.start <= error.span.end && error.span.end <= source.len());
        assert!(
            source.is_char_boundary(error.span.start) && source.is_char_boundary(error.span.end)
        );
    }
}

#[test]
fn nesting_limit_is_a_resource_error_and_flat_chains_are_iterative() {
    let source = format!("a = {{ {}ANY{} }}", "(".repeat(1000), ")".repeat(1000));
    assert_eq!(parse(&source).unwrap_err().code, "G006");
    assert_eq!(
        parse_with_options("a = { ANY }", ParseOptions { max_nesting: 0 })
            .unwrap_err()
            .code,
        "G006"
    );
    let flat = format!("a = {{ {} }}", vec!["ANY"; 10_000].join(" ~ "));
    assert_eq!(parse(&flat).unwrap().expressions.len(), 19_999);
}

#[test]
fn arena_is_topologically_ordered_and_ranges_are_valid_utf8() {
    let source = r#"main = { !("😀" | 'α'..'ω') ~ PUSH(ANY+) ~ PEEK ~ EOI }"#;
    let syntax = parse(source).unwrap();
    for (id, expr) in syntax.expressions.iter().enumerate() {
        assert!(source.get(expr.span.start..expr.span.end).is_some());
        let children: Vec<_> = match expr.kind {
            ExprKind::Sequence { left, right } | ExprKind::Choice { left, right } => {
                vec![left, right]
            }
            ExprKind::Repeat { expression, .. } | ExprKind::Predicate { expression, .. } => {
                vec![expression]
            }
            ExprKind::Group(child) | ExprKind::Push(child) => vec![child],
            _ => vec![],
        };
        for child in children {
            assert!(child < id);
            assert!(expr.span.start <= syntax.expressions[child].span.start);
            assert!(expr.span.end >= syntax.expressions[child].span.end);
        }
    }
}

#[test]
fn diagnostic_rendering_uses_unicode_columns_and_eof_positions() {
    let source = "a = { \"😀\" }\nb = { missing }";
    let error = compile(source).unwrap_err().remove(0);
    assert_eq!(
        error.render("sample.pest", source),
        "sample.pest:2:7: G010: undefined rule or unsupported builtin 'missing'"
    );
    let source = "a = { \"😀\" ~";
    let error = parse(source).unwrap_err();
    assert_eq!(error.span.start, source.len());
    assert!(error.render("sample", source).starts_with("sample:1:12:"));
}

#[test]
fn compilation_is_deterministic_reentrant_and_thread_independent() {
    let source = r#"a = { "α" ~ b } b = { ANY? }"#;
    let first = compile(source).unwrap();
    compile("fragment = { EOI }").unwrap();
    assert_eq!(first.syntax(), compile(source).unwrap().syntax());
    let expected = first.syntax().clone();
    for thread in (0..4)
        .map(|_| std::thread::spawn(move || compile(source).unwrap().syntax().clone()))
        .collect::<Vec<_>>()
    {
        assert_eq!(thread.join().unwrap(), expected);
    }
}
