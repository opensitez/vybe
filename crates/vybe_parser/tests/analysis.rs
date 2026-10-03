use vybe_parser::{analysis::Effects, compile, parse, resolve};

#[test]
fn direct_and_indirect_left_recursion_are_located_errors() {
    for source in [
        "a = { a }",
        "a = { b } b = { a }",
        r#"a = { "" ~ a }"#,
        "a = { &a ~ ANY }",
    ] {
        let errors = compile(source).unwrap_err();
        assert!(
            errors.iter().any(|error| error.code == "G011"),
            "{source}: {errors:?}"
        );
        let cycle = errors.iter().find(|error| error.code == "G011").unwrap();
        assert!(source.get(cycle.span.start..cycle.span.end).is_some());
        assert!(!cycle.related.is_empty());
    }
}

#[test]
fn uncertain_recursion_is_warned_and_consuming_recursion_is_accepted() {
    let uncertain = compile(r#"a = { "x" | a }"#).unwrap();
    assert!(
        uncertain
            .analysis()
            .warnings
            .iter()
            .any(|warning| warning.code == "W001")
    );
    let consumes = compile(r#"a = { "x" ~ a? }"#).unwrap();
    assert!(
        !consumes
            .analysis()
            .warnings
            .iter()
            .any(|warning| warning.code == "W001")
    );
    assert_eq!(consumes.recognize("a", "xxx").unwrap().consumed, 3);
}

#[test]
fn definite_empty_loops_are_errors_but_input_dependent_ones_are_warned() {
    for source in [
        r#"a = { ""* }"#,
        r#"a = { empty+ } empty = { "" }"#,
        r#"WHITESPACE = _{ "" } a = { ANY ~ ANY }"#,
    ] {
        assert!(
            compile(source)
                .unwrap_err()
                .iter()
                .any(|error| error.code == "G012")
        );
    }
    let uncertain = compile("a = { (&ANY)* }").unwrap();
    assert!(
        uncertain
            .analysis()
            .warnings
            .iter()
            .any(|warning| warning.code == "W002")
    );
    assert_eq!(
        uncertain.recognize("a", "x").unwrap_err().kind,
        vybe_parser::ParseErrorKind::NonProgress
    );
    assert!(compile(r#"a = { ""{3} }"#).is_ok());
}

#[test]
fn effects_propagate_through_indirect_calls_and_implicit_skip_rules() {
    let grammar =
        compile(r#"WHITESPACE = _{ PUSH(" ") } a = { b ~ b } b = { "a" } c = @{ POP } d = { &c }"#)
            .unwrap();
    let facts = &grammar.analysis().rules;
    let a = facts[grammar.rule_id("a").unwrap()].effects;
    assert!(a.contains(Effects::STACK_WRITE));
    assert!(a.contains(Effects::CAPTURE));
    let d = facts[grammar.rule_id("d").unwrap()].effects;
    assert!(d.contains(Effects::STACK_READ));
    assert!(d.contains(Effects::STACK_WRITE));
    assert!(d.contains(Effects::LOOKAHEAD));
    assert!(d.contains(Effects::MODE_CHANGE));
}

#[test]
fn stack_changing_recursion_is_uncertain_not_unconditionally_rejected() {
    let grammar = compile("a = { DROP ~ a | ANY }").unwrap();
    assert!(grammar.analysis().errors.is_empty());
    assert!(
        grammar
            .analysis()
            .warnings
            .iter()
            .any(|warning| warning.code == "W001")
    );
    assert_eq!(grammar.recognize("a", "x").unwrap().consumed, 1);
}

#[test]
fn low_level_resolution_keeps_analysis_findings_and_order_is_deterministic() {
    let source = r#"a = { a ~ ""* }"#;
    let grammar = resolve(parse(source).unwrap()).unwrap();
    assert_eq!(
        grammar.analysis().errors,
        resolve(parse(source).unwrap()).unwrap().analysis().errors
    );
    assert!(grammar.analysis().errors.len() >= 2);
}
