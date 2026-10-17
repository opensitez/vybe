use grameo::{
    MatchOptions, Span,
    builder::Builder,
    engine,
    program::{CompactScope, Instruction, Program, RuleSpec, Scope},
};
#[derive(Debug)]
struct Certified<'a>(&'a grameo::CompiledGrammar);
impl Program for Certified<'_> {
    fn rule_id(&self, name: &str) -> Option<usize> {
        self.0.rule_id(name)
    }
    fn rule(&self, id: usize) -> RuleSpec<'_> {
        Program::rule(self.0, id)
    }
    fn instruction(&self, id: usize) -> Instruction<'_> {
        Program::instruction(self.0, id)
    }
    fn scope(&self, id: usize) -> Scope {
        Program::scope(self.0, id)
    }
    fn compact_capture_scope(&self, id: usize) -> CompactScope {
        CompactScope::new(id)
    }
}
#[derive(Default)]
struct Events(Vec<(usize, usize, usize)>);
impl Builder for Events {
    fn checkpoint(&self) -> usize {
        self.0.len()
    }
    fn rollback(&mut self, mark: usize) {
        self.0.truncate(mark)
    }
    fn begin(&mut self, rule: usize, start: usize) {
        self.0.push((rule, start, start));
    }
    fn finish(&mut self, rule: usize, span: Span, _: &str) {
        self.0.push((rule, span.start, span.end));
    }
}
fn shape(tree: grameo::ParseTree<'_, '_>) -> String {
    let mut pending: Vec<_> = tree.pairs().collect();
    let mut nodes = Vec::new();
    while let Some(pair) = pending.pop() {
        nodes.push((pair.as_rule(), pair.as_span()));
        pending.extend(pair.into_inner());
    }
    format!("{nodes:?}")
}
#[test]
fn passive_predicates_and_trivia_restore_stack_without_builder_hooks() {
    for (grammar, source) in [
        (
            r#"entry={SOI~&probe~"a"~!DROP~EOI} probe={PUSH_LITERAL("sentinel")~"a"}"#,
            "a",
        ),
        (
            r#"WHITESPACE={" "~PUSH_LITERAL("w")~" "} entry={SOI~"a"~PEEK~DROP~"b"~EOI}"#,
            "a  wb",
        ),
        (
            r#"WHITESPACE={" "~PUSH_LITERAL("bad")~"x"} entry={SOI~"a"~" "~"b"~!DROP~EOI}"#,
            "a b",
        ),
    ] {
        let grammar = grameo::compile(grammar).unwrap();
        let program = Certified(&grammar);
        let expected =
            engine::parse_program(&program, "entry", source, MatchOptions::default()).unwrap();
        let expected_shape = shape(expected);
        for compact in [false, true] {
            let mut events = Events::default();
            let tree = if compact {
                engine::build_compact_capture_program(
                    &program,
                    "entry",
                    source,
                    MatchOptions::default(),
                    &mut events,
                )
            } else {
                engine::build_capture_program(
                    &program,
                    "entry",
                    source,
                    MatchOptions::default(),
                    &mut events,
                )
            }
            .unwrap();
            assert_eq!(shape(tree), expected_shape);
            assert_eq!(events.0.len(), 2);
            assert_eq!(events.0[0].0, grammar.rule_id("entry").unwrap());
            assert_eq!(events.0[1].0, grammar.rule_id("entry").unwrap());
        }
    }
}

#[test]
fn passive_budget_failures_restore_every_journal() {
    for (grammar, source) in [
        (
            r#"entry={SOI~&probe~"a"~!DROP~EOI} probe={PUSH_LITERAL("sentinel")~"a"}"#,
            "a",
        ),
        (
            r#"WHITESPACE={" "~PUSH_LITERAL("w")~" "} entry={SOI~"a"~PEEK~DROP~"b"~EOI}"#,
            "a  wb",
        ),
        (
            r#"WHITESPACE={" "~PUSH_LITERAL("bad")~"x"} entry={SOI~"a"~" "~"b"~!DROP~EOI}"#,
            "a b",
        ),
    ] {
        let grammar = grameo::compile(grammar).unwrap();
        let program = Certified(&grammar);
        for max_steps in 0..96 {
            for max_rule_depth in 0..6 {
                let options = MatchOptions {
                    max_steps,
                    max_rule_depth,
                };
                let expected = engine::parse_program(&program, "entry", source, options).map(shape);
                for compact in [false, true] {
                    let mut events = Events::default();
                    let actual = if compact {
                        engine::build_compact_capture_program(
                            &program,
                            "entry",
                            source,
                            options,
                            &mut events,
                        )
                    } else {
                        engine::build_capture_program(
                            &program,
                            "entry",
                            source,
                            options,
                            &mut events,
                        )
                    }
                    .map(shape);
                    assert_eq!(
                        actual, expected,
                        "steps={max_steps} depth={max_rule_depth} compact={compact}"
                    );
                    if actual.is_err() {
                        assert!(events.0.is_empty());
                    }
                }
            }
        }
    }
}
