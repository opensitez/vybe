use grameo::{
    MatchOptions, RuleMode, Span,
    builder::Builder,
    program::{Instruction, Program, RuleSpec},
};
#[derive(Debug)]
struct MaximumRule;
impl Program for MaximumRule {
    fn rule_id(&self, name: &str) -> Option<usize> {
        (name == "entry").then_some(usize::MAX)
    }
    fn rule(&self, id: usize) -> RuleSpec<'_> {
        assert_eq!(id, usize::MAX);
        RuleSpec {
            name: "entry",
            mode: RuleMode::Normal,
            expression: 0,
            scope: 0,
        }
    }
    fn instruction(&self, id: usize) -> Instruction<'_> {
        assert_eq!(id, 0);
        Instruction::Literal {
            text: "a",
            insensitive: false,
        }
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
    fn begin(&mut self, rule: usize, offset: usize) {
        self.0.push((rule, offset, offset))
    }
    fn finish(&mut self, rule: usize, span: Span, _: &str) {
        self.0.push((rule, span.start, span.end))
    }
}
#[test]
fn semantic_completion_retains_maximum_rule_id() {
    let mut events = Events::default();
    let r = grameo::engine::build_program(
        &MaximumRule,
        "entry",
        "a",
        MatchOptions::default(),
        &mut events,
    )
    .unwrap();
    assert_eq!(r.consumed, 1);
    assert_eq!(events.0, [(usize::MAX, 0, 0), (usize::MAX, 0, 1)]);
    let mut events = Events::default();
    assert!(
        grameo::engine::build_program(
            &MaximumRule,
            "entry",
            "b",
            MatchOptions::default(),
            &mut events
        )
        .is_err()
    );
    assert!(events.0.is_empty());
}

#[derive(Debug)]
struct SilentMaximumRule;
impl Program for SilentMaximumRule {
    fn rule_id(&self, name: &str) -> Option<usize> {
        MaximumRule.rule_id(name)
    }
    fn rule(&self, id: usize) -> RuleSpec<'_> {
        let mut rule = MaximumRule.rule(id);
        rule.mode = RuleMode::Silent;
        rule
    }
    fn instruction(&self, id: usize) -> Instruction<'_> {
        MaximumRule.instruction(id)
    }
}
#[test]
fn hybrid_completion_retains_maximum_observed_rule_id() {
    for (source, accepted) in [("a", true), ("b", false)] {
        let mut events = Events::default();
        let result = grameo::engine::build_capture_program(
            &SilentMaximumRule,
            "entry",
            source,
            MatchOptions::default(),
            &mut events,
        );
        assert_eq!(result.is_ok(), accepted);
        if accepted {
            assert_eq!(events.0, [(usize::MAX, 0, 0), (usize::MAX, 0, 1)]);
        } else {
            assert!(events.0.is_empty());
        }
    }
}
