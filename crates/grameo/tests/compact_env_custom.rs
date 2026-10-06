use grameo::{
    MatchOptions, RuleMode, engine,
    program::{Instruction, Program, RuleSpec, Scope},
};
#[derive(Debug)]
struct MaximumScope;
impl Program for MaximumScope {
    fn rule_id(&self, name: &str) -> Option<usize> {
        (name == "entry").then_some(usize::MAX)
    }
    fn rule(&self, id: usize) -> RuleSpec<'_> {
        assert_eq!(id, usize::MAX);
        RuleSpec {
            name: "entry",
            mode: RuleMode::Normal,
            expression: usize::MAX,
            scope: usize::MAX,
        }
    }
    fn instruction(&self, id: usize) -> Instruction<'_> {
        assert_eq!(id, usize::MAX);
        Instruction::Literal {
            text: "x",
            insensitive: false,
        }
    }
    fn scope(&self, id: usize) -> Scope {
        assert!(id == 0 || id == usize::MAX);
        Scope::default()
    }
}
#[test]
fn custom_maximum_scope_stays_wide() {
    let program = MaximumScope;
    for source in ["x", "", "y"] {
        let tree = engine::parse_program(&program, "entry", source, MatchOptions::default());
        let recognized =
            engine::recognize_program(&program, "entry", source, MatchOptions::default());
        assert_eq!(tree.is_ok(), source == "x");
        assert_eq!(tree.is_ok(), recognized.is_ok());
        if let Ok(tree) = tree {
            assert_eq!(tree.consumed(), 1);
        }
    }
}
#[test]
#[should_panic(expected = "scope does not fit compact capture environment")]
fn too_large_compact_scope_is_rejected() {
    let _ = grameo::program::CompactScope::new(usize::MAX);
}
