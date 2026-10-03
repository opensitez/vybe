use vybe_ast::{ExprKind, Expression, Literal, Statement, StmtKind};
use vybe_parser::program::Program;
use vybe_parser::{MatchOptions, Span, builder::Builder, engine::build_program};
use vybe_parser_generated_tests::ast::{Parser, Rule};

enum Undo {
    PopValue,
    RestoreValue(Expression),
    PopStatement,
}
#[derive(Default)]
struct AstBuilder {
    values: Vec<Expression>,
    body: Vec<Statement>,
    undo: Vec<Undo>,
    number: usize,
    statement: usize,
}
impl AstBuilder {
    fn new() -> Self {
        Self {
            number: Parser.rule_id(Rule::number.name()).unwrap(),
            statement: Parser.rule_id(Rule::return_stmt.name()).unwrap(),
            ..Self::default()
        }
    }
}
impl Builder for AstBuilder {
    fn checkpoint(&self) -> usize {
        self.undo.len()
    }
    fn rollback(&mut self, mark: usize) {
        while self.undo.len() > mark {
            match self.undo.pop().unwrap() {
                Undo::PopValue => {
                    self.values.pop().unwrap();
                }
                Undo::RestoreValue(value) => self.values.push(value),
                Undo::PopStatement => {
                    self.body.pop().unwrap();
                }
            }
        }
    }
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, rule: usize, span: Span, source: &str) {
        if rule == self.number {
            let value = source[span.start..span.end].parse::<i64>().unwrap();
            self.values.push(Expression::int(value));
            self.undo.push(Undo::PopValue);
        } else if rule == self.statement {
            let value = self.values.pop().unwrap();
            // Test adapter retains an undo value. Production builders can use
            // persistent arena IDs to avoid cloning owned common AST nodes.
            self.undo.push(Undo::RestoreValue(value.clone()));
            self.body
                .push(Statement::new(StmtKind::Return(Some(value))));
            self.undo.push(Undo::PopStatement);
        }
    }
}

#[test]
fn common_ast_is_built_during_matching_without_capture_tree_or_walker() {
    let mut builder = AstBuilder::new();
    let source = "return 7; return 42;";
    let matched = build_program(
        &Parser,
        "program",
        source,
        MatchOptions::default(),
        &mut builder,
    )
    .unwrap();
    assert_eq!(matched.consumed, source.len());
    assert!(builder.values.is_empty());
    assert_eq!(builder.body.len(), 2);
    for (statement, expected) in builder.body.iter().zip([7, 42]) {
        assert!(
            matches!(&statement.kind, StmtKind::Return(Some(Expression { kind: ExprKind::Lit(Literal::Int(value)), .. })) if *value == expected)
        );
    }
    // Both lookahead and the failed branch would create duplicate values and
    // statements if either suppression or transactional rollback were broken.
}

#[test]
fn syntax_and_resource_failures_restore_existing_builder_state() {
    let mut builder = AstBuilder::new();
    build_program(
        &Parser,
        "program",
        "return 3;",
        MatchOptions::default(),
        &mut builder,
    )
    .unwrap();
    let mark = builder.checkpoint();
    for (source, options) in [
        ("return 8; return", MatchOptions::default()),
        (
            "return 8;",
            MatchOptions {
                max_steps: 15,
                ..MatchOptions::default()
            },
        ),
    ] {
        assert!(build_program(&Parser, "program", source, options, &mut builder).is_err());
        assert_eq!(builder.checkpoint(), mark);
        assert_eq!(builder.body.len(), 1);
        assert!(builder.values.is_empty());
    }
}
