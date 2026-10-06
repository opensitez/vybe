use vybe_ast::{ExprKind, Expression, Lang, Literal, Module, Statement, StmtKind};
use grameo::program::Program;
use grameo::{MatchOptions, Span, builder::Builder, engine::build_program};
use vybe_parser_generated_tests::ast::{Parser, Rule};

enum Value {
    Expression(Expression),
    Statement(Box<Statement>),
    Module(Box<Module>),
}
struct AstBuilder<'a> {
    arena: grameo::arena::OwnedArena<Value>,
    index: grameo::source::Index<'a>,
    number: usize,
    statement: usize,
    program: usize,
}
impl<'a> AstBuilder<'a> {
    fn new(source: &'a str) -> Self {
        Self {
            arena: Default::default(),
            index: grameo::source::Index::new(source),
            number: Parser.rule_id(Rule::number.name()).unwrap(),
            statement: Parser.rule_id(Rule::return_stmt.name()).unwrap(),
            program: Parser.rule_id(Rule::program.name()).unwrap(),
        }
    }
    fn span(&self, span: Span) -> vybe_ast::Span {
        use grameo::source::Encoding;
        let start = self.index.position(span.start, Encoding::Scalar).unwrap();
        let end = self.index.position(span.end, Encoding::Scalar).unwrap();
        vybe_ast::Span {
            start_line: start.line + 1,
            start_col: start.character + 1,
            end_line: end.line + 1,
            end_col: end.character + 1,
        }
    }
}
impl Builder for AstBuilder<'_> {
    fn checkpoint(&self) -> usize {
        self.arena.checkpoint()
    }
    fn rollback(&mut self, mark: usize) {
        self.arena.rollback(mark);
    }
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, _: usize, _: Span, _: &str) {}
    fn try_finish(
        &mut self,
        rule: usize,
        span: Span,
        source: &str,
    ) -> Result<(), grameo::builder::BuildError> {
        if rule == self.number {
            let value = source[span.start..span.end].parse::<i64>().map_err(|_| {
                grameo::builder::BuildError {
                    span,
                    message: "integer literal is out of range".into(),
                }
            })?;
            let value = self.arena.alloc(Value::Expression(Expression::with_span(
                ExprKind::Lit(Literal::Int(value)),
                self.span(span),
            )));
            self.arena.push(value);
        } else if rule == self.statement {
            let value = self.arena.pop().unwrap();
            let range = self.span(span);
            let statement = self.arena.unary(
                value,
                |value| {
                    let Value::Expression(value) = value else {
                        unreachable!("expected expression")
                    };
                    Value::Statement(Box::new(Statement::with_span(
                        StmtKind::Return(Some(value)),
                        range,
                    )))
                },
                |value| {
                    let Value::Statement(statement) = value else {
                        unreachable!("expected return statement")
                    };
                    let StmtKind::Return(Some(value)) = statement.kind else {
                        unreachable!("expected return value")
                    };
                    Value::Expression(value)
                },
            );
            self.arena.push(statement);
        } else if rule == self.program {
            let mut children = Vec::new();
            while self
                .arena
                .values()
                .last()
                .is_some_and(|id| matches!(self.arena.get(*id), Value::Statement(_)))
            {
                children.push(self.arena.pop().unwrap());
            }
            children.reverse();
            let module = self.arena.list(
                children,
                |values| {
                    let body = values
                        .into_iter()
                        .map(|value| {
                            let Value::Statement(statement) = value else {
                                unreachable!("expected statement")
                            };
                            *statement
                        })
                        .collect();
                    Value::Module(Box::new(Module {
                        name: "fixture".into(),
                        language: Lang::Unknown,
                        body,
                        imports: Vec::new(),
                        directives: Default::default(),
                        canon: Default::default(),
                    }))
                },
                |value| {
                    let Value::Module(module) = value else {
                        unreachable!("expected module")
                    };
                    module
                        .body
                        .into_iter()
                        .map(|statement| Value::Statement(Box::new(statement)))
                        .collect()
                },
            );
            self.arena.push(module);
        }
        Ok(())
    }
}

#[test]
fn common_ast_is_built_during_matching_without_capture_tree_or_walker() {
    let source = "return 7; return 42;";
    let mut builder = AstBuilder::new(source);
    let matched = build_program(
        &Parser,
        "module_fallback",
        source,
        MatchOptions::default(),
        &mut builder,
    )
    .unwrap();
    assert_eq!(matched.consumed, source.len());
    assert_eq!(builder.arena.values().len(), 1);
    let root = builder.arena.values()[0];
    let Value::Module(module) = builder.arena.into_root(root) else {
        panic!("expected owned module")
    };
    let module: Module = *module;
    assert_eq!(module.body.len(), 2);
    for (statement, expected) in module.body.iter().zip([7, 42]) {
        assert!(
            matches!(&statement.kind, StmtKind::Return(Some(Expression { kind: ExprKind::Lit(Literal::Int(value)), .. })) if *value == expected)
        );
    }
    // Both lookahead and the failed branch would create duplicate values and
    // statements if either suppression or transactional rollback were broken.
}

#[test]
fn syntax_and_resource_failures_restore_existing_builder_state() {
    let mut builder = AstBuilder::new("return 3;");
    build_program(
        &Parser,
        "program",
        "return 3;",
        MatchOptions::default(),
        &mut builder,
    )
    .unwrap();
    let mark = builder.checkpoint();
    for (entry, source, options) in [
        ("program", "return 8; return", MatchOptions::default()),
        ("module_failure", "return 8;", MatchOptions::default()),
        (
            "program",
            "return 999999999999999999999999;",
            MatchOptions::default(),
        ),
        (
            "program",
            "return 8;",
            MatchOptions {
                max_steps: 15,
                ..MatchOptions::default()
            },
        ),
    ] {
        builder.index = grameo::source::Index::new(source);
        assert!(build_program(&Parser, entry, source, options, &mut builder).is_err());
        assert_eq!(builder.checkpoint(), mark);
        assert_eq!(builder.arena.values().len(), 1);
        let Value::Module(module) = builder.arena.get(builder.arena.values()[0]) else {
            panic!("expected module")
        };
        assert_eq!(module.body.len(), 1);
        assert!(matches!(
            &module.body[0].kind,
            StmtKind::Return(Some(Expression {
                kind: ExprKind::Lit(Literal::Int(3)),
                ..
            }))
        ));
    }
    let root = builder.arena.values()[0];
    let Value::Module(module) = builder.arena.into_root(root) else {
        panic!("expected owned module")
    };
    let module: Module = *module;
    let statement = module.body.into_iter().next().unwrap();
    assert!(matches!(
        statement.kind,
        StmtKind::Return(Some(Expression {
            kind: ExprKind::Lit(Literal::Int(3)),
            ..
        }))
    ));
}

#[test]
fn direct_owned_module_preserves_large_body_order_and_indexed_spans() {
    for source in [
        String::new(),
        (0..200)
            .map(|value| format!("return {value};\n"))
            .collect::<String>(),
    ] {
        let mut builder = AstBuilder::new(&source);
        build_program(
            &Parser,
            "module_fallback",
            &source,
            MatchOptions::default(),
            &mut builder,
        )
        .unwrap();
        let root = builder.arena.values()[0];
        let Value::Module(module) = builder.arena.into_root(root) else {
            panic!("expected module")
        };
        assert_eq!(module.name, "fixture");
        assert_eq!(module.language, Lang::Unknown);
        assert!(module.imports.is_empty());
        assert_eq!(module.body.len(), if source.is_empty() { 0 } else { 200 });
        for (index, statement) in module.body.iter().enumerate() {
            assert_eq!(
                (statement.span.start_line, statement.span.start_col),
                (index as u32 + 1, 1)
            );
            let StmtKind::Return(Some(expression)) = &statement.kind else {
                panic!("expected return statement")
            };
            assert!(
                matches!(expression.kind, ExprKind::Lit(Literal::Int(value)) if value == index as i64)
            );
            assert_eq!(
                (expression.span.start_line, expression.span.start_col),
                (index as u32 + 1, 8)
            );
            assert_eq!(
                (expression.span.end_line, expression.span.end_col),
                (index as u32 + 1, 8 + index.to_string().len() as u32)
            );
        }
    }
}
