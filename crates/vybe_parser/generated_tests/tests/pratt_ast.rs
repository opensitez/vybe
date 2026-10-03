use vybe_ast::{BinOp, ExprKind, Expression, Literal, UnaryOp};
use vybe_parser::{
    MatchOptions, Span,
    builder::Builder,
    engine::build_program,
    pratt::{self, Build, Error, Fixity, Limits, Operator, Table, Token, TokenKind},
};
use vybe_parser_generated_tests::expression::{Parser, Rule};

const OPS: &[Operator<'static>] = &[
    Operator {
        symbol: "+",
        precedence: 10,
        fixity: Fixity::InfixLeft,
    },
    Operator {
        symbol: "-",
        precedence: 10,
        fixity: Fixity::InfixLeft,
    },
    Operator {
        symbol: "*",
        precedence: 20,
        fixity: Fixity::InfixLeft,
    },
    Operator {
        symbol: "-",
        precedence: 25,
        fixity: Fixity::Prefix,
    },
    Operator {
        symbol: "^",
        precedence: 30,
        fixity: Fixity::InfixRight,
    },
];
#[derive(Default)]
struct Tokens(Vec<Token<'static>>);
impl Builder for Tokens {
    fn checkpoint(&self) -> usize {
        self.0.len()
    }
    fn rollback(&mut self, mark: usize) {
        self.0.truncate(mark);
    }
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, rule: usize, span: Span, source: &str) {
        let kind = match Rule::from_capture(vybe_parser::CaptureRule::User(rule)) {
            Some(Rule::number) => TokenKind::Atom,
            Some(Rule::open) => TokenKind::Open,
            Some(Rule::close) => TokenKind::Close,
            Some(Rule::operator) => TokenKind::Symbol(match &source[span.start..span.end] {
                "+" => "+",
                "-" => "-",
                "*" => "*",
                "^" => "^",
                _ => unreachable!(),
            }),
            _ => return,
        };
        self.0.push(Token { kind, span });
    }
}
struct Ast<'a> {
    source: &'a str,
    index: vybe_parser::source::Index<'a>,
}
impl Build for Ast<'_> {
    type Value = Expression;
    fn checkpoint(&self) -> usize {
        0
    }
    fn rollback(&mut self, _: usize) {}
    fn atom(&mut self, span: Span) -> Result<Expression, Error> {
        let value = self.source[span.start..span.end]
            .parse::<i64>()
            .map_err(|_| Error {
                span,
                message: "integer literal is out of range",
            })?;
        let start = self
            .index
            .position(span.start, vybe_parser::source::Encoding::Scalar)
            .unwrap();
        let end = self
            .index
            .position(span.end, vybe_parser::source::Encoding::Scalar)
            .unwrap();
        Ok(Expression::with_span(
            ExprKind::Lit(Literal::Int(value)),
            vybe_ast::Span {
                start_line: start.line + 1,
                start_col: start.character + 1,
                end_line: end.line + 1,
                end_col: end.character + 1,
            },
        ))
    }
    fn prefix(&mut self, _: usize, span: Span, right: Expression) -> Result<Expression, Error> {
        let start = self
            .index
            .position(span.start, vybe_parser::source::Encoding::Scalar)
            .unwrap();
        let range = vybe_ast::Span {
            start_line: start.line + 1,
            start_col: start.character + 1,
            ..right.span
        };
        Ok(Expression::with_span(
            ExprKind::Unary {
                op: UnaryOp::Neg,
                expr: Box::new(right),
            },
            range,
        ))
    }
    fn infix(
        &mut self,
        id: usize,
        _: Span,
        left: Expression,
        right: Expression,
    ) -> Result<Expression, Error> {
        let op = match OPS[id].symbol {
            "+" => BinOp::Add,
            "-" => BinOp::Sub,
            "*" => BinOp::Mul,
            "^" => BinOp::Pow,
            _ => unreachable!(),
        };
        let range = vybe_ast::Span {
            start_line: left.span.start_line,
            start_col: left.span.start_col,
            end_line: right.span.end_line,
            end_col: right.span.end_col,
        };
        Ok(Expression::with_span(
            ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            },
            range,
        ))
    }
    fn postfix(&mut self, _: usize, span: Span, _: Expression) -> Result<Expression, Error> {
        Err(Error {
            span,
            message: "no postfix operator in this fixture",
        })
    }
}
fn expression(source: &str) -> Result<Expression, Error> {
    let mut tokens = Tokens::default();
    build_program(
        &Parser,
        "expression",
        source,
        MatchOptions::default(),
        &mut tokens,
    )
    .unwrap();
    pratt::parse(
        &Table::new(OPS).unwrap(),
        &tokens.0,
        &mut Ast {
            source,
            index: vybe_parser::source::Index::new(source),
        },
        Limits::default(),
    )
}
fn shape(expression: &Expression) -> String {
    match &expression.kind {
        ExprKind::Lit(Literal::Int(n)) => n.to_string(),
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => format!("(neg {})", shape(expr)),
        ExprKind::Binary { op, left, right } => {
            format!("({op:?} {} {})", shape(left), shape(right))
        }
        _ => panic!("unexpected AST"),
    }
}

#[test]
fn generated_grammar_tokens_feed_pratt_directly_into_common_ast() {
    assert_eq!(
        shape(&expression("-2 ^ 2 + 3 * 4").unwrap()),
        "(Add (neg (Pow 2 2)) (Mul 3 4))"
    );
    assert_eq!(
        shape(&expression("2 ^ 3 ^ 4").unwrap()),
        "(Pow 2 (Pow 3 4))"
    );
    assert_eq!(
        shape(&expression("(1 + 2) * 3").unwrap()),
        "(Mul (Add 1 2) 3)"
    );
    let span = expression("1 +\n 23").unwrap().span;
    assert_eq!(
        (span.start_line, span.start_col, span.end_line, span.end_col),
        (1, 1, 2, 4)
    );
    assert!(expression("1 +").is_err());
    assert!(expression("999999999999999999999999999999").is_err());
}

/// Native-source path: no token vector, capture tree, AST cloning or final walk.
/// Composition inverses restore moved common-AST children on speculation.
struct NativeAst<'a> {
    ast: Ast<'a>,
    values: vybe_parser::arena::OwnedArena<Expression>,
}
impl Builder for NativeAst<'_> {
    fn supports_pratt(&self, rule: usize) -> bool {
        use vybe_parser::program::Program;
        Some(rule) == vybe_parser_generated_tests::islands::Parser.rule_id("expr")
    }
    fn checkpoint(&self) -> usize {
        self.values.checkpoint()
    }
    fn rollback(&mut self, mark: usize) {
        self.values.rollback(mark);
    }
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, _: usize, _: Span, _: &str) {}
    fn try_finish(
        &mut self,
        rule: usize,
        span: Span,
        _: &str,
    ) -> Result<(), vybe_parser::builder::BuildError> {
        use vybe_parser::program::Program;
        if Some(rule) == vybe_parser_generated_tests::islands::Parser.rule_id("number") {
            let value = self
                .ast
                .atom(span)
                .map_err(|e| vybe_parser::builder::BuildError {
                    span: e.span,
                    message: e.message.into(),
                })?;
            let id = self.values.alloc(value);
            self.values.push(id);
        }
        Ok(())
    }
    fn try_reduce(
        &mut self,
        rule: usize,
        fixity: Fixity,
        span: Span,
    ) -> Result<(), vybe_parser::builder::BuildError> {
        use vybe_parser::program::Program;
        fn split_prefix(parent: Expression) -> Expression {
            match parent.kind {
                ExprKind::Unary { expr, .. } => *expr,
                _ => unreachable!("invalid composition inverse"),
            }
        }
        fn split_binary(parent: Expression) -> (Expression, Expression) {
            match parent.kind {
                ExprKind::Binary { left, right, .. } => (*left, *right),
                _ => unreachable!("invalid composition inverse"),
            }
        }
        if fixity == Fixity::Postfix {
            return Err(vybe_parser::builder::BuildError {
                span,
                message: "postfix is outside this common AST fixture".into(),
            });
        }
        let right = self.values.pop().unwrap();
        let id = match fixity {
            Fixity::Prefix => {
                let start = self
                    .ast
                    .index
                    .position(span.start, vybe_parser::source::Encoding::Scalar)
                    .unwrap();
                self.values.unary(
                    right,
                    |right| {
                        let range = vybe_ast::Span {
                            start_line: start.line + 1,
                            start_col: start.character + 1,
                            ..right.span
                        };
                        Expression::with_span(
                            ExprKind::Unary {
                                op: UnaryOp::Neg,
                                expr: Box::new(right),
                            },
                            range,
                        )
                    },
                    split_prefix,
                )
            }
            _ => {
                let left = self.values.pop().unwrap();
                let name = vybe_parser_generated_tests::islands::Parser.rule(rule).name;
                let op = match name {
                    "plus" => BinOp::Add,
                    "minus" => BinOp::Sub,
                    "star" => BinOp::Mul,
                    "power" => BinOp::Pow,
                    _ => unreachable!(),
                };
                self.values.binary(
                    left,
                    right,
                    |left, right| {
                        let range = vybe_ast::Span {
                            start_line: left.span.start_line,
                            start_col: left.span.start_col,
                            end_line: right.span.end_line,
                            end_col: right.span.end_col,
                        };
                        Expression::with_span(
                            ExprKind::Binary {
                                op,
                                left: Box::new(left),
                                right: Box::new(right),
                            },
                            range,
                        )
                    },
                    split_binary,
                )
            }
        };
        self.values.push(id);
        Ok(())
    }
}

#[test]
fn generated_native_pratt_constructs_common_ast_and_restores_failed_reductions() {
    use vybe_parser::program::Program;
    use vybe_parser_generated_tests::islands::Parser;
    let target = Parser.rule_id("expr").unwrap();
    assert_eq!(Parser.source_pratt(target).unwrap().operators.len(), 6);
    for (source, expected) in [
        ("-2 ^ 2 + 3 * 4", "(Add (neg (Pow 2 2)) (Mul 3 4))"),
        ("2 ^ 3 ^ 4", "(Pow 2 (Pow 3 4))"),
        ("(1 + 2) * 3", "(Mul (Add 1 2) 3)"),
        ("1 +\n 23", "(Add 1 23)"),
    ] {
        let mut builder = NativeAst {
            ast: Ast {
                source,
                index: vybe_parser::source::Index::new(source),
            },
            values: Default::default(),
        };
        // First alternative builds/reduces then fails; lookahead suppresses
        // construction. Both paths must leave exactly one common AST root.
        for entry in ["fallback", "probe"] {
            builder.values = Default::default();
            builder.ast = Ast {
                source,
                index: vybe_parser::source::Index::new(source),
            };
            build_program(
                &Parser,
                entry,
                source,
                MatchOptions::default(),
                &mut builder,
            )
            .unwrap();
            assert_eq!(builder.values.values().len(), 1);
            let expression = builder.values.get(builder.values.values()[0]);
            assert_eq!(shape(expression), expected);
            if source.contains('\n') {
                assert_eq!((expression.span.end_line, expression.span.end_col), (2, 4));
            }
            let mark = builder.checkpoint();
            builder.ast = Ast {
                source: "1*2+",
                index: vybe_parser::source::Index::new("1*2+"),
            };
            assert!(
                build_program(
                    &Parser,
                    "program",
                    "1*2+",
                    MatchOptions::default(),
                    &mut builder
                )
                .is_err()
            );
            assert_eq!(builder.checkpoint(), mark);
            assert_eq!(builder.values.values().len(), 1);
            assert_eq!(
                shape(builder.values.get(builder.values.values()[0])),
                expected
            );
            let arena = std::mem::take(&mut builder.values);
            let root = arena.values()[0];
            let owned_expression: Expression = arena.into_root(root);
            assert_eq!(shape(&owned_expression), expected);
        }
    }
}

#[test]
fn native_common_ast_errors_are_located_and_transactional() {
    use vybe_parser_generated_tests::islands::Parser;
    for (source, offset, message) in [
        (
            "1 + 999999999999999999999999",
            4,
            "integer literal is out of range",
        ),
        ("1!", 1, "postfix is outside this common AST fixture"),
    ] {
        let mut builder = NativeAst {
            ast: Ast {
                source,
                index: vybe_parser::source::Index::new(source),
            },
            values: Default::default(),
        };
        let sentinel = builder.values.alloc(Expression::int(42));
        builder.values.push(sentinel);
        let mark = builder.checkpoint();
        let error = build_program(
            &Parser,
            "program",
            source,
            MatchOptions::default(),
            &mut builder,
        )
        .unwrap_err();
        assert_eq!(error.kind, vybe_parser::ParseErrorKind::BuildFailure);
        assert_eq!(error.offset, offset);
        assert_eq!(error.message, message);
        assert_eq!(builder.checkpoint(), mark);
        assert_eq!(builder.values.values(), &[sentinel]);
        assert_eq!(shape(builder.values.get(sentinel)), "42");
    }
}
