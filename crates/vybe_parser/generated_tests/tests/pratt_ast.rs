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
