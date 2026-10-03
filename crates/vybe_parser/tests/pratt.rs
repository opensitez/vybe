use vybe_parser::{
    Span,
    pratt::{self, Build, Error, Fixity, Limits, Operator, Table, Token, TokenKind},
};

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
    Operator {
        symbol: "!",
        precedence: 40,
        fixity: Fixity::Postfix,
    },
];
struct Shape<'a>(&'a str);
impl Build for Shape<'_> {
    type Value = String;
    fn checkpoint(&self) -> usize {
        0
    }
    fn rollback(&mut self, _: usize) {}
    fn atom(&mut self, span: Span) -> Result<String, Error> {
        Ok(self.0[span.start..span.end].into())
    }
    fn prefix(&mut self, _: usize, _: Span, right: String) -> Result<String, Error> {
        Ok(format!("(neg {right})"))
    }
    fn infix(
        &mut self,
        operator: usize,
        _: Span,
        left: String,
        right: String,
    ) -> Result<String, Error> {
        Ok(format!("({} {left} {right})", OPS[operator].symbol))
    }
    fn postfix(&mut self, _: usize, _: Span, left: String) -> Result<String, Error> {
        Ok(format!("(! {left})"))
    }
}
fn tokens(source: &str) -> Vec<Token<'_>> {
    source
        .char_indices()
        .filter(|(_, ch)| !ch.is_whitespace())
        .map(|(start, ch)| {
            let span = Span {
                start,
                end: start + ch.len_utf8(),
            };
            Token {
                span,
                kind: match ch {
                    '(' => TokenKind::Open,
                    ')' => TokenKind::Close,
                    '0'..='9' => TokenKind::Atom,
                    _ => TokenKind::Symbol(&source[span.start..span.end]),
                },
            }
        })
        .collect()
}
fn shape(source: &str) -> Result<String, Error> {
    pratt::parse(
        &Table::new(OPS).unwrap(),
        &tokens(source),
        &mut Shape(source),
        Limits::default(),
    )
}

#[test]
fn precedence_associativity_prefix_binding_postfix_and_groups() {
    for (source, expected) in [
        ("1+2*3", "(+ 1 (* 2 3))"),
        ("1-2-3", "(- (- 1 2) 3)"),
        ("2^3^4", "(^ 2 (^ 3 4))"),
        ("-2^2", "(neg (^ 2 2))"),
        ("2^-3", "(^ 2 (neg 3))"),
        ("(-2)^2", "(^ (neg 2) 2)"),
        ("--2!!", "(neg (neg (! (! 2))))"),
        ("(1+2)*3!", "(* (+ 1 2) (! 3))"),
    ] {
        assert_eq!(shape(source).unwrap(), expected, "{source}");
    }
}

#[test]
fn malformed_expressions_return_located_errors_without_panicking() {
    for source in [
        "", "()", "1 2", "1+", "*1", "1)", "(1", "1+(2*)", "1(2)", "!1", "1?2",
    ] {
        let error = shape(source).unwrap_err();
        assert!(
            error.span.start <= error.span.end && error.span.end <= source.len(),
            "{source}"
        );
    }
    assert_eq!(shape("(1").unwrap_err().span, Span { start: 0, end: 1 });
    assert_eq!(shape("1+").unwrap_err().span, Span { start: 2, end: 2 });
}

#[test]
fn grouping_and_prefix_nesting_use_explicit_stack_and_obey_limits() {
    let source = format!("{}1{}", "(".repeat(1000), ")".repeat(1000));
    assert_eq!(shape(&source).unwrap(), "1");
    let tokens = tokens(&source);
    let table = Table::new(OPS).unwrap();
    assert_eq!(
        pratt::parse(
            &table,
            &tokens,
            &mut Shape(&source),
            Limits {
                max_stack: 100,
                ..Limits::default()
            }
        )
        .unwrap_err()
        .message,
        "expression stack limit exceeded"
    );
    assert_eq!(
        pratt::parse(
            &table,
            &tokens,
            &mut Shape(&source),
            Limits {
                max_tokens: 100,
                ..Limits::default()
            }
        )
        .unwrap_err()
        .message,
        "expression token limit exceeded"
    );
}

#[test]
fn operator_tables_reject_ambiguous_following_operators_but_allow_contextual_prefix() {
    assert!(Table::new(OPS).is_ok());
    assert!(
        Table::new(&[
            Operator {
                symbol: "+",
                precedence: 1,
                fixity: Fixity::InfixLeft
            },
            Operator {
                symbol: "+",
                precedence: 1,
                fixity: Fixity::Postfix
            },
        ])
        .is_err()
    );
    assert!(
        Table::new(&[Operator {
            symbol: "",
            precedence: 1,
            fixity: Fixity::Prefix
        }])
        .is_err()
    );
}

#[test]
fn syntax_resource_and_semantic_errors_restore_the_builder_journal() {
    struct Journal(usize);
    impl Build for Journal {
        type Value = usize;
        fn checkpoint(&self) -> usize {
            self.0
        }
        fn rollback(&mut self, mark: usize) {
            self.0 = mark;
        }
        fn atom(&mut self, span: Span) -> Result<usize, Error> {
            self.0 += 1;
            Ok(span.start)
        }
        fn prefix(&mut self, _: usize, _: Span, value: usize) -> Result<usize, Error> {
            self.0 += 1;
            Ok(value)
        }
        fn infix(&mut self, _: usize, span: Span, _: usize, _: usize) -> Result<usize, Error> {
            self.0 += 1;
            Err(Error {
                span,
                message: "semantic binding rejected operands",
            })
        }
        fn postfix(&mut self, _: usize, _: Span, value: usize) -> Result<usize, Error> {
            Ok(value)
        }
    }
    let mut builder = Journal(7);
    let table = Table::new(OPS).unwrap();
    for (source, limits) in [
        ("1+", Limits::default()),
        ("1+2", Limits::default()),
        (
            "1+((2))",
            Limits {
                max_stack: 1,
                ..Limits::default()
            },
        ),
    ] {
        assert!(pratt::parse(&table, &tokens(source), &mut builder, limits).is_err());
        assert_eq!(builder.0, 7);
    }
}
