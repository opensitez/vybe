use vybe_parser::{
    MatchOptions, Span,
    arena::{Arena, Id},
    builder::Builder,
    engine::build_program,
    islands::Operator,
    pratt::Fixity,
};

const GRAMMAR: &str = include_str!("islands.grammar");

fn grammar(bound: bool) -> vybe_parser::CompiledGrammar {
    let mut grammar = vybe_parser::compile(GRAMMAR).unwrap();
    if bound {
        grammar
            .bind_pratt(
                "expr",
                "atom",
                &[
                    Operator {
                        rule: "minus",
                        precedence: 25,
                        fixity: Fixity::Prefix,
                    },
                    Operator {
                        rule: "power",
                        precedence: 30,
                        fixity: Fixity::InfixRight,
                    },
                    Operator {
                        rule: "bang",
                        precedence: 40,
                        fixity: Fixity::Postfix,
                    },
                    Operator {
                        rule: "star",
                        precedence: 20,
                        fixity: Fixity::InfixLeft,
                    },
                    Operator {
                        rule: "plus",
                        precedence: 10,
                        fixity: Fixity::InfixLeft,
                    },
                    Operator {
                        rule: "minus",
                        precedence: 10,
                        fixity: Fixity::InfixLeft,
                    },
                ],
                vybe_parser::islands::TrailingTrivia::Consume,
            )
            .unwrap();
    }
    grammar
}

#[derive(Debug)]
enum Node {
    Number(String),
    Unary(&'static str, Id),
    Binary(&'static str, Id, Id),
}
struct Ast {
    arena: Arena<Node>,
    target: usize,
    number: usize,
    names: Vec<String>,
}
impl Ast {
    fn new(grammar: &vybe_parser::CompiledGrammar) -> Self {
        Self {
            arena: Arena::default(),
            target: grammar.rule_id("expr").unwrap(),
            number: grammar.rule_id("number").unwrap(),
            names: grammar
                .syntax()
                .rules
                .iter()
                .map(|r| r.name.clone())
                .collect(),
        }
    }
    fn shape(&self, id: Id) -> String {
        match self.arena.get(id) {
            Node::Number(n) => n.clone(),
            Node::Unary(op, a) => format!("({op} {})", self.shape(*a)),
            Node::Binary(op, a, b) => format!("({op} {} {})", self.shape(*a), self.shape(*b)),
        }
    }
}
impl Builder for Ast {
    fn supports_pratt(&self, rule: usize) -> bool {
        rule == self.target
    }
    fn checkpoint(&self) -> usize {
        self.arena.checkpoint()
    }
    fn rollback(&mut self, mark: usize) {
        self.arena.rollback(mark);
    }
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, rule: usize, span: Span, source: &str) {
        if rule == self.number {
            let id = self
                .arena
                .alloc(Node::Number(source[span.start..span.end].into()));
            self.arena.push(id);
        }
    }
    fn reduce(&mut self, rule: usize, fixity: Fixity, _: Span) {
        let op = match self.names[rule].as_str() {
            "plus" => "+",
            "minus" if fixity == Fixity::Prefix => "neg",
            "minus" => "-",
            "star" => "*",
            "power" => "^",
            "bang" => "!",
            _ => unreachable!(),
        };
        let right = self.arena.pop().expect("operator RHS");
        let node = match fixity {
            Fixity::Prefix | Fixity::Postfix => Node::Unary(op, right),
            _ => Node::Binary(op, self.arena.pop().expect("operator LHS"), right),
        };
        let id = self.arena.alloc(node);
        self.arena.push(id);
    }
}

#[test]
fn native_atoms_and_operators_build_transactional_ast() {
    let grammar = grammar(true);
    for (source, expected) in [
        ("1+2*3", "(+ 1 (* 2 3))"),
        ("1-2-3", "(- (- 1 2) 3)"),
        ("2^3^4", "(^ 2 (^ 3 4))"),
        ("-2^2", "(neg (^ 2 2))"),
        ("2^-3", "(^ 2 (neg 3))"),
        ("(-2)^2", "(^ (neg 2) 2)"),
        ("--2!!", "(neg (neg (! (! 2))))"),
        (" (1 + 2) * 3! ", "(* (+ 1 2) (! 3))"),
    ] {
        for entry in ["program", "probe", "fallback"] {
            let mut ast = Ast::new(&grammar);
            build_program(&grammar, entry, source, MatchOptions::default(), &mut ast).unwrap();
            assert_eq!(ast.arena.values().len(), 1, "{entry}: {source}");
            assert_eq!(
                ast.shape(ast.arena.values()[0]),
                expected,
                "{entry}: {source}"
            );
        }
    }
}

#[test]
fn failed_rhs_restores_eager_reductions_and_operator_stack() {
    let grammar = grammar(true);
    for (source, consumed, expected) in [
        ("1+", 1, "1"),
        ("1*2+", 3, "(* 1 2)"),
        ("-1^2 + -", 5, "(neg (^ 1 2))"),
        ("1+2*", 3, "(+ 1 2)"),
    ] {
        let mut ast = Ast::new(&grammar);
        let parsed =
            build_program(&grammar, "expr", source, MatchOptions::default(), &mut ast).unwrap();
        assert_eq!(parsed.consumed, consumed, "{source}");
        assert_eq!(ast.arena.values().len(), 1);
        assert_eq!(ast.shape(ast.arena.values()[0]), expected, "{source}");
    }
    for source in ["", "-", "1+", "1+(2*)"] {
        let mut ast = Ast::new(&grammar);
        assert!(
            build_program(
                &grammar,
                "program",
                source,
                MatchOptions::default(),
                &mut ast
            )
            .is_err()
        );
        assert_eq!(ast.arena.node_count(), 0);
        assert!(ast.arena.values().is_empty());
    }
}

#[test]
fn nesting_and_work_limits_restore_semantic_state() {
    let grammar = grammar(true);
    let source = format!("{}1{}", "(".repeat(500), ")".repeat(500));
    assert_eq!(
        grammar.recognize("program", &source).unwrap().consumed,
        source.len()
    );
    for options in [
        MatchOptions {
            max_rule_depth: 20,
            ..MatchOptions::default()
        },
        MatchOptions {
            max_steps: 20,
            ..MatchOptions::default()
        },
    ] {
        let mut ast = Ast::new(&grammar);
        assert!(build_program(&grammar, "program", &source, options, &mut ast).is_err());
        assert_eq!(ast.arena.node_count(), 0);
        assert!(ast.arena.values().is_empty());
    }
    let prefix = format!("{}1", "-".repeat(100));
    assert!(
        grammar
            .recognize_with_options(
                "program",
                &prefix,
                MatchOptions {
                    max_rule_depth: 20,
                    ..MatchOptions::default()
                }
            )
            .is_err()
    );
}

#[test]
fn checked_rule_and_reduction_hooks_restore_even_partially_mutated_state() {
    use vybe_parser::builder::BuildError;
    struct Failing {
        ast: Ast,
        stage: u8,
    }
    impl Builder for Failing {
        fn supports_pratt(&self, rule: usize) -> bool {
            self.ast.supports_pratt(rule)
        }
        fn checkpoint(&self) -> usize {
            self.ast.checkpoint()
        }
        fn rollback(&mut self, mark: usize) {
            self.ast.rollback(mark);
        }
        fn begin(&mut self, _: usize, _: usize) {}
        fn finish(&mut self, rule: usize, span: Span, source: &str) {
            self.ast.finish(rule, span, source);
        }
        fn try_begin(&mut self, rule: usize, offset: usize) -> Result<(), BuildError> {
            if self.stage == 0 && rule == self.ast.number {
                let node = self.ast.arena.alloc(Node::Number("temporary".into()));
                self.ast.arena.push(node);
                return Err(BuildError {
                    span: Span {
                        start: offset,
                        end: offset,
                    },
                    message: "begin failed".into(),
                });
            }
            Ok(())
        }
        fn try_finish(&mut self, rule: usize, span: Span, source: &str) -> Result<(), BuildError> {
            self.finish(rule, span, source);
            if self.stage == 1 && rule == self.ast.number {
                return Err(BuildError {
                    span,
                    message: "finish failed".into(),
                });
            }
            Ok(())
        }
        fn try_reduce(
            &mut self,
            rule: usize,
            fixity: Fixity,
            span: Span,
        ) -> Result<(), BuildError> {
            self.ast.reduce(rule, fixity, span);
            Err(BuildError {
                span,
                message: "reduce failed".into(),
            })
        }
    }
    let grammar = grammar(true);
    for (stage, offset) in [(0, 0), (1, 0), (2, 1)] {
        let mut builder = Failing {
            ast: Ast::new(&grammar),
            stage,
        };
        let sentinel = builder.ast.arena.alloc(Node::Number("99".into()));
        builder.ast.arena.push(sentinel);
        let mark = builder.checkpoint();
        let error = build_program(
            &grammar,
            "fallback",
            "1+2",
            MatchOptions::default(),
            &mut builder,
        )
        .unwrap_err();
        assert_eq!(error.kind, vybe_parser::ParseErrorKind::BuildFailure);
        assert_eq!(error.offset, offset);
        assert_eq!(builder.checkpoint(), mark);
        assert_eq!(builder.ast.arena.node_count(), 1);
        assert_eq!(builder.ast.arena.values(), &[sentinel]);
    }
}

#[test]
fn invalid_bindings_are_rejected_without_replacing_valid_profile() {
    use vybe_parser::{islands::TrailingTrivia, program::Program};
    let mut grammar = grammar(true);
    let target = grammar.rule_id("expr").unwrap();
    for (target_name, atom) in [("unknown", "atom"), ("expr", "expr"), ("expr", "missing")] {
        assert_eq!(
            grammar
                .bind_pratt(target_name, atom, &[], TrailingTrivia::Consume)
                .unwrap_err()[0]
                .code,
            "G014"
        );
    }
    let repeated = Operator {
        rule: "plus",
        precedence: 10,
        fixity: Fixity::InfixLeft,
    };
    assert!(
        grammar
            .bind_pratt(
                "expr",
                "atom",
                &[repeated, repeated],
                TrailingTrivia::Consume
            )
            .is_err()
    );
    assert_eq!(grammar.source_pratt(target).unwrap().operators.len(), 6);
    let mut stack = vybe_parser::compile(r#"atom = { PUSH("a") } expr = { atom }"#).unwrap();
    assert!(
        stack
            .bind_pratt("expr", "atom", &[], TrailingTrivia::Preserve)
            .is_err()
    );
    let mut nullable = vybe_parser::compile(r#"atom = { "a"? } expr = { atom }"#).unwrap();
    assert!(
        nullable
            .bind_pratt("expr", "atom", &[], TrailingTrivia::Preserve)
            .is_err()
    );
}

#[test]
fn lua_expression_island_preserves_real_grammar_recognition() {
    use vybe_parser::islands::TrailingTrivia;
    let source = include_str!("../../../languages/lua/src/grammar.pest");
    let ordinary = vybe_parser::compile(source).unwrap();
    let mut bound = vybe_parser::compile(source).unwrap();
    let mut operators = vec![Operator {
        rule: "unop",
        precedence: 25,
        fixity: Fixity::Prefix,
    }];
    for (rule, precedence, fixity) in [
        ("pow_op", 30, Fixity::InfixRight),
        ("mul_op", 20, Fixity::InfixLeft),
        ("additive_op", 19, Fixity::InfixLeft),
        ("CONCAT", 18, Fixity::InfixRight),
        ("shift_op", 17, Fixity::InfixLeft),
        ("AMP", 16, Fixity::InfixLeft),
        ("TILDE", 15, Fixity::InfixLeft),
        ("PIPE", 14, Fixity::InfixLeft),
        ("compare_op", 13, Fixity::InfixLeft),
        ("KW_AND", 12, Fixity::InfixLeft),
        ("KW_OR", 11, Fixity::InfixLeft),
    ] {
        operators.push(Operator {
            rule,
            precedence,
            fixity,
        });
    }
    bound
        .bind_pratt("expr", "postfix", &operators, TrailingTrivia::Consume)
        .unwrap();
    for source in [
        "return 1 + 2 * 3",
        "local a = -2^2; return a",
        "return 2^-3^4",
        "return f(1+2, (4-5)/6).x[2]",
        "return not a and b or c",
        "return 'a' .. 'b' .. 'c'",
        "return 1 << 2 & 3 ~ 4 | 5",
        "return a <= b and c ~= d",
        "return 1 -- comment\n + 2",
        "return [==[hello]==] .. [[world]]",
        "return 1 +",
        "return f(1+)",
        "return not",
        "return (1+2",
        "return 1..2",
        "return 1 // 2",
    ] {
        let a = ordinary
            .recognize("chunk", source)
            .map(|r| r.consumed)
            .map_err(|e| (e.kind, e.offset));
        let b = bound
            .recognize("chunk", source)
            .map(|r| r.consumed)
            .map_err(|e| (e.kind, e.offset));
        assert_eq!(a, b, "{source}");
    }
}

#[test]
fn recognition_matches_ordinary_peg_for_bounded_expression_corpus() {
    fn captures(tree: vybe_parser::ParseTree<'_, '_>) -> Vec<(vybe_parser::CaptureRule, Span)> {
        let mut pending: Vec<_> = tree.pairs().collect();
        let mut out = Vec::new();
        while let Some(pair) = pending.pop() {
            out.push((pair.as_rule(), pair.as_span()));
            pending.extend(pair.into_inner());
        }
        out
    }
    let ordinary = grammar(false);
    let bound = grammar(true);
    let alphabet = ["1", "+", "-", "*", "^", "!", "(", ")", " "];
    let mut inputs = vec![String::new()];
    let mut frontier = inputs.clone();
    for _ in 0..4 {
        frontier = frontier
            .iter()
            .flat_map(|prefix| {
                alphabet
                    .iter()
                    .map(move |suffix| format!("{prefix}{suffix}"))
            })
            .collect();
        inputs.extend(frontier.iter().cloned());
    }
    for source in inputs {
        for entry in ["expr", "program"] {
            let a = ordinary
                .recognize(entry, &source)
                .map(|r| r.consumed)
                .map_err(|e| (e.kind, e.offset));
            let b = bound
                .recognize(entry, &source)
                .map(|r| r.consumed)
                .map_err(|e| (e.kind, e.offset));
            assert_eq!(a, b, "{entry}: {source:?}");
        }
        // Capture parsing preserves the original grammar and its wrapper nodes.
        let a = ordinary.parse_source("program", &source).map(captures);
        let b = bound.parse_source("program", &source).map(captures);
        assert_eq!(a, b, "captures: {source:?}");
    }
}
