use grameo::{
    MatchOptions, Span,
    arena::{Arena, Id},
    builder::Builder,
    engine::build_program,
    islands::Operator,
    pratt::Fixity,
};

const GRAMMAR: &str = include_str!("islands.grammar");

#[test]
fn unchanged_builder_cursor_is_reused_across_peg_backtracking() {
    #[derive(Default)]
    struct Inactive { checkpoints: std::cell::Cell<usize>, rollbacks: usize }
    impl Builder for Inactive {
        fn observes_rule(&self, _: usize) -> bool { false }
        fn checkpoint(&self) -> usize {
            self.checkpoints.set(self.checkpoints.get() + 1);
            0
        }
        fn rollback(&mut self, _: usize) { self.rollbacks += 1; }
        fn begin(&mut self, _: usize, _: usize) {}
        fn finish(&mut self, _: usize, _: Span, _: &str) {}
    }
    let grammar = grameo::compile(r#"program = { SOI ~ (("a" ~ "b") | "a")* ~ EOI }"#).unwrap();
    let source = "a".repeat(128);
    let mut builder = Inactive::default();
    let tree = grameo::engine::build_capture_program(
        &grammar, "program", &source, MatchOptions::default(), &mut builder,
    ).unwrap();
    assert_eq!(tree.consumed(), source.len());
    assert_eq!(builder.checkpoints.get(), 2); // Outer failure guard and initial cursor.
    assert_eq!(builder.rollbacks, 0); // Matching backtracks; semantic state never changes.
}

#[test]
fn input_selection_is_queried_only_for_configured_islands() {
    struct Selection { target: usize, calls: std::cell::Cell<usize> }
    impl Builder for Selection {
        fn observes_rule(&self, _: usize) -> bool { false }
        fn supports_pratt_at(&self, rule: usize, _: usize, _: &str) -> bool {
            assert_eq!(rule, self.target, "ordinary rules need no input-selection callback");
            self.calls.set(self.calls.get() + 1);
            false // Exercise capture fallback without constructing semantic values.
        }
        fn checkpoint(&self) -> usize { 0 }
        fn rollback(&mut self, _: usize) {}
        fn begin(&mut self, _: usize, _: usize) {}
        fn finish(&mut self, _: usize, _: Span, _: &str) {}
    }
    let mut grammar = grameo::compile(GRAMMAR).unwrap();
    grammar.bind_pratt("expr", "atom", &[Operator {
        rule: "plus", precedence: 10, fixity: Fixity::InfixLeft,
    }], grameo::islands::TrailingTrivia::Preserve).unwrap();
    let mut builder = Selection {
        target: grammar.rule_id("expr").unwrap(), calls: std::cell::Cell::new(0),
    };
    grameo::engine::build_capture_program(
        &grammar, "program", "1+2", MatchOptions::default(), &mut builder,
    ).unwrap();
    assert_eq!(builder.calls.get(), 1);
    builder.calls.set(0);
    build_program(&grammar, "program", "1+2", MatchOptions::default(), &mut builder).unwrap();
    assert_eq!(builder.calls.get(), 1);
}

#[derive(Default)]
struct Boundaries {
    ends: Vec<(usize, usize)>,
    reject: bool,
}
impl Builder for Boundaries {
    fn supports_pratt(&self, _: usize) -> bool {
        true
    }
    fn checkpoint(&self) -> usize {
        self.ends.len()
    }
    fn rollback(&mut self, mark: usize) {
        self.ends.truncate(mark);
    }
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, _: usize, _: Span, _: &str) {}
    fn try_pratt_boundary(
        &mut self,
        rule: usize,
        offset: usize,
    ) -> Result<(), grameo::builder::BuildError> {
        self.ends.push((rule, offset));
        if self.reject {
            Err(grameo::builder::BuildError {
                span: Span {
                    start: offset,
                    end: offset,
                },
                message: "boundary failure".into(),
            })
        } else {
            Ok(())
        }
    }
}

#[test]
fn pratt_operand_boundaries_use_matched_trivia_and_rollback_policy() {
    use grameo::islands::TrailingTrivia;
    for trivia in [TrailingTrivia::Consume, TrailingTrivia::Preserve] {
        let mut grammar = grameo::compile(GRAMMAR).unwrap();
        grammar
            .bind_pratt(
                "expr",
                "atom",
                &[Operator {
                    rule: "plus",
                    precedence: 10,
                    fixity: Fixity::InfixLeft,
                }],
                trivia,
            )
            .unwrap();
        let target = grammar.rule_id("expr").unwrap();
        let mut builder = Boundaries::default();
        let parsed = build_program(
            &grammar,
            "expr",
            "1 \n",
            MatchOptions::default(),
            &mut builder,
        )
        .unwrap();
        if trivia == TrailingTrivia::Consume {
            assert_eq!(parsed.consumed, 3);
            assert_eq!(builder.ends, [(target, 3)]);
        } else {
            assert_eq!(parsed.consumed, 1);
            assert!(builder.ends.is_empty());
        }
        let mut builder = Boundaries::default();
        build_program(
            &grammar,
            "expr",
            "1 + 2 \n",
            MatchOptions::default(),
            &mut builder,
        )
        .unwrap();
        let expected = if trivia == TrailingTrivia::Consume {
            vec![(target, 2), (target, 7)]
        } else {
            vec![(target, 2)]
        };
        assert_eq!(builder.ends, expected);
        let mut builder = Boundaries::default();
        build_program(
            &grammar,
            "expr",
            "1 + ",
            MatchOptions::default(),
            &mut builder,
        )
        .unwrap();
        assert_eq!(
            builder.ends,
            if trivia == TrailingTrivia::Consume {
                vec![(target, 2)]
            } else {
                Vec::new()
            }
        );
    }
}

#[test]
fn pratt_boundary_failure_is_located_fatal_and_restores_reference_adapter() {
    let grammar = grammar(true);
    let mut builder = Boundaries {
        reject: true,
        ..Default::default()
    };
    // Exercise the forwarding implementation for &mut B too.
    let error = build_program(
        &grammar,
        "fallback",
        "1 ",
        MatchOptions::default(),
        &mut &mut builder,
    )
    .unwrap_err();
    assert_eq!(error.kind, grameo::ParseErrorKind::BuildFailure);
    assert_eq!(error.offset, 2);
    assert!(builder.ends.is_empty());
    let mut builder = Boundaries::default();
    build_program(
        &grammar,
        "probe",
        "1 ",
        MatchOptions::default(),
        &mut builder,
    )
    .unwrap();
    assert_eq!(builder.ends, [(grammar.rule_id("expr").unwrap(), 2)]);
}

fn grammar(bound: bool) -> grameo::CompiledGrammar {
    let mut grammar = grameo::compile(GRAMMAR).unwrap();
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
                grameo::islands::TrailingTrivia::Consume,
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
    fn new(grammar: &grameo::CompiledGrammar) -> Self {
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
    use grameo::builder::BuildError;
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
        assert_eq!(error.kind, grameo::ParseErrorKind::BuildFailure);
        assert_eq!(error.offset, offset);
        assert_eq!(builder.checkpoint(), mark);
        assert_eq!(builder.ast.arena.node_count(), 1);
        assert_eq!(builder.ast.arena.values(), &[sentinel]);
    }
}

#[test]
fn invalid_bindings_are_rejected_without_replacing_valid_profile() {
    use grameo::{islands::TrailingTrivia, program::Program};
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
    let mut stack = grameo::compile(r#"atom = { "a" } op = { PUSH("+") } expr = { atom ~ (op ~ atom)* }"#).unwrap();
    assert!(
        stack
            .bind_pratt("expr", "atom", &[Operator { rule: "op", precedence: 10, fixity: Fixity::InfixLeft }], TrailingTrivia::Preserve)
            .is_err()
    );
    let mut nullable = grameo::compile(r#"atom = { "a"? } expr = { atom }"#).unwrap();
    assert!(
        nullable
            .bind_pratt("expr", "atom", &[], TrailingTrivia::Preserve)
            .is_err()
    );
}

#[test]
fn pratt_atoms_preserve_grammar_stack_and_restore_failed_operands() {
    use grameo::islands::TrailingTrivia;
    // Atoms pop an existing value in a failed choice, push their own value,
    // and leave it for the caller. A failed RHS must undo both pop and push.
    let source = r#"
root = { PUSH("z") ~ expr ~ "+"? ~ POP ~ POP ~ EOI }
expr = { atom ~ (plus ~ atom)* }
atom = { (POP ~ "!") | (PUSH("a" | "b") ~ ";") }
plus = { "+" }
probe = { PUSH("z") ~ &expr ~ expr ~ POP ~ POP ~ EOI }
fallback = { PUSH("z") ~ ((expr ~ "!") | (expr ~ POP ~ POP)) ~ EOI }
"#;
    let ordinary = grameo::compile(source).unwrap();
    let mut bound = grameo::compile(source).unwrap();
    bound.bind_pratt("expr", "atom", &[Operator {
        rule: "plus", precedence: 10, fixity: Fixity::InfixLeft,
    }], TrailingTrivia::Preserve).unwrap();
    for (entry, input) in [
        ("root", "za;az"), ("root", "za;+az"),
        ("root", "za;+b;ba"), ("root", "za;+zz"),
        ("root", "za;+z!az"), ("probe", "za;az"),
        ("fallback", "za;az"), ("fallback", "za;+b;ba"),
    ] {
        let expected = ordinary.recognize(entry, input);
        let actual = bound.recognize(entry, input);
        assert_eq!(actual.as_ref().ok().map(|r| r.consumed),
                   expected.as_ref().ok().map(|r| r.consumed), "{entry}: {input}");
        assert_eq!(actual.is_ok(), expected.is_ok(), "{entry}: {input}");
    }
    assert!(bound.recognize("root", "za;az").is_ok());
    assert!(bound.recognize("root", "za;+az").is_ok());
    assert!(bound.recognize("fallback", "za;az").is_ok());
}

#[test]
fn pratt_atoms_can_read_caller_stack_with_transactional_builder_hooks() {
    use grameo::islands::TrailingTrivia;
    let mut grammar = grameo::compile(r#"
root = { PUSH("END") ~ expr ~ POP ~ EOI }
expr = { atom ~ (plus ~ atom)* }
atom = { PEEK ~ "!" }
plus = { "+" }
"#).unwrap();
    grammar.bind_pratt("expr", "atom", &[Operator {
        rule: "plus", precedence: 10, fixity: Fixity::InfixLeft,
    }], TrailingTrivia::Preserve).unwrap();
    let input = "ENDEND!+END!END";
    let mut builder = Boundaries::default();
    assert_eq!(grammar.build_ast("root", input, &mut builder).unwrap(), input.len());
    // Semantic failure inside a stack-reading atom's island restores the
    // caller's pre-existing builder values, even after boundary mutations.
    let initial = vec![(usize::MAX, 123)];
    let mut builder = Boundaries { ends: initial.clone(), reject: true };
    assert_eq!(grammar.build_ast("root", input, &mut builder).unwrap_err().kind,
               grameo::ParseErrorKind::BuildFailure);
    assert_eq!(builder.ends, initial);
}

#[test]
fn lua_expression_island_preserves_real_grammar_recognition() {
    use grameo::islands::TrailingTrivia;
    let source = include_str!("fixtures/lua.pest");
    let ordinary = grameo::compile(source).unwrap();
    let mut bound = grameo::compile(source).unwrap();
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
    fn captures(tree: grameo::ParseTree<'_, '_>) -> Vec<(grameo::CaptureRule, Span)> {
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

#[test]
fn hybrid_capture_builds_pratt_ast_and_restores_both_arenas() {
    let grammar = grammar(true);
    for entry in ["program", "probe", "fallback"] {
        let mut ast = Ast::new(&grammar);
        let tree = grameo::engine::build_capture_program(
            &grammar, entry, " (1 + 2) * 3! ", MatchOptions::default(), &mut ast,
        ).unwrap();
        assert_eq!(ast.arena.values().len(), 1);
        assert_eq!(ast.shape(ast.arena.values()[0]), "(* (+ 1 2) (! 3))");
        fn leaves(pair: grameo::Pair<'_, '_, '_>, found: &mut Vec<String>) {
            if pair.rule_name() == "number" { found.push(pair.as_str().into()); }
            for child in pair.into_inner() { leaves(child, found); }
        }
        let mut found = Vec::new();
        for root in tree.pairs() { leaves(root, &mut found); }
        // Lookahead and the rejected fallback branch cannot duplicate atoms.
        assert_eq!(found, ["1", "2", "3"]);

    }
    let mut ast = Ast::new(&grammar);
    let before = ast.arena.checkpoint();
    assert!(grameo::engine::build_capture_program(
        &grammar, "program", "1+2*", MatchOptions::default(), &mut ast,
    ).is_err());
    assert_eq!(ast.arena.checkpoint(), before);
}

#[test]
fn hybrid_without_islands_preserves_captures_and_atomic_semantic_hooks() {
    #[derive(Default)]
    struct Hooks(Vec<(usize, Span)>);
    impl Builder for Hooks {
        fn checkpoint(&self) -> usize { self.0.len() }
        fn rollback(&mut self, mark: usize) { self.0.truncate(mark); }
        fn begin(&mut self, _: usize, _: usize) {}
        fn finish(&mut self, rule: usize, span: Span, _: &str) { self.0.push((rule, span)); }
    }
    fn shape(tree: &grameo::ParseTree<'_, '_>) -> Vec<(String, Span)> {
        fn visit(pair: grameo::Pair<'_, '_, '_>, output: &mut Vec<(String, Span)>) {
            output.push((pair.rule_name().into(), pair.as_span()));
            for child in pair.into_inner() { visit(child, output); }
        }
        let mut output = Vec::new();
        for pair in tree.pairs() { visit(pair, &mut output); }
        output
    }
    let grammar = grameo::compile(r#"
        inner = { "a" }
        atomic = @{ inner }
        silent = _{ atomic }
        program = { SOI ~ ((silent ~ "x") | silent) ~ EOI }
    "#).unwrap();
    let mut hooks = Hooks::default();
    let tree = grameo::engine::build_capture_program(
        &grammar, "program", "a", MatchOptions::default(), &mut hooks,
    ).unwrap();
    assert_eq!(shape(&tree), shape(&grammar.parse_source("program", "a").unwrap()));
    let names: Vec<_> = hooks.0.iter().map(|(rule, _)| grammar.syntax().rules[*rule].name.as_str()).collect();
    assert_eq!(names, ["inner", "atomic", "silent", "program"]);
    for options in [MatchOptions { max_steps: 1, ..MatchOptions::default() }, MatchOptions { max_rule_depth: 1, ..MatchOptions::default() }] {
        let mark = hooks.checkpoint();
        assert!(grameo::engine::build_capture_program(&grammar, "program", "a", options, &mut hooks).is_err());
        assert_eq!(hooks.checkpoint(), mark);
    }
}

#[test]
fn semantic_capture_pruning_preserves_pratt_values_and_speculative_rollback() {
    struct Pruning(Ast);
    impl Builder for Pruning {
        fn supports_pratt(&self, rule: usize) -> bool { self.0.supports_pratt(rule) }
        fn checkpoint(&self) -> usize { self.0.checkpoint() }
        fn rollback(&mut self, mark: usize) { self.0.rollback(mark); }
        fn begin(&mut self, rule: usize, offset: usize) { self.0.begin(rule, offset); }
        fn finish(&mut self, rule: usize, span: Span, source: &str) { self.0.finish(rule, span, source); }
        fn reduce(&mut self, rule: usize, fixity: Fixity, span: Span) { self.0.reduce(rule, fixity, span); }
        fn prune_capture_children(&self, rule: usize, _: Span) -> bool { rule == self.0.target }
    }
    let grammar = grammar(true);
    for entry in ["program", "probe", "fallback"] {
        let mut ast = Pruning(Ast::new(&grammar));
        let tree = grammar.build_captures(entry, "(1+2)*3", &mut ast).unwrap();
        assert_eq!(ast.0.shape(ast.0.arena.values()[0]), "(* (+ 1 2) 3)");
        let expression = tree.pairs().next().unwrap().into_inner().next().unwrap();
        assert_eq!(expression.rule_name(), "expr");
        assert_eq!(expression.as_str(), "(1+2)*3");
        assert_eq!(expression.into_inner().count(), 0);
    }
    let mut ast = Pruning(Ast::new(&grammar));
    let mark = ast.checkpoint();
    assert!(grammar.build_captures("program", "(1+2)*3+", &mut ast).is_err());
    assert_eq!(ast.checkpoint(), mark);
}

#[test]
fn suppressed_pratt_children_preserve_owned_values_anchors_and_rollback() {
    struct Owned(Ast, bool);
    impl Builder for Owned {
        fn supports_pratt(&self, rule: usize) -> bool { self.0.supports_pratt(rule) }
        fn suppress_pratt_captures(&self, rule: usize) -> bool { self.1 && rule == self.0.target }
        fn checkpoint(&self) -> usize { self.0.checkpoint() }
        fn rollback(&mut self, mark: usize) { self.0.rollback(mark); }
        fn begin(&mut self, rule: usize, offset: usize) { self.0.begin(rule, offset); }
        fn finish(&mut self, rule: usize, span: Span, source: &str) { self.0.finish(rule, span, source); }
        fn reduce(&mut self, rule: usize, fixity: Fixity, span: Span) { self.0.reduce(rule, fixity, span); }
    }
    let grammar = grammar(true);
    for suppress in [false, true] {
        for entry in ["program", "probe", "fallback"] {
            let mut ast = Owned(Ast::new(&grammar), suppress);
            let tree = grammar.build_captures(entry, "(1+2)*3", &mut &mut ast).unwrap();
            assert_eq!(ast.0.arena.values().len(), 1);
            assert_eq!(ast.0.shape(ast.0.arena.values()[0]), "(* (+ 1 2) 3)");
            let expression = tree.pairs().next().unwrap().into_inner().next().unwrap();
            assert_eq!(expression.rule_name(), "expr");
            assert_eq!(expression.as_str(), "(1+2)*3");
            assert_eq!(expression.into_inner().count() == 0, suppress);
        }
        for options in [
            MatchOptions::default(),
            MatchOptions { max_steps: 12, ..MatchOptions::default() },
            MatchOptions { max_rule_depth: 2, ..MatchOptions::default() },
        ] {
            let mut ast = Owned(Ast::new(&grammar), suppress);
            let mark = ast.checkpoint();
            assert!(grameo::engine::build_capture_program(
                &grammar, "program", "(1+2)*3+", options, &mut ast,
            ).is_err());
            assert_eq!(ast.checkpoint(), mark);
        }
    }
}

#[test]
fn unobserved_wrappers_retain_child_hooks_reductions_and_rollback() {
    struct Selective(Ast);
    impl Builder for Selective {
        fn observes_rule(&self, rule: usize) -> bool { rule == self.0.number }
        fn supports_pratt(&self, rule: usize) -> bool { self.0.supports_pratt(rule) }
        fn checkpoint(&self) -> usize { self.0.checkpoint() }
        fn rollback(&mut self, mark: usize) { self.0.rollback(mark); }
        fn begin(&mut self, rule: usize, _: usize) { assert_eq!(rule, self.0.number); }
        fn finish(&mut self, rule: usize, span: Span, source: &str) {
            assert_eq!(rule, self.0.number);
            self.0.finish(rule, span, source);
        }
        fn reduce(&mut self, rule: usize, fixity: Fixity, span: Span) { self.0.reduce(rule, fixity, span); }
    }
    let grammar = grammar(true);
    for entry in ["program", "probe", "fallback"] {
        let mut ast = Selective(Ast::new(&grammar));
        let tree = grammar.build_captures(entry, "1+2*3", &mut &mut ast).unwrap();
        assert_eq!(tree.consumed(), 5);
        assert_eq!(ast.0.shape(ast.0.arena.values()[0]), "(+ 1 (* 2 3))");
    }
    let mut ast = Selective(Ast::new(&grammar));
    assert!(grammar.build_captures("program", "1+2*", &mut ast).is_err());
    assert_eq!(ast.0.arena.node_count(), 0);
}

#[test]
fn hybrid_rejected_semantic_island_restores_original_peg_captures() {
    #[derive(Default)]
    struct Contextual {
        events: Vec<(usize, bool)>,
    }
    impl Builder for Contextual {
        fn supports_pratt(&self, _: usize) -> bool { true }
        fn accepts_pratt(&self, _: usize, _: Span) -> bool { false }
        fn suppress_pratt_captures(&self, _: usize) -> bool { true }
        fn checkpoint(&self) -> usize { self.events.len() }
        fn rollback(&mut self, mark: usize) { self.events.truncate(mark); }
        fn begin(&mut self, rule: usize, _: usize) { self.events.push((rule, false)); }
        fn finish(&mut self, rule: usize, _: Span, _: &str) { self.events.push((rule, true)); }
        fn reduce(&mut self, rule: usize, _: Fixity, _: Span) { self.events.push((rule, true)); }
    }
    fn captures(tree: &grameo::ParseTree<'_, '_>) -> Vec<(String, Span)> {
        tree.pairs().flat_map(|pair| {
            fn visit(pair: grameo::Pair<'_, '_, '_>, out: &mut Vec<(String, Span)>) {
                out.push((pair.rule_name().to_owned(), pair.as_span()));
                for child in pair.into_inner() { visit(child, out); }
            }
            let mut out = Vec::new();
            visit(pair, &mut out);
            out
        }).collect()
    }
    for mode in ["", "@", "$", "!"] {
        let source = format!(r#"
            WHITESPACE = _{{ " " }}
            number = @{{ ASCII_DIGIT+ }}
            atom = {{ number | "(" ~ expr ~ ")" }}
            mul = {{ "*" }}
            plus = {{ "+" }}
            product = {{ atom ~ (mul ~ atom)* }}
            expr = {mode}{{ product ~ (plus ~ product)* }}
            program = {{ SOI ~ ((&expr ~ expr ~ "x") | expr) ~ EOI }}
        "#);
        let mut grammar = grameo::compile(&source).unwrap();
        grammar.bind_pratt("expr", "atom", &[
            Operator { rule: "mul", precedence: 20, fixity: Fixity::InfixLeft },
            Operator { rule: "plus", precedence: 10, fixity: Fixity::InfixLeft },
        ], grameo::islands::TrailingTrivia::Consume).unwrap();
        for input in ["1+2*3", "(1+2)*3", "1 + 2 * 3 ", "1+2x"] {
            if matches!(mode, "@" | "$") && input.contains(' ') { continue; }
            let mut contextual = Contextual::default();
            let actual = grameo::engine::build_capture_program(
                &grammar, "program", input, MatchOptions::default(), &mut contextual,
            ).unwrap();
            let expected = grammar.parse_source("program", input).unwrap();
            assert_eq!(captures(&actual), captures(&expected), "mode={mode} input={input}");
            // All speculative child construction is rolled back. The PEG
            // fallback emits captures, leaving only the surrounding hook pair.
            let program = grammar.rule_id("program").unwrap();
            assert_eq!(contextual.events, [(program, false), (program, true)]);
        }
        let mut contextual = Contextual::default();
        let mark = contextual.checkpoint();
        assert!(grameo::engine::build_capture_program(
            &grammar, "program", "1+2*", MatchOptions::default(), &mut contextual,
        ).is_err());
        assert_eq!(contextual.checkpoint(), mark);
    }
}

#[test]
fn rejected_pratt_island_restores_grammar_stack_before_peg_fallback() {
    struct Contextual;
    impl Builder for Contextual {
        fn supports_pratt(&self, _: usize) -> bool { true }
        fn accepts_pratt(&self, _: usize, _: Span) -> bool { false }
        fn suppress_pratt_captures(&self, _: usize) -> bool { true }
        fn checkpoint(&self) -> usize { 0 }
        fn rollback(&mut self, _: usize) {}
        fn begin(&mut self, _: usize, _: usize) {}
        fn finish(&mut self, _: usize, _: Span, _: &str) {}
    }
    let mut grammar = grameo::compile(r#"
        atom = { PUSH(ASCII_DIGIT+) }
        plus = { "+" }
        expr = { atom ~ (plus ~ atom)* }
        program = { expr ~ ":" ~ POP ~ POP ~ PEEK_ALL ~ "z" ~ EOI }
    "#).unwrap();
    grammar.bind_pratt("expr", "atom", &[
        Operator { rule: "plus", precedence: 10, fixity: Fixity::InfixLeft },
    ], grameo::islands::TrailingTrivia::Consume).unwrap();
    let expected = grammar.parse_source("program", "1+2:21z").unwrap();
    let actual = grameo::engine::build_capture_program(
        &grammar, "program", "1+2:21z", MatchOptions::default(), &mut Contextual,
    ).unwrap();
    let spans = |tree: &grameo::ParseTree<'_, '_>| {
        fn visit(pair: grameo::Pair<'_, '_, '_>, output: &mut Vec<(String, Span)>) {
            output.push((pair.rule_name().to_owned(), pair.as_span()));
            for child in pair.into_inner() { visit(child, output); }
        }
        let mut output = Vec::new();
        for pair in tree.pairs() { visit(pair, &mut output); }
        output
    };
    assert_eq!(spans(&actual), spans(&expected));
}


#[test]
fn rejected_pratt_observations_remain_owned_by_enclosing_transactions() {
    struct Observations { atom: usize, receipts: Vec<Span>, replayed: usize }
    impl Builder for Observations {
        fn observes_rule(&self, rule: usize) -> bool { rule == self.atom }
        fn supports_pratt(&self, _: usize) -> bool { true }
        fn accepts_pratt(&self, _: usize, _: Span) -> bool { false }
        fn suppress_pratt_captures(&self, _: usize) -> bool { true }
        fn checkpoint(&self) -> usize { self.receipts.len() }
        fn rollback(&mut self, mark: usize) { self.receipts.truncate(mark); }
        fn rollback_pratt(&mut self, mark: usize, _: usize, span: Span) {
            let retained = self.receipts[mark..].iter().copied()
                .filter(|receipt| receipt.start >= span.start && receipt.end <= span.end)
                .collect::<Vec<_>>();
            self.rollback(mark);
            self.receipts.extend(retained);
            self.replayed += 1; // Test-only instrumentation, not semantic state.
        }
        fn begin(&mut self, _: usize, _: usize) {}
        fn finish(&mut self, rule: usize, span: Span, _: &str) {
            if rule == self.atom { self.receipts.push(span); }
        }
    }
    let mut grammar = grameo::compile(r#"
        atom = { PUSH(ASCII_DIGIT+) }
        plus = { "+" }
        expr = { atom ~ (plus ~ atom)* }
        program = { SOI ~ ((&expr ~ expr ~ "x") | expr) ~ ":" ~ POP ~ POP ~ PEEK_ALL ~ "z" ~ EOI }
    "#).unwrap();
    grammar.bind_pratt("expr", "atom", &[Operator {
        rule: "plus", precedence: 10, fixity: Fixity::InfixLeft,
    }], grameo::islands::TrailingTrivia::Consume).unwrap();
    let atom = grammar.rule_id("atom").unwrap();
    let mut owner = Observations { atom, receipts: vec![], replayed: 0 };
    let source = "1+2:21z";
    let expected = grammar.parse_source("program", source).unwrap();
    let actual = grameo::engine::build_capture_program(
        &grammar, "program", source, MatchOptions::default(), &mut &mut owner,
    ).unwrap();
    fn captures(tree: &grameo::ParseTree<'_, '_>) -> Vec<(String, Span)> {
        fn visit(pair: grameo::Pair<'_, '_, '_>, out: &mut Vec<(String, Span)>) {
            out.push((pair.rule_name().to_owned(), pair.as_span()));
            for child in pair.into_inner() { visit(child, out); }
        }
        let mut out=vec![];
        for pair in tree.pairs() { visit(pair, &mut out); }
        out
    }
    assert_eq!(captures(&actual), captures(&expected));
    assert_eq!(owner.receipts, [Span { start: 0, end: 1 }, Span { start: 2, end: 3 }]);
    assert_eq!(owner.replayed, 2, "failed outer choice must discard its observations; predicate must publish none");
    // Preserve a caller-owned prefix while deleting all receipts on a later
    // replay/outer suffix failure. The builder can be reused after the error.
    for bad in ["1+2:12z", "1+2:21", "1+2*:21z"] {
        let prefix = Span { start: 90, end: 91 };
        owner.receipts = vec![prefix];
        assert!(grameo::engine::build_capture_program(
            &grammar, "program", bad, MatchOptions::default(), &mut owner,
        ).is_err());
        assert_eq!(owner.receipts, [prefix], "{bad}");
    }
    // Finite resource failures after preservation must also restore the root
    // guard. No budget may leak receipt entries from the abandoned island.
    for max_steps in 0..160 {
        let mut owner = Observations { atom, receipts: vec![], replayed: 0 };
        let result = grameo::engine::build_capture_program(&grammar, "program", source,
            MatchOptions { max_steps, max_rule_depth: 4096 }, &mut owner);
        if result.is_err() { assert!(owner.receipts.is_empty(), "max_steps={max_steps}"); }
        else { assert_eq!(owner.receipts.len(), 2); }
    }
}
