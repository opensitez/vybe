use grameo::{
    MatchOptions, ParseErrorKind, Span,
    builder::{BuildError, Builder},
    streaming::{Boundary, CompletedUnits, Sink, StreamingBuilder},
};
use std::{cell::RefCell, rc::Rc};

#[derive(Clone, Debug, PartialEq)]
enum Expr {
    Number(i64),
    Add(Box<Expr>, Box<Expr>),
}
#[derive(Clone, Debug, PartialEq)]
struct Unit {
    expression: Expr,
    span: Span,
    scope: &'static str,
    directives: Vec<&'static str>,
}
#[derive(Clone, Debug, Default, PartialEq)]
struct State {
    operands: Vec<Expr>,
    units: Vec<Unit>,
    events: Vec<(usize, Span)>,
}
struct Producer {
    state: State,
    undo: Vec<State>,
    number: usize,
    expression: usize,
    statement: usize,
    plus: usize,
    pratt: bool,
    commits: usize,
    retain_receipt: bool,
}
impl Producer {
    fn new(g: &grameo::CompiledGrammar, pratt: bool) -> Self {
        Self {
            state: State::default(),
            undo: vec![],
            number: g.rule_id("number").unwrap(),
            expression: g.rule_id("expression").unwrap(),
            statement: g.rule_id("statement").unwrap(),
            plus: g.rule_id("plus").unwrap(),
            pratt,
            commits: 0,
            retain_receipt: false,
        }
    }
    fn save(&mut self) {
        self.undo.push(self.state.clone());
    }
}
impl Builder for Producer {
    fn observes_rule(&self, r: usize) -> bool {
        r == self.number || r == self.expression || r == self.statement
    }
    fn supports_pratt(&self, r: usize) -> bool {
        self.pratt && r == self.expression
    }
    fn checkpoint(&self) -> usize {
        self.undo.len()
    }
    fn rollback(&mut self, n: usize) {
        while self.undo.len() > n {
            self.state = self.undo.pop().unwrap();
        }
    }
    fn rollback_pratt(&mut self, n: usize, _: usize, span: Span) {
        self.rollback(n);
        if self.retain_receipt {
            self.save();
            self.state.events.push((usize::MAX, span));
        }
    }
    fn begin(&mut self, r: usize, p: usize) {
        self.save();
        self.state.events.push((r, Span { start: p, end: p }));
    }
    fn finish(&mut self, r: usize, s: Span, source: &str) {
        self.save();
        self.state.events.push((r, s));
        if r == self.number {
            self.state
                .operands
                .push(Expr::Number(source[s.start..s.end].parse().unwrap()));
        }
        if r == self.expression && !self.pratt {
            while self.state.operands.len() > 1 {
                let right = self.state.operands.pop().unwrap();
                let left = self.state.operands.pop().unwrap();
                self.state
                    .operands
                    .push(Expr::Add(Box::new(left), Box::new(right)));
            }
        }
        if r == self.statement {
            let expression = self.state.operands.pop().unwrap();
            self.state.units.push(Unit {
                expression,
                span: s,
                scope: "caller lexical namespace",
                directives: vec!["slots=descriptor", "class=canonical", "strict=true"],
            });
        }
    }
    fn reduce(&mut self, r: usize, _: grameo::pratt::Fixity, _: Span) {
        assert_eq!(r, self.plus);
        self.save();
        let right = self.state.operands.pop().unwrap();
        let left = self.state.operands.pop().unwrap();
        self.state
            .operands
            .push(Expr::Add(Box::new(left), Box::new(right)));
    }
    fn committed(&mut self, _: &str, _: usize) {
        self.commits += 1;
    }
}
impl CompletedUnits for Producer {
    type Unit = Unit;
    fn is_emission_boundary(&self, r: usize) -> bool {
        r == self.statement
    }
    fn completed_unit(&self, b: Boundary) -> Result<Option<&Unit>, BuildError> {
        Ok(self.state.units.last().filter(|u| u.span == b.span))
    }
}
#[derive(Clone, Debug, PartialEq)]
enum Op {
    Const(i64),
    Add,
    Print,
}
#[derive(Clone, Debug, Default, PartialEq)]
struct Code {
    ops: Vec<Op>,
    metadata: Vec<(&'static str, Vec<&'static str>, Span)>,
    finished: bool,
}
struct VmSink {
    code: Code,
    attempts: Rc<RefCell<Vec<Boundary>>>,
    reject: bool,
    finish_error: bool,
}
impl Default for VmSink {
    fn default() -> Self {
        Self {
            code: Code::default(),
            attempts: Rc::default(),
            reject: false,
            finish_error: false,
        }
    }
}
impl VmSink {
    fn expression(&mut self, e: &Expr) {
        match e {
            Expr::Number(n) => self.code.ops.push(Op::Const(*n)),
            Expr::Add(a, b) => {
                self.expression(a);
                self.expression(b);
                self.code.ops.push(Op::Add)
            }
        }
    }
    fn run(&self) -> Vec<i64> {
        assert!(self.code.finished);
        let mut stack = vec![];
        let mut output = vec![];
        for op in &self.code.ops {
            match *op {
                Op::Const(n) => stack.push(n),
                Op::Add => {
                    let b = stack.pop().unwrap();
                    let a = stack.pop().unwrap();
                    stack.push(a + b)
                }
                Op::Print => output.push(stack.pop().unwrap()),
            }
        }
        output
    }
}
impl Sink<Unit> for VmSink {
    type Checkpoint = Code;
    fn checkpoint(&self) -> Code {
        self.code.clone()
    }
    fn rollback(&mut self, c: Code) {
        self.code = c;
    }
    fn emit(&mut self, u: &Unit, b: Boundary, _: &str) -> Result<(), BuildError> {
        self.attempts.borrow_mut().push(b);
        self.expression(&u.expression);
        self.code.ops.push(Op::Print);
        self.code
            .metadata
            .push((u.scope, u.directives.clone(), u.span));
        if self.reject {
            return Err(BuildError {
                span: b.span,
                message: "sink rejected semantic unit".into(),
            });
        }
        Ok(())
    }
    fn finish(&mut self, _: &str, consumed: usize) -> Result<(), BuildError> {
        self.code.finished = true;
        if self.finish_error {
            return Err(BuildError {
                span: Span {
                    start: consumed,
                    end: consumed,
                },
                message: "sink finalization failed".into(),
            });
        }
        Ok(())
    }
}
fn grammar(entry: &str) -> grameo::CompiledGrammar {
    let source = format!(
        r#"
// @grameo pratt expression atom=number trailing=consume
// @grameo operator expression plus precedence=10 fixity=infix_left
entry={{{entry}}}
statement={{"emit"~expression}}
expression={{number~(plus~number)*}}
plus=@{{"+"}}
number=@{{ASCII_DIGIT+}}
WHITESPACE=_{{" "|"\n"}}
"#
    );
    grameo::compile_with_options(
        &source,
        grameo::CompileOptions {
            optimizations: true,
            ..Default::default()
        },
    )
    .unwrap()
}
#[test]
fn actual_typed_ast_to_vm_code_emits_during_matching_without_a_tree_or_module() {
    let g = grammar("SOI~(statement~\";\")*~EOI");
    let mut bridge = StreamingBuilder::new(Producer::new(&g, true), VmSink::default());
    let consumed = grameo::engine::build_ast_program(
        &g,
        "entry",
        "emit 2+3+4; emit 7;",
        MatchOptions::default(),
        &mut bridge,
    )
    .unwrap();
    assert_eq!(consumed, 19);
    assert_eq!(bridge.sink().run(), [9, 7]);
    assert_eq!(bridge.sink().attempts.borrow().len(), 2);
    assert_eq!(bridge.builder().commits, 1);
    assert!(
        bridge
            .sink()
            .code
            .metadata
            .iter()
            .all(
                |(scope, directives, _)| *scope == "caller lexical namespace"
                    && directives.len() == 3
            )
    );
    let (producer, sink) = bridge.into_parts();
    assert_eq!(producer.state.units.len(), 2);
    assert_eq!(sink.run(), [9, 7]);
}
#[test]
fn peg_capture_ast_hooks_spans_and_scope_directives_are_preserved() {
    let g = grammar("SOI~(statement~\";\")*~EOI");
    let s = "emit 2+3;";
    let mut original = Producer::new(&g, false);
    let original_tree = grameo::engine::build_capture_program(
        &g,
        "entry",
        s,
        MatchOptions::default(),
        &mut original,
    )
    .unwrap();
    let mut bridge = StreamingBuilder::new(Producer::new(&g, false), VmSink::default());
    let tree =
        grameo::engine::build_capture_program(&g, "entry", s, MatchOptions::default(), &mut bridge)
            .unwrap();
    fn snapshot(t: &grameo::ParseTree<'_, '_>) -> String {
        let mut all = vec![];
        let mut pending = t.pairs().map(|p| (p, 0)).collect::<Vec<_>>();
        while let Some((p, d)) = pending.pop() {
            all.push((p.as_rule(), p.as_span(), p.as_str().to_owned(), d));
            pending.extend(p.into_inner().map(|p| (p, d + 1)));
        }
        format!("{all:?}")
    }
    assert_eq!(snapshot(&tree), snapshot(&original_tree));
    assert_eq!(bridge.builder().state, original.state);
    assert_eq!(bridge.sink().run(), [5]);
}
#[test]
fn failed_alternative_retracts_already_emitted_instructions_before_fallback() {
    let g = grammar("SOI~((statement~\"x\")|statement)~\";\"~EOI");
    let mut bridge = StreamingBuilder::new(Producer::new(&g, true), VmSink::default());
    grameo::engine::build_ast_program(&g, "entry", "emit 2;", MatchOptions::default(), &mut bridge)
        .unwrap();
    assert!(
        bridge.sink().attempts.borrow().len() >= 2,
        "emission must actually happen before the failed enclosing suffix"
    );
    assert_eq!(bridge.sink().run(), [2]);
    assert_eq!(bridge.builder().state.units.len(), 1);
}
#[test]
fn syntax_failure_restores_both_initial_states() {
    let g = grammar("SOI~(statement~\";\")*~EOI");
    let mut sink = VmSink::default();
    sink.code.ops.push(Op::Const(99));
    let initial = sink.code.clone();
    let mut bridge = StreamingBuilder::new(Producer::new(&g, true), sink);
    let error = grameo::engine::build_ast_program(
        &g,
        "entry",
        "emit 2; broken",
        MatchOptions::default(),
        &mut bridge,
    )
    .unwrap_err();
    assert_eq!(error.kind, ParseErrorKind::Syntax);
    assert_eq!(bridge.sink().code, initial);
    assert_eq!(bridge.builder().state, State::default());
    assert_eq!(bridge.builder().commits, 0);
}
#[test]
fn checked_sink_error_is_fatal_not_an_ordered_choice_fallback() {
    let g = grammar("SOI~(statement|statement)~\";\"~EOI");
    let mut bridge = StreamingBuilder::new(
        Producer::new(&g, true),
        VmSink {
            reject: true,
            ..Default::default()
        },
    );
    let e = grameo::engine::build_ast_program(
        &g,
        "entry",
        "emit 2;",
        MatchOptions::default(),
        &mut bridge,
    )
    .unwrap_err();
    assert_eq!(e.kind, ParseErrorKind::BuildFailure);
    assert_eq!(bridge.sink().attempts.borrow().len(), 1);
    assert_eq!(bridge.sink().code, Code::default());
    assert_eq!(bridge.builder().state, State::default());
}
#[test]
fn checked_unit_finalization_error_rolls_back_instructions_and_semantics() {
    let g = grammar("SOI~statement~\";\"~EOI");
    let mut bridge = StreamingBuilder::new(
        Producer::new(&g, true),
        VmSink {
            finish_error: true,
            ..Default::default()
        },
    );
    let e = grameo::engine::build_ast_program(
        &g,
        "entry",
        "emit 2;",
        MatchOptions::default(),
        &mut bridge,
    )
    .unwrap_err();
    assert_eq!(e.kind, ParseErrorKind::BuildFailure);
    assert_eq!(e.offset, 7);
    assert_eq!(bridge.sink().code, Code::default());
    assert_eq!(bridge.builder().state, State::default());
    assert_eq!(bridge.builder().commits, 0);
}
#[test]
fn lookahead_never_emits_code() {
    let g = grammar("SOI~&statement~statement~\";\"~EOI");
    let mut bridge = StreamingBuilder::new(Producer::new(&g, true), VmSink::default());
    grameo::engine::build_ast_program(&g, "entry", "emit 2;", MatchOptions::default(), &mut bridge)
        .unwrap();
    assert_eq!(bridge.sink().attempts.borrow().len(), 1);
    assert_eq!(bridge.sink().run(), [2]);
}
#[test]
fn all_bounded_work_and_depth_failures_restore_initial_state() {
    let g = grammar("SOI~(statement~\";\")*~EOI");
    for depth in 0..10 {
        for work in 0..160 {
            let mut bridge = StreamingBuilder::new(Producer::new(&g, true), VmSink::default());
            let result = grameo::engine::build_ast_program(
                &g,
                "entry",
                "emit 2+3; emit 7;",
                MatchOptions {
                    max_rule_depth: depth,
                    max_steps: work,
                },
                &mut bridge,
            );
            if result.is_err() {
                assert_eq!(
                    bridge.sink().code,
                    Code::default(),
                    "depth={depth} work={work}"
                );
                assert_eq!(bridge.builder().state, State::default());
                assert_eq!(bridge.builder().commits, 0)
            } else {
                assert_eq!(bridge.sink().run(), [5, 7]);
            }
        }
    }
}
#[test]
fn absolute_parse_at_emission_span_is_not_rebased() {
    let g = grammar("SOI~statement~EOI");
    let mut bridge = StreamingBuilder::new(Producer::new(&g, true), VmSink::default());
    let tree = grameo::engine::build_capture_program_at(
        &g,
        "statement",
        "xxemit 2tail",
        2,
        MatchOptions::default(),
        &mut bridge,
    )
    .unwrap();
    assert_eq!(tree.consumed(), 8);
    assert_eq!(bridge.sink().code.metadata[0].2, Span { start: 2, end: 8 });
}

#[test]
fn rejected_pratt_preserves_granted_receipts_but_never_emitted_code() {
    let g = grammar("SOI~statement~EOI");
    let mut producer = Producer::new(&g, true);
    producer.retain_receipt = true;
    let number = producer.number;
    let statement = producer.statement;
    let mut bridge = StreamingBuilder::new(producer, VmSink::default());
    let outer = bridge.checkpoint();
    bridge.try_begin(number, 0).unwrap();
    bridge
        .try_finish(number, Span { start: 0, end: 1 }, "2")
        .unwrap();
    bridge
        .try_finish(statement, Span { start: 0, end: 1 }, "2")
        .unwrap();
    assert!(!bridge.sink().code.ops.is_empty());
    bridge.rollback_pratt(outer, statement, Span { start: 0, end: 1 });
    assert_eq!(
        bridge.builder().state.events,
        [(usize::MAX, Span { start: 0, end: 1 })]
    );
    assert_eq!(bridge.sink().code, Code::default());
    assert!(bridge.checkpoint() > outer);
    bridge.rollback(outer);
    assert_eq!(bridge.builder().state, State::default());
    assert_eq!(bridge.sink().code, Code::default());
}
#[test]
fn later_source_unit_failure_preserves_previously_validated_unit() {
    let g = grammar(r#"SOI~statement~";"~EOI"#);
    let mut bridge = StreamingBuilder::new(Producer::new(&g, true), VmSink::default());
    grameo::engine::build_ast_program(&g, "entry", "emit 2;", MatchOptions::default(), &mut bridge)
        .unwrap();
    let code = bridge.sink().code.clone();
    let state = bridge.builder().state.clone();
    assert!(
        grameo::engine::build_ast_program(
            &g,
            "entry",
            "emit 7; error",
            MatchOptions::default(),
            &mut bridge
        )
        .is_err()
    );
    assert_eq!(bridge.sink().code, code);
    assert_eq!(bridge.builder().state, state);
    grameo::engine::build_ast_program(&g, "entry", "emit 3;", MatchOptions::default(), &mut bridge)
        .unwrap();
    assert_eq!(bridge.sink().run(), [2, 3]);
    assert_eq!(bridge.builder().commits, 2);
}

#[test]
fn checked_commit_failure_restores_plain_builders_in_every_entry_lane() {
    #[derive(Default)]
    struct FinalFailure(Vec<usize>);
    impl Builder for FinalFailure {
        fn checkpoint(&self) -> usize {
            self.0.len()
        }
        fn rollback(&mut self, n: usize) {
            self.0.truncate(n)
        }
        fn begin(&mut self, r: usize, _: usize) {
            self.0.push(r)
        }
        fn finish(&mut self, r: usize, _: Span, _: &str) {
            self.0.push(r)
        }
        fn try_committed(&mut self, _: &str, consumed: usize) -> Result<(), BuildError> {
            self.0.push(usize::MAX);
            Err(BuildError {
                span: Span {
                    start: consumed,
                    end: consumed,
                },
                message: "final semantic validation".into(),
            })
        }
    }
    let g = grameo::compile("entry={\"x\"}").unwrap();
    // This literal-only fixture has one small scope and no grammar stack.
    // Compact entry points require this certificate, unlike arbitrary Program.
    #[derive(Debug)]
    struct Compact<'a>(&'a grameo::CompiledGrammar);
    impl grameo::program::Program for Compact<'_> {
        fn rule_id(&self, n: &str) -> Option<usize> {
            self.0.rule_id(n)
        }
        fn rule(&self, r: usize) -> grameo::program::RuleSpec<'_> {
            grameo::program::Program::rule(self.0, r)
        }
        fn instruction(&self, r: usize) -> grameo::program::Instruction<'_> {
            grameo::program::Program::instruction(self.0, r)
        }
        fn scope(&self, r: usize) -> grameo::program::Scope {
            grameo::program::Program::scope(self.0, r)
        }
        fn compact_capture_scope(&self, r: usize) -> grameo::program::CompactScope {
            grameo::program::CompactScope::new(r)
        }
    }
    let compact = Compact(&g);
    for lane in 0..5 {
        let mut b = FinalFailure(vec![99]);
        let e = match lane {
            0 => grameo::engine::build_program(&g, "entry", "x", MatchOptions::default(), &mut b)
                .unwrap_err(),
            1 => {
                grameo::engine::build_ast_program(&g, "entry", "x", MatchOptions::default(), &mut b)
                    .unwrap_err()
            }
            2 => grameo::engine::build_capture_program(
                &g,
                "entry",
                "x",
                MatchOptions::default(),
                &mut b,
            )
            .unwrap_err(),
            3 => grameo::engine::build_compact_capture_program(
                &compact,
                "entry",
                "x",
                MatchOptions::default(),
                &mut b,
            )
            .unwrap_err(),
            _ => grameo::engine::build_stackfree_compact_capture_program(
                &compact,
                "entry",
                "x",
                MatchOptions::default(),
                &mut b,
            )
            .unwrap_err(),
        };
        assert_eq!(e.kind, ParseErrorKind::BuildFailure);
        assert_eq!(e.offset, 1);
        assert_eq!(b.0, [99], "lane={lane}");
    }
}

#[test]
fn borrowed_producer_and_sink_use_the_same_checked_path() {
    let g = grammar(r#"SOI~statement~";"~EOI"#);
    let mut producer = Producer::new(&g, true);
    let mut sink = VmSink::default();
    {
        let mut bridge = StreamingBuilder::new(&mut producer, &mut sink);
        grameo::engine::build_ast_program(
            &g,
            "entry",
            "emit 6;",
            MatchOptions::default(),
            &mut bridge,
        )
        .unwrap();
    }
    assert_eq!(producer.state.units.len(), 1);
    assert_eq!(sink.run(), [6]);
}
