use grameo::{
    MatchOptions,
    builder::Builder,
    program::{CompactScope, Instruction, NativeChoicePrefix, Program, RuleSpec, Scope},
};
use std::cell::{Cell, RefCell};
#[allow(dead_code)]
mod native {
    include!("fixtures/native_choices.rs");
}
struct Profile {
    enabled: bool,
    fast: bool,
    tails: bool,
    events: RefCell<Vec<String>>,
    queries: Cell<usize>,
}
impl std::fmt::Debug for Profile {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str("Profile")
    }
}
impl Program for Profile {
    fn rule_id(&self, n: &str) -> Option<usize> {
        native::Parser.rule_id(n)
    }
    fn rule(&self, id: usize) -> RuleSpec<'_> {
        native::Parser.rule(id)
    }
    fn instruction(&self, id: usize) -> Instruction<'_> {
        native::Parser.instruction(id)
    }
    fn scope(&self, id: usize) -> Scope {
        native::Parser.scope(id)
    }
    fn compact_capture_scope(&self, id: usize) -> CompactScope {
        native::Parser.compact_capture_scope(id)
    }
    fn trivia_failure_prefix(&self, id: usize) -> Option<grameo::lexical::FailurePrefix> {
        native::Parser.trivia_failure_prefix(id)
    }
    fn borrowed_failure_prefixes(&self) -> bool {
        true
    }
    fn trivia_failure_prefix_ref(&self, id: usize) -> Option<&grameo::lexical::FailurePrefix> {
        native::Parser.trivia_failure_prefix_ref(id)
    }
    fn failure_terminals(&self, id: usize) -> &[usize] {
        native::Parser.failure_terminals(id)
    }
    fn fast_rule_failures(&self) -> bool {
        self.fast
    }
    fn tail_dispatch(&self) -> bool {
        self.tails
    }
    fn tail_continuations(&self) -> bool {
        self.tails
    }
    fn native_hybrid_choices(&self) -> bool {
        self.enabled
    }
    fn native_hybrid_rule_head(&self, rule: usize) -> Option<&grameo::program::NativeRuleHead<'_>> {
        native::Parser.native_hybrid_rule_head(rule)
    }
    fn native_hybrid_choice_prefix(&self, id: usize) -> Option<&NativeChoicePrefix<'_>> {
        self.queries.set(self.queries.get() + 1);
        native::Parser.native_hybrid_choice_prefix(id)
    }
    fn observe_expression(&self, id: usize, offset: usize) {
        self.events.borrow_mut().push(format!("E{id}@{offset}"));
    }
    fn observe_rule_attempt(&self, id: usize, offset: usize, flags: u8) {
        self.events
            .borrow_mut()
            .push(format!("R{id}@{offset}:{flags}"));
    }
}
#[derive(Default)]
struct Journal {
    nodes: Vec<(usize, usize, usize)>,
    events: Vec<String>,
}
impl Builder for Journal {
    fn observes_rule(&self, _: usize) -> bool {
        true
    }
    fn checkpoint(&self) -> usize {
        self.nodes.len()
    }
    fn begin(&mut self, r: usize, o: usize) {
        self.nodes.push((r, o, o));
        self.events.push(format!("B{r}@{o}"));
    }
    fn finish(&mut self, r: usize, s: grameo::Span, _: &str) {
        self.nodes.push((r, s.start, s.end));
        self.events.push(format!("F{r}@{}:{}", s.start, s.end));
    }
    fn rollback(&mut self, m: usize) {
        self.nodes.truncate(m);
        self.events.push(format!("U{m}"));
    }
}
fn run(
    enabled: bool,
    fast: bool,
    compact: bool,
    tails: bool,
    s: &str,
    opts: MatchOptions,
) -> (
    String,
    Vec<String>,
    Vec<String>,
    Vec<(usize, usize, usize)>,
    usize,
) {
    let p = Profile {
        enabled,
        fast,
        tails,
        events: RefCell::new(Vec::new()),
        queries: Cell::new(0),
    };
    let mut b = Journal::default();
    let result = if compact {
        grameo::engine::build_compact_capture_program(&p, "entry", s, opts, &mut b)
    } else {
        grameo::engine::build_capture_program(&p, "entry", s, opts, &mut b)
    };
    let text = format!("{result:?}");
    if result.is_err() {
        assert!(b.nodes.is_empty(), "failure restores builder");
    }
    drop(result);
    (
        text,
        p.events.into_inner(),
        b.events,
        b.nodes,
        p.queries.get(),
    )
}
#[test]
fn generated_choice_prefixes_preserve_hooks_traces_stack_and_all_small_limits() {
    let sources = [
        "c",
        "axx",
        "axxc",
        "bzzc",
        "axxbzz",
        "a x x #note\nc",
        "a x x",
        "bzz",
        "b z z",
        "qq",
        "q",
        "a",
        "b",
        "z",
        "",
        " #note\nc",
        " #note\nqq",
    ];
    let mut cases = 0;
    let mut requests = 0;
    for compact in [false, true] {
        for fast in [false, true] {
            for tails in [false, true] {
                for s in sources {
                    for steps in 0..160 {
                        for depth in 0..9 {
                            let opts = MatchOptions {
                                max_steps: steps,
                                max_rule_depth: depth,
                            };
                            let a = run(false, fast, compact, tails, s, opts);
                            let b = run(true, fast, compact, tails, s, opts);
                            assert_eq!(
                                (&a.0, &a.1, &a.2, &a.3),
                                (&b.0, &b.1, &b.2, &b.3),
                                "source={s:?} steps={steps} depth={depth} compact={compact} fast={fast}"
                            );
                            if !fast {
                                assert_eq!(
                                    b.4, 0,
                                    "disabled original failure certificates disable native fusion"
                                );
                            }
                            requests += b.4;
                            cases += 1;
                        }
                    }
                    let a = run(false, fast, compact, tails, s, MatchOptions::default());
                    let b = run(true, fast, compact, tails, s, MatchOptions::default());
                    assert_eq!(
                        (&a.0, &a.1, &a.2, &a.3),
                        (&b.0, &b.1, &b.2, &b.3),
                        "unlimited {s:?}"
                    );
                }
            }
        }
    }
    assert!(requests > 0);
    println!("cases={cases} native_prefix_requests={requests}");
    let a = run(true, true, false, true, "c", MatchOptions::default());
    assert!(
        // A failed observed alternative still enters its semantic rule.
        // Checked begin errors must not disappear through a lexical shortcut.
        a.2.iter()
            .any(|x| x.starts_with(&format!("B{}@", native::Rule::a as usize)))
    );
    assert!(
        native::Parser
            .trivia_failure_prefix(native::Rule::unsafe_first as usize)
            .is_none(),
        "leading stack mutations are not certified"
    );
}
#[test]
fn ordinary_peg_ignores_native_hybrid_opt_in() {
    for s in ["c", "axx", "qq", "q", "z", ""] {
        for steps in 0..160 {
            let p = Profile {
                enabled: true,
                fast: true,
                tails: true,
                events: RefCell::new(Vec::new()),
                queries: Cell::new(0),
            };
            let _ = grameo::engine::parse_program(
                &p,
                "entry",
                s,
                MatchOptions {
                    max_steps: steps,
                    max_rule_depth: 8,
                },
            );
            assert_eq!(p.queries.get(), 0);
        }
    }
}
#[derive(Debug)]
struct WideIds {
    enabled: bool,
    events: RefCell<Vec<(usize, usize)>>,
}
impl Program for WideIds {
    fn rule_id(&self, n: &str) -> Option<usize> {
        (n == "entry").then_some(0)
    }
    fn rule(&self, id: usize) -> RuleSpec<'_> {
        RuleSpec {
            name: if id == 0 { "entry" } else { "wide" },
            mode: grameo::RuleMode::Silent,
            expression: if id == 0 { 0 } else { 3 },
            scope: if id == 0 { 0 } else { usize::MAX },
        }
    }
    fn scope(&self, _: usize) -> Scope {
        Scope::default()
    }
    fn instruction(&self, id: usize) -> Instruction<'_> {
        match id {
            0 => Instruction::Choice { left: 1, right: 2 },
            1 => Instruction::Call(usize::MAX),
            2 => Instruction::Literal {
                text: "z",
                insensitive: false,
            },
            3 => Instruction::Literal {
                text: "a",
                insensitive: false,
            },
            _ => panic!(),
        }
    }
    fn borrowed_failure_prefixes(&self) -> bool {
        true
    }
    fn trivia_failure_prefix_ref(&self, r: usize) -> Option<&grameo::lexical::FailurePrefix> {
        const P: grameo::lexical::FailurePrefix = grameo::lexical::FailurePrefix {
            bytes: [0, 1u64 << (b'a' - 64), 0, 0],
            steps: 1,
            depth: 0,
        };
        (r == usize::MAX).then_some(&P)
    }
    fn failure_terminals(&self, r: usize) -> &[usize] {
        if r == usize::MAX { &[3] } else { &[] }
    }
    fn native_hybrid_choices(&self) -> bool {
        self.enabled
    }
    fn native_hybrid_choice_prefix(&self, id: usize) -> Option<&NativeChoicePrefix<'_>> {
        const P: NativeChoicePrefix<'static> = NativeChoicePrefix {
            bytes: [0, 1u64 << (b'a' - 64), 0, 0],
            steps: 4,
            depth: 1,
            diagnostic_complete: true,
            events: &[grameo::program::NativeChoiceEvent::Call {
                expression: 1,
                rule: usize::MAX,
            }],
        };
        (id == 0).then_some(&P)
    }
    fn observe_rule_attempt(&self, id: usize, o: usize, _: u8) {
        self.events.borrow_mut().push((id, o));
    }
}
#[test]
fn native_choice_observers_and_fallback_preserve_full_width_rule_and_scope_ids() {
    for s in ["z", "a", "", "x"] {
        for steps in 0..96 {
            for depth in 0..5 {
                let mut results = Vec::new();
                for enabled in [false, true] {
                    let p = WideIds {
                        enabled,
                        events: RefCell::new(Vec::new()),
                    };
                    let mut b = Journal::default();
                    let result = grameo::engine::build_capture_program(
                        &p,
                        "entry",
                        s,
                        MatchOptions {
                            max_steps: steps,
                            max_rule_depth: depth,
                        },
                        &mut b,
                    );
                    let text = if let Ok(ref t) = result {
                        format!("{}:{}", t.consumed(), t.capture_count())
                    } else {
                        format!("{:?}", result.unwrap_err())
                    };
                    results.push((text, p.events.into_inner(), b.events, b.nodes));
                }
                assert_eq!(results[0], results[1], "{s:?} steps={steps} depth={depth}");
            }
        }
    }
}
struct Reject {
    journal: Journal,
    rule: usize,
    on_finish: bool,
}
impl Builder for Reject {
    fn observes_rule(&self, _: usize) -> bool {
        true
    }
    fn checkpoint(&self) -> usize {
        self.journal.checkpoint()
    }
    fn rollback(&mut self, m: usize) {
        self.journal.rollback(m)
    }
    fn begin(&mut self, r: usize, o: usize) {
        self.journal.begin(r, o)
    }
    fn finish(&mut self, r: usize, s: grameo::Span, text: &str) {
        self.journal.finish(r, s, text)
    }
    fn try_begin(&mut self, r: usize, o: usize) -> Result<(), grameo::builder::BuildError> {
        if r == self.rule && !self.on_finish {
            return Err(grameo::builder::BuildError {
                span: grameo::Span { start: o, end: o },
                message: "rejected begin".into(),
            });
        }
        self.begin(r, o);
        Ok(())
    }
    fn try_finish(
        &mut self,
        r: usize,
        s: grameo::Span,
        text: &str,
    ) -> Result<(), grameo::builder::BuildError> {
        if r == self.rule && self.on_finish {
            return Err(grameo::builder::BuildError {
                span: s,
                message: "rejected finish".into(),
            });
        }
        self.finish(r, s, text);
        Ok(())
    }
}
#[test]
fn native_choices_preserve_possible_branch_semantic_rejection_and_rollback() {
    for source in ["c", "axx", "q"] {
        for rule in [native::Rule::a as usize, native::Rule::c as usize] {
            for on_finish in [false, true] {
                for steps in 0..128 {
                    for depth in 0..6 {
                        let mut values = Vec::new();
                        for enabled in [false, true] {
                            let p = Profile {
                                enabled,
                                fast: true,
                                tails: true,
                                events: RefCell::new(Vec::new()),
                                queries: Cell::new(0),
                            };
                            let mut b = Reject {
                                journal: Journal::default(),
                                rule,
                                on_finish,
                            };
                            let r = grameo::engine::build_capture_program(
                                &p,
                                "entry",
                                source,
                                MatchOptions {
                                    max_steps: steps,
                                    max_rule_depth: depth,
                                },
                                &mut b,
                            );
                            values.push((format!("{r:?}"), b.journal.nodes, b.journal.events));
                        }
                        assert_eq!(
                            values[0], values[1],
                            "{source:?} rule={rule} finish={on_finish} steps={steps} depth={depth}"
                        );
                    }
                }
            }
        }
    }
}
#[derive(Debug)]
struct WideHead {
    enabled: bool,
    events: RefCell<Vec<(bool, usize, usize)>>,
}
impl Program for WideHead {
    fn rule_id(&self, name: &str) -> Option<usize> {
        (name == "entry").then_some(0)
    }
    fn rule(&self, id: usize) -> RuleSpec<'_> {
        RuleSpec {
            name: if id == 0 { "entry" } else { "wide" },
            mode: grameo::RuleMode::Silent,
            expression: if id == 0 { usize::MAX - 1 } else { 1 },
            scope: usize::MAX,
        }
    }
    fn scope(&self, _: usize) -> Scope {
        Scope::default()
    }
    fn instruction(&self, id: usize) -> Instruction<'_> {
        match id {
            x if x == usize::MAX - 1 => Instruction::Sequence {
                left: usize::MAX,
                right: usize::MAX - 2,
            },
            usize::MAX => Instruction::Call(usize::MAX),
            x if x == usize::MAX - 2 => Instruction::Literal {
                text: "z",
                insensitive: false,
            },
            1 => Instruction::Literal {
                text: "a",
                insensitive: false,
            },
            _ => panic!("unexpected instruction {id}"),
        }
    }
    fn native_hybrid_choices(&self) -> bool {
        self.enabled
    }
    fn native_hybrid_rule_head(&self, id: usize) -> Option<&grameo::program::NativeRuleHead<'_>> {
        static HEAD: grameo::program::NativeRuleHead<'static> = grameo::program::NativeRuleHead {
            sequences: &[(usize::MAX - 1, usize::MAX - 2)],
            call_expression: usize::MAX,
            callee: usize::MAX,
        };
        (id == 0).then_some(&HEAD)
    }
    fn tail_dispatch(&self) -> bool {
        true
    }
    fn tail_continuations(&self) -> bool {
        true
    }
    fn observe_expression(&self, id: usize, offset: usize) {
        self.events.borrow_mut().push((false, id, offset));
    }
    fn observe_rule_attempt(&self, id: usize, offset: usize, _: u8) {
        self.events.borrow_mut().push((true, id, offset));
    }
}
#[test]
fn native_heads_preserve_full_width_expression_rule_scope_and_suffix_limits() {
    for source in ["az", "a", "z", "", "x"] {
        for steps in 0..96 {
            for depth in 0..5 {
                let mut results = Vec::new();
                for enabled in [false, true] {
                    let p = WideHead {
                        enabled,
                        events: RefCell::new(Vec::new()),
                    };
                    let mut b = Journal::default();
                    let result = grameo::engine::build_capture_program(
                        &p,
                        "entry",
                        source,
                        MatchOptions {
                            max_steps: steps,
                            max_rule_depth: depth,
                        },
                        &mut b,
                    );
                    let text = match result {
                        Ok(t) => format!("{}:{}", t.consumed(), t.capture_count()),
                        Err(e) => format!("{e:?}"),
                    };
                    results.push((text, p.events.into_inner(), b.events, b.nodes));
                }
                assert_eq!(
                    results[0], results[1],
                    "source={source:?} steps={steps} depth={depth}"
                );
            }
        }
    }
}
