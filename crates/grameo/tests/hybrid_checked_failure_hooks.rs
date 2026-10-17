use grameo::{
    CompiledGrammar, MatchOptions, Span,
    builder::{BuildError, Builder},
    program::{Instruction, NativeChoiceEvent, NativeChoicePrefix, Program, RuleSpec, Scope},
};
#[derive(Debug)]
struct P<'a> {
    g: &'a CompiledGrammar,
    fast: bool,
    native: bool,
}
static PREFIX: NativeChoicePrefix<'static> = NativeChoicePrefix {
    bytes: [0, 1u64 << (b'a' - 64), 0, 0],
    steps: 4,
    depth: 1,
    diagnostic_complete: true,
    events: &[
        NativeChoiceEvent::Call {
            expression: 0,
            rule: 1,
        },
        NativeChoiceEvent::Expression(3),
    ],
};
impl Program for P<'_> {
    fn rule_id(&self, n: &str) -> Option<usize> {
        self.g.rule_id(n)
    }
    fn rule(&self, r: usize) -> RuleSpec<'_> {
        self.g.rule(r)
    }
    fn instruction(&self, e: usize) -> Instruction<'_> {
        self.g.instruction(e)
    }
    fn scope(&self, s: usize) -> Scope {
        self.g.scope(s)
    }
    fn fast_rule_failures(&self) -> bool {
        self.fast
    }
    fn fast_trivia_failures(&self) -> bool {
        false
    }
    fn fast_predicate_failures(&self) -> bool {
        false
    }
    fn trivia_failure_prefix(&self, r: usize) -> Option<grameo::lexical::FailurePrefix> {
        self.g.trivia_failure_prefix(r)
    }
    fn failure_terminals(&self, r: usize) -> &[usize] {
        self.g.failure_terminals(r)
    }
    fn native_hybrid_choices(&self) -> bool {
        self.native
    }
    fn native_hybrid_choice_prefix(&self, e: usize) -> Option<&NativeChoicePrefix<'_>> {
        (self.native && e == 2).then_some(&PREFIX)
    }
}
#[derive(Default, Debug)]
struct B {
    observed: usize,
    fail: bool,
    inert: bool,
    attempts: Vec<(usize, usize)>,
    journal: Vec<(usize, Span)>,
}
impl Builder for B {
    fn observes_rule(&self, r: usize) -> bool {
        r == self.observed
    }
    fn failure_prefix_hooks_are_inert(&self, r: usize) -> bool {
        self.inert && r == self.observed
    }
    fn checkpoint(&self) -> usize {
        self.journal.len()
    }
    fn rollback(&mut self, n: usize) {
        self.journal.truncate(n)
    }
    fn begin(&mut self, r: usize, p: usize) {
        self.journal.push((r, Span { start: p, end: p }))
    }
    fn finish(&mut self, r: usize, s: Span, _: &str) {
        self.journal.push((r, s))
    }
    fn try_begin(&mut self, r: usize, p: usize) -> Result<(), BuildError> {
        self.attempts.push((r, p));
        if self.fail {
            Err(BuildError {
                span: Span { start: p, end: p },
                message: format!("checked rule {r}"),
            })
        } else {
            self.begin(r, p);
            Ok(())
        }
    }
}
fn snapshot(t: grameo::ParseTree<'_, '_>) -> (usize, Vec<(usize, Span, usize)>) {
    let mut out = Vec::new();
    let mut tasks: Vec<_> = t.pairs().map(|p| (p, 0)).collect();
    while let Some((p, d)) = tasks.pop() {
        out.push((
            match p.as_rule() {
                grameo::CaptureRule::User(r) => r,
                _ => usize::MAX,
            },
            p.as_span(),
            d,
        ));
        tasks.extend(p.into_inner().map(|p| (p, d + 1)));
    }
    (t.consumed(), out)
}
#[test]
fn plain_default_peg_checked_error_matches_semantic_only() {
    for source in [
        "entry={a|b}\na={\"a\"}\nb={\"b\"}",
        "entry={outer|b}\nouter={child}\nchild={\"a\"}\nb={\"b\"}",
        "entry=@{a|b}\na={\"a\"}\nb={\"b\"}",
    ] {
        let g = grameo::compile(source).unwrap();
        let observed = g
            .rule_id(if source.contains("child") {
                "child"
            } else {
                "a"
            })
            .unwrap();
        let mut a = B {
            observed,
            fail: true,
            ..Default::default()
        };
        let mut b = B {
            observed,
            fail: true,
            ..Default::default()
        };
        let x = grameo::engine::build_program(&g, "entry", "b", MatchOptions::default(), &mut a)
            .unwrap_err();
        let y = grameo::engine::build_capture_program(
            &g,
            "entry",
            "b",
            MatchOptions::default(),
            &mut b,
        )
        .unwrap_err();
        assert_eq!(x, y);
        assert_eq!(a.attempts, b.attempts);
        assert_eq!(x.kind, grameo::ParseErrorKind::BuildFailure);
    }
}
#[test]
fn hybrid_default_hooks_errors_and_limits_equal_no_shortcuts() {
    let mut cases = 0;
    for source in [
        "entry={a|b}\na={\"a\"}\nb={\"b\"}",
        "entry={outer|b}\nouter={child}\nchild={\"a\"}\nb={\"b\"}",
        "entry=@{a|b}\na={\"a\"}\nb={\"b\"}",
    ] {
        let g = grameo::compile(source).unwrap();
        let observed = g
            .rule_id(if source.contains("child") {
                "child"
            } else {
                "a"
            })
            .unwrap();
        for input in ["a", "b", "c", "", "aa", "α"] {
            for fail in [false, true] {
                for depth in 0..=6 {
                    for work in (0..=120).step_by(3).chain([usize::MAX]) {
                        let options = MatchOptions {
                            max_steps: work,
                            max_rule_depth: depth,
                        };
                        let mut a = B {
                            observed,
                            fail,
                            ..Default::default()
                        };
                        let mut b = B {
                            observed,
                            fail,
                            ..Default::default()
                        };
                        let x = grameo::engine::build_capture_program(
                            &P {
                                g: &g,
                                fast: false,
                                native: false,
                            },
                            "entry",
                            input,
                            options,
                            &mut a,
                        )
                        .map(snapshot);
                        let y = grameo::engine::build_capture_program(
                            &P {
                                g: &g,
                                fast: true,
                                native: false,
                            },
                            "entry",
                            input,
                            options,
                            &mut b,
                        )
                        .map(snapshot);
                        assert_eq!(
                            x, y,
                            "source={source} input={input:?} depth={depth} work={work} fail={fail}"
                        );
                        assert_eq!(a.attempts, b.attempts);
                        assert_eq!(a.journal, b.journal);
                        cases += 1;
                    }
                }
            }
        }
    }
    println!("exact hybrid reference cases={cases}");
}
#[test]
fn optimized_choice_prefix_preserves_checked_error_and_attempts() {
    let g = grameo::compile("entry={a|b}\na={\"a\"}\nb={\"b\"}").unwrap();
    for fail in [false, true] {
        for input in ["a", "b", "c", ""] {
            for work in [0, 1, 4, 8, 16, 32, usize::MAX] {
                for depth in [0, 1, 2, 3, 9] {
                    let mut a = B {
                        observed: 1,
                        fail,
                        ..Default::default()
                    };
                    let mut b = B {
                        observed: 1,
                        fail,
                        ..Default::default()
                    };
                    let o = MatchOptions {
                        max_steps: work,
                        max_rule_depth: depth,
                    };
                    let x = grameo::engine::build_capture_program(
                        &P {
                            g: &g,
                            fast: false,
                            native: false,
                        },
                        "entry",
                        input,
                        o,
                        &mut a,
                    )
                    .map(snapshot);
                    let y = grameo::engine::build_capture_program(
                        &P {
                            g: &g,
                            fast: true,
                            native: true,
                        },
                        "entry",
                        input,
                        o,
                        &mut b,
                    )
                    .map(snapshot);
                    assert_eq!(x, y, "input{input:?} work{work} depth{depth} fail{fail}");
                    assert_eq!(a.attempts, b.attempts);
                    assert_eq!(a.journal, b.journal);
                }
            }
        }
    }
}
#[test]
fn explicit_inert_permission_can_still_omit_checked_hook_query() {
    let g = grameo::compile("entry={a|b}\na={\"a\"}\nb={\"b\"}").unwrap();
    let mut b = B {
        observed: 1,
        inert: true,
        ..Default::default()
    };
    let p = P {
        g: &g,
        fast: true,
        native: true,
    };
    let tree =
        grameo::engine::build_capture_program(&p, "entry", "b", MatchOptions::default(), &mut b)
            .unwrap();
    assert_eq!(tree.consumed(), 1);
    assert!(b.attempts.is_empty());
}
