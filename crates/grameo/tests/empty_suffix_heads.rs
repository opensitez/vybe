#![allow(dead_code, mismatched_lifetime_syntaxes)]
include!("fixtures/empty_suffix_heads.rs");
use grameo::{MatchOptions, Span, builder::Builder, program::Program};
fn shape(t: &grameo::ParseTree) -> Vec<(String, Span, usize)> {
    fn visit(p: grameo::tree::Pair, d: usize, o: &mut Vec<(String, Span, usize)>) {
        o.push((p.rule_name().into(), p.as_span(), d));
        for c in p.into_inner() {
            visit(c, d + 1, o)
        }
    }
    let mut v = vec![];
    for p in t.pairs() {
        visit(p, 0, &mut v)
    }
    v
}
#[derive(Default)]
struct B {
    allow: bool,
    observe: bool,
    query: std::cell::RefCell<Vec<(u8, usize)>>,
    events: Vec<(usize, Span)>,
}
impl Builder for B {
    fn empty_suffix_is_inert(&self, _: usize) -> bool {
        self.allow
    }
    fn observes_rule(&self, r: usize) -> bool {
        if !self.allow {
            self.query.borrow_mut().push((0, r));
        }
        self.observe
    }
    fn observes_capture(&self, r: usize) -> bool {
        if !self.allow {
            self.query.borrow_mut().push((1, r));
        }
        false
    }
    fn checkpoint(&self) -> usize {
        if !self.allow {
            self.query.borrow_mut().push((2, self.events.len()));
        }
        self.events.len()
    }
    fn rollback(&mut self, m: usize) {
        if !self.allow {
            self.query.borrow_mut().push((3, m));
        }
        self.events.truncate(m)
    }
    fn begin(&mut self, r: usize, o: usize) {
        self.events.push((r, Span { start: o, end: o }))
    }
    fn finish(&mut self, r: usize, s: Span, _: &str) {
        self.events.push((r, s))
    }
}
#[test]
fn exact_selected_denied_observed_and_resources() {
    let mut comparisons = 0;
    for allow in [false, true] {
        for observe in [false, true] {
            for entry in [
                "root",
                "stack",
                "predicate",
                "negative",
                "backtrack",
                "atomic",
                "compound",
            ] {
                for source in [
                    "", "a!", "a?z!", "a ? z !", "a;?z!", "a\n?z!", "a/*x*/!", "a//x\n!", "a?",
                    "a x!", "a!held", "a!x", "é!", "\na!", "a\r\n!", "a! ",
                ] {
                    for steps in [0, 1, 32, 128, usize::MAX] {
                        for depth in [0, 1, 2, 8, 64, 70, 4096] {
                            let o = MatchOptions {
                                max_steps: steps,
                                max_rule_depth: depth,
                            };
                            let mut x = B {
                                allow,
                                observe,
                                ..Default::default()
                            };
                            let mut y = B {
                                allow,
                                observe,
                                ..Default::default()
                            };
                            let aa = grameo::engine::build_capture_program(
                                &a::Parser,
                                entry,
                                source,
                                o,
                                &mut x,
                            );
                            let bb = grameo::engine::build_capture_program(
                                &b::Parser,
                                entry,
                                source,
                                o,
                                &mut y,
                            );
                            match (aa, bb) {
                                (Ok(a), Ok(b)) => {
                                    assert_eq!(shape(&a), shape(&b), "{entry}{source:?}")
                                }
                                (Err(a), Err(b)) => assert_eq!(a, b, "{entry}{source:?}"),
                                (a, b) => panic!("{entry}{source:?}{a:?}{b:?}"),
                            };
                            assert_eq!(x.events, y.events);
                            if !allow {
                                assert_eq!(x.query.into_inner(), y.query.into_inner());
                            }
                            comparisons += 1;
                        }
                    }
                }
            }
        }
    }
    println!("{comparisons} exact comparisons");
}
#[derive(Debug)]
struct Forward<P> {
    inner: P,
    change: u8,
    events: std::cell::RefCell<Vec<(u8, usize, usize, usize)>>,
    permission_queries: std::cell::Cell<usize>,
}
impl<P: Program> Program for Forward<P> {
    fn rule_id(&self, n: &str) -> Option<usize> {
        self.inner.rule_id(n)
    }
    fn rule(&self, r: usize) -> grameo::program::RuleSpec {
        self.inner.rule(r)
    }
    fn instruction(&self, id: usize) -> grameo::program::Instruction {
        let v = self.inner.instruction(id);
        if self.change == 1 {
            if let grameo::program::Instruction::Literal { text: "?", .. } = v {
                return grameo::program::Instruction::Literal {
                    text: ":",
                    insensitive: false,
                };
            }
        }
        v
    }
    fn scope(&self, s: usize) -> grameo::program::Scope {
        let mut scope = self.inner.scope(s);
        if self.change == 2 {
            scope.whitespace = None;
        }
        scope
    }
    fn source_pratt(&self, r: usize) -> Option<grameo::program::SourcePratt> {
        if self.change == 3 && r == a::Rule::pad as usize {
            Some(grameo::program::SourcePratt {
                atom: a::Rule::word as usize,
                operators: &[],
                trailing_trivia: grameo::islands::TrailingTrivia::Preserve,
            })
        } else {
            self.inner.source_pratt(r)
        }
    }
    fn observe_rule_attempt(&self, r: usize, o: usize, m: u8) {
        self.events.borrow_mut().push((0, r, o, m as usize));
    }
    fn observe_expression(&self, id: usize, o: usize) {
        self.events.borrow_mut().push((1, id, o, 0));
    }
    fn empty_suffix_candidate(&self, id: usize) -> bool {
        self.inner.empty_suffix_candidate(id)
    }
    fn empty_suffix_head(&self, id: usize) -> Option<&grameo::program::EmptySuffixHead> {
        self.inner.empty_suffix_head(id)
    }
    fn match_observers_are_noop(&self) -> bool {
        self.permission_queries
            .set(self.permission_queries.get() + 1);
        true
    }
}
#[test]
fn forwarded_metadata_without_permission_preserves_queries() {
    for s in ["a!", "a?z!", "a?", "", "a/*x*/!"] {
        let x = Forward {
            inner: a::Parser,
            change: 0,
            events: Default::default(),
            permission_queries: Default::default(),
        };
        let y = Forward {
            inner: b::Parser,
            change: 0,
            events: Default::default(),
            permission_queries: Default::default(),
        };
        let mut bx = B {
            allow: true,
            ..Default::default()
        };
        let mut by = B {
            allow: true,
            ..Default::default()
        };
        let aa =
            grameo::engine::build_capture_program(&x, "root", s, MatchOptions::default(), &mut bx);
        let bb =
            grameo::engine::build_capture_program(&y, "root", s, MatchOptions::default(), &mut by);
        match (aa, bb) {
            (Ok(a), Ok(b)) => assert_eq!(shape(&a), shape(&b)),
            (Err(a), Err(b)) => assert_eq!(a, b),
            _ => panic!("parity"),
        };
        assert_eq!(x.permission_queries.get(), y.permission_queries.get());
    }
}

#[derive(Default)]
struct Tokens {
    fail: bool,
    events: Vec<(usize, Span, Option<grameo::tree::CaptureHandle>)>,
}
impl Builder for Tokens {
    fn empty_suffix_is_inert(&self, _: usize) -> bool {
        true
    }
    fn observes_rule(&self, r: usize) -> bool {
        r != a::Rule::pad as usize
    }
    fn observes_capture(&self, r: usize) -> bool {
        r == a::Rule::word as usize || r == a::Rule::root as usize
    }
    fn checkpoint(&self) -> usize {
        self.events.len()
    }
    fn rollback(&mut self, m: usize) {
        self.events.truncate(m)
    }
    fn begin(&mut self, r: usize, o: usize) {
        self.events.push((r, Span { start: o, end: o }, None))
    }
    fn finish(&mut self, r: usize, s: Span, _: &str) {
        self.events.push((r, s, None))
    }
    fn try_finish_capture(
        &mut self,
        r: usize,
        s: Span,
        _: &str,
        c: Option<grameo::tree::CaptureHandle>,
    ) -> Result<(), grameo::builder::BuildError> {
        if self.fail && r == a::Rule::root as usize {
            return Err(grameo::builder::BuildError {
                span: s,
                message: "checked root".into(),
            });
        }
        self.events.push((r, s, c));
        Ok(())
    }
}
#[test]
fn issued_handles_and_checked_error_journals() {
    let mut n = 0;
    for fail in [false, true] {
        for entry in [
            "root",
            "stack",
            "predicate",
            "negative",
            "backtrack",
            "atomic",
            "compound",
        ] {
            for s in [
                "", "a!", "a?z!", "a;?z!", "a/*x*/!", "a\n?z!", "a x!", "a!x", "a!held",
            ] {
                for steps in [0, 128, usize::MAX] {
                    for depth in [2, 70, 4096] {
                        let o = MatchOptions {
                            max_steps: steps,
                            max_rule_depth: depth,
                        };
                        let mut x = Tokens {
                            fail,
                            ..Default::default()
                        };
                        let mut y = Tokens {
                            fail,
                            ..Default::default()
                        };
                        let aa =
                            grameo::engine::build_capture_program(&a::Parser, entry, s, o, &mut x);
                        let bb =
                            grameo::engine::build_capture_program(&b::Parser, entry, s, o, &mut y);
                        match (aa, bb) {
                            (Ok(a), Ok(b)) => assert_eq!(shape(&a), shape(&b)),
                            (Err(a), Err(b)) => assert_eq!(a, b),
                            _ => panic!("token outcome {entry}{s:?}"),
                        };
                        assert_eq!(x.events, y.events, "tokens {entry}{s:?}");
                        n += 1;
                    }
                }
            }
        }
    }
    println!("{n} exact issued-handle/checked-error/journal comparisons");
}

#[test]
fn transformed_geometry_scope_pratt_observers_keep_vm() {
    let mut count = 0;
    for change in [1, 2, 3] {
        for s in [
            "", "a!", "a:z!", "a?z!", "a ? z!", "a/*x*/!", "a//x\n!", "a\n!", "a;!",
        ] {
            for steps in [0, 32, usize::MAX] {
                for depth in [2, 70, 4096] {
                    let x = Forward {
                        inner: a::Parser,
                        change,
                        events: Default::default(),
                        permission_queries: Default::default(),
                    };
                    let y = Forward {
                        inner: b::Parser,
                        change,
                        events: Default::default(),
                        permission_queries: Default::default(),
                    };
                    let mut bx = B {
                        allow: true,
                        ..Default::default()
                    };
                    let mut by = B {
                        allow: true,
                        ..Default::default()
                    };
                    let o = MatchOptions {
                        max_steps: steps,
                        max_rule_depth: depth,
                    };
                    let aa = grameo::engine::build_capture_program(&x, "root", s, o, &mut bx);
                    let bb = grameo::engine::build_capture_program(&y, "root", s, o, &mut by);
                    match (aa, bb) {
                        (Ok(a), Ok(b)) => assert_eq!(shape(&a), shape(&b)),
                        (Err(a), Err(b)) => assert_eq!(a, b),
                        _ => panic!("transformed {change} {s:?}"),
                    };
                    assert_eq!(
                        x.events.into_inner(),
                        y.events.into_inner(),
                        "observer {change}{s:?}"
                    );
                    assert_eq!(x.permission_queries.get(), y.permission_queries.get());
                    count += 1;
                }
            }
        }
    }
    println!("{count} transformed geometry/scope/SourcePratt/observer comparisons");
}
#[test]
fn later_pratt_configuration_invalidates_compiled_certificate() {
    let source = r#"// @grameo empty_suffix_heads root
root={word ~ (pad ~ "?" ~ word)?}
word=@{ASCII_ALPHA+}
pad=_{";"*}
"#;
    let mut g = grameo::compile_with_options(
        source,
        grameo::CompileOptions {
            optimizations: true,
            ..Default::default()
        },
    )
    .unwrap();
    let owner = g.rule_id("root").unwrap();
    let mut pending = vec![g.rule(owner).expression];
    let mut selected = vec![];
    while let Some(id) = pending.pop() {
        if g.empty_suffix_candidate(id) {
            selected.push(id);
        }
        match g.instruction(id) {
            grameo::program::Instruction::Group(x) => pending.push(x),
            grameo::program::Instruction::Sequence { left, right }
            | grameo::program::Instruction::Choice { left, right } => {
                pending.push(left);
                pending.push(right);
            }
            _ => {}
        }
    }
    assert_eq!(selected.len(), 1);
    g.bind_pratt(
        "root",
        "word",
        &[],
        grameo::islands::TrailingTrivia::Preserve,
    )
    .unwrap();
    assert!(selected.into_iter().all(|id| !g.empty_suffix_candidate(id)));
}

#[test]
fn explicit_permissions_default_to_false() {
    use grameo::builder::Builder;
    let rule = a::Parser.rule_id("root").unwrap();
    assert!(!B::default().empty_suffix_is_inert(rule));
    let mut stack = vec![a::Parser.rule(rule).expression];
    let mut seen = false;
    while let Some(id) = stack.pop() {
        assert!(!a::Parser.empty_suffix_candidate(id));
        if b::Parser.empty_suffix_candidate(id) {
            assert!(b::Parser.empty_suffix_metadata_is_current());
            seen = true;
        }
        match a::Parser.instruction(id) {
            grameo::program::Instruction::Sequence { left, right }
            | grameo::program::Instruction::Choice { left, right } => {
                stack.push(left);
                stack.push(right);
            }
            grameo::program::Instruction::Group(x) => stack.push(x),
            _ => {}
        }
    }
    assert!(seen);
}
