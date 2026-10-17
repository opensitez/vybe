use grameo::{CompileOptions, MatchOptions, Span, builder::Builder, program::Program};
const BODY: &str = r#"root={SOI ~ word ~ pad ~ "!" ~ EOI}
word=@{ASCII_ALPHA+}
pad=_{("\n"|"\r\n"|";")*}
stack={PUSH("x") ~ pad ~ POP}
predicate={&pad ~ word}
backtrack={(pad ~ "bad") | (pad ~ word)}
atomic=@{pad ~ word}
compound=${pad ~ word}
"#;
fn grammars() -> (grameo::CompiledGrammar, grameo::CompiledGrammar) {
    let o = CompileOptions {
        optimizations: true,
        ..Default::default()
    };
    (
        grameo::compile_with_options(BODY, o).unwrap(),
        grameo::compile_with_options(&format!("// @grameo empty_success_heads pad\n{BODY}"), o)
            .unwrap(),
    )
}
fn shape(t: &grameo::ParseTree<'_, '_>) -> Vec<(String, Span, usize)> {
    fn visit(p: grameo::tree::Pair<'_, '_, '_>, d: usize, o: &mut Vec<(String, Span, usize)>) {
        o.push((p.rule_name().into(), p.as_span(), d));
        for c in p.into_inner() {
            visit(c, d + 1, o)
        }
    }
    let mut o = vec![];
    for p in t.pairs() {
        visit(p, 0, &mut o)
    }
    o
}
#[derive(Default)]
struct Denied {
    trace: std::cell::RefCell<Vec<(u8, usize)>>,
}
impl Builder for Denied {
    fn observes_rule(&self, r: usize) -> bool {
        self.trace.borrow_mut().push((0, r));
        false
    }
    fn observes_capture(&self, r: usize) -> bool {
        self.trace.borrow_mut().push((1, r));
        false
    }
    fn opaque_rule(&self, r: usize) -> bool {
        self.trace.borrow_mut().push((2, r));
        false
    }
    fn checkpoint(&self) -> usize {
        self.trace.borrow_mut().push((3, 0));
        0
    }
    fn rollback(&mut self, m: usize) {
        self.trace.borrow_mut().push((4, m));
    }
    fn begin(&mut self, r: usize, _: usize) {
        self.trace.borrow_mut().push((5, r));
    }
    fn finish(&mut self, r: usize, _: Span, _: &str) {
        self.trace.borrow_mut().push((6, r));
    }
}
#[test]
fn denied_builder_preserves_every_legacy_query() {
    let (a, b) = grammars();
    for entry in [
        "root",
        "pad",
        "stack",
        "predicate",
        "backtrack",
        "atomic",
        "compound",
    ] {
        for source in ["", "abc!", "xx", "x\nx", "\nabc!", "abc?", "α"] {
            for steps in [0, 1, 8, 64, usize::MAX] {
                let opts = MatchOptions {
                    max_steps: steps,
                    max_rule_depth: 4096,
                };
                let mut x = Denied::default();
                let mut y = Denied::default();
                let aa = grameo::engine::build_capture_program(&a, entry, source, opts, &mut x);
                let bb = grameo::engine::build_capture_program(&b, entry, source, opts, &mut y);
                match (aa, bb) {
                    (Ok(a), Ok(b)) => assert_eq!(shape(&a), shape(&b)),
                    (Err(a), Err(b)) => assert_eq!(a, b),
                    (a, b) => panic!("{entry}{source:?}{a:?}{b:?}"),
                };
                assert_eq!(
                    x.trace.into_inner(),
                    y.trace.into_inner(),
                    "{entry}{source:?}"
                );
            }
        }
    }
}
struct Pure;
impl Builder for Pure {
    fn empty_success_is_inert(&self, _: usize) -> bool {
        true
    }
    fn observes_rule(&self, _: usize) -> bool {
        false
    }
    fn observes_capture(&self, _: usize) -> bool {
        false
    }
    fn checkpoint(&self) -> usize {
        0
    }
    fn rollback(&mut self, _: usize) {}
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, _: usize, _: Span, _: &str) {}
}
#[test]
fn selected_captures_stack_errors_and_limits_remain_exact() {
    let (a, b) = grammars();
    for e in [
        "root",
        "pad",
        "stack",
        "predicate",
        "backtrack",
        "atomic",
        "compound",
    ] {
        for s in ["", "abc!", "xx", "x\nx", "\nabc!", "abc?", "α"] {
            for max_steps in [0, 1, 8, 64, usize::MAX] {
                for max_rule_depth in [1, 2, 8, 4096] {
                    let o = MatchOptions {
                        max_steps,
                        max_rule_depth,
                    };
                    let aa = grameo::engine::build_capture_program(&a, e, s, o, &mut Pure);
                    let bb = grameo::engine::build_capture_program(&b, e, s, o, &mut Pure);
                    match (aa, bb) {
                        (Ok(a), Ok(b)) => assert_eq!(shape(&a), shape(&b)),
                        (Err(a), Err(b)) => assert_eq!(a, b),
                        (a, b) => panic!("{e}{s:?}{a:?}{b:?}"),
                    };
                }
            }
        }
    }
}
#[test]
fn permission_is_explicit_and_default_grammar_unchanged() {
    let (a, b) = grammars();
    let r = b.rule_id("pad").unwrap();
    assert!(!a.empty_success_candidate(r));
    assert!(b.empty_success_candidate(r));
    let raw = grameo::compile(&format!("// @grameo empty_success_heads pad\n{BODY}")).unwrap();
    assert!(!raw.empty_success_candidate(r));
    assert!(!Denied::default().empty_success_is_inert(r));
}

#[derive(Debug)]
struct PermissionTrace<'a> {
    inner: &'a grameo::CompiledGrammar,
    queries: std::cell::Cell<usize>,
}
impl Program for PermissionTrace<'_> {
    fn rule_id(&self, n: &str) -> Option<usize> {
        self.inner.rule_id(n)
    }
    fn rule(&self, r: usize) -> grameo::program::RuleSpec<'_> {
        self.inner.rule(r)
    }
    fn instruction(&self, e: usize) -> grameo::program::Instruction<'_> {
        self.inner.instruction(e)
    }
    fn scope(&self, s: usize) -> grameo::program::Scope {
        self.inner.scope(s)
    }
    fn empty_success_candidate(&self, r: usize) -> bool {
        self.inner.empty_success_candidate(r)
    }
    fn empty_success_head(&self, r: usize) -> Option<&grameo::program::EmptySuccessHead> {
        self.inner.empty_success_head(r)
    }
    fn match_observers_are_noop(&self) -> bool {
        self.queries.set(self.queries.get() + 1);
        true
    }
}
#[test]
fn metadata_forwarding_without_freshness_preserves_legacy_permission_queries() {
    let (a, b) = grammars();
    for source in ["", "abc!", "xx", "\n", "!"] {
        let x = PermissionTrace {
            inner: &a,
            queries: Default::default(),
        };
        let y = PermissionTrace {
            inner: &b,
            queries: Default::default(),
        };
        let aa = grameo::engine::build_capture_program(
            &x,
            "root",
            source,
            MatchOptions::default(),
            &mut Pure,
        );
        let bb = grameo::engine::build_capture_program(
            &y,
            "root",
            source,
            MatchOptions::default(),
            &mut Pure,
        );
        match (aa, bb) {
            (Ok(a), Ok(b)) => assert_eq!(shape(&a), shape(&b)),
            (Err(a), Err(b)) => assert_eq!(a, b),
            (a, b) => panic!("{a:?}{b:?}"),
        };
        assert_eq!(x.queries.get(), y.queries.get());
    }
}
