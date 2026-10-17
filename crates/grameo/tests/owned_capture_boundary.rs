use grameo::{
    CompiledGrammar, MatchOptions, Span,
    builder::{BuildError, Builder},
    engine,
    program::{Instruction, OwnedCaptureBoundary, Program, RuleSpec, Scope},
};
use std::cell::{Cell, RefCell};
#[derive(Debug)]
struct P<'g> {
    g: &'g CompiledGrammar,
    target: usize,
    cert: bool,
    current: bool,
    noop: bool,
    rules: &'static [usize],
    queries: Cell<usize>,
    events: RefCell<Vec<(usize, usize, Option<u8>)>>,
}
impl Program for P<'_> {
    fn rule_id(&self, n: &str) -> Option<usize> {
        self.g.rule_id(n)
    }
    fn rule(&self, r: usize) -> RuleSpec<'_> {
        self.g.rule(r)
    }
    fn instruction(&self, r: usize) -> Instruction<'_> {
        self.g.instruction(r)
    }
    fn scope(&self, s: usize) -> Scope {
        self.g.scope(s)
    }
    fn match_observers_are_noop(&self) -> bool {
        self.noop
    }
    fn owned_capture_boundary_metadata_is_current(&self) -> bool {
        self.current
    }
    fn owned_capture_boundary(&self, r: usize) -> Option<OwnedCaptureBoundary<'_>> {
        self.queries.set(self.queries.get() + 1);
        (self.cert && r == self.target).then(|| {
            let d = self.g.rule(r);
            OwnedCaptureBoundary {
                expression: d.expression,
                scope: d.scope,
                capture_rules: self.rules,
            }
        })
    }
    fn observe_expression(&self, id: usize, pos: usize) {
        if !self.noop {
            self.events.borrow_mut().push((id, pos, None))
        }
    }
    fn observe_rule_attempt(&self, id: usize, pos: usize, flags: u8) {
        if !self.noop {
            self.events.borrow_mut().push((id, pos, Some(flags)))
        }
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
enum Ev {
    Begin(usize, usize),
    Finish(usize, Span),
}
struct B {
    target: usize,
    permission: bool,
    bad_end: bool,
    observed: Option<usize>,
    checked: bool,
    events: Vec<Ev>,
    ends: Cell<usize>,
}
impl Builder for B {
    fn observes_rule(&self, _: usize) -> bool {
        true
    }
    fn checkpoint(&self) -> usize {
        self.events.len()
    }
    fn rollback(&mut self, n: usize) {
        self.events.truncate(n)
    }
    fn begin(&mut self, r: usize, pos: usize) {
        self.events.push(Ev::Begin(r, pos))
    }
    fn finish(&mut self, r: usize, s: Span, _: &str) {
        self.events.push(Ev::Finish(r, s))
    }
    fn try_finish(&mut self, r: usize, s: Span, source: &str) -> Result<(), BuildError> {
        if self.checked && r == self.target && source[s.start..s.end].contains("bad") {
            return Err(BuildError {
                span: s,
                message: "checked owned constructor".into(),
            });
        }
        self.finish(r, s, source);
        Ok(())
    }
    fn observes_capture(&self, r: usize) -> bool {
        self.observed == Some(r)
    }
    fn supports_owned_capture_boundary(&self, r: usize) -> bool {
        self.permission && r == self.target
    }
    fn owned_capture_end(&self, _: usize, start: usize, source: &str) -> Option<usize> {
        self.ends.set(self.ends.get() + 1);
        if self.bad_end {
            return Some(start);
        }
        let b = source.as_bytes();
        let mut i = start;
        while b.get(i).is_some_and(|c| c.is_ascii_alphabetic()) {
            i += 1
        }
        if i == start {
            return None;
        }
        loop {
            match b.get(i) {
                Some(b'.') => {
                    i += 1;
                    let j = i;
                    while b.get(i).is_some_and(|c| c.is_ascii_alphabetic()) {
                        i += 1
                    }
                    if i == j {
                        return None;
                    }
                }
                Some(b'(') => {
                    if b.get(i + 1) != Some(&b')') {
                        return None;
                    }
                    i += 2
                }
                _ => break,
            }
        }
        Some(i)
    }
}
fn gram(mode: &str) -> CompiledGrammar {
    grameo::compile(&format!(r#"WHITESPACE=_{{" "}} root={{SOI~parent~EOI}} parent={{(owned~"!")|(owned~("+"~owned)*)}} owned={mode}{{name~suffix*}} name={{ASCII_ALPHA+}} suffix={{"."~name|"("~args?~")"}} args={{name~(","~name)*}}"#)).unwrap()
}
fn b(g: &CompiledGrammar, permission: bool) -> B {
    B {
        target: g.rule_id("owned").unwrap(),
        permission,
        bad_end: false,
        observed: None,
        checked: false,
        events: vec![],
        ends: Cell::new(0),
    }
}
fn p(g: &CompiledGrammar, cert: bool, current: bool, noop: bool) -> P<'_> {
    let rules: Vec<_> = [
        "WHITESPACE",
        "root",
        "parent",
        "owned",
        "name",
        "suffix",
        "args",
    ]
    .iter()
    .map(|n| g.rule_id(n).unwrap())
    .collect();
    P {
        g,
        target: g.rule_id("owned").unwrap(),
        cert,
        current,
        noop,
        rules: Box::leak(rules.into_boxed_slice()),
        queries: Cell::new(0),
        events: RefCell::new(vec![]),
    }
}
fn snap(t: &grameo::ParseTree<'_, '_>) -> Vec<(String, Span)> {
    let mut pending: Vec<_> = t.pairs().collect();
    let mut out = vec![];
    while let Some(p) = pending.pop() {
        out.push((p.rule_name().into(), p.as_span()));
        pending.extend(p.into_inner())
    }
    out
}
#[test]
fn owned_roots_retain_anchors_and_all_callbacks_with_fewer_captures() {
    let g = gram("");
    let a = p(&g, false, true, true);
    let c = p(&g, true, true, true);
    for source in ["a", "a.b()", "a.b()+c()", "a!", "a+b.c()"] {
        let (mut x, mut y) = (b(&g, false), b(&g, true));
        let old =
            engine::build_capture_program(&a, "root", source, MatchOptions::default(), &mut x)
                .unwrap();
        let new =
            engine::build_capture_program(&c, "root", source, MatchOptions::default(), &mut y)
                .unwrap();
        assert_eq!(x.events, y.events, "{source}");
        assert!(old.capture_count() > new.capture_count(), "{source}");
        let olds = snap(&old);
        let news = snap(&new);
        for (name, span) in &news {
            assert!(olds.contains(&(name.clone(), *span)))
        }
        assert_eq!(
            olds.iter()
                .filter(|(n, _)| n == "owned")
                .collect::<Vec<_>>(),
            news.iter()
                .filter(|(n, _)| n == "owned")
                .collect::<Vec<_>>()
        );
        eprintln!(
            "{source} captures {} -> {}",
            old.capture_count(),
            new.capture_count()
        );
    }
}
#[test]
fn mismatch_replays_from_post_begin_without_duplicate_semantics() {
    let g = gram("");
    let a = p(&g, false, true, true);
    let c = p(&g, true, true, true);
    for source in ["a.b()", "a(x)", "a.b()+c()", "a(x)+b(y)"] {
        let (mut x, mut y) = (b(&g, false), b(&g, true));
        y.bad_end = true;
        let old =
            engine::build_capture_program(&a, "root", source, MatchOptions::default(), &mut x)
                .map(|t| snap(&t));
        let new =
            engine::build_capture_program(&c, "root", source, MatchOptions::default(), &mut y)
                .map(|t| snap(&t));
        assert_eq!(old, new, "{source}");
        assert_eq!(x.events, y.events, "{source}");
        assert!(y.ends.get() > 0);
    }
}
#[test]
fn default_work_observers_child_capture_hooks_and_modes_keep_original() {
    let mut cases = 0;
    for mode in ["", "_", "@", "$", "!"] {
        let g = gram(mode);
        for source in [
            "a",
            "a.b()",
            "a(x)",
            "a+",
            "a(",
            "",
            "a.b()+c!",
            "雪",
            "a . b ( )",
        ] {
            for (work, depth) in [
                (0, 0),
                (1, 1),
                (20, 4),
                (200, 20),
                (usize::MAX, 4),
                (200, usize::MAX),
                (usize::MAX, usize::MAX),
            ] {
                for denial in 0..4 {
                    let a = p(&g, false, false, denial != 1);
                    let c = p(&g, true, denial != 0, denial != 1);
                    let (mut x, mut y) = (b(&g, false), b(&g, true));
                    if denial == 2 {
                        y.permission = false
                    }
                    if denial == 3 {
                        let observed = g.rule_id("name");
                        x.observed = observed;
                        y.observed = observed
                    }
                    let opts = MatchOptions {
                        max_steps: work,
                        max_rule_depth: depth,
                    };
                    let old = engine::build_capture_program(&a, "root", source, opts, &mut x)
                        .map(|t| snap(&t));
                    let new = engine::build_capture_program(&c, "root", source, opts, &mut y)
                        .map(|t| snap(&t));
                    if old.is_err() || work != usize::MAX || mode != "" || denial <= 3 {
                        assert_eq!(old, new, "{source:?} mode={mode} denial={denial} {opts:?}");
                    }
                    assert_eq!(x.events, y.events);
                    if denial == 1 {
                        assert_eq!(*a.events.borrow(), *c.events.borrow())
                    }
                    if denial <= 1 || denial == 2 || mode != "" || work != usize::MAX {
                        assert_eq!(y.ends.get(), 0, "prediction queried on denied path")
                    };
                    cases += 1;
                }
            }
        }
    }
    eprintln!("default/resource comparisons={cases}");
}
#[test]
fn checked_errors_restore_all_journals() {
    let g = gram("");
    for source in ["bad", "a+bad", "bad()", "a.b()+bad"] {
        let a = p(&g, false, true, true);
        let c = p(&g, true, true, true);
        let (mut x, mut y) = (b(&g, false), b(&g, true));
        x.checked = true;
        y.checked = true;
        let old =
            engine::build_capture_program(&a, "root", source, MatchOptions::default(), &mut x)
                .map(|t| snap(&t));
        let new =
            engine::build_capture_program(&c, "root", source, MatchOptions::default(), &mut &mut y)
                .map(|t| snap(&t));
        assert_eq!(old, new);
        assert!(old.is_err());
        assert!(x.events.is_empty() && y.events.is_empty());
    }
}

#[test]
fn default_depth_is_not_raised_and_depth_failures_are_exact() {
    let g = gram("");
    let a = p(&g, false, true, true);
    let c = p(&g, true, true, true);
    for source in ["a.b()+c()", "a!", "a(x)", "a+", "a("] {
        for depth in 0..10 {
            let (mut x, mut y) = (b(&g, false), b(&g, true));
            let opts = MatchOptions {
                max_steps: usize::MAX,
                max_rule_depth: depth,
            };
            let old = engine::build_capture_program(&a, "root", source, opts, &mut x);
            let new = engine::build_capture_program(&c, "root", source, opts, &mut y);
            match (old, new) {
                (Err(x), Err(y)) => assert_eq!(x, y, "{source} depth={depth}"),
                (Ok(_), Ok(_)) => {}
                _ => panic!("success mismatch depth={depth}"),
            };
            assert_eq!(x.events, y.events);
        }
    }
}
#[test]
fn retained_root_capture_hooks_receive_valid_exact_span_handles() {
    use grameo::tree::CaptureHandle;
    struct Root {
        inner: B,
        tokens: Vec<(usize, CaptureHandle)>,
    }
    impl Builder for Root {
        fn observes_rule(&self, r: usize) -> bool {
            self.inner.observes_rule(r)
        }
        fn checkpoint(&self) -> usize {
            self.inner.checkpoint()
        }
        fn rollback(&mut self, m: usize) {
            self.inner.rollback(m);
            self.tokens.retain(|(i, _)| *i < m)
        }
        fn begin(&mut self, r: usize, o: usize) {
            self.inner.begin(r, o)
        }
        fn finish(&mut self, r: usize, s: Span, t: &str) {
            self.inner.finish(r, s, t)
        }
        fn supports_owned_capture_boundary(&self, r: usize) -> bool {
            self.inner.supports_owned_capture_boundary(r)
        }
        fn owned_capture_end(&self, r: usize, o: usize, t: &str) -> Option<usize> {
            self.inner.owned_capture_end(r, o, t)
        }
        fn observes_capture(&self, r: usize) -> bool {
            r == self.inner.target
        }
        fn try_finish_capture(
            &mut self,
            r: usize,
            s: Span,
            t: &str,
            h: Option<CaptureHandle>,
        ) -> Result<(), BuildError> {
            if let Some(h) = h {
                self.tokens.push((self.inner.events.len(), h))
            }
            self.inner.try_finish(r, s, t)
        }
    }
    #[derive(Clone, Copy)]
    struct Id;
    impl grameo::compat::RuleIdentity for Id {
        fn from_capture(_: grameo::CaptureRule) -> Option<Self> {
            Some(Self)
        }
    }
    let g = gram("");
    let c = p(&g, true, true, true);
    for source in ["a", "a.b()+c()", "a(x)"] {
        let mut b = Root {
            inner: b(&g, true),
            tokens: vec![],
        };
        let tree =
            engine::build_capture_program(&c, "root", source, MatchOptions::default(), &mut &mut b)
                .unwrap();
        let root = tree.into_typed_pairs::<Id>().next().unwrap();
        assert!(!b.tokens.is_empty());
        for (_, h) in b.tokens {
            let pair = root.resolve_capture(h).expect("retained root token");
            assert!(
                source
                    .get(pair.as_span().start()..pair.as_span().end())
                    .is_some()
            );
        }
    }
}
mod generated {
    include!("support/owned_capture_boundary/optimized.rs");
}
mod generated_default {
    include!("support/owned_capture_boundary/peg.rs");
}
#[test]
fn generated_program_default_denials_resources_and_rollbacks() {
    let old = generated_default::Parser;
    let new = generated::Parser;
    let g = gram("");
    let mut cases = 0;
    for source in [
        "a",
        "a.b()",
        "a.b()+c()",
        "a(x)",
        "a . b ( )",
        "a+",
        "a(",
        "雪",
        "",
    ] {
        for (work, depth) in [
            (usize::MAX, 4096),
            (usize::MAX, 1),
            (usize::MAX, 5),
            (0, 0),
            (200, 20),
        ] {
            for denial in 0..3 {
                let (mut x, mut y) = (b(&g, false), b(&g, true));
                if denial == 0 {
                    y.permission = false
                }
                if denial == 1 {
                    x.observed = g.rule_id("name");
                    y.observed = x.observed
                }
                if denial == 2 {
                    y.bad_end = true
                }
                let opts = MatchOptions {
                    max_steps: work,
                    max_rule_depth: depth,
                };
                let a = engine::build_capture_program(&old, "root", source, opts, &mut x)
                    .map(|t| snap(&t));
                let c = engine::build_capture_program(&new, "root", source, opts, &mut y)
                    .map(|t| snap(&t));
                assert_eq!(a, c, "{source} {opts:?} {denial}");
                assert_eq!(x.events, y.events);
                cases += 1;
            }
        }
    }
    eprintln!("generated exact default/denial/resource/replay cases={cases}");
}
#[test]
fn generated_program_owned_and_runtime_metadata_agree() {
    let source = "a.b()+c()";
    let g = gram("");
    let (mut x, mut y) = (b(&g, false), b(&g, true));
    let old = engine::build_capture_program(
        &generated_default::Parser,
        "root",
        source,
        MatchOptions::default(),
        &mut x,
    )
    .unwrap();
    let new = engine::build_capture_program(
        &generated::Parser,
        "root",
        source,
        MatchOptions::default(),
        &mut y,
    )
    .unwrap();
    assert_eq!(x.events, y.events);
    assert_eq!(old.capture_count(), 11);
    assert_eq!(new.capture_count(), 5);
    let text = "// @grameo owned_capture_boundary owned\nWHITESPACE=_{\" \"} root={SOI~parent~EOI} parent={(owned~\"!\")|(owned~(\"+\"~owned)*)} owned={name~suffix*} name={ASCII_ALPHA+} suffix={\".\"~name|\"(\"~args?~\")\"} args={name~(\",\"~name)*}";
    let runtime = grameo::compile_with_options(
        text,
        grameo::CompileOptions {
            optimizations: true,
            ..Default::default()
        },
    )
    .unwrap();
    let mut z = b(&runtime, true);
    let t =
        engine::build_capture_program(&runtime, "root", source, MatchOptions::default(), &mut z)
            .unwrap();
    assert_eq!(snap(&new), snap(&t));
    assert_eq!(y.events, z.events);
}
#[test]
fn generated_forwarding_wrappers_default_to_no_permission() {
    #[derive(Debug)]
    struct Wrapper {
        p: generated::Parser,
        events: RefCell<Vec<(usize, usize)>>,
    }
    impl Program for Wrapper {
        fn rule_id(&self, n: &str) -> Option<usize> {
            self.p.rule_id(n)
        }
        fn rule(&self, r: usize) -> RuleSpec<'_> {
            self.p.rule(r)
        }
        fn instruction(&self, e: usize) -> Instruction<'_> {
            self.p.instruction(e)
        }
        fn scope(&self, s: usize) -> Scope {
            self.p.scope(s)
        }
        fn owned_capture_boundary(&self, r: usize) -> Option<OwnedCaptureBoundary<'_>> {
            self.p.owned_capture_boundary(r)
        }
        fn observe_expression(&self, e: usize, o: usize) {
            self.events.borrow_mut().push((e, o));
        }
    }
    let g = gram("");
    let p = Wrapper {
        p: generated::Parser,
        events: RefCell::new(vec![]),
    };
    let (mut x, mut y) = (b(&g, false), b(&g, true));
    let a = engine::build_capture_program(&p, "root", "a.b()", MatchOptions::default(), &mut x)
        .map(|t| snap(&t));
    let events = p.events.replace(vec![]);
    let c = engine::build_capture_program(&p, "root", "a.b()", MatchOptions::default(), &mut y)
        .map(|t| snap(&t));
    assert_eq!(a, c);
    assert_eq!(events, *p.events.borrow());
    assert_eq!(x.events, y.events);
    assert_eq!(y.ends.get(), 0);
}
