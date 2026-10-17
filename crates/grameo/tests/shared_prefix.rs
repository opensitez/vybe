use grameo::{
    CaptureRule, CompileOptions, Span,
    program::{Instruction, Program, RuleSpec, Scope, SharedPrefixChoice, SourcePratt},
};
use std::cell::RefCell;
#[derive(Debug, Clone, Copy)]
enum Mutation {
    None,
    Pratt(usize),
    WrapperBody(usize),
    Scope,
    ScopeData,
    Observe,
}
#[derive(Debug)]
struct Forward<'a> {
    base: &'a grameo::CompiledGrammar,
    mutation: Mutation,
    wrapper: usize,
    branch: usize,
    other: usize,
    events: RefCell<Vec<(u8, usize, usize)>>,
}
impl Program for Forward<'_> {
    fn rule_id(&self, n: &str) -> Option<usize> {
        self.base.rule_id(n)
    }
    fn rule(&self, id: usize) -> RuleSpec<'_> {
        let mut s = self.base.rule(id);
        if matches!(self.mutation, Mutation::Scope) && id == self.wrapper {
            s.scope = 1;
        }
        s
    }
    fn instruction(&self, id: usize) -> Instruction<'_> {
        if matches!(self.mutation,Mutation::WrapperBody(n) if n==id) {
            Instruction::Call(self.other)
        } else {
            self.base.instruction(id)
        }
    }
    fn source_pratt(&self, id: usize) -> Option<SourcePratt<'_>> {
        if matches!(self.mutation,Mutation::Pratt(n) if n==id) {
            Some(SourcePratt {
                atom: self.other,
                operators: &[],
                trailing_trivia: grameo::islands::TrailingTrivia::Consume,
            })
        } else {
            self.base.source_pratt(id)
        }
    }
    fn scope(&self, id: usize) -> Scope {
        let mut scope = self.base.scope(if id == 1 { 0 } else { id });
        if matches!(self.mutation, Mutation::ScopeData) {
            scope.whitespace = Some(self.other);
            scope.whitespace_class = None;
            scope.whitespace_prefix_class = None;
        }
        scope
    }
    fn shared_prefix_choice(&self, id: usize) -> Option<SharedPrefixChoice> {
        self.base.shared_prefix_choice(id)
    }
    fn match_observers_are_noop(&self) -> bool {
        !matches!(self.mutation, Mutation::Observe)
    }
    fn observe_rule_attempt(&self, id: usize, pos: usize, _: u8) {
        if matches!(self.mutation, Mutation::Observe) {
            self.events.borrow_mut().push((0, id, pos));
        }
    }
    fn observe_expression(&self, id: usize, pos: usize) {
        if matches!(self.mutation, Mutation::Observe) {
            self.events.borrow_mut().push((1, id, pos));
        }
    }
}
struct B {
    enabled: bool,
    value: usize,
    pratt: Option<usize>,
    tokens: bool,
    handles: Vec<grameo::tree::CaptureHandle>,
}
impl grameo::builder::Builder for B {
    fn supports_shared_prefix(&self, _: usize) -> bool {
        self.enabled
    }
    fn observes_rule(&self, r: usize) -> bool {
        self.tokens && r == self.value
    }
    fn observes_capture(&self, r: usize) -> bool {
        self.tokens && r == self.value
    }
    fn supports_pratt(&self, r: usize) -> bool {
        self.pratt == Some(r)
    }
    fn checkpoint(&self) -> usize {
        self.handles.len()
    }
    fn rollback(&mut self, n: usize) {
        self.handles.truncate(n)
    }
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, _: usize, _: Span, _: &str) {}
    fn try_finish_capture(
        &mut self,
        _: usize,
        _: Span,
        _: &str,
        h: Option<grameo::tree::CaptureHandle>,
    ) -> Result<(), grameo::builder::BuildError> {
        if let Some(h) = h {
            self.handles.push(h)
        }
        Ok(())
    }
}
#[derive(Clone, Copy)]
struct Identity;
impl grameo::compat::RuleIdentity for Identity {
    fn from_capture(_: CaptureRule) -> Option<Self> {
        Some(Self)
    }
}
fn shape(t: &grameo::ParseTree<'_, '_>) -> Vec<(CaptureRule, Span, usize)> {
    fn v(p: grameo::Pair<'_, '_, '_>, d: usize, out: &mut Vec<(CaptureRule, Span, usize)>) {
        out.push((p.as_rule(), p.as_span(), d));
        for c in p.into_inner() {
            v(c, d + 1, out)
        }
    }
    let mut out = vec![];
    for p in t.pairs() {
        v(p, 0, &mut out)
    }
    out
}
#[test]
fn shared_prefix_preserves_forwarded_program_contracts_and_capture_handles() {
    let mut cases = 0;
    for mode in ["{", "@{", "${"] {
        let source=format!("// @grameo shared_prefix arg branch=gen prefix=value\nWHITESPACE = _{{ \\\" \\\" }}\nprogram = {mode} arg ~ EOI }}\narg = {{ first | gen | value }}\nfirst = {{ \\\"!\\\" }}\ngen = {{ value ~ suffix }}\nvalue = {{ \\\"x\\\" }}\nsuffix = {{ \\\"?\\\" }}\nother = {{ \\\"y\\\" }}\n").replace("\\\"","\"");
        let g = grameo::compile_with_options(
            &source,
            CompileOptions {
                optimizations: true,
                ..Default::default()
            },
        )
        .unwrap();
        let target = g.rule_id("arg").unwrap();
        let wrapper = g.rule_id("gen").unwrap();
        let value = g.rule_id("value").unwrap();
        let suffix = g.rule_id("suffix").unwrap();
        let other = g.rule_id("other").unwrap();
        let branch = g.rule(wrapper).expression;
        for mutation in [
            Mutation::None,
            Mutation::Pratt(target),
            Mutation::Pratt(wrapper),
            Mutation::Pratt(value),
            Mutation::Pratt(suffix),
            Mutation::WrapperBody(branch),
            Mutation::Scope,
            Mutation::ScopeData,
            Mutation::Observe,
        ] {
            for input in ["x", "x?", "x ?", "y", "xy", "xy?", "x y", "!", "?", ""] {
                for depth in [1, 2, 3, 4, 5, 6, 4096] {
                    for tokens in [false, true] {
                        let p = Forward {
                            base: &g,
                            mutation,
                            wrapper,
                            branch,
                            other,
                            events: RefCell::new(vec![]),
                        };
                        let pratt = if let Mutation::Pratt(r) = mutation {
                            Some(r)
                        } else {
                            None
                        };
                        let mut a = B {
                            enabled: false,
                            value,
                            pratt,
                            tokens,
                            handles: vec![],
                        };
                        let mut b = B {
                            enabled: true,
                            value,
                            pratt,
                            tokens,
                            handles: vec![],
                        };
                        let opts = grameo::MatchOptions {
                            max_steps: usize::MAX,
                            max_rule_depth: depth,
                        };
                        let x = grameo::engine::build_capture_program(
                            &p, "program", input, opts, &mut a,
                        );
                        let events = p.events.replace(vec![]);
                        let y = grameo::engine::build_capture_program(
                            &p, "program", input, opts, &mut b,
                        );
                        assert_eq!(
                            events,
                            *p.events.borrow(),
                            "observer {mode} {mutation:?} {input}"
                        );
                        match (x, y) {
                            (Ok(x), Ok(y)) => {
                                assert_eq!(
                                    shape(&x),
                                    shape(&y),
                                    "{mode} {mutation:?} {input} depth{depth}"
                                );
                                for (tree, handles) in [(x, a.handles), (y, b.handles)] {
                                    let root = tree.into_typed_pairs::<Identity>().next().unwrap();
                                    for h in handles {
                                        assert!(
                                            root.resolve_capture(h).is_some(),
                                            "invalidated capture handle {mode} {input}"
                                        );
                                    }
                                }
                            }
                            (Err(x), Err(y)) => {
                                assert_eq!(x, y, "{mode} {mutation:?} {input} depth{depth}")
                            }
                            _ => panic!("acceptance {mode} {mutation:?} {input} depth{depth}"),
                        };
                        cases += 1;
                    }
                }
            }
        }
    }
    let source = "// @grameo shared_prefix arg branch=gen prefix=value\nprogram = { seed ~ arg ~ EOI }\nseed = { \"s\" }\narg = { first | gen | value }\nfirst = { \"!\" }\ngen = { value ~ suffix }\nvalue = { \"x\" }\nsuffix = { \"?\" }\n";
    let g = grameo::compile_with_options(
        source,
        CompileOptions {
            optimizations: true,
            ..Default::default()
        },
    )
    .unwrap();
    for input in ["sx", "sx?", "s!", "sxz"] {
        for depth in [1, 2, 3, 4, 5, 6, 4096] {
            let seed = g.rule_id("seed").unwrap();
            let opts = grameo::MatchOptions {
                max_steps: usize::MAX,
                max_rule_depth: depth,
            };
            let mut a = B {
                enabled: false,
                value: seed,
                pratt: None,
                tokens: true,
                handles: vec![],
            };
            let mut b = B {
                enabled: true,
                value: seed,
                pratt: None,
                tokens: true,
                handles: vec![],
            };
            let x = grameo::engine::build_capture_program(&g, "program", input, opts, &mut a);
            let y = grameo::engine::build_capture_program(&g, "program", input, opts, &mut b);
            match (x, y) {
                (Ok(x), Ok(y)) => {
                    assert_eq!(shape(&x), shape(&y));
                    for (tree, handles) in [(x, a.handles), (y, b.handles)] {
                        let root = tree.into_typed_pairs::<Identity>().next().unwrap();
                        for h in handles {
                            assert!(
                                root.resolve_capture(h).is_some(),
                                "earlier seed token invalidated"
                            );
                        }
                    }
                }
                (Err(x), Err(y)) => assert_eq!(x, y),
                _ => panic!("seed acceptance mismatch"),
            };
            cases += 1;
        }
    }
    assert_eq!(cases, 3808);
}

#[test]
fn shared_prefix_defaults_validation_stack_and_limits() {
    let source = "// @grameo shared_prefix arg branch=gen prefix=value\narg = { first | gen | value }\nfirst = { \"!\" }\ngen = { value ~ \"?\" }\nvalue = { \"x\" }\n";
    let active = grameo::compile_with_options(
        source,
        CompileOptions {
            optimizations: true,
            ..Default::default()
        },
    )
    .unwrap();
    let raw = grameo::compile(source).unwrap();
    assert!((0..raw.syntax().expressions.len()).all(|id| raw.shared_prefix_choice(id).is_none()));

    assert_eq!(
        (0..active.syntax().expressions.len())
            .filter(|&id| active.shared_prefix_choice(id).is_some())
            .count(),
        1
    );
    for replacement in [
        "shared_prefix missing branch=gen prefix=value",
        "shared_prefix arg branch=missing prefix=value",
        "shared_prefix arg branch=gen prefix=missing",
        "shared_prefix arg branch=gen",
        "shared_prefix arg branch=first prefix=value",
    ] {
        let bad = source.replace("shared_prefix arg branch=gen prefix=value", replacement);
        assert!(grameo::compile(&bad).is_ok());
        assert!(
            grameo::compile_with_options(
                &bad,
                CompileOptions {
                    optimizations: true,
                    ..Default::default()
                }
            )
            .is_err(),
            "{bad}"
        );
    }
    for bad in [
        source.replace("gen = {", "gen = @{"),
        source.replace("gen = { value ~", "gen = { first ~"),
        format!("// @grameo shared_prefix arg branch=gen prefix=value\n{source}"),
    ] {
        assert!(grameo::compile(&bad).is_ok());
        assert!(
            grameo::compile_with_options(
                &bad,
                CompileOptions {
                    optimizations: true,
                    ..Default::default()
                }
            )
            .is_err(),
            "{bad}"
        );
    }
    struct B {
        enabled: bool,
    }
    impl grameo::builder::Builder for B {
        fn supports_shared_prefix(&self, _: usize) -> bool {
            self.enabled
        }
        fn observes_rule(&self, _: usize) -> bool {
            false
        }
        fn checkpoint(&self) -> usize {
            0
        }
        fn rollback(&mut self, _: usize) {}
        fn begin(&mut self, _: usize, _: usize) {}
        fn finish(&mut self, _: usize, _: grameo::Span, _: &str) {}
    }
    fn shape(tree: &grameo::ParseTree<'_, '_>) -> Vec<(String, usize, usize, usize)> {
        fn visit(
            p: grameo::Pair<'_, '_, '_>,
            d: usize,
            v: &mut Vec<(String, usize, usize, usize)>,
        ) {
            let s = p.as_span();
            v.push((format!("{:?}", p.as_rule()), s.start, s.end, d));
            for c in p.into_inner() {
                visit(c, d + 1, v)
            }
        }
        let mut out = vec![];
        for p in tree.pairs() {
            visit(p, 0, &mut out)
        }
        out
    }
    for (prefix, program, prior, inputs) in [
        (
            "PUSH(\"x\")",
            "arg ~ POP ~ EOI",
            "\"!\"",
            vec!["xx", "xz", "xy"],
        ),
        (
            "PEEK",
            "PUSH(\"a\") ~ arg ~ POP ~ EOI",
            "\"!\"",
            vec!["aaa", "aaz", "aay"],
        ),
        (
            "POP",
            "PUSH(\"a\") ~ arg ~ EOI",
            "\"!\"",
            vec!["aa", "aaz", "aay"],
        ),
        (
            "PEEK",
            "PUSH(\"a\") ~ arg ~ POP ~ EOI",
            "POP ~ \"!\"",
            vec!["aaa", "aa!a", "aaz", "aay"],
        ),
    ] {
        let grammar = format!(
            "// @grameo shared_prefix arg branch=gen prefix=value\nprogram = {{ {program} }}\narg = {{ first | gen | value }}\nfirst = {{ {prior} }}\ngen = {{ value ~ \"z\" }}\nvalue = {{ {prefix} }}\n"
        );
        let g = grameo::compile_with_options(
            &grammar,
            CompileOptions {
                optimizations: true,
                ..Default::default()
            },
        )
        .unwrap();
        for input in inputs {
            for (steps, depth) in [
                (usize::MAX, 4096),
                (usize::MAX, 1),
                (usize::MAX, 2),
                (usize::MAX, 3),
                (0, 4096),
                (10, 4096),
                (50, 4096),
                (100, 4096),
            ] {
                let options = grameo::MatchOptions {
                    max_steps: steps,
                    max_rule_depth: depth,
                };
                let mut control = B { enabled: false };
                let a = grameo::engine::build_capture_program(
                    &g,
                    "program",
                    input,
                    options,
                    &mut control,
                );
                let mut candidate = B { enabled: true };
                let b = grameo::engine::build_capture_program(
                    &g,
                    "program",
                    input,
                    options,
                    &mut candidate,
                );
                match (a, b) {
                    (Ok(a), Ok(b)) => assert_eq!(shape(&a), shape(&b), "{grammar} {input}"),
                    (Err(a), Err(b)) => assert_eq!(a, b, "{grammar} {input}"),
                    _ => panic!("accept mismatch {grammar} {input}"),
                }
            }
        }
    }
    println!("directive positive/default-ignore/default-generation/8 rejection checks passed");
}
