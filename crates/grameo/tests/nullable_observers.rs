//! Forwarding certificate metadata must not silently disable custom tracing.
use grameo::{
    MatchOptions,
    builder::Builder,
    program::{FragmentExpected, Instruction, NullableFailurePrefix, Program, RuleSpec, Scope},
};
use std::cell::{Cell, RefCell};
struct Profile {
    grammar: grameo::CompiledGrammar,
    metadata: bool,
    noop: bool,
    queries: Cell<usize>,
    events: RefCell<Vec<String>>,
}
impl std::fmt::Debug for Profile {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str("Profile")
    }
}
impl Program for Profile {
    fn rule_id(&self, name: &str) -> Option<usize> {
        self.grammar.rule_id(name)
    }
    fn rule(&self, id: usize) -> RuleSpec<'_> {
        self.grammar.rule(id)
    }
    fn instruction(&self, id: usize) -> Instruction<'_> {
        self.grammar.instruction(id)
    }
    fn scope(&self, id: usize) -> Scope {
        self.grammar.scope(id)
    }
    fn match_observers_are_noop(&self) -> bool {
        self.noop
    }
    fn nullable_failure_prefix(
        &self,
        _: usize,
        _: u8,
        first: Option<u8>,
    ) -> Option<&NullableFailurePrefix<'_>> {
        self.queries.set(self.queries.get() + 1);
        static CERT: NullableFailurePrefix<'static> = NullableFailurePrefix {
            steps: 3,
            depth: 1,
            expected: &[FragmentExpected::Literal("x")],
            unobserved_rules: &[0],
        };
        (self.metadata && first == Some(b'z')).then_some(&CERT)
    }
    fn observe_expression(&self, id: usize, offset: usize) {
        if !self.noop {
            self.events.borrow_mut().push(format!("E{id}@{offset}"));
        }
    }
    fn observe_rule_attempt(&self, id: usize, offset: usize, flags: u8) {
        if !self.noop {
            self.events
                .borrow_mut()
                .push(format!("R{id}@{offset}:{flags}"));
        }
    }
}
#[derive(Debug)]
struct Hooks;
impl Builder for Hooks {
    fn supports_unobserved_failure_prefixes(&self) -> bool {
        true
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
fn profile(metadata: bool, noop: bool) -> Profile {
    Profile {
        grammar: grameo::compile("root = { \"x\" }").unwrap(),
        metadata,
        noop,
        queries: Cell::new(0),
        events: RefCell::new(Vec::new()),
    }
}
#[test]
fn forwarding_metadata_preserves_observers_errors_and_resource_boundaries() {
    for input in ["z", "x", "π", ""] {
        for depth in 0..8 {
            for work in 0..32 {
                let a = profile(false, false);
                let b = profile(true, false);
                let opts = MatchOptions {
                    max_steps: work,
                    max_rule_depth: depth,
                };
                let x = grameo::engine::build_capture_program(&a, "root", input, opts, &mut Hooks);
                let y = grameo::engine::build_capture_program(&b, "root", input, opts, &mut Hooks);
                assert_eq!(
                    format!("{x:?}"),
                    format!("{y:?}"),
                    "{input:?} depth={depth} work={work}"
                );
                assert_eq!(
                    *a.events.borrow(),
                    *b.events.borrow(),
                    "{input:?} depth={depth} work={work}"
                );
                assert_eq!(
                    b.queries.get(),
                    0,
                    "observer admission precedes metadata lookup"
                );
            }
        }
    }
}
#[test]
fn explicit_noop_permission_keeps_the_certificate_route_active() {
    let a = profile(false, true);
    let b = profile(true, true);
    let opts = MatchOptions::default();
    let x = grameo::engine::build_capture_program(&a, "root", "z", opts, &mut Hooks).unwrap_err();
    let y = grameo::engine::build_capture_program(&b, "root", "z", opts, &mut Hooks).unwrap_err();
    assert_eq!(x, y);
    assert!(b.queries.get() > 0);
}
#[test]
fn custom_program_observer_permission_defaults_to_false() {
    #[derive(Debug)]
    struct Plain;
    impl Program for Plain {
        fn rule_id(&self, _: &str) -> Option<usize> {
            None
        }
        fn rule(&self, _: usize) -> RuleSpec<'_> {
            unreachable!()
        }
        fn instruction(&self, _: usize) -> Instruction<'_> {
            unreachable!()
        }
    }
    assert!(!Plain.match_observers_are_noop());
}

#[derive(Default)]
struct FinishOnly {
    permit: bool,
    reject_begin: bool,
    events: Vec<(usize, grameo::Span, bool)>,
}
impl Builder for FinishOnly {
    fn supports_unobserved_failure_prefixes(&self) -> bool { true }
    fn observes_rule(&self, _:usize) -> bool { true }
    fn observes_capture(&self, _:usize) -> bool { true }
    fn failure_prefix_hooks_are_inert(&self, _:usize) -> bool { self.permit && !self.reject_begin }
    fn checkpoint(&self)->usize { self.events.len() }
    fn rollback(&mut self,n:usize) { self.events.truncate(n); }
    fn begin(&mut self,_:usize,_:usize) {}
    fn finish(&mut self,_:usize,_:grameo::Span,_:&str) {}
    fn try_begin(&mut self,_:usize,n:usize)->Result<(),grameo::builder::BuildError> {
        if self.reject_begin { Err(grameo::builder::BuildError {span:grameo::Span{start:n,end:n},message:"checked begin".into()}) } else { Ok(()) }
    }
    fn try_finish_capture(&mut self,r:usize,s:grameo::Span,_:&str,c:Option<grameo::tree::CaptureHandle>)->Result<(),grameo::builder::BuildError> {
        self.events.push((r,s,c.is_some())); Ok(())
    }
}
#[test]
fn inert_finish_only_permission_preserves_receipts_errors_and_observer_fallback() {
    for input in ["z","x","π",""] { for depth in 0..8 { for work in 0..32 {
        for permit in [false,true] { for noop in [false,true] { for reject_begin in [false,true] {
            let a=profile(false,noop);let b=profile(true,noop);
            let mut x=FinishOnly{permit,reject_begin,..Default::default()};let mut y=FinishOnly{permit,reject_begin,..Default::default()};
            let opts=MatchOptions{max_steps:work,max_rule_depth:depth};
            let old=grameo::engine::build_capture_program(&a,"root",input,opts,&mut &mut x);
            let new=grameo::engine::build_capture_program(&b,"root",input,opts,&mut &mut y);
            assert_eq!(format!("{old:?}"),format!("{new:?}"),"{input:?} depth={depth} work={work} permit={permit} noop={noop} reject={reject_begin}");
            assert_eq!(x.events,y.events);assert_eq!(*a.events.borrow(),*b.events.borrow());
            if !noop { assert_eq!(b.queries.get(),0); }
        }}} 
    }}}
}
#[test]
fn ordinary_builders_do_not_grant_inert_prefix_permission() {
    assert!(!Hooks.failure_prefix_hooks_are_inert(0));
}
