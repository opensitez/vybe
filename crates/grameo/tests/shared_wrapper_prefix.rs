use grameo::{CompileOptions, MatchOptions, Span, builder::Builder, program::Program};
const BODY: &str = r#"
entry = { atom ~ postfix* }
atom = { "o" | "x" }
postfix = { call | read }
call = { "." ~ name ~ suffix }
read = { "." ~ name }
name = { "x" | "{" ~ entry ~ "}" }
suffix = { "(" ~ ")" }
"#;
const ANNOTATION: &str = "// @grameo shared_wrapper_prefix postfix branch=call fallback=read\n";
#[derive(Default)]
struct Hooks { enabled: bool, observed: bool, pratt: bool, events: Vec<(usize,Span)> }
impl Builder for Hooks {
    fn supports_pratt(&self, _:usize)->bool {self.pratt}
    fn supports_shared_wrapper_prefix(&self, _:usize, _:usize)->bool {self.enabled}
    fn observes_rule(&self,_:usize)->bool {self.observed}
    fn checkpoint(&self)->usize {self.events.len()}
    fn rollback(&mut self,n:usize){self.events.truncate(n)}
    fn begin(&mut self,r:usize,o:usize){self.events.push((r,Span{start:o,end:o}))}
    fn finish(&mut self,r:usize,s:Span,_:&str){self.events.push((r,s))}
}
fn shape(tree: &grameo::ParseTree<'_, '_>)->Vec<(String,Span,usize)> {
    fn visit(pair:grameo::tree::Pair<'_, '_, '_>,depth:usize,out:&mut Vec<(String,Span,usize)>){
        out.push((pair.rule_name().to_owned(),pair.as_span(),depth));
        for child in pair.into_inner(){visit(child,depth+1,out)}
    }
    let mut out=Vec::new();for p in tree.pairs(){visit(p,0,&mut out)}out
}
#[test]
fn capture_error_and_hook_parity_with_resource_limits(){
    let grammar=grameo::compile_with_options(&format!("{ANNOTATION}{BODY}"),CompileOptions{optimizations:true,..Default::default()}).unwrap();
    let mut sources=vec!["o.x".to_owned(),"o.x()".into(),"o.x(".into(),"o.{x}".into(),"o.{x}()".into(),"o.{x}(".into(),"o.".into()];
    let mut nested="x".to_owned();for _ in 0..8{nested=format!("o.{{{nested}}}");sources.push(nested.clone());}
    let mut cases=0;
    for source in sources{for observed in [false,true]{for depth in [1,2,4,8,16,4096]{for work in [1,32,256,usize::MAX]{
        let mut a=Hooks{observed,..Default::default()};let mut b=Hooks{enabled:true,observed,..Default::default()};
        let options=MatchOptions{max_rule_depth:depth,max_steps:work,..Default::default()};
        let left=grameo::engine::build_capture_program(&grammar,"entry",&source,options,&mut a);
        let right=grameo::engine::build_capture_program(&grammar,"entry",&source,options,&mut b);
        match(left,right){(Ok(a),Ok(b))=>assert_eq!(shape(&a),shape(&b),"{source}"),(Err(a),Err(b))=>assert_eq!(a,b,"{source}"),(a,b)=>panic!("admission {source}: {a:?} {b:?}")}
        assert_eq!(a.events,b.events);cases+=1;
    }}}}assert_eq!(cases,720);
}
#[test]
fn directive_is_inert_by_default_and_rejects_non_equivalent_prefixes(){
    let source=format!("{ANNOTATION}{BODY}");let normal=grameo::compile(&source).unwrap();
    assert!((0..normal.syntax().expressions.len()).all(|id|normal.shared_wrapper_prefix(id).is_none()));
    for invalid in [source.replace("read = { \".\" ~ name }","read = { \"!\" ~ name }"),source.replace("call = {", "call = @{"),source.replace("branch=call","branch=missing"),format!("{ANNOTATION}{source}")] {
        assert!(grameo::compile(&invalid).is_ok());
        let errors=grameo::compile_with_options(&invalid,CompileOptions{optimizations:true,..Default::default()}).unwrap_err();assert_eq!(errors[0].code,"G015");
    }
}

// Immutable geometry is not a proof that later SourcePratt annotations are absent.
#[derive(Debug)]
struct CertifiedProbe<'a> {
    grammar: &'a grameo::CompiledGrammar,
    body: usize,
    body_queries: std::cell::Cell<usize>,
    changed_prefix: Option<usize>,
}
impl Program for CertifiedProbe<'_> {
    fn rule_id(&self, name: &str) -> Option<usize> { self.grammar.rule_id(name) }
    fn rule(&self, rule: usize) -> grameo::program::RuleSpec<'_> { self.grammar.rule(rule) }
    fn scope(&self, scope: usize) -> grameo::program::Scope { self.grammar.scope(scope) }
    fn instruction(&self, expression: usize) -> grameo::program::Instruction<'_> {
        if expression == self.body { self.body_queries.set(self.body_queries.get() + 1); }
        if self.changed_prefix == Some(expression) {
            grameo::program::Instruction::Literal { text: "!", insensitive: false }
        } else { self.grammar.instruction(expression) }
    }
    fn source_pratt(&self, rule: usize) -> Option<grameo::program::SourcePratt<'_>> {
        self.grammar.source_pratt(rule)
    }
    fn shared_wrapper_prefix(&self, expression: usize) -> Option<grameo::program::SharedWrapperPrefix> {
        self.grammar.shared_wrapper_prefix(expression)
    }
    fn shared_wrapper_prefix_metadata_is_current(&self) -> bool { true }
    fn match_observers_are_noop(&self) -> bool { true }
}
// A forwarding program may change instructions; it must inherit false unless
// it independently establishes the new immutable metadata contract.
#[derive(Debug)]
struct StaleProbe<'a>(CertifiedProbe<'a>);
impl Program for StaleProbe<'_> {
    fn rule_id(&self, name: &str) -> Option<usize> { self.0.rule_id(name) }
    fn rule(&self, rule: usize) -> grameo::program::RuleSpec<'_> { self.0.rule(rule) }
    fn scope(&self, scope: usize) -> grameo::program::Scope { self.0.scope(scope) }
    fn instruction(&self, expression: usize) -> grameo::program::Instruction<'_> {
        self.0.instruction(expression)
    }
    fn source_pratt(&self, rule: usize) -> Option<grameo::program::SourcePratt<'_>> {
        self.0.source_pratt(rule)
    }
    fn shared_wrapper_prefix(&self, expression: usize) -> Option<grameo::program::SharedWrapperPrefix> {
        self.0.shared_wrapper_prefix(expression)
    }
    fn match_observers_are_noop(&self) -> bool { true }
}
const PRATT_BODY: &str = "entry={postfix~EOI}\npostfix={call|read}\ncall={\".\"~name~\"()\"}\nread={\".\"~name}\nname={word}\nword=@{ASCII_ALPHA+}\n";
#[test]
fn immutable_wrapper_geometry_preserves_later_pratt_exclusions() {
    for target in ["call", "read", "name"] {
        // apply() certifies the wrapper first and registers Pratt bindings later.
        let source = format!("// @grameo shared_wrapper_prefix postfix branch=call fallback=read\n// @grameo pratt {target} atom=word trailing=consume\n{PRATT_BODY}");
        let grammar = grameo::compile_with_options(&source, CompileOptions { optimizations: true, ..Default::default() }).unwrap();
        assert!(grammar.source_pratt(grammar.rule_id(target).unwrap()).is_some());
        for input in [".x", ".x()", ".x(", ".xyz"] {
            for active_pratt in [false, true] {
                let program = CertifiedProbe { grammar: &grammar,
                    body: grammar.rule(grammar.rule_id("name").unwrap()).expression,
                    body_queries: std::cell::Cell::new(0), changed_prefix: None };
                let mut reference = Hooks { pratt: active_pratt, ..Default::default() };
                let mut candidate = Hooks { enabled: true, pratt: active_pratt, ..Default::default() };
                let options = MatchOptions { max_steps: usize::MAX, ..Default::default() };
                let a = grameo::engine::build_capture_program(&program, "entry", input, options, &mut reference).map(|t| (t.consumed(), shape(&t)));
                let reference_queries = program.body_queries.replace(0);
                let b = grameo::engine::build_capture_program(&program, "entry", input, options, &mut candidate).map(|t| (t.consumed(), shape(&t)));
                assert_eq!(a, b, "{target}: {input}, Pratt={active_pratt}");
                assert_eq!(reference.events, candidate.events);
                assert_eq!(reference_queries, program.body_queries.get(), "Pratt override must retain ordinary prefix replay: {target}");
            }
        }
    }
}
#[test]
fn forwarding_mutated_wrapper_certificate_revalidates_by_default() {
    let source = format!("// @grameo shared_wrapper_prefix postfix branch=call fallback=read\n{PRATT_BODY}");
    let grammar = grameo::compile_with_options(&source, CompileOptions { optimizations: true, ..Default::default() }).unwrap();
    let certificate = (0..grammar.syntax().expressions.len()).find_map(|id| grammar.shared_wrapper_prefix(id)).unwrap();
    let program = StaleProbe(CertifiedProbe { grammar: &grammar,
        body: grammar.rule(grammar.rule_id("name").unwrap()).expression,
        body_queries: std::cell::Cell::new(0), changed_prefix: Some(certificate.prefix) });
    assert!(!program.shared_wrapper_prefix_metadata_is_current());
    for input in [".x", ".x()", ".x(", "!()"] {
        let mut reference = Hooks::default();
        let mut candidate = Hooks { enabled: true, ..Default::default() };
        let options = MatchOptions { max_steps: usize::MAX, ..Default::default() };
        let a = grameo::engine::build_capture_program(&program, "entry", input, options, &mut reference).map(|t| (t.consumed(), shape(&t)));
        let b = grameo::engine::build_capture_program(&program, "entry", input, options, &mut candidate).map(|t| (t.consumed(), shape(&t)));
        assert_eq!(a, b, "stale certificate: {input}");
        assert_eq!(reference.events, candidate.events);
    }
}
