use grameo::{builder::{Builder,BuildError},engine,source_units::WithBuilder,CompileOptions,MatchOptions,Span};
struct Ast {rule:usize,values:Vec<String>,error:bool}
impl Builder for Ast {
    fn checkpoint(&self)->usize {self.values.len()}
    fn rollback(&mut self,mark:usize) {self.values.truncate(mark);}
    fn observes_rule(&self,rule:usize)->bool {rule==self.rule}
    fn begin(&mut self,_:usize,_:usize) {self.values.push("begin1".into());self.values.push("begin2".into());}
    fn finish(&mut self,_:usize,span:Span,source:&str) {self.values.push(source[span.start..span.end].into());}
    fn try_finish(&mut self,rule:usize,span:Span,source:&str)->Result<(),BuildError> {
        self.finish(rule,span,source);
        if self.error {Err(BuildError {span,message:"AST error".into()})} else {Ok(())}
    }
    fn supports_pratt(&self,_:usize)->bool {true}
    fn opaque_rule(&self,_:usize)->bool {true}
}
fn grammar()->grameo::CompiledGrammar {
    grameo::compile_with_options("// @grameo source_unit load target=target loader=source phase=execute repeat=once missing=error scope=caller\nroot = { &load ~ (load ~ \"!\" | load) ~ EOI } load = { \"[\" ~ target ~ \"]\" } target = { load | name } name = @{ ASCII_ALPHA+ }",CompileOptions {optimizations:true,..Default::default()}).unwrap()
}
#[test]
fn one_parse_preserves_ast_and_collects_nested_intents_with_exact_rollback() {
    let grammar=grammar();let name=grammar.rule_id("name").unwrap();
    let input="[[file]]";
    let mut control=Ast {rule:name,values:Vec::new(),error:false};
    engine::build_capture_program(&grammar,"root",input,MatchOptions::default(),&mut control).unwrap();
    let mut combined=WithBuilder::new(Ast {rule:name,values:Vec::new(),error:false},grammar.source_unit_bindings());
    engine::build_capture_program(&grammar,"root",input,MatchOptions::default(),&mut combined).unwrap();
    assert_eq!(combined.ast.values,control.values);
    assert_eq!(combined.collector.owned_requests().iter().map(|r|r.target.as_str()).collect::<Vec<_>>(),["file","[file]"]);
    let before=combined.ast.values.clone();let requests=combined.collector.requests().to_vec();
    assert!(engine::build_capture_program(&grammar,"root","[[file]]?",MatchOptions::default(),&mut combined).is_err());
    assert_eq!(combined.ast.values,before);assert_eq!(combined.collector.requests(),requests);
}
#[test]
fn work_depth_and_checked_ast_errors_restore_both_journals() {
    let grammar=grammar();let name=grammar.rule_id("name").unwrap();
    for steps in 0..100 {for depth in [1,2,3,64] {
        let mut combined=WithBuilder::new(Ast {rule:name,values:Vec::new(),error:false},grammar.source_unit_bindings());
        let result=engine::build_capture_program(&grammar,"root","[[file]]",MatchOptions {max_steps:steps,max_rule_depth:depth},&mut combined);
        if result.is_err() {
            assert!(combined.ast.values.is_empty());assert!(combined.collector.requests().is_empty());
            engine::build_capture_program(&grammar,"root","[[file]]",MatchOptions::default(),&mut combined).unwrap();
        }
        assert_eq!(combined.collector.requests().len(),2);
    }}
    let mut combined=WithBuilder::new(Ast {rule:name,values:Vec::new(),error:true},grammar.source_unit_bindings());
    assert!(engine::build_capture_program(&grammar,"root","[[file]]",MatchOptions::default(),&mut combined).is_err());
    assert!(combined.ast.values.is_empty());assert!(combined.collector.owned_requests().is_empty());
}
#[test]
fn shortcut_permissions_require_explicit_admission_only_when_intents_are_bound() {
    let grammar=grammar();let name=grammar.rule_id("name").unwrap();
    let mut combined=WithBuilder::new(Ast {rule:name,values:Vec::new(),error:false},grammar.source_unit_bindings());
    assert!(!combined.supports_pratt(name));assert!(!combined.opaque_rule(name));
    combined.allow_pratt_rules([name]);combined.allow_opaque_rules([name]);
    assert!(combined.supports_pratt(name));assert!(combined.opaque_rule(name));
    let empty=WithBuilder::new(Ast {rule:name,values:Vec::new(),error:false},[]);
    assert!(empty.supports_pratt(name));assert!(empty.opaque_rule(name));
}

#[test]
fn collector_observation_does_not_invent_unobserved_ast_callbacks() {
    struct CapturePolicy;
    impl Builder for CapturePolicy {
        fn checkpoint(&self)->usize {0}
        fn rollback(&mut self,_:usize) {}
        fn observes_rule(&self,_:usize)->bool {false}
        fn observes_capture(&self,_:usize)->bool {true}
        fn begin(&mut self,_:usize,_:usize) {panic!("unobserved AST begin");}
        fn finish(&mut self,_:usize,_:Span,_:&str) {panic!("unobserved AST finish");}
    }
    let grammar=grammar();let mut combined=WithBuilder::new(CapturePolicy,grammar.source_unit_bindings());
    engine::build_capture_program(&grammar,"root","[file]",MatchOptions::default(),&mut combined).unwrap();
    assert_eq!(combined.collector.owned_requests()[0].target,"file");
}

#[test]
fn lexical_admission_preserves_observed_children_even_under_atomic_rules() {
    use grameo::source_units::opaque_rules_preserving_requests;
    let source="// @grameo source_unit load target=target loader=source phase=execute repeat=once missing=error scope=caller\nroot={load~EOI} load=@{\"[\"~target~\"]\"} target=@{name} name=@{ASCII_ALPHA+}";
    let grammar=grameo::compile_with_options(source,CompileOptions {optimizations:true,..Default::default()}).unwrap();
    let id=|name|grammar.rule_id(name).unwrap();let bindings=grammar.source_unit_bindings().collect::<Vec<_>>();
    let admitted=opaque_rules_preserving_requests(&grammar,&[id("root")],&bindings);
    assert!(!admitted.contains(&id("load")),"load's atomic scan would omit its observed target");
    assert!(admitted.contains(&id("target")),"own target callback is preserved");
    assert!(admitted.contains(&id("name")));
    let mut combined=WithBuilder::new(Ast {rule:id("name"),values:Vec::new(),error:false},bindings);
    // AST owner itself admits this opaque rule; name callbacks are irrelevant
    // to that permission. Include target callbacks must still be preserved.
    combined.allow_opaque_rules(admitted);
    engine::build_capture_program(&grammar,"root","[file]",MatchOptions::default(),&mut combined).unwrap();
    assert_eq!(combined.collector.owned_requests()[0].target,"file");
}

#[test]
fn source_pratt_admission_is_pure_and_still_requires_owner_permission() {
    let grammar=grammar();let name=grammar.rule_id("name").unwrap();
    let mut combined=WithBuilder::new(Ast {rule:name,values:Vec::new(),error:false},grammar.source_unit_bindings());
    combined.set_pratt_admission(|_,_,offset,source|source.as_bytes().get(offset)==Some(&b'1'));
    let mark=combined.checkpoint();
    assert!(combined.supports_pratt_at(name,0,"1 + 2"));
    assert!(!combined.supports_pratt_at(name,0,"include file"));
    assert!(!combined.supports_pratt(name));
    assert_eq!(combined.checkpoint(),mark);
}

#[test]
fn deferred_context_rolls_back_with_scopes_without_dropping_loads() {
    use grameo::source_units::PreparationContext::{Immediate,Deferred};
    let source="// @grameo source_unit load target=name loader=source phase=execute repeat=always missing=error scope=caller\nroot={&deferred~(deferred~\"!\"|deferred)~load~EOI} deferred={\"f\"~load~deferred?} load={\"[\"~name~\"]\"} name=@{ASCII_ALPHA+}";
    let grammar=grameo::compile_with_options(source,CompileOptions {optimizations:true,..Default::default()}).unwrap();
    let name=grammar.rule_id("name").unwrap();let deferred=grammar.rule_id("deferred").unwrap();
    for steps in [0,5,20,60,120,usize::MAX] {
        let mut combined=WithBuilder::new(Ast {rule:name,values:Vec::new(),error:false},grammar.source_unit_bindings());
        combined.collector.defer_within([deferred]);
        let result=engine::build_capture_program(&grammar,"root","f[a]f[b][c]",MatchOptions {max_steps:steps,..Default::default()},&mut combined);
        if result.is_err() {
            assert!(combined.collector.owned_requests().is_empty());assert!(combined.ast.values.is_empty());
            engine::build_capture_program(&grammar,"root","f[a]f[b][c]",MatchOptions::default(),&mut combined).unwrap();
        }
        assert_eq!(combined.collector.owned_requests().iter().map(|request|request.preparation).collect::<Vec<_>>(),[Deferred,Deferred,Immediate]);
        assert_eq!(combined.collector.owned_requests().iter().map(|request|request.target.as_str()).collect::<Vec<_>>(),["a","b","c"]);
    }
    let mut combined=WithBuilder::new(Ast {rule:name,values:Vec::new(),error:true},grammar.source_unit_bindings());
    combined.collector.defer_within([deferred]);
    assert!(engine::build_capture_program(&grammar,"root","f[a]f[b][c]",MatchOptions::default(),&mut combined).is_err());
    assert!(combined.collector.owned_requests().is_empty());
    combined.ast.error=false;
    engine::build_capture_program(&grammar,"root","f[a]f[b][c]",MatchOptions::default(),&mut combined).unwrap();
    assert_eq!(combined.collector.owned_requests().last().unwrap().preparation,Immediate);
}
