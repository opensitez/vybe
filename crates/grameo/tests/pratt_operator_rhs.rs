use grameo::{builder::Builder,program::Program,MatchOptions,Span};
#[derive(Default)]struct Hooks{target:usize,events:Vec<(usize,usize,usize,bool)>}
impl Builder for Hooks{
 fn supports_pratt(&self,rule:usize)->bool{rule==self.target}
 fn checkpoint(&self)->usize{self.events.len()}
 fn rollback(&mut self,checkpoint:usize){self.events.truncate(checkpoint)}
 fn begin(&mut self,rule:usize,start:usize){self.events.push((rule,start,start,false))}
 fn finish(&mut self,rule:usize,span:Span,_:&str){self.events.push((rule,span.start,span.end,true))}
}
const SOURCE:&str=r#"
// @grameo pratt expr atom=atom trailing=consume
// @grameo operator expr plus precedence=30 fixity=infix_left
// @grameo operator expr eq precedence=20 fixity=infix_left rhs_atom=eq_rhs
// @grameo operator expr and precedence=10 fixity=infix_left rhs_atom=and_rhs
entry={SOI~expr~EOI}
expr={equality~(and~and_rhs)*}
equality={sum~(eq~eq_rhs)*}
sum={atom~(plus~atom)*}
eq_rhs={sum~("="~expr)?}
and_rhs={equality~("="~expr)?}
atom={"a"|"b"|"c"|"d"}
plus=@{"+"}
eq=@{"=="}
and=@{"&&"}
empty={""}
"#;
fn compile(source:&str)->grameo::CompiledGrammar{grameo::compile_with_options(source,grameo::CompileOptions{optimizations:true,..Default::default()}).unwrap()}
#[test]fn each_operator_uses_its_registered_operand(){
 let g=compile(SOURCE);let target=g.rule_id("expr").unwrap();
 assert_eq!(g.source_pratt_operator_rhs_atom(target,g.rule_id("eq").unwrap()),g.rule_id("eq_rhs"));
 assert_eq!(g.source_pratt_operator_rhs_atom(target,g.rule_id("and").unwrap()),g.rule_id("and_rhs"));
 assert_eq!(g.source_pratt_operator_rhs_atom(target,g.rule_id("plus").unwrap()),None);
 let peg=grameo::compile(SOURCE).unwrap();assert!(peg.source_pratt(target).is_none());
 for input in ["a+b+c","a==b+c=d","a&&b=c==d","a==b=c&&d","a+b==c&&d","a&&b==c"]{
  let mut h=Hooks{target,..Default::default()};let actual=grameo::engine::build_capture_program(&g,"entry",input,MatchOptions::default(),&mut h).unwrap();
  let reference=grameo::engine::parse_program(&peg,"entry",input,MatchOptions::default()).unwrap();assert_eq!(actual.consumed(),reference.consumed(),"{input}");
  assert_eq!(h.events.iter().filter(|e|e.0==g.rule_id("and_rhs").unwrap()&&e.3).count(),input.matches("&&").count(),"{input}");
 }
}
#[test]fn per_operator_operands_are_opt_in_and_validated(){
 for (old,new) in [("rhs_atom=eq_rhs","rhs_atom=missing"),("rhs_atom=eq_rhs","rhs_atom=empty"),("rhs_atom=eq_rhs","rhs_atom=expr"),("rhs_atom=eq_rhs","rhs_atom="),("fixity=infix_left rhs_atom=eq_rhs","fixity=prefix rhs_atom=eq_rhs"),("rhs_atom=eq_rhs","unknown=eq_rhs")]{
  let source=SOURCE.replace(old,new);assert!(grameo::compile(&source).is_ok());assert!(grameo::compile_with_options(&source,grameo::CompileOptions{optimizations:true,..Default::default()}).is_err(),"{new}");
 }
}
#[test]fn incomplete_operands_restore_all_builder_events(){
 let g=compile(SOURCE);let target=g.rule_id("expr").unwrap();
 for input in ["a&&b=","a==?","a&&b==?","a==b=c&&?"]{for depth in 0..20{for steps in 0..512{
  let mut h=Hooks{target,..Default::default()};assert!(grameo::engine::build_capture_program(&g,"entry",input,MatchOptions{max_rule_depth:depth,max_steps:steps},&mut h).is_err());assert!(h.events.is_empty(),"{input} depth={depth} steps={steps}");
 }}}
}
#[test]fn per_operator_operand_overrides_target_default(){
 let source=SOURCE.replace("trailing=consume","trailing=consume rhs_atom=atom rhs_operator=eq");let g=compile(&source);let target=g.rule_id("expr").unwrap();
 assert_eq!(g.source_pratt_rhs_atom(target),g.rule_id("atom"));
 assert_eq!(g.source_pratt_operator_rhs_atom(target,g.rule_id("and").unwrap()),g.rule_id("and_rhs"));
 let mut h=Hooks{target,..Default::default()};assert!(grameo::engine::build_capture_program(&g,"entry","a&&b=c==d",MatchOptions::default(),&mut h).is_ok());
 let mut h=Hooks{target,..Default::default()};assert!(grameo::engine::build_capture_program(&g,"entry","a==b+c=d",MatchOptions::default(),&mut h).is_ok());
}

#[test]fn a_prefix_declaration_cannot_set_an_existing_infix_operand(){
 let source=SOURCE.replace("// @grameo operator expr eq precedence=20", "// @grameo operator expr eq precedence=40 fixity=prefix rhs_atom=eq_rhs\n// @grameo operator expr eq precedence=20");
 let errors=grameo::compile_with_options(&source,grameo::CompileOptions{optimizations:true,..Default::default()}).unwrap_err();
 assert!(errors.iter().any(|error|error.message=="operator RHS atom requires an infix declaration"));
 assert!(grameo::compile(&source).is_ok());
}
