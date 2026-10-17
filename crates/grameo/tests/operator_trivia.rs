use grameo::{builder::Builder,program::Program,MatchOptions,Span};
#[derive(Default)]struct Hooks{target:usize,events:Vec<(usize,usize,usize,bool)>}
impl Builder for Hooks {
 fn supports_pratt(&self,r:usize)->bool{r==self.target}
 fn checkpoint(&self)->usize{self.events.len()}
 fn rollback(&mut self,n:usize){self.events.truncate(n);}
 fn begin(&mut self,r:usize,s:usize){self.events.push((r,s,s,false));}
 fn finish(&mut self,r:usize,s:Span,_:&str){self.events.push((r,s.start,s.end,true));}
}
fn compile(source:&str)->grameo::CompiledGrammar{grameo::compile_with_options(source,grameo::CompileOptions{optimizations:true,..Default::default()}).unwrap()}
#[test]fn defaults_ignore_new_and_malformed_properties(){
 let base="expr={atom~(plus~nls~atom)*}\natom={\"a\"}\nplus=@{\"+\"}\nnls=_{\"\\n\"*}";
 for declaration in ["operator_trivia=nls","operator_trivia=missing","wrong=property","operator_trivia="] {
  let source=format!("// @grameo pratt expr atom=atom trailing=consume {declaration}\n// @grameo operator expr plus precedence=10 fixity=infix_left\n{base}");
  let defaults=grameo::compile(&source).unwrap();assert!(defaults.source_pratt(defaults.rule_id("expr").unwrap()).is_none());
  if declaration=="operator_trivia=nls"{let opted=compile(&source);assert_eq!(opted.source_pratt_operator_trivia(opted.rule_id("expr").unwrap()),opted.rule_id("nls"));}
  else {assert!(grameo::compile_with_options(&source,grameo::CompileOptions{optimizations:true,..Default::default()}).is_err());}
 }
}
#[test]fn separators_run_hooks_and_preserve_stack_rollback_and_limits(){
 for separator in ["nls=_{\"\\n\"*}","nls=_{PUSH(\"x\")~POP}"] {
  let source=format!("// @grameo pratt expr atom=atom trailing=consume operator_trivia=nls\n// @grameo operator expr plus precedence=10 fixity=infix_left\nentry={{SOI~expr~EOI}}\nexpr={{atom~(plus~nls~atom)*}}\natom={{\"a\"}}\nplus=@{{\"+\"}}\n{separator}");
  let grammar=compile(&source);let target=grammar.rule_id("expr").unwrap();let nls=grammar.rule_id("nls").unwrap();
  let valid=if separator.contains("PUSH"){"a+xxa+xxa"}else{"a+\n\na+\na"};
  let mut hooks=Hooks{target,..Default::default()};let tree=grameo::engine::build_capture_program(&grammar,"entry",valid,MatchOptions::default(),&mut hooks).unwrap();assert_eq!(tree.consumed(),valid.len());assert_eq!(hooks.events.iter().filter(|e|e.0==nls&&e.3).count(),2);
  let invalid=if separator.contains("PUSH"){"a+xya"}else{"a+\n?"};
  for depth in 0..16 {for work in 0..256 {
   let mut hooks=Hooks{target,..Default::default()};let mark=hooks.checkpoint();
   let result=grameo::engine::build_capture_program(&grammar,"entry",invalid,MatchOptions{max_rule_depth:depth,max_steps:work},&mut hooks);
   assert!(result.is_err());assert_eq!(hooks.checkpoint(),mark,"depth={depth} work={work}");
  }}
 }
}
