use grameo::{builder::Builder, program::Program, MatchOptions, Span};
#[derive(Default)]
struct Hooks { target:usize, events:Vec<(usize,usize,usize,bool)> }
impl Builder for Hooks {
 fn supports_pratt(&self,r:usize)->bool {r==self.target}
 fn checkpoint(&self)->usize {self.events.len()}
 fn rollback(&mut self,n:usize) {self.events.truncate(n);}
 fn begin(&mut self,r:usize,s:usize) {self.events.push((r,s,s,false));}
 fn finish(&mut self,r:usize,s:Span,_:&str) {self.events.push((r,s.start,s.end,true));}
}
fn source(option:&str)->String {format!(r#"
// @grameo pratt expr atom=atom trailing=consume {option}
// @grameo operator expr eq precedence=10 fixity=infix_left
entry={{SOI~expr~EOI}}
expr={{atom~(eq~rhs)*}}
rhs={{atom~("="~atom)?}}
atom={{"a"|"b"|"c"}}
eq=@{{"=="}}
empty={{""}}
"#)}
fn compile(src:&str)->grameo::CompiledGrammar {grameo::compile_with_options(src,grameo::CompileOptions{optimizations:true,..Default::default()}).unwrap()}
#[test] fn explicit_rhs_preserves_first_operand_and_assignments() {
 let g=compile(&source("rhs_atom=rhs")); let target=g.rule_id("expr").unwrap(); let rhs=g.rule_id("rhs").unwrap();
 assert_eq!(g.source_pratt_rhs_atom(target),Some(rhs));
 for input in ["a","a==b","a==b=c","a==b=c==a"] {
  let mut h=Hooks{target,..Default::default()};
  let tree=grameo::engine::build_capture_program(&g,"entry",input,MatchOptions::default(),&mut h).unwrap();
  assert_eq!(tree.consumed(),input.len());
  assert_eq!(h.events.iter().filter(|e|e.0==rhs&&e.3).count(),input.matches("==").count());
 }
 let mut h=Hooks{target,..Default::default()};
 assert!(grameo::engine::build_capture_program(&g,"entry","a=b",MatchOptions::default(),&mut h).is_err());
 assert!(h.events.is_empty());
}
#[test] fn rhs_options_are_opt_in_and_validated() {
 for opt in ["rhs_atom=rhs","rhs_atom=missing","rhs_atom=empty","rhs_atom=expr","rhs_atom=","rhs_atom=rhs rhs_atom=atom"] {
  let src=source(opt);let g=grameo::compile(&src).unwrap();assert!(g.source_pratt(g.rule_id("expr").unwrap()).is_none());
  let result=grameo::compile_with_options(&src,grameo::CompileOptions{optimizations:true,..Default::default()});
  assert_eq!(result.is_ok(),opt=="rhs_atom=rhs", "{opt}");
 }
}
#[test] fn failed_rhs_restores_input_stack_captures_and_builder() {
 let src=source("rhs_atom=rhs");let g=compile(&src);let target=g.rule_id("expr").unwrap();
 for input in ["a==?","a==b=", "a==b==?"] {for depth in 0..16 {for work in 0..256 {
  let mut h=Hooks{target,..Default::default()};
  assert!(grameo::engine::build_capture_program(&g,"entry",input,MatchOptions{max_rule_depth:depth,max_steps:work},&mut h).is_err());
  assert!(h.events.is_empty(),"{input} depth={depth} work={work}");
 }}}
}
#[test] fn rhs_combines_with_prefixes_and_operator_trivia() {
 let src=source("rhs_atom=rhs operator_trivia=nls")
  .replace("expr={atom~(eq~rhs)*}","expr={neg*~atom~(eq~nls~neg*~rhs)*}")
  .replace("// @grameo operator expr eq", "// @grameo operator expr neg precedence=20 fixity=prefix\n// @grameo operator expr eq")
  + "\nneg=@{\"!\"}\nnls=_{\"\\n\"*}\n";
 let g=compile(&src);let target=g.rule_id("expr").unwrap();let rhs=g.rule_id("rhs").unwrap();
 let mut h=Hooks{target,..Default::default()};
 let input="!a==\n!b=c";
 let tree=grameo::engine::build_capture_program(&g,"entry",input,MatchOptions::default(),&mut h).unwrap();
 assert_eq!(tree.consumed(),input.len());
 assert_eq!(h.events.iter().filter(|e|e.0==rhs&&e.3).count(),1);
 let src=src.replace("rhs_atom=rhs operator_trivia=nls","operator_trivia=nls rhs_atom=rhs");
 assert!(grameo::compile_with_options(&src,grameo::CompileOptions{optimizations:true,..Default::default()}).is_ok());
}

#[test] fn explicit_rhs_operator_preserves_other_operands() {
 let src="// @grameo pratt expr atom=atom trailing=consume rhs_atom=rhs rhs_operator=eq\n// @grameo operator expr plus precedence=20 fixity=infix_left\n// @grameo operator expr eq precedence=10 fixity=infix_left\nentry={SOI~expr~EOI}\nexpr={sum~(eq~rhs)*}\nsum={atom~(plus~atom)*}\nrhs={sum~(\"=\"~expr)?}\natom={\"a\"|\"b\"|\"c\"|\"d\"}\nplus=@{\"+\"}\neq=@{\"==\"}";
 let g=compile(src);let target=g.rule_id("expr").unwrap();let rhs=g.rule_id("rhs").unwrap();
 assert_eq!(g.source_pratt_rhs_operator(target),g.rule_id("eq"));
 for input in ["a+b+c","a==b+c=d","a+b==c=d"] {
  let mut h=Hooks{target,..Default::default()};let tree=grameo::engine::build_capture_program(&g,"entry",input,MatchOptions::default(),&mut h).unwrap();assert_eq!(tree.consumed(),input.len());
  assert_eq!(h.events.iter().filter(|e|e.0==rhs&&e.3).count(),input.matches("==").count());
 }
 for modified in [src.replace("rhs_operator=eq","rhs_operator=missing"),src.replace("rhs_atom=rhs ",""),src.replace("rhs_operator=eq","rhs_operator=atom")] {assert!(grameo::compile_with_options(&modified,grameo::CompileOptions{optimizations:true,..Default::default()}).is_err());}
}
