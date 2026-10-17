use std::cell::RefCell;
use grameo::{program::{Instruction,RuleSpec,Scope,SourcePratt},CompileOptions,MatchOptions,Span,builder::Builder,program::Program};
const BODY:&str=r##"
WHITESPACE = _{ " " | "\t" }
COMMENT = _{ "#" ~ (!NEWLINE ~ ANY)* }
entry = { expression ~ EOI }
expression = { primary ~ ("+" ~ primary)* }
primary = { "(" ~ gen ~ ")" | "(" ~ list? ~ ")" | ident }
gen = { expression ~ clause+ }
clause = { "for" ~ ident ~ "in" ~ ident }
list = { expression ~ ("," ~ expression)* ~ ","? }
ident = @{ ASCII_ALPHA+ }
"##;
const ANN:&str="// @grameo shared_group_prefix primary branch=gen fallback=list prefix=expression\n";
#[derive(Default)]struct Hooks{on:bool,observed:bool,pratt:bool,selected:Option<usize>,events:Vec<(usize,Span)>}
impl Builder for Hooks{
 fn supports_shared_group_prefix(&self,_:usize,_:usize)->bool{self.on}
 fn observes_rule(&self,r:usize)->bool{self.observed||self.selected==Some(r)}
 fn supports_pratt(&self,_:usize)->bool{self.pratt}
 fn checkpoint(&self)->usize{self.events.len()}
 fn rollback(&mut self,n:usize){self.events.truncate(n)}
 fn begin(&mut self,r:usize,o:usize){self.events.push((r,Span{start:o,end:o}))}
 fn finish(&mut self,r:usize,s:Span,_:&str){self.events.push((r,s))}
}
fn shape(tree:&grameo::ParseTree<'_,'_>)->Vec<(String,Span,usize)>{
 fn visit(p:grameo::tree::Pair<'_,'_,'_>,n:usize,out:&mut Vec<(String,Span,usize)>){out.push((p.rule_name().into(),p.as_span(),n));for c in p.into_inner(){visit(c,n+1,out)}}
 let mut v=Vec::new();for p in tree.pairs(){visit(p,0,&mut v)}v
}
#[test]fn resource_capture_error_hooks(){
 let g=grameo::compile_with_options(&format!("{ANN}{BODY}"),CompileOptions{optimizations:true,..Default::default()}).unwrap();
 let mut sources:Vec<String>=["x","()","(x)","(x,)","(x,y)","(x,y,)","(x for y in z)","(x for y in z for a in b)","(x for)","(x for y in z,)","(x", "(","(,)","(x,,y)","(x +)","(x) + y","(x # c\n)","( )"].into_iter().map(String::from).collect();
 for depth in [2,4,6,8,12]{sources.push(format!("{}x{}","(".repeat(depth),")".repeat(depth)));sources.push(format!("{}x{}","(".repeat(depth),")".repeat(depth-1)));}
 let mut count=0;
 for source in sources{for selected in [None,g.rule_id("expression"),g.rule_id("ident"),g.rule_id("clause"),g.rule_id("gen"),g.rule_id("list")]{for observed in [false,true]{for depth in [1,2,4,8,16,64,4096]{for steps in [1,32,256,usize::MAX]{
 let options=MatchOptions{max_rule_depth:depth,max_steps:steps,..Default::default()};let(mut a,mut b)=(Hooks{selected,observed,..Default::default()},Hooks{on:true,selected,observed,..Default::default()});
 let left=grameo::engine::build_capture_program(&g,"entry",&source,options,&mut a);let right=grameo::engine::build_capture_program(&g,"entry",&source,options,&mut b);
 match(left,right){(Ok(a),Ok(b))=>assert_eq!(shape(&a),shape(&b),"{source} depth={depth} steps={steps}"),(Err(a),Err(b))=>assert_eq!(a,b,"{source} depth={depth} steps={steps}"),(a,b)=>panic!("{source} {depth} {steps}: {a:?} {b:?}")}
 assert_eq!(a.events,b.events);count+=1;
 }}}}}println!("{count} exact capture/error/resource comparisons");
}
#[test]fn default_inert(){let g=grameo::compile(&format!("{ANN}{BODY}")).unwrap();assert!((0..g.syntax().expressions.len()).all(|id|g.shared_group_prefix(id).is_none()));}

struct TokenHooks{on:bool,rule:usize,handles:Vec<grameo::tree::CaptureHandle>}
impl Builder for TokenHooks{
 fn supports_shared_group_prefix(&self,_:usize,_:usize)->bool{self.on}
 fn observes_rule(&self,r:usize)->bool{r==self.rule}
 fn observes_capture(&self,r:usize)->bool{r==self.rule}
 fn checkpoint(&self)->usize{self.handles.len()}
 fn rollback(&mut self,n:usize){self.handles.truncate(n)}
 fn begin(&mut self,_:usize,_:usize){} fn finish(&mut self,_:usize,_:Span,_:&str){}
 fn try_finish_capture(&mut self,_:usize,_:Span,_:&str,h:Option<grameo::tree::CaptureHandle>)->Result<(),grameo::builder::BuildError>{if let Some(h)=h{self.handles.push(h)}Ok(())}
}
#[test]fn retained_prefix_handles_remain_valid(){
 let g=grameo::compile_with_options(&format!("{ANN}{BODY}"),CompileOptions{optimizations:true,..Default::default()}).unwrap();
 for source in ["(x)","((x,y))","(x for y in z)","((x),y,)"]{
  for name in ["expression","ident","primary","gen","list"]{
   let rule=g.rule_id(name).unwrap();let mut a=TokenHooks{on:false,rule,handles:vec![]};let mut b=TokenHooks{on:true,rule,handles:vec![]};
   let ta=grameo::engine::build_capture_program(&g,"entry",source,MatchOptions::default(),&mut a).unwrap();let tb=grameo::engine::build_capture_program(&g,"entry",source,MatchOptions::default(),&mut b).unwrap();
   let anchor_a=ta.into_typed_pairs::<Identity>().next().unwrap();let anchor_b=tb.into_typed_pairs::<Identity>().next().unwrap();
   let ra:Vec<_>=a.handles.into_iter().map(|h|anchor_a.resolve_capture(h).map(|p|(p.as_rule(),p.as_str().to_owned()))).collect();
   let rb:Vec<_>=b.handles.into_iter().map(|h|anchor_b.resolve_capture(h).map(|p|(p.as_rule(),p.as_str().to_owned()))).collect();
   assert_eq!(ra,rb,"{source} {name}");assert!(ra.iter().all(Option::is_some));
  }
 }
}

#[derive(Clone,Copy,Debug,PartialEq,Eq)]struct Identity(grameo::CaptureRule);
impl grameo::compat::RuleIdentity for Identity{fn from_capture(r:grameo::CaptureRule)->Option<Self>{Some(Self(r))}}

#[test]fn identical_prefix_pratt_override_keeps_scoped_captures(){
 let source=format!("{ANN}// @grameo pratt expression atom=primary trailing=consume\n{BODY}");
 let g=grameo::compile_with_options(&source,CompileOptions{optimizations:true,..Default::default()}).unwrap();
 for source in ["x","()","(x)","((x))","(x,y,)","(x for y in z)","(x for)","(x", "((x,,y))"]{
  for max_rule_depth in [1,2,4,8,16,64,4096]{let options=MatchOptions{max_rule_depth,..Default::default()};let(mut a,mut b)=(Hooks{pratt:true,..Default::default()},Hooks{pratt:true,on:true,..Default::default()});
   match(grameo::engine::build_capture_program(&g,"entry",source,options,&mut a),grameo::engine::build_capture_program(&g,"entry",source,options,&mut b)){
    (Ok(a),Ok(b))=>assert_eq!(shape(&a),shape(&b),"{source} {max_rule_depth}"),(Err(a),Err(b))=>assert_eq!(a,b,"{source} {max_rule_depth}"),(a,b)=>panic!("{source} {max_rule_depth}: {a:?} {b:?}")
   }
  }
 }
}
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
    fn shared_group_prefix(&self, id: usize) -> Option<grameo::program::SharedGroupPrefix> {
        self.base.shared_group_prefix(id)
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

#[test]fn forwarded_current_geometry_scope_and_observers(){
 let g=grameo::compile_with_options(&format!("{ANN}{BODY}"),CompileOptions{optimizations:true,..Default::default()}).unwrap();
 let spec=(0..g.syntax().expressions.len()).find_map(|id|g.shared_group_prefix(id)).unwrap();
 for mutation in [Mutation::None,Mutation::Pratt(spec.target),Mutation::Pratt(spec.branch),Mutation::Pratt(spec.fallback_rule),Mutation::Pratt(spec.prefix_rule),Mutation::WrapperBody(spec.open),Mutation::WrapperBody(spec.branch_expression),Mutation::WrapperBody(spec.fallback_expression),Mutation::WrapperBody(spec.prefix),Mutation::WrapperBody(spec.commas),Mutation::WrapperBody(spec.trailing),Mutation::Scope,Mutation::ScopeData,Mutation::Observe]{
  let p=Forward{base:&g,mutation,wrapper:spec.branch,branch:spec.target,other:g.rule_id("ident").unwrap(),events:std::cell::RefCell::new(vec![])};
  for source in ["x","()","(x)","((x))","(x,y,)","(x for y in z)","(x for)","(x"]{
   for max_rule_depth in [2,8,32,4096]{let options=MatchOptions{max_rule_depth,..Default::default()};let(mut a,mut b)=(Hooks::default(),Hooks{on:true,..Default::default()});
    let x=grameo::engine::build_capture_program(&p,"entry",source,options,&mut a);let events=p.events.take();let y=grameo::engine::build_capture_program(&p,"entry",source,options,&mut b);assert_eq!(events,p.events.take(),"observer {source} {mutation:?}");
    match(x,y){(Ok(a),Ok(b))=>assert_eq!(shape(&a),shape(&b),"{source} {mutation:?}"),(Err(a),Err(b))=>assert_eq!(a,b,"{source} {mutation:?}"),(a,b)=>panic!("{source} {mutation:?}: {a:?} {b:?}")}
   }
  }
 }
}
