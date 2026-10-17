use grameo::{builder::{Builder,BuildError},Span,MatchOptions};
#[derive(Default)]
struct B { selected:usize, values:Vec<(Span,Vec<(String,Span)>,usize)>, consumed:usize, fail:bool }
impl Builder for B {
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, _: usize, _: grameo::Span, _: &str) {}
 fn observes_rule(&self,r:usize)->bool{r==self.selected}
 fn wants_completed_capture(&self,r:usize)->bool{r==self.selected}
 fn checkpoint(&self)->usize{self.values.len()}
 fn rollback(&mut self,m:usize){self.values.truncate(m)}
 fn try_completed_capture(&mut self,_:usize,span:Span,tree:grameo::ParseTree<'_,'_>)->Result<(),BuildError>{
  if self.fail{return Err(BuildError{span,message:"consumer rejected".into()})}
  let count=tree.capture_count();let root=tree.pairs().next().unwrap();assert_eq!(root.as_span(),span);assert_eq!(tree.consumed(),span.end);
  let children=root.into_inner().map(|p|(p.rule_name().to_owned(),p.as_span())).collect();
  self.values.push((span,children,count));Ok(())
 }
 fn committed(&mut self,_:&str,n:usize){self.consumed=n}
}
#[test]fn callbacks_run_per_completed_subtree_with_absolute_spans(){
 let g=grameo::compile("root={SOI~stmt*~EOI} stmt={name~\";\"} name=@{ASCII_ALPHA+}").unwrap();let mut b=B{selected:g.rule_id("stmt").unwrap(),..B::default()};
 grameo::engine::build_capture_program(&g,"root","one;two;three;",MatchOptions::default(),&mut b).unwrap();
 assert_eq!(b.values.len(),3);assert_eq!(b.values.iter().map(|v|v.2).collect::<Vec<_>>(),[2,2,2]);assert_eq!(b.values[1].0,Span{start:4,end:8});assert_eq!(b.values[1].1[0],("name".into(),Span{start:4,end:7}));assert_eq!(b.consumed,14);
}
#[test]fn speculative_completed_subtree_is_rolled_back(){
 let g=grameo::compile("root={SOI~((stmt~\"!\")|(stmt~\"?\"))~EOI} stmt={\"a\"}").unwrap();let mut b=B{selected:g.rule_id("stmt").unwrap(),..B::default()};
 grameo::engine::build_capture_program(&g,"root","a?",MatchOptions::default(),&mut b).unwrap();assert_eq!(b.values.len(),1);
}
#[test]fn rejected_callback_restores_initial_transaction(){
 let g=grameo::compile("root={stmt~EOI} stmt={\"a\"}").unwrap();let mut b=B{selected:g.rule_id("stmt").unwrap(),fail:true,..B::default()};
 let e=grameo::engine::build_capture_program(&g,"root","a",MatchOptions::default(),&mut b).unwrap_err();assert!(e.to_string().contains("consumer rejected"));assert!(b.values.is_empty());assert_eq!(b.consumed,0);
}
#[test]fn failed_unit_restores_completed_callbacks_and_no_commit(){
 let g=grameo::compile("root={stmt*~EOI} stmt={\"a\"}").unwrap();let mut b=B{selected:g.rule_id("stmt").unwrap(),..B::default()};
 assert!(grameo::engine::build_capture_program(&g,"root","a?",MatchOptions::default(),&mut b).is_err());assert!(b.values.is_empty());assert_eq!(b.consumed,0);
}
#[test]fn finite_budgets_preserve_transactional_callback_results(){
 let g=grameo::compile("root={stmt*~EOI} stmt={name~\";\"} name=@{ASCII_ALPHA+}").unwrap();
 for max_steps in [0,1,2,8,16,32,64,128,usize::MAX]{for max_rule_depth in [0,1,2,4,16]{let mut b=B{selected:g.rule_id("stmt").unwrap(),..B::default()};let result=grameo::engine::build_capture_program(&g,"root","a;b;",MatchOptions{max_steps,max_rule_depth},&mut b);if result.is_err(){assert!(b.values.is_empty())}else{assert_eq!(b.values.len(),2)}}}
}
#[test]fn streaming_adapter_emits_snapshot_value_once_and_journals_sink(){
 use grameo::streaming::{Boundary,CompletedUnits,Sink,StreamingBuilder};
 impl CompletedUnits for B {
  type Unit=(Span,Vec<(String,Span)>,usize);
  fn is_emission_boundary(&self,r:usize)->bool{r==self.selected}
  fn completed_unit(&self,_:Boundary)->Result<Option<&Self::Unit>,BuildError>{Ok(self.values.last())}
 }
 #[derive(Default)]struct Output(Vec<Span>);
 impl Sink<(Span,Vec<(String,Span)>,usize)> for Output {
  type Checkpoint=usize;
  fn checkpoint(&self)->usize{self.0.len()}
  fn rollback(&mut self,m:usize){self.0.truncate(m)}
  fn emit(&mut self,u:&(Span,Vec<(String,Span)>,usize),_:Boundary,_:&str)->Result<(),BuildError>{self.0.push(u.0);Ok(())}
 }
 let g=grameo::compile("root={stmt*~EOI} stmt={\"a\"}").unwrap();let builder=B{selected:g.rule_id("stmt").unwrap(),..B::default()};let mut stream=StreamingBuilder::new(builder,Output::default());
 grameo::engine::build_capture_program(&g,"root","aa",MatchOptions::default(),&mut stream).unwrap();assert_eq!(stream.sink().0,[Span{start:0,end:1},Span{start:1,end:2}]);
}
