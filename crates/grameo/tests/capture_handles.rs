use grameo::{builder::Builder, Span, MatchOptions};
use grameo::tree::CaptureHandle;
#[derive(Default)]struct Hooks{events:Vec<(usize,Span,Option<CaptureHandle>)>,requested:Option<usize>}
impl Builder for Hooks {
 fn checkpoint(&self)->usize{self.events.len()}
 fn rollback(&mut self,n:usize){self.events.truncate(n)}
 fn begin(&mut self,_:usize,_:usize){}
 fn finish(&mut self,r:usize,s:Span,_:&str){self.events.push((r,s,None))}
 fn observes_capture(&self,r:usize)->bool{self.requested==Some(r)}
 fn try_finish_capture(&mut self,r:usize,s:Span,_:&str,c:Option<CaptureHandle>)->Result<(),grameo::builder::BuildError>{self.events.push((r,s,c));Ok(())}
}
#[derive(Clone,Copy)]struct Rule;
impl grameo::compat::RuleIdentity for Rule {fn from_capture(_:grameo::CaptureRule)->Option<Self>{Some(Self)}}
#[test]fn only_selected_rule_receives_a_resolvable_token_and_mut_borrows_forward_selection(){
 let g=grameo::compile("root={word~EOI} word=@{ASCII_ALPHA+}").unwrap();let word=g.rule_id("word").unwrap();
 let mut h=Hooks{requested:Some(word),..Default::default()};let tree=g.build_captures("root","hello",&mut &mut h).unwrap();let root=tree.into_typed_pairs::<Rule>().next().unwrap();
 let mut handles=0;for (r,_,c) in &h.events {if let Some(c)=c {assert_eq!(*r,word);assert_eq!(root.resolve_capture(*c).unwrap().as_str(),"hello");handles+=1;}}
 assert_eq!(handles,1);
 let other_source="world";let other=g.build_captures("root",other_source,&mut Hooks::default()).unwrap().into_typed_pairs::<Rule>().next().unwrap();assert!(other.resolve_capture(h.events.iter().find_map(|(_,_,c)|*c).unwrap()).is_none());
}
#[test]fn default_capture_admission_matches_legacy_finish_events_without_tokens(){
 let g=grameo::compile("root={word~EOI} word=@{ASCII_ALPHA+}").unwrap();let mut no=Hooks::default();let mut selected=Hooks{requested:g.rule_id("word"),..Default::default()};
 let a=g.build_captures("root","hello",&mut no).unwrap();let b=g.build_captures("root","hello",&mut selected).unwrap();assert_eq!(a.capture_count(),b.capture_count());
 assert!(no.events.iter().all(|(_,_,c)|c.is_none()));assert_eq!(no.events.iter().map(|(r,s,_)|(*r,*s)).collect::<Vec<_>>(),selected.events.iter().map(|(r,s,_)|(*r,*s)).collect::<Vec<_>>());
}
#[test]fn failed_speculation_and_budget_failure_restore_all_receipts(){
 let g=grameo::compile("root={(word~\"!\"|word)~EOI} word=@{ASCII_ALPHA+}").unwrap();
 for input in ["hello","hello?","hello!"] {for max_steps in (0..120).chain([usize::MAX]) {for max_rule_depth in 0..8 {
  let mut h=Hooks{requested:g.rule_id("word"),..Default::default()};let result=grameo::engine::build_capture_program(&g,"root",input,MatchOptions{max_steps,max_rule_depth},&mut h);
  if let Ok(tree)=result{let root=tree.into_typed_pairs::<Rule>().next().unwrap();for (_,_,c) in &h.events {if let Some(c)=c{assert_eq!(root.resolve_capture(*c).unwrap().as_str(),"hello");}}assert_eq!(h.events.iter().filter(|(_,_,c)|c.is_some()).count(),1);}else{assert!(h.events.is_empty());}
 }}}
}
#[test]fn builders_without_new_hooks_keep_their_existing_finish_policy(){
 #[derive(Default)]struct Legacy(Vec<(usize,Span)>);
 impl Builder for Legacy{fn checkpoint(&self)->usize{self.0.len()}fn rollback(&mut self,n:usize){self.0.truncate(n)}fn begin(&mut self,_:usize,_:usize){}fn finish(&mut self,r:usize,s:Span,_:&str){self.0.push((r,s));}}
 let g=grameo::compile("root={word~EOI} word=@{ASCII_ALPHA+}").unwrap();let mut old=Legacy::default();let mut new=Hooks::default();
 let a=g.build_captures("root","hello",&mut old).unwrap();let b=g.build_captures("root","hello",&mut new).unwrap();assert_eq!(a.capture_count(),b.capture_count());assert_eq!(old.0,new.events.iter().map(|(r,s,_)|(*r,*s)).collect::<Vec<_>>());
 let prior=old.0.clone();assert!(g.build_captures("root","hello?",&mut old).is_err());assert_eq!(old.0,prior);
}
