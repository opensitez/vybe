use grameo::{CompileOptions,MatchOptions,Span,builder::Builder,program::Program};
const BODY:&str=r##"
WHITESPACE=_{" "|"\n"}
COMMENT=_{"#"~(!"\n"~ANY)*}
entry={expression~EOI}
expression={primary~("+"~primary)*}
primary={"["~list~"]"|ident~("["~index~"]")*}
list={expression~clause+|(expression~(","~expression)*)?~","?}
clause={"for"~ident~"in"~ident}
index={expression?~":"~expression?~(":"~expression?)?|expression}
ident=@{ASCII_ALPHA+}
"##;
const ANN:&str="// @grameo shared_inner_prefix list prefix=expression kind=list\n// @grameo shared_inner_prefix index prefix=expression kind=index\n";
#[derive(Default)]struct Hooks{on:bool,observe:bool,events:Vec<(usize,Span)>}
impl Builder for Hooks{
 fn supports_shared_inner_prefix(&self,_:usize)->bool{self.on}
 fn observes_rule(&self,_:usize)->bool{self.observe}
 fn checkpoint(&self)->usize{self.events.len()}
 fn rollback(&mut self,n:usize){self.events.truncate(n)}
 fn begin(&mut self,r:usize,o:usize){self.events.push((r,Span{start:o,end:o}))}
 fn finish(&mut self,r:usize,s:Span,_:&str){self.events.push((r,s))}
}
fn shape(t:&grameo::ParseTree<'_,'_>)->Vec<(String,Span,usize)>{fn v(p:grameo::tree::Pair<'_,'_,'_>,n:usize,o:&mut Vec<(String,Span,usize)>){o.push((p.rule_name().into(),p.as_span(),n));for c in p.into_inner(){v(c,n+1,o)}}let mut o=vec![];for p in t.pairs(){v(p,0,&mut o)}o}
#[test]fn parity(){let g=grameo::compile_with_options(&format!("{ANN}{BODY}"),CompileOptions{optimizations:true,..Default::default()}).unwrap();let mut sources:Vec<String>=["[]","[x]","[x,]","[x,y,]","[x for y in z]","[x for]","[x,,y]","[x", "[", "[,]","a[x]","a[:]","a[::]","a[x:]","a[:y]","a[x:y:z]","a[x::z]","a[x:y:]","a[]","a[x y]","a[x:y:z:q]","a[x+ ]","[x #c\n, y]"].into_iter().map(String::from).collect();for n in [2,4,6,8]{sources.push(format!("{}x{}","[".repeat(n),"]".repeat(n)));sources.push(format!("{}x{}","a[".repeat(n),"]".repeat(n)));}
 let mut count=0;for s in sources{for observe in [false,true]{for depth in [1,2,4,8,16,64,4096]{for steps in [1,32,256,usize::MAX]{let o=MatchOptions{max_rule_depth:depth,max_steps:steps,..Default::default()};let(mut a,mut b)=(Hooks{observe,..Default::default()},Hooks{on:true,observe,..Default::default()});let x=grameo::engine::build_capture_program(&g,"entry",&s,o,&mut a);let y=grameo::engine::build_capture_program(&g,"entry",&s,o,&mut b);match(x,y){(Ok(x),Ok(y))=>assert_eq!(shape(&x),shape(&y),"{s} {depth} {steps}"),(Err(x),Err(y))=>assert_eq!(x,y,"{s} {depth} {steps}"),(x,y)=>panic!("{s} {depth} {steps} {x:?} {y:?}")}assert_eq!(a.events,b.events,"{s}");count+=1;}}}}println!("{count} exact comparisons");}
#[test]fn default_inert(){let g=grameo::compile(&format!("{ANN}{BODY}")).unwrap();assert!((0..g.syntax().expressions.len()).all(|id|g.shared_inner_prefix(id).is_none()));}
#[derive(Debug)]struct Forward<'a>{g:&'a grameo::CompiledGrammar,replace:Option<usize>,target:usize,pratt:bool,observer:bool}
impl Program for Forward<'_>{
 fn rule_id(&self,n:&str)->Option<usize>{self.g.rule_id(n)}
 fn rule(&self,r:usize)->grameo::program::RuleSpec<'_>{self.g.rule(r)}
 fn instruction(&self,e:usize)->grameo::program::Instruction<'_>{if self.replace==Some(e){grameo::program::Instruction::Literal{text:"!",insensitive:false}}else{self.g.instruction(e)}}
 fn scope(&self,s:usize)->grameo::program::Scope{self.g.scope(s)}
 fn source_pratt(&self,r:usize)->Option<grameo::program::SourcePratt<'_>>{if self.pratt&&r==self.target{Some(grameo::program::SourcePratt{atom:self.g.rule_id("ident").unwrap(),operators:&[],trailing_trivia:grameo::islands::TrailingTrivia::Consume})}else{self.g.source_pratt(r)}}
 fn shared_inner_prefix(&self,e:usize)->Option<grameo::program::SharedInnerPrefix>{self.g.shared_inner_prefix(e)}
 fn match_observers_are_noop(&self)->bool{!self.observer}
}
#[test]fn forwarded_geometry_guards(){let g=grameo::compile_with_options(&format!("{ANN}{BODY}"),CompileOptions{optimizations:true,..Default::default()}).unwrap();let mut count=0;for id in 0..g.syntax().expressions.len(){if let Some(s)=g.shared_inner_prefix(id){for replace in [None,Some(s.left),Some(s.right),Some(s.prefix),Some(s.first),Some(s.second),Some(s.third)]{for pratt in [false,true]{for observer in [false,true]{let f=Forward{g:&g,replace,target:s.target,pratt,observer};for src in ["[x]","[x,y,]","[x for y in z]","a[x]","a[x:y]","a[:]","a[x y]","[x for]"]{let(mut a,mut b)=(Hooks::default(),Hooks{on:true,..Default::default()});match(grameo::engine::build_capture_program(&f,"entry",src,MatchOptions::default(),&mut a),grameo::engine::build_capture_program(&f,"entry",src,MatchOptions::default(),&mut b)){(Ok(x),Ok(y))=>assert_eq!(shape(&x),shape(&y)),(Err(x),Err(y))=>assert_eq!(x,y),(x,y)=>panic!("{src} {s:?} {replace:?} {pratt} {x:?} {y:?}")}count+=1;}}}}}}println!("{count} forwarded geometry/Pratt/observer comparisons");}
struct Tokens{on:bool,rule:usize,handles:Vec<grameo::tree::CaptureHandle>}
impl Builder for Tokens{
 fn supports_shared_inner_prefix(&self,_:usize)->bool{self.on}
 fn observes_rule(&self,r:usize)->bool{r==self.rule}
 fn observes_capture(&self,r:usize)->bool{r==self.rule}
 fn checkpoint(&self)->usize{self.handles.len()}
 fn rollback(&mut self,n:usize){self.handles.truncate(n)}
 fn begin(&mut self,_:usize,_:usize){}fn finish(&mut self,_:usize,_:Span,_:&str){}
 fn try_finish_capture(&mut self,_:usize,_:Span,_:&str,h:Option<grameo::tree::CaptureHandle>)->Result<(),grameo::builder::BuildError>{if let Some(h)=h{self.handles.push(h)}Ok(())}
}
#[derive(Clone,Copy,Debug,PartialEq,Eq)]struct Identity(grameo::CaptureRule);
impl grameo::compat::RuleIdentity for Identity{fn from_capture(r:grameo::CaptureRule)->Option<Self>{Some(Self(r))}}
#[test]fn handles_survive_retained_prefix(){let g=grameo::compile_with_options(&format!("{ANN}{BODY}"),CompileOptions{optimizations:true,..Default::default()}).unwrap();for src in ["[x]","[[x,y]]","[x for y in z]","a[a[x]:y:z]"]{for name in ["expression","ident","primary","index","list"]{let rule=g.rule_id(name).unwrap();let(mut a,mut b)=(Tokens{on:false,rule,handles:vec![]},Tokens{on:true,rule,handles:vec![]});let ta=grameo::engine::build_capture_program(&g,"entry",src,MatchOptions::default(),&mut a).unwrap();let tb=grameo::engine::build_capture_program(&g,"entry",src,MatchOptions::default(),&mut b).unwrap();let pa=ta.into_typed_pairs::<Identity>().next().unwrap();let pb=tb.into_typed_pairs::<Identity>().next().unwrap();let ra:Vec<_>=a.handles.into_iter().map(|h|pa.resolve_capture(h).map(|p|(p.as_rule(),p.as_str().to_owned()))).collect();let rb:Vec<_>=b.handles.into_iter().map(|h|pb.resolve_capture(h).map(|p|(p.as_rule(),p.as_str().to_owned()))).collect();assert_eq!(ra,rb,"{src} {name}");assert!(ra.iter().all(Option::is_some));}}}

#[test]fn owned_boundary_combines_with_inner_prefix_without_replay_or_resource_drift(){
 struct Owned{hooks:Hooks,target:usize,own:bool,bad:bool}
 impl Builder for Owned{
 fn supports_shared_inner_prefix(&self,_:usize)->bool{self.hooks.on}
 fn supports_owned_capture_boundary(&self,r:usize)->bool{self.own&&r==self.target}
 fn owned_capture_end(&self,_:usize,start:usize,source:&str)->Option<usize>{Some(if self.bad{start}else{source.len()})}
 fn observes_rule(&self,_:usize)->bool{true}
 fn checkpoint(&self)->usize{self.hooks.checkpoint()}
 fn rollback(&mut self,n:usize){self.hooks.rollback(n)}
 fn begin(&mut self,r:usize,o:usize){self.hooks.begin(r,o)}
 fn finish(&mut self,r:usize,s:Span,t:&str){self.hooks.finish(r,s,t)}
 }
 let source=format!("// @grameo owned_capture_boundary primary\n{ANN}{BODY}");let g=grameo::compile_with_options(&source,CompileOptions{optimizations:true,..Default::default()}).unwrap();let target=g.rule_id("primary").unwrap();let mut comparisons=0;
 for source in ["[x]","[[x]]","[x for y in z]","[x,]","a[x]","a[x:y]","a[:]","[x,,]","a[x:y:z:q]","[x"]{for own in [false,true]{for bad in [false,true]{for depth in [2,8,64]{for steps in [1,32,usize::MAX]{let options=MatchOptions{max_rule_depth:depth,max_steps:steps};let(mut a,mut b)=(Owned{hooks:Hooks::default(),target,own,bad},Owned{hooks:Hooks{on:true,..Default::default()},target,own,bad});let x=grameo::engine::build_capture_program(&g,"entry",source,options,&mut a);let y=grameo::engine::build_capture_program(&g,"entry",source,options,&mut b);match(x,y){(Ok(x),Ok(y))=>assert_eq!(shape(&x),shape(&y),"{source}"),(Err(x),Err(y))=>assert_eq!(x,y,"{source}"),(x,y)=>panic!("{source}: {x:?} {y:?}")}assert_eq!(a.hooks.events,b.hooks.events,"{source}");comparisons+=1;}}}}}println!("{comparisons} combined ownership/prefix captures/hooks/resource comparisons");
}
