//! Explicit intra-rule deterministic prefix replay elimination. No captures move.
use crate::{CompiledGrammar,RuleMode,program::{Program,Instruction,SharedInnerPrefix}};
fn ug<P:Program>(p:&P,mut e:usize)->Option<usize>{for _ in 0..128{if let Instruction::Group(n)=p.instruction(e){e=n}else{return Some(e)}}None}
fn seq<P:Program>(p:&P,e:usize)->Option<(usize,usize)>{match p.instruction(ug(p,e)?){Instruction::Sequence{left,right}=>Some((left,right)),_=>None}}
fn call<P:Program>(p:&P,e:usize,r:usize)->bool{ug(p,e).is_some_and(|e|matches!(p.instruction(e),Instruction::Call(n) if n==r))}
fn lit<P:Program>(p:&P,e:usize,s:&str)->bool{ug(p,e).is_some_and(|e|matches!(p.instruction(e),Instruction::Literal{text,insensitive:false} if text==s))}
fn optional<P:Program>(p:&P,e:usize)->Option<usize>{match p.instruction(ug(p,e)?){Instruction::Repeat{expression,min:0,max:Some(1)}=>Some(expression),_=>None}}
fn shape<P:Program>(p:&P,target:usize,prefix_rule:usize,kind:usize)->Option<SharedInnerPrefix>{
 let target_expression=p.rule(target).expression;let choice=ug(p,target_expression)?;
 let Instruction::Choice{left,right}=p.instruction(choice)else{return None};
 let (prefix,first,second,third)=match kind{
  0=>{
   let (prefix,first)=seq(p,left)?;if !call(p,prefix,prefix_rule){return None;}
   let (opt,third)=seq(p,right)?;let (other,second)=seq(p,optional(p,opt)?)?;
   if !call(p,other,prefix_rule)||!lit(p,optional(p,third)?,","){return None;}
   let Instruction::Repeat{expression,min:0,max:None}=p.instruction(ug(p,second)?)else{return None};
   let (comma,operand)=seq(p,expression)?;if !lit(p,comma,",")||!call(p,operand,prefix_rule){return None;}
   let Instruction::Repeat{min:1,max:None,..}=p.instruction(ug(p,first)?)else{return None};
   (prefix,first,second,third)
  },
  1=>{
   let (inner,third)=seq(p,left)?;let (head,second)=seq(p,inner)?;let (opt,first)=seq(p,head)?;
   let prefix=optional(p,opt)?;
   if !call(p,prefix,prefix_rule)||!call(p,right,prefix_rule)||!lit(p,first,":")||!call(p,optional(p,second)?,prefix_rule){return None;}
   let (colon,bound)=seq(p,optional(p,third)?)?;if !lit(p,colon,":")||!call(p,optional(p,bound)?,prefix_rule){return None;}
   (prefix,first,second,third)
  },_=>return None,
 };
 Some(SharedInnerPrefix{target,target_expression,choice,left,right,prefix_rule,prefix,first,second,third,kind})
}
pub(crate)fn valid<P:Program>(p:&P,s:SharedInnerPrefix,id:usize,left:usize,right:usize,scope:usize)->bool{
 let t=p.rule(s.target);let sc=p.scope(scope);
 id==s.choice&&left==s.left&&right==s.right&&t.mode==RuleMode::Normal&&t.scope==scope&&p.source_pratt(s.target).is_none()
 &&! [sc.whitespace,sc.comment].into_iter().flatten().any(|r|r==s.target)
 &&shape(p,s.target,s.prefix_rule,s.kind).is_some_and(|a|a==s)
}
pub(crate)fn certify(p:&CompiledGrammar,target:&str,prefix:&str,kind:&str)->Result<SharedInnerPrefix,String>{
 let r=|n:&str|p.rule_id(n).ok_or_else(||format!("unknown shared-inner-prefix rule '{n}'"));
 let target=r(target)?;let prefix=r(prefix)?;let kind=match kind{"list"=>0,"index"=>1,_=>return Err("shared-inner-prefix kind must be list or index".into())};
 let s=shape(p,target,prefix,kind).ok_or("shared-inner-prefix requires exact expression/comprehension/list or optional-slice/index shape")?;
 if !valid(p,s,s.choice,s.left,s.right,p.rule(target).scope){return Err("shared-inner-prefix target must retain normal same-scope PEG geometry".into());}Ok(s)
}
