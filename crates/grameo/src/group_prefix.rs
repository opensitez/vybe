//! Explicit certificate for delimited alternatives sharing their first expression.
use crate::{CompiledGrammar, RuleMode, program::{Program,Instruction,SharedGroupPrefix}};
fn ug<P:Program>(p:&P,mut e:usize)->Option<usize>{for _ in 0..128{if let Instruction::Group(n)=p.instruction(e){e=n}else{return Some(e)}}None}
fn seq<P:Program>(p:&P,e:usize)->Option<(usize,usize)>{match p.instruction(ug(p,e)?){Instruction::Sequence{left,right}=>Some((left,right)),_=>None}}
fn call<P:Program>(p:&P,e:usize,r:usize)->bool{ug(p,e).is_some_and(|e|matches!(p.instruction(e),Instruction::Call(n) if n==r))}
fn literal<P:Program>(p:&P,e:usize,s:&str)->bool{ug(p,e).is_some_and(|e|matches!(p.instruction(e),Instruction::Literal{text,insensitive:false} if text==s))}
pub(crate) fn valid<P:Program>(p:&P,s:SharedGroupPrefix,id:usize,left:usize,right:usize,scope:usize)->bool {
 let (Some(optional),Some(commas),Some(trailing))=(ug(p,s.optional),ug(p,s.commas),ug(p,s.trailing))else{return false;};
 let t=p.rule(s.target);let b=p.rule(s.branch);let f=p.rule(s.fallback_rule);
 if id!=s.choice||left!=s.left||right!=s.right||t.expression!=s.target_expression||b.expression!=s.branch_expression||f.expression!=s.fallback_expression
  ||[t.mode,b.mode,f.mode].into_iter().any(|m|m!=RuleMode::Normal)||[t.scope,b.scope,f.scope].into_iter().any(|v|v!=scope)
  ||[s.target,s.branch,s.fallback_rule].into_iter().any(|r|p.source_pratt(r).is_some()) {return false;}
 let sc=p.scope(scope);if [sc.whitespace,sc.comment].into_iter().flatten().any(|r|r==s.branch||r==s.fallback_rule){return false;}
 seq(p,left)==Some((s.left_inner,s.close))&&seq(p,s.left_inner)==Some((s.open,s.branch_call))&&call(p,s.branch_call,s.branch)
  &&seq(p,right)==Some((s.right_inner,s.fallback_close))&&seq(p,s.right_inner)==Some((s.fallback_open,s.optional))
  &&literal(p,s.open,"(")&&literal(p,s.fallback_open,"(")&&literal(p,s.close,")")&&literal(p,s.fallback_close,")")
  &&matches!(p.instruction(optional),Instruction::Repeat{expression,min:0,max:Some(1)} if call(p,expression,s.fallback_rule))
  &&seq(p,b.expression)==Some((s.prefix,s.branch_tail))&&call(p,s.prefix,s.prefix_rule)
  &&seq(p,f.expression)==Some((s.fallback_inner,s.trailing))&&seq(p,s.fallback_inner)==Some((s.fallback_prefix,s.commas))&&call(p,s.fallback_prefix,s.prefix_rule)
  &&matches!(p.instruction(commas),Instruction::Repeat{expression,min:0,max:None} if seq(p,expression).is_some_and(|(a,b)|literal(p,a,",")&&call(p,b,s.prefix_rule)))
  &&matches!(p.instruction(trailing),Instruction::Repeat{expression,min:0,max:Some(1)} if literal(p,expression,","))
}
pub(crate) fn certify(p:&CompiledGrammar,target:&str,branch:&str,fallback:&str,prefix:&str)->Result<SharedGroupPrefix,String>{
 let r=|name:&str|p.rule_id(name).ok_or_else(||format!("unknown shared-group-prefix rule '{name}'"));
 let target=r(target)?;let branch=r(branch)?;let fallback_rule=r(fallback)?;let prefix_rule=r(prefix)?;
 if target==branch||target==fallback_rule||branch==fallback_rule{return Err("shared-group-prefix rules must be distinct".into());}
 fn first_pair<P:Program>(p:&P,e:usize)->Option<(usize,usize,usize)>{let e=ug(p,e)?;let Instruction::Choice{left,right}=p.instruction(e)else{return None};if matches!(p.instruction(ug(p,left)?),Instruction::Choice{..}){first_pair(p,left)}else{Some((e,left,right))}}
 let (choice,left,right)=first_pair(p,p.rule(target).expression).ok_or("shared-group-prefix requires ordered alternatives")?;
 let bad=||"shared-group-prefix requires '(branch)' | '(fallback?)' with shared expression and list tails".to_string();
 let (left_inner,close)=seq(p,left).ok_or_else(bad)?;let (open,branch_call)=seq(p,left_inner).ok_or_else(bad)?;
 let (right_inner,fallback_close)=seq(p,right).ok_or_else(bad)?;let (fallback_open,optional)=seq(p,right_inner).ok_or_else(bad)?;
 let (prefix,branch_tail)=seq(p,p.rule(branch).expression).ok_or_else(bad)?;
 let (fallback_inner,trailing)=seq(p,p.rule(fallback_rule).expression).ok_or_else(bad)?;let (fallback_prefix,commas)=seq(p,fallback_inner).ok_or_else(bad)?;
 let s=SharedGroupPrefix{target,target_expression:p.rule(target).expression,choice,left,right,left_inner,open,branch_call,close,right_inner,fallback_open,optional,fallback_close,branch,branch_expression:p.rule(branch).expression,fallback_rule,fallback_expression:p.rule(fallback_rule).expression,prefix_rule,prefix,branch_tail,fallback_inner,fallback_prefix,commas,trailing};
 if !valid(p,s,choice,left,right,p.rule(target).scope){return Err(bad());}Ok(s)
}
