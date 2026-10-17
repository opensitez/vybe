//! Explicit successful-empty rule certificate; original PEG IR is untouched.
use crate::{CompiledGrammar,RuleMode,program::{Program,Instruction,EmptySuccessHead}};
pub(crate) fn certify(grammar:&CompiledGrammar,rule:usize)->Option<EmptySuccessHead>{
 let d=grammar.rule(rule);if d.mode!=RuleMode::Silent||grammar.source_pratt(rule).is_some(){return None;}
 let mut root=d.expression;let mut budget=grammar.syntax().expressions.len().checked_mul(2)?;
 while let Instruction::Group(child)=grammar.instruction(root){budget=budget.checked_sub(1)?;root=child;}
 let Instruction::Repeat{expression,min:0,max}=grammar.instruction(root)else{return None};
 if max==Some(0){return None;}
 let mut pending=vec![expression];let mut first=[0u64;4];
 while let Some(id)=pending.pop(){budget=budget.checked_sub(1)?;match grammar.instruction(id){
  Instruction::Group(c)=>pending.push(c),
  Instruction::Choice{left,right}=>{pending.push(right);pending.push(left);},
  Instruction::Literal{text,insensitive}=>{let b=*text.as_bytes().first()?;first[usize::from(b/64)]|=1u64<<(b%64);if insensitive&&b.is_ascii_alphabetic(){let b=if b.is_ascii_lowercase(){b.to_ascii_uppercase()}else{b.to_ascii_lowercase()};first[usize::from(b/64)]|=1u64<<(b%64);}},
  _=>return None,
 }}
 Some(EmptySuccessHead{expression:d.expression,scope:d.scope,first})
}
