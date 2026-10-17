//! Conservative, grammar-authored certificate for two capture-preserving alternatives.
use crate::{CompiledGrammar, RuleMode, program::{Program, Instruction, SharedWrapperPrefix}};
fn ungroup<P:Program>(p:&P,mut id:usize)->Result<usize,&'static str>{
    for _ in 0..128 { if let Instruction::Group(child)=p.instruction(id){id=child;}else{return Ok(id);} }
    Err("shared-wrapper-prefix group chain is too deep")
}
fn equal<P:Program>(p:&P,a:usize,b:usize,depth:usize)->bool{
    if depth>16{return false;}
    let (Ok(a),Ok(b))=(ungroup(p,a),ungroup(p,b))else{return false;};
    match (p.instruction(a),p.instruction(b)) {
        (Instruction::Literal{text:a,insensitive:ai},Instruction::Literal{text:b,insensitive:bi})=>a==b&&ai==bi,
        (Instruction::Call(a),Instruction::Call(b))=>a==b,
        (Instruction::Sequence{left:a,right:b},Instruction::Sequence{left:c,right:d})|
        (Instruction::Choice{left:a,right:b},Instruction::Choice{left:c,right:d})=>equal(p,a,c,depth+1)&&equal(p,b,d,depth+1),
        _=>false,
    }
}
pub(crate) fn certify(p:&CompiledGrammar,target:&str,branch:&str,fallback:&str)->Result<SharedWrapperPrefix,String>{
    let rule=|name:&str|p.rule_id(name).ok_or_else(||format!("unknown shared-wrapper-prefix rule '{name}'"));
    let target=rule(target)?;let method=rule(branch)?;let property=rule(fallback)?;
    let t=p.rule(target);let m=p.rule(method);let f=p.rule(property);
    if target==method||target==property||method==property{return Err("shared-wrapper-prefix rules must be distinct".into());}
    if t.mode!=RuleMode::Normal||m.mode!=RuleMode::Normal||f.mode!=RuleMode::Normal||t.scope!=m.scope||m.scope!=f.scope{
        return Err("shared-wrapper-prefix rules require normal mode and the same scope".into());
    }
    let scope=p.scope(m.scope);
    if [scope.whitespace,scope.comment].into_iter().flatten().any(|r|r==method||r==property){
        return Err("shared-wrapper-prefix wrappers cannot be designated trivia rules".into());
    }
    let Instruction::Sequence{left:prefix,right:suffix}=p.instruction(ungroup(p,m.expression)?) else{return Err("shared-wrapper-prefix branch requires prefix ~ suffix".into());};
    if !equal(p,prefix,f.expression,0){return Err("shared-wrapper-prefix fallback must structurally equal the branch prefix".into());}
    // Search only the target's ordered-choice spine. Never descend into rule calls,
    // sequences, predicates or repeats; those have a different matching boundary.
    fn find<P:Program>(p:&P,id:usize,method:usize,property:usize,out:&mut Vec<(usize,usize)>,depth:usize)->Result<(),String>{
        if depth>128{return Err("shared-wrapper-prefix choice nesting is too deep".into());}
        let id=ungroup(p,id)?;
        if let Instruction::Choice{left,right}=p.instruction(id){
            if matches!(p.instruction(ungroup(p,left)?),Instruction::Call(r) if r==method)
                && matches!(p.instruction(ungroup(p,right)?),Instruction::Call(r) if r==property){out.push((id,right));}
            find(p,left,method,property,out,depth+1)?;find(p,right,method,property,out,depth+1)?;
        }
        Ok(())
    }
    let mut choices=Vec::new();find(p,t.expression,method,property,&mut choices,0)?;
    if choices.len()!=1{return Err("shared-wrapper-prefix requires one direct branch | fallback pair in the target choice".into());}
    let (choice,fallback)=choices[0];
    Ok(SharedWrapperPrefix{choice,branch:method,fallback_rule:property,prefix,suffix,fallback})
}
