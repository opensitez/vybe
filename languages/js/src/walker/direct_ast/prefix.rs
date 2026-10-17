//! Prefix construction over exact owned grammar operands.
use super::*;
#[derive(Clone,Copy)]
enum Op { Unary(UnaryOp), TypeOf, Void }
fn operator(start: usize, source: &str) -> Option<(Op,usize)> {
    let tail=source.get(start..)?;
    let (op,len)=if tail.starts_with('-') && !tail.starts_with("--") && !tail.starts_with("-=") {(Op::Unary(UnaryOp::Neg),1)}
    else if tail.starts_with('+') && !tail.starts_with("++") && !tail.starts_with("+=") {(Op::Unary(UnaryOp::Pos),1)}
    else if tail.starts_with('!') && !tail.starts_with("!=") {(Op::Unary(UnaryOp::Not),1)}
    else if tail.starts_with('~') {(Op::Unary(UnaryOp::BitNot),1)}
    else if tail.starts_with("typeof") {(Op::TypeOf,6)}
    else if tail.starts_with("void") {(Op::Void,4)} else {return None};
    if len>1 && tail.as_bytes().get(len).is_some_and(|b|b.is_ascii_alphanumeric()||*b==b'_') {return None;}
    Some((op,start+len))
}
pub(super) fn selected(start:usize,source:&str)->bool {operator(start,source).is_some()}
pub(super) fn complete<const S:bool,const B:bool>(ast:&IslandAst<'_,S,B>,rule:usize,raw:grameo::Span)->bool {
    rule==Rule::unary as usize && ast.arena.values().last().is_some_and(|id|
        matches!(ast.arena.get(*id),Value::Ready(_,meta) if meta.rule==Rule::unary && meta.anchor==Some(raw)))
}
pub(super) fn finish<const S:bool,const B:bool>(ast:&mut IslandAst<'_,S,B>,raw:grameo::Span,source:&str) {
    let Some((op,operand_start))=operator(raw.start,source) else{return};
    let mut child=None;let mut count=0;
    while let Some(id)=ast.arena.pop() {
        if matches!(ast.arena.get(id),Value::RegionStart(rule) if *rule==Rule::unary as usize){break;}
        child=Some(id);count+=1;
    }
    let Some(child)=child.filter(|id|count==1 && matches!(ast.arena.get(*id),Value::Ready(_,meta)
        if meta.raw.start>=operand_start && meta.raw.end<=raw.end
        && source[operand_start..meta.raw.start].trim().is_empty()
        && source[meta.raw.end..raw.end].trim().is_empty())) else {ast.arena.push(ast.gap.unwrap());return};
    let result=ast.arena.list_with_metadata(vec![child],|mut values| {
        let Value::Ready(mut expr,meta)=values.pop().unwrap() else{unreachable!()};
        let oldspan=expr.span;
        expr.span=span(grameo::Span{start:meta.raw.start,end:raw.end});
        let kind=match op{Op::Unary(op)=>ExprKind::Unary{op,expr:Box::new(expr)},Op::TypeOf=>ExprKind::TypeOf(Box::new(expr)),Op::Void=>ExprKind::Void(Box::new(expr))};
        (Value::Ready(Expression::with_span(kind,span(raw)),Metadata{raw,anchor:Some(raw),layer:0,facts:meta.facts,rule:Rule::unary,boundary:raw.end}),Undo::Prefix(meta,oldspan))
    },|value,undo| {
        let (Value::Ready(expr,_),Undo::Prefix(meta,oldspan))=(value,undo) else{unreachable!()};
        let mut operand=match expr.kind{ExprKind::Unary{expr,..}|ExprKind::TypeOf(expr)|ExprKind::Void(expr)=>*expr,_=>unreachable!()};
        operand.span=oldspan;
        vec![Value::Ready(operand,meta)]
    });
    ast.arena.push(result);
}
