//! Complete owned array regions, including the grammar's explicit elisions.
use super::*;
const SPREAD:u8=128;
pub(super) enum Route {Expression(Metadata),Elision(grameo::Span)}
pub(super) fn publish(expr:Expression,mut meta:Metadata,caches:&mut body::Caches) {
    meta.layer&=!SPREAD;
    if let Some(anchor)=meta.anchor {caches.expressions.insert((anchor.start,anchor.end),Cached::new_at(expr,meta.raw,anchor,meta.facts,meta.rule));}
}
pub(super) fn finish<const S:bool,const B:bool>(ast:&mut IslandAst<'_,S,B>,rule:usize,raw:grameo::Span,source:&str) {
    if rule==Rule::array_elision as usize {
        let id=ast.arena.alloc(Value::ArrayElision(raw));ast.arena.push(id);return;
    }
    let mut ids=Vec::new();while let Some(id)=ast.arena.pop() {
        if matches!(ast.arena.get(id),Value::RegionStart(r) if *r==rule){break;}
        ids.push(id);
    }
    ids.reverse();
    if rule==Rule::array_element as usize {
        if ids.len()!=1 {ast.arena.push(ast.gap.unwrap());return;}
        let id=ids[0];let Value::Ready(_,meta)=ast.arena.get(id) else {ast.arena.push(ast.gap.unwrap());return;};
        let prefix=&source[raw.start..meta.raw.start];
        if meta.anchor.is_none() || !matches!(prefix.trim(),""|"...") || meta.boundary>raw.end || !source[meta.boundary..raw.end].trim().is_empty() {
            ast.arena.push(ast.gap.unwrap());return;
        }
        let spread=prefix.trim()=="...";
        // Reuse the expression slot and metadata's unused high layer bit while
        // it is an array carrier. Arithmetic only sees Ready values. The edit
        // restores the exact original layer and variant on rollback.
        ast.arena.edit(id,|value| {
            let Value::Ready(expr,mut meta)=std::mem::replace(value,Value::Gap) else {unreachable!()};
            let old=meta.layer;if spread {meta.layer|=SPREAD;}
            *value=Value::ArrayItem(expr,meta);EditUndo::ArrayItem(old)
        },|value,undo| {
            let (Value::ArrayItem(expr,mut meta),EditUndo::ArrayItem(layer))=(std::mem::replace(value,Value::Gap),undo) else {unreachable!()};
            meta.layer=layer;*value=Value::Ready(expr,meta);
        });
        ast.arena.push(id);return;
    }
    if !ids.iter().all(|id|matches!(ast.arena.get(*id),Value::ArrayItem(_,_)|Value::ArrayElision(_))) {
        ast.arena.push(ast.gap.unwrap());return;
    }
    let id=ast.arena.list_with_metadata(ids,|values| {
        let mut elements=Vec::with_capacity(values.len());let mut routes=Vec::with_capacity(values.len());let mut facts=0;
        for value in values {
            match value {
                Value::ArrayItem(expr,meta)=>{
                    facts|=meta.facts;
                    elements.push(ArrayElement{key:None,value:expr,spread:meta.layer&SPREAD!=0,by_ref:false});routes.push(Route::Expression(meta));
                },
                Value::ArrayElision(elision)=>{elements.push(js_array_elision_marker());routes.push(Route::Elision(elision));},
                _=>unreachable!(),
            }
        }
        (Value::Ready(Expression::with_span(ExprKind::Array(elements),span(raw)),Metadata {raw,anchor:Some(raw),layer:0,facts,rule:Rule::array_literal,boundary:raw.end}),Undo::Array(routes))
    },split);
    ast.arena.push(id);
}
fn split(value:Value,undo:Undo)->Vec<Value> {
    let (Value::Ready(Expression{kind:ExprKind::Array(elements),..},_),Undo::Array(routes))=(value,undo) else {unreachable!()};
    elements.into_iter().zip(routes).map(|(element,route)|match route {
        Route::Expression(meta)=>Value::ArrayItem(element.value,meta),
        Route::Elision(raw)=>Value::ArrayElision(raw),
    }).collect()
}
#[cfg(test)]mod tests {
    use super::*;
    #[test]fn complete_array_slots_preserve_holes_spreads_and_contextual_children() {
        for source in [
            "const a=[]; const b=[,]; const c=[,,]; const d=[1,,2,];",
            "const x=[a+2,obj.foo(3),true,null,undefined,...xs]; f(x,[this.a,super.b]);",
            "const x=[[1,,],...foo.bar(2),[,3]]; if(x[0])return x.length;",
            "const x=[obj?.foo(),function(){return this.x;},()=>2,{a:1},`tag`];",
            "const x=[Promise.resolve(1),document]; const y=obj[Symbol.iterator]([1]);",
        ] {assert_eq!(format!("{:?}",super::super::super::parse(source)),format!("{:?}",super::super::super::parse_direct_bodies(source)),"{source}");}
    }
    #[test]fn array_buffer_moves_and_inverse_restores_spreads_and_zero_width_elisions() {
        let source="[1,,foo(2),...xs,undefined]";let _lines=vybe_ast::line_index::LineIndex::install(source);
        let mut ast=BodyAst::new(source);let mark=ast.checkpoint();
        for _ in 0..3 {super::super::super::super::native_grammar::Parser.build_captures(Rule::array_literal,source,&mut ast).unwrap();ast.rollback(mark);}
        let tree=super::super::super::super::native_grammar::Parser.build_captures(Rule::array_literal,source,&mut ast).unwrap();
        assert!(tree.capture_count()<10,"complete array owns children");
        let mut caches=ast.into_caches();let cache=caches.expressions.get_mut(&(0,source.len())).expect("owned array root");
        let Cached::Tree{first:Some(expr),..}=cache else {panic!("array stored directly")};
        let ExprKind::Array(elements)=&expr.kind else {unreachable!()};let pointer=elements.as_ptr();
        let first=cache.take(source);let ExprKind::Array(elements)=&first.kind else {unreachable!()};assert_eq!(pointer,elements.as_ptr());
        assert_eq!(format!("{first:?}"),format!("{:?}",cache.take(source)));
    }
}
