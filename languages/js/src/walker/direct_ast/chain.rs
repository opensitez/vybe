//! Owned ordinary chain carriers. Captures survive until the complete parent
//! expression is certified, so fallback can reuse argument/index expressions.
use super::*;
use std::rc::Rc;

pub(super) enum Link {
    Member(grameo::Span),
    Index(Box<Expression>, Metadata),
    Call(Vec<Argument>, Rc<Vec<Metadata>>),
}
pub(super) enum LinkUndo {
    Member(grameo::Span, bool),
    Index(Metadata),
    Call(Rc<Vec<Metadata>>),
}
pub(super) struct Step {
    undo: LinkUndo,
    raw: grameo::Span,
    facts: u8,
    before: Span,
}
pub(super) fn property_expected<const S:bool,const B:bool>(ast:&IslandAst<'_,S,B>)->bool {
    ast.arena.values().last().is_some_and(|id|matches!(ast.arena.get(*id),Value::RegionStart(rule) if *rule==Rule::call_chain as usize))
}
pub(super) fn property_leaf<const S:bool,const B:bool>(ast:&IslandAst<'_,S,B>)->bool {
    ast.arena.values().last().is_some_and(|id|matches!(ast.arena.get(*id),Value::RegionStart(rule) if *rule==Rule::call_chain as usize))
}
pub(super) fn publish(link:Link,caches:&mut body::Caches) {
    let mut put=|expr:Expression,meta:Metadata| {
        if let Some(anchor)=meta.anchor {
            caches.expressions.insert((anchor.start,anchor.end),Cached::new_at(expr,meta.raw,anchor,meta.facts,meta.rule));
        }
    };
    match link {
        Link::Member(_)=>{},
        Link::Index(expr,meta)=>put(*expr,meta),
        Link::Call(args,metas)=>for (arg,meta) in args.into_iter().zip(metas.iter()) {put(arg.value,*meta)},
    }
}
pub(super) fn finish_property<const S:bool,const B:bool>(ast:&mut IslandAst<'_,S,B>,raw:grameo::Span) {
    let id=ast.arena.alloc(Value::Property(raw));ast.arena.push(id);
}
pub(super) fn finish_link<const S:bool,const B:bool>(ast:&mut IslandAst<'_,S,B>,raw:grameo::Span,source:&str) {
    let mut ids=Vec::new();
    while let Some(id)=ast.arena.pop() {
        if matches!(ast.arena.get(id),Value::RegionStart(rule) if *rule==Rule::call_chain as usize) {break;}
        ids.push(id);
    }
    ids.reverse();
    let text=source[raw.start..raw.end].trim_start();
    if text.starts_with('.') && !text.starts_with("?.") && ids.len()==1 {
        if let Value::Property(name)=ast.arena.get(ids[0]) {
            let name=*name;
            let id=ids[0];
            ast.arena.edit(id,|value| {
                *value=Value::Link(Link::Member(name),raw,0);EditUndo::Property(name)
            },|value,undo| {
                let EditUndo::Property(name)=undo else {unreachable!()};*value=Value::Property(name)
            });
            ast.arena.push(id);return;
        }
    }
    if text.starts_with('[') && ids.len()==1 && matches!(ast.arena.get(ids[0]),Value::Ready(_,meta) if meta.anchor.is_some()) {
        let id=ast.arena.list_with_metadata(ids,|mut values| {
            let Value::Ready(expr,meta)=values.pop().unwrap() else {unreachable!()};
            (Value::Link(Link::Index(Box::new(expr),meta),raw,meta.facts),Undo::LinkIndex(meta))
        },split_index);
        ast.arena.push(id);return;
    }
    if text.starts_with('(') && ids.iter().all(|id|matches!(ast.arena.get(*id),Value::Ready(_,meta) if meta.anchor.is_some())) {
        let complete=if ids.is_empty() {text.trim()=="()"} else {
            let Value::Ready(_,first)=ast.arena.get(ids[0]) else {unreachable!()};
            let Value::Ready(_,last)=ast.arena.get(*ids.last().unwrap()) else {unreachable!()};
            source[raw.start..first.raw.start].trim()=="("
                && matches!(source[last.boundary..raw.end].trim(),")"|",)")
                && ids.windows(2).all(|pair| {
                    let (Value::Ready(_,a),Value::Ready(_,b))=(ast.arena.get(pair[0]),ast.arena.get(pair[1])) else {unreachable!()};
                    a.boundary<=b.raw.start && source[a.boundary..b.raw.start].trim()=="," 
                })
        };
        if complete {
            let id=ast.arena.list_with_metadata(ids,|values| {
                let mut args=Vec::with_capacity(values.len());let mut metas=Vec::with_capacity(values.len());let mut facts=0;
                for value in values {let Value::Ready(expr,meta)=value else {unreachable!()};facts|=meta.facts;args.push(Argument::positional(expr));metas.push(meta);}
                let metas=Rc::new(metas);
                (Value::Link(Link::Call(args,Rc::clone(&metas)),raw,facts),Undo::LinkArgs(metas))
            },split_arguments);
            ast.arena.push(id);return;
        }
    }
    ast.arena.push(ast.gap.unwrap());
}
fn split_index(value:Value,undo:Undo)->Vec<Value> {
    let (Value::Link(Link::Index(expr,_),_,_),Undo::LinkIndex(meta))=(value,undo) else {unreachable!()};
    vec![Value::Ready(*expr,meta)]
}
fn split_arguments(value:Value,undo:Undo)->Vec<Value> {
    let (Value::Link(Link::Call(args,_),_,_),Undo::LinkArgs(metas))=(value,undo) else {unreachable!()};
    args.into_iter().zip(metas.iter()).map(|(arg,meta)|Value::Ready(arg.value,*meta)).collect()
}
pub(super) fn valid<const S:bool,const B:bool>(ast:&IslandAst<'_,S,B>,ids:&[Id],raw:grameo::Span,source:&str)->bool {
    let Some(Value::Ready(_,first))=ids.first().map(|id|ast.arena.get(*id)) else {return false;};
    if first.raw.start!=raw.start || first.anchor.is_none() || ids.len()<2 {return false;}
    let mut end=first.raw.end;
    for id in &ids[1..] {
        let Value::Link(link,part,_)=ast.arena.get(*id) else {return false;};
        if end>part.start || !source[end..part.start].trim().is_empty() {return false;}
        // These policies change routing or introduce generated counter names;
        // their complete parent stays captured until its inverse is modeled.
        match link {
            Link::Member(name) if matches!(&source[name.start..name.end],"isArray"|"supportedLocalesOf"|"concat")=>return false,
            Link::Index(expr,_) if js_well_known_symbol_alias(expr).is_some()=>return false,
            _=>{},
        }
        end=part.end;
    }
    end<=raw.end && source[end..raw.end].trim().is_empty()
}
pub(super) fn construct(values:Vec<Value>,raw:grameo::Span,source:&str)->(Value,Undo) {
    let mut values=values.into_iter();
    let Value::Ready(mut expr,meta)=values.next().unwrap() else {unreachable!()};
    let original=expr.span;let mut facts=meta.facts;let mut steps=Vec::with_capacity(values.len());
    for value in values {
        let Value::Link(link,part,link_facts)=value else {unreachable!()};
        facts|=link_facts;let before=expr.span;
        let undo=match link {
            Link::Member(name)=>{
                let length=&source[name.start..name.end]=="length";
                expr=if length {Expression::new(ExprKind::Call {callee:Box::new(Expression::ident("__len__")),args:vec![Argument::positional(expr)],optional:false})}
                    else {Expression::new(ExprKind::Member {object:Box::new(expr),field:source[name.start..name.end].to_owned(),null_safe:false})};
                LinkUndo::Member(name,length)
            },
            Link::Index(index,index_meta)=>{
                expr=Expression::new(ExprKind::Index {object:Box::new(expr),index,null_safe:false});LinkUndo::Index(index_meta)
            },
            Link::Call(args,metas)=>{
                // Empty member calls reset the member callee span in the
                // production regex helper. New expressions are not owned yet.
                if args.is_empty() && matches!(expr.kind,ExprKind::Member{..}) {expr.span=Span::default();}
                expr=Expression::new(ExprKind::Call {callee:Box::new(expr),args,optional:false});LinkUndo::Call(metas)
            },
        };
        steps.push(Step {undo,raw:part,facts:link_facts,before});
    }
    expr.span=span(raw);
    (Value::Ready(expr,Metadata {raw,anchor:Some(raw),layer:0,facts,rule:Rule::call_expression,boundary:raw.end}),Undo::Chain(meta,original,steps))
}
pub(super) fn split(value:Value,undo:Undo)->Vec<Value> {
    let (Value::Ready(mut expr,_),Undo::Chain(meta,original,steps))=(value,undo) else {unreachable!()};
    let mut links=Vec::with_capacity(steps.len());
    for step in steps.into_iter().rev() {
        let (previous,link)=match step.undo {
            LinkUndo::Member(name,length)=>{
                let previous=if length {let ExprKind::Call {mut args,..}=expr.kind else {unreachable!()};args.pop().unwrap().value}
                    else {let ExprKind::Member {object,..}=expr.kind else {unreachable!()};*object};
                (previous,Link::Member(name))
            },
            LinkUndo::Index(index_meta)=>{
                let ExprKind::Index {object,index,..}=expr.kind else {unreachable!()};(*object,Link::Index(index,index_meta))
            },
            LinkUndo::Call(metas)=>{
                let ExprKind::Call {callee,args,..}=expr.kind else {unreachable!()};(*callee,Link::Call(args,metas))
            },
        };
        expr=previous;expr.span=step.before;
        links.push(Value::Link(link,step.raw,step.facts));
    }
    links.reverse();expr.span=original;
    let mut values=Vec::with_capacity(links.len()+1);values.push(Value::Ready(expr,meta));values.extend(links);values
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn ordinary_chains_preserve_policy_and_span_semantics() {
        for source in [
            "const x=obj.foo(a,2).bar[3].baz(); const n=obj.length;",
            "function f(x){let a=x[0].value; if(a.foo(2))return this.items.length; return false;}",
            "const x=obj .\n foo(1)[a+2].default; const y=obj.foo(1) < obj.bar(2);",
            "class C extends B {constructor(x){super(x);this.x=x;}m(){return super.foo(this.x);}}",
            "const x=Promise.resolve(1); document.createElement('div');",
            "const x=obj[Symbol.iterator](); Array.isArray(obj[0]); obj.concat(a,b); Intl.NumberFormat.supportedLocalesOf(a);",
            "const x=obj?.foo(a).bar; obj.foo?.(a); const y=new RegExp('a').toString();",
            "const x=true&&false||null??undefined; const y=f(true,null,this);",
            "obj[#bad]; const x=(f())['x']; obj.foo(...xs); obj.foo({a:1}); obj.foo`tag`;",
        ] {
            assert_eq!(format!("{:?}",super::super::super::parse(source)),format!("{:?}",super::super::super::parse_direct_bodies(source)),"{source}");
        }
    }
    #[test]
    fn full_chain_moves_once_recaptures_and_restores_all_carrier_routes() {
        for source in ["obj.length.foo(1)[2].bar()","this.foo(true)[a+2].length","obj[Symbol.iterator]()","obj.foo({a:1})"] {
            let _lines=vybe_ast::line_index::LineIndex::install(source);
            let mut ast=BodyAst::new(source);let mark=ast.checkpoint();
            for _ in 0..3 {
                super::super::super::super::native_grammar::Parser.build_captures(Rule::call_expression,source,&mut ast).unwrap();
                ast.rollback(mark);
            }
            super::super::super::super::native_grammar::Parser.build_captures(Rule::call_expression,source,&mut ast).unwrap();
            let mut caches=ast.into_caches();
            if !source.contains("Symbol") && !source.contains('{') {
                let cache=caches.expressions.get_mut(&(0,source.len())).expect("complete chain owns root");
                let first=cache.take(source);assert_eq!(format!("{first:?}"),format!("{:?}",cache.take(source)));
            }
            assert_eq!(format!("{:?}",super::super::super::parse(source)),format!("{:?}",super::super::super::parse_direct_bodies(source)));
        }
    }
}
