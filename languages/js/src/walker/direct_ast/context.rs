//! Deferred contextual atoms; first lowering retains the live source scope.
use super::*;
pub(in super::super) enum Recipe {
    Deferred(Box<super::recipes::Recipe>),
    Owned(Expression),
    Atom { raw:grameo::Span, span:Span },
    Binary { op:BinOp, left:Box<Recipe>,right:Box<Recipe>,span:Span },
}
impl Recipe {
    pub(super) fn is_binary(&self)->bool {matches!(self,Self::Binary{..})}
    pub(super) fn span(&self)->Span {match self {Self::Deferred(recipe)=>recipe.output_span(),Self::Owned(expr)=>expr.span,Self::Atom{span,..}|Self::Binary{span,..}=>*span}}
    pub(super) fn set_span(&mut self,value:Span) {match self {Self::Deferred(_)=>unreachable!("deferred private root uses byte metadata"),Self::Owned(expr)=>expr.span=value,Self::Atom{span,..}|Self::Binary{span,..}=>*span=value}}
    pub(super) fn resolve(self,walker:&mut JsWalker,root:&Pair<Rule>)->Expression {
        let this=match self{Self::Deferred(recipe)=>return recipe.resolve(walker,root),this=>this};
        let mut atoms=HashMap::new();let mut pending=vec![root.clone()];
        while let Some(pair)=pending.pop() {
            if pair.as_rule()==Rule::unary {let raw=pair.as_span();atoms.insert((raw.start(),raw.end()),pair);}
            else {pending.extend(pair.into_inner());}
        }
        this.resolve_atoms(walker,&mut atoms)
    }
    fn resolve_atoms<'i>(self,walker:&mut JsWalker,atoms:&mut HashMap<(usize,usize),Pair<'i,Rule>>)->Expression {
        match self {
            Self::Deferred(_)=>unreachable!("deferred root resolves independently"),
            Self::Owned(expr)=>expr,
            Self::Atom{raw,span}=>{
                let atom=atoms.remove(&(raw.start,raw.end)).expect("contextual atom retains its complete original capture");
                let kind=walk_expr_kind(walker,collapse_passthrough_expression(atom).expect("complete contextual operand")).expect("certified contextual operand");
                normalize_optional_chain(walker,Expression::with_span(kind,span))
            },
            Self::Binary{op,left,right,span}=>{let expr=normalize_typeof_strict_eq(Expression::new(ExprKind::Binary{op,left:Box::new(left.resolve_atoms(walker,atoms)),right:Box::new(right.resolve_atoms(walker,atoms))}));Expression::with_span(expr.kind,span)},
        }
    }
}
fn metadata(value:&Value)->Option<Metadata> {match value {Value::Ready(_,meta)|Value::Context(_,meta)=>Some(*meta),_=>None}}
fn edit_span(value:&mut Value,new:Span)->Span {match value {Value::Ready(expr,_)=>std::mem::replace(&mut expr.span,new),Value::Context(recipe,_)=>{let old=recipe.span();recipe.set_span(new);old},_=>unreachable!()}}
fn meta_mut(value:&mut Value)->&mut Metadata {match value {Value::Ready(_,meta)|Value::Context(_,meta)=>meta,_=>unreachable!()}}
pub(super) fn boundary<const S:bool,const B:bool>(ast:&mut IslandAst<'_,S,B>,end:usize)->bool {
    let Some(id)=ast.arena.values().last().copied() else{return false};
    let Value::Context(_,meta)=ast.arena.get(id) else{return false};
    if meta.boundary==end && (meta.layer!=0 || matches!(ast.arena.get(id),Value::Context(recipe,_) if same_span(recipe.span(),span(grameo::Span{start:meta.raw.start,end})))) {return true;}
    ast.arena.edit(id,|value| {
        let meta=metadata(value).unwrap();let old=if meta.layer==0 {edit_span(value,span(grameo::Span{start:meta.raw.start,end}))}else {match value{Value::Context(recipe,_)=>recipe.span(),_=>unreachable!()}};
        meta_mut(value).boundary=end;EditUndo::Boundary(meta.boundary,old)
    },|value,undo| {let EditUndo::Boundary(boundary,old)=undo else{unreachable!()};meta_mut(value).boundary=boundary;edit_span(value,old);});
    true
}
pub(super) fn finish<const S:bool,const B:bool>(ast:&mut IslandAst<'_,S,B>,id:Id,raw:grameo::Span,source:&str,owner:Rule)->bool {
    let Value::Context(_,meta)=ast.arena.get(id) else{return false};
    if meta.raw.start!=raw.start || meta.raw.end>raw.end || !trivia(&source[meta.raw.end..raw.end]) {return true;}
    ast.arena.edit(id,|value|{let meta=meta_mut(value);let old=EditUndo::Anchor(meta.anchor.replace(raw));old},|value,undo| {let EditUndo::Anchor(old)=undo else{unreachable!()};meta_mut(value).anchor=old;});
    ast.arena.edit(id,|value|{let meta=meta_mut(value);let old=EditUndo::Rule(meta.rule);meta.rule=owner;old},|value,undo|{let EditUndo::Rule(old)=undo else{unreachable!()};meta_mut(value).rule=old;});
    ast.arena.edit(id,|value|{let old=edit_span(value,span(raw));let meta=meta_mut(value);let undo=EditUndo::Boundary(meta.boundary,old);meta.boundary=raw.end;undo},|value,undo|{let EditUndo::Boundary(boundary,old)=undo else{unreachable!()};meta_mut(value).boundary=boundary;edit_span(value,old);});
    true
}
pub(super) fn atom<const S:bool,const B:bool>(ast:&mut IslandAst<'_,S,B>,raw:grameo::Span,source:&str) {
    if prefix::selected(raw.start,source) {
        prefix::finish(ast,raw,source);
        if prefix::complete(ast,Rule::unary as usize,raw) {return;}
        // Incomplete prefix leaves its children live for the contextual walker.
        if ast.arena.values().last().is_some_and(|id|matches!(ast.arena.get(*id),Value::Gap)){ast.arena.pop();}
    } else {
        let values=ast.arena.values();
        if values.len()>=2 && matches!(ast.arena.get(values[values.len()-2]),Value::RegionStart(r) if *r==Rule::unary as usize)
            && matches!(ast.arena.get(*values.last().unwrap()),Value::Ready(_,meta) if meta.raw.start==raw.start && meta.raw.end<=raw.end && trivia(&source[meta.raw.end..raw.end])) {
            ast.arena.remove_value_below_top();
            // The ordinary child's anchor remains available if a contextual
            // enclosing atom needs its pruned capture. Pratt's boundary/owner
            // already certify the outer operand span, so no unary metadata edit
            // or extra capture suppression is required for this identity.
            return;
        }
        while let Some(id)=ast.arena.pop() {if matches!(ast.arena.get(id),Value::RegionStart(r) if *r==Rule::unary as usize){break;}}
    }
    let id=ast.arena.alloc(Value::Context(Box::new(Recipe::Atom{raw,span:span(raw)}),Metadata{raw,anchor:Some(raw),layer:0,facts:0,rule:Rule::unary,boundary:raw.end}));ast.arena.push(id);
}
fn trivia(mut text:&str)->bool {
    loop {
        text=text.trim_start();if text.is_empty(){return true;}
        if let Some(rest)=text.strip_prefix("/*") {let Some(end)=rest.find("*/") else{return false};text=&rest[end+2..];}
        else if let Some(rest)=text.strip_prefix("//") {text=rest.find(['\r','\n']).map_or("",|end|&rest[end..]);}
        else{return false;}
    }
}
fn recipe(value:Value)->(Box<Recipe>,Metadata,bool) {match value {Value::Ready(expr,meta)=>(Box::new(Recipe::Owned(expr)),meta,false),Value::Context(recipe,meta)=>(recipe,meta,true),_=>unreachable!()}}
pub(super) fn reduce<const S:bool,const B:bool>(ast:&mut IslandAst<'_,S,B>,op:BinOp,layer:u8,operator:grameo::Span)->bool {
    let values=ast.arena.values();if values.len()<2{return false;}
    let (left,right)=(values[values.len()-2],values[values.len()-1]);
    let (Some(lm),Some(rm))=(metadata(ast.arena.get(left)),metadata(ast.arena.get(right))) else{return false};
    if !matches!(ast.arena.get(left),Value::Context(_,_)) && !matches!(ast.arena.get(right),Value::Context(_,_)) && !(op==BinOp::StrictEq && [left,right].iter().any(|id|matches!(ast.arena.get(*id),Value::Ready(expr,_) if matches!(expr.kind,ExprKind::TypeOf(_))))){return false;}
    if lm.raw.end>operator.start || rm.raw.start<operator.end || !trivia(&ast.source[lm.raw.end..operator.start]) || !trivia(&ast.source[operator.end..rm.raw.start]) {return true;}
    ast.arena.pop();ast.arena.pop();
    let joined=grameo::Span{start:lm.raw.start,end:rm.raw.end};
    let id=ast.arena.binary_with_metadata(left,right,|left,right| {
        let (mut left,mut lm,lc)=recipe(left);let (mut right,mut rm,rc)=recipe(right);
        let (ls,rs)=(left.span(),right.span());
        if lm.layer>layer {left.set_span(span(grameo::Span{start:lm.raw.start,end:lm.boundary}));}
        if rm.layer>layer || op==BinOp::Pow && rm.layer==layer {right.set_span(span(grameo::Span{start:rm.raw.start,end:rm.boundary}));}
        let meta=Metadata{raw:joined,anchor:None,layer,facts:lm.facts|rm.facts,rule:Rule::additive,boundary:rm.boundary};
        if lc {lm.layer|=128;}if rc {rm.layer|=128;}
        (Value::Context(Box::new(Recipe::Binary{op,left,right,span:Span::default()}),meta),Undo::ContextBinary(lm,rm,ls,rs))
    },|value,undo| {
        let (Value::Context(recipe,_),Undo::ContextBinary(mut lm,mut rm,ls,rs))=(value,undo) else{unreachable!()};
        let Recipe::Binary{mut left,mut right,..}=*recipe else{unreachable!()};left.set_span(ls);right.set_span(rs);
        let (lc,rc)=(lm.layer&128!=0,rm.layer&128!=0);lm.layer&=127;rm.layer&=127;
        let restore=|recipe:Box<Recipe>,meta,context|if context {Value::Context(recipe,meta)}else {let Recipe::Owned(expr)=*recipe else{unreachable!()};Value::Ready(expr,meta)};
        (restore(left,lm,lc),restore(right,rm,rc))
    });ast.arena.push(id);true
}

#[cfg(test)]mod tests {
 use super::*;
 #[test]fn contextual_atoms_keep_live_scope_normalization_and_newlines(){
  for source in [
   "function f(a){return a+g({x:2});}",
   "let x=document+f({a:Promise});",
   "function f(a){return g(a)+\n h(a)*2;}",
   "let x=obj?.f()+new Foo(2);",
   "let x=typeof f({a:2})===\"number\";",
   "function f(){let Promise=2;return -Promise+g({});}",
   "let x=(f({})+2)*g([1]);",
   "let x=f({})/*before*/+/*after*/2;",
   "let x=-await f({})+2;",
  ] {assert_eq!(format!("{:?}",super::super::super::parse(source)),format!("{:?}",super::super::super::parse_direct_bodies(source)),"{source}");}
 }
 #[test]fn captured_roots_lower_first_use_and_replay_in_original_scope(){
  let source="f({a:2})+g([1])*2";let _lines=vybe_ast::line_index::LineIndex::install(source);
  let mut builder=BodyAst::new(source);let tree=super::super::super::super::native_grammar::Parser.build_captures(Rule::additive,source,&mut builder).unwrap();
  let caches=builder.into_caches();assert!(matches!(caches.expressions.get(&(0,source.len())),Some(Cached::Context{..})),"context root");
  let pair=tree.into_typed_pairs::<Rule>().next().unwrap();
  let mut walker=JsWalker::default();walker.direct_input=Some((source.as_ptr() as usize,source.len()));walker.direct_expressions=caches.expressions;
  let expected=walk_expression(&mut JsWalker::default(),JsParser::parse(Rule::additive,source).unwrap().next().unwrap()).unwrap();
  let first=take_for_pair(&mut walker,&pair).unwrap();let replay=take_for_pair(&mut walker,&pair).unwrap();
  assert_eq!(format!("{expected:?}"),format!("{first:?}"));assert_eq!(format!("{first:?}"),format!("{replay:?}"));
 }
 #[test]fn contextual_and_owned_inverses_restore_all_small_limits(){
  for source in ["f({})+g(2)","-f({})+\n g(2)","f({})+g(","(a+f({}))*2"] {for depth in 0..24 {for work in 0..256 {
   let mut builder=BodyAst::new(source);builder.retain_test_journal=true;let mark=builder.checkpoint();
   let result=grameo::engine::build_stackfree_compact_capture_program(&super::super::super::super::native_grammar::Parser,"expression",source,grameo::MatchOptions{max_rule_depth:depth,max_steps:work},&mut builder);
   if result.is_err(){assert_eq!(builder.checkpoint(),mark,"{source} depth={depth} work={work}");assert!(builder.arena.values().is_empty());assert!(!builder.active());assert!(builder.into_caches().expressions.is_empty());}
  }}}
 }
}

#[cfg(test)] mod whole_root_tests {
 use super::*;
 #[test]fn whole_root_owned_and_contextual_replays_keep_scope_and_exact_spans(){
  for source in ["function f(document){return document < g({}) && Promise > h({});}","const x=typeof f({})===\"number\"||g({})>2;","const x=2+f({})<g({})===3&&h({})??z;","const x=obj?.f({})==null||new F({})>2;","const x=(f({})<g({}))&&h({});","const x=f({})<\n g({})&&h({});","const x=a /*lhs*/ < /*rhs*/ b || c;","const x=a|b^c&d;","const x=4<<1>3===true;"] {
   assert_eq!(format!("{:?}",crate::walker::parse(source)),format!("{:?}",crate::walker::parse_direct_bodies(source)),"{source}");
  }
  for source in ["f({}) < g({}) === 2","f({}) < 2 && g({}) > 3","typeof f({}) === \"number\" || g({})>2"] {
   let _lines=vybe_ast::line_index::LineIndex::install(source);let mut builder=BodyAst::new(source);
   let tree=crate::native_grammar::Parser.build_captures(Rule::logical_expr,source,&mut builder).unwrap();
   let caches=builder.into_caches();assert!(matches!(caches.expressions.get(&(0,source.len())),Some(Cached::Context{rule:Rule::logical_expr,..})));
   let pair=tree.into_typed_pairs::<Rule>().next().unwrap();let mut walker=JsWalker::default();walker.direct_input=Some((source.as_ptr() as usize,source.len()));walker.direct_expressions=caches.expressions;
   let expected=walk_expression(&mut JsWalker::default(),JsParser::parse(Rule::logical_expr,source).unwrap().next().unwrap()).unwrap();
   for _ in 0..3{let actual=take_for_pair(&mut walker,&pair).unwrap();assert_eq!(format!("{expected:?}"),format!("{actual:?}"));}
  }
 }
 #[test]fn whole_root_inverses_restore_budget_failures(){
  for source in ["f({})<g({})&&h({})", "f({})<g({})===2", "f({})<g({})&&h(","typeof f({})===\"number\"||z", "a+2*b<c**2&&d", "f({})<\n g({})&&h({})"]{for depth in 0..24{for work in 0..256{
   let mut builder=BodyAst::new(source);builder.retain_test_journal=true;let mark=builder.checkpoint();
   let result=grameo::engine::build_stackfree_compact_capture_program(&crate::native_grammar::Parser,"expression",source,grameo::MatchOptions{max_rule_depth:depth,max_steps:work},&mut builder);
   if result.is_err(){assert_eq!(builder.checkpoint(),mark,"{source} depth={depth} work={work}");assert!(builder.arena.values().is_empty());assert!(!builder.active());assert!(builder.into_caches().expressions.is_empty());}
  }}}
 }
}
