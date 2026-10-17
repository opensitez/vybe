//! Matching owns immutable syntax recipes; only live-scope consumption builds AST.
use super::*;
use grameo::arena::EditArena;

#[derive(Clone, Copy)]
struct Meta { raw: grameo::Span, anchor: Option<grameo::Span>, layer: u8, boundary: usize }
enum Node { Scope(usize), Owner(bool), Unary, Postfix, Gap, Atom(Meta), Expr(Box<Recipe>, Meta) }
impl Node {
    fn meta(&self) -> Option<&Meta> { match self { Self::Atom(meta) | Self::Expr(_, meta) => Some(meta), _ => None } }
    fn into_recipe(self) -> (Box<Recipe>, Meta, bool) {
        match self {
            Self::Atom(meta) => (Box::new(Recipe::Atom { raw: meta.raw, span: Some(grameo::Span { start: meta.raw.start, end: meta.boundary }) }), meta, true),
            Self::Expr(recipe, meta) => (recipe, meta, false),
            _ => unreachable!(),
        }
    }
}
struct Inverse { left: Meta, right: Meta, left_span: Option<grameo::Span>, right_span: Option<grameo::Span>, left_inline: bool, right_inline: bool }
enum Edit { Scope(usize), Expr(Meta, Option<grameo::Span>), Postfix(Meta, Option<grameo::Span>) }
pub(in super::super) struct RecipeAst<'i> {
    source: &'i str, arena: EditArena<Node, Inverse, Edit>,
    scope: Id, owner: Id, nested: Id, unary: Id, postfix: Id, gap: Id,
    #[cfg(test)] retain_journal: bool,
}
impl<'i> RecipeAst<'i> {
    pub(in super::super) fn new(source: &'i str) -> Self {
        let mut arena = EditArena::default();
        let scope = arena.alloc(Node::Scope(0));
        let owner = arena.alloc(Node::Owner(true));
        let nested = arena.alloc(Node::Owner(false));
        let unary = arena.alloc(Node::Unary);
        let postfix = arena.alloc(Node::Postfix);
        let gap = arena.alloc(Node::Gap);
        arena.commit();
        Self { source, arena, scope, owner, nested, unary, postfix, gap, #[cfg(test)] retain_journal: false }
    }
    fn active(&self) -> bool { matches!(self.arena.get(self.scope), Node::Scope(n) if *n != 0) }
    fn activity(&mut self, enabled: bool) {
        self.arena.edit(self.scope, |node| {
            let Node::Scope(n) = node else { unreachable!() };
            let old = Edit::Scope(*n); *n = usize::from(enabled); old
        }, |node, undo| {
            let (Node::Scope(n), Edit::Scope(old)) = (node, undo) else { unreachable!() }; *n = old;
        });
    }
    fn update(&mut self, id: Id, anchor: Option<grameo::Span>, end: usize) {
        self.arena.edit(id, |node| {
            let (meta, recipe) = match node { Node::Atom(meta) => (meta, None), Node::Expr(recipe, meta) => (meta, Some(recipe)), _ => unreachable!() };
            let old = Edit::Expr(*meta, recipe.as_ref().and_then(|recipe| recipe.span()));
            if let Some(raw) = anchor { meta.anchor = Some(raw); if let Some(recipe) = recipe { recipe.set_span(Some(raw)); } }
            else if meta.layer == 0 { if let Some(recipe) = recipe { recipe.set_span(Some(grameo::Span { start: meta.raw.start, end })); } }
            meta.boundary = end; old
        }, |node, undo| {
            let Edit::Expr(old, old_span) = undo else { unreachable!() };
            match node { Node::Atom(meta) => *meta = old, Node::Expr(recipe, meta) => { *meta = old; recipe.set_span(old_span); }, _ => unreachable!() }
        });
    }
    pub(in super::super) fn finish(self) -> HashMap<(Rule, usize, usize), Cached> {
        let mut caches = HashMap::new();
        for node in self.arena.into_live_nodes() {
            if let Node::Expr(recipe, Meta { anchor: Some(raw), .. }) = node {
                // Single operands are consumed through transparent captures, never this cache.
                if matches!(&*recipe, Recipe::Atom { .. }) { continue; }
                let rule=if matches!(&*recipe, Recipe::Postfix(_)) {
                    record_published(); Rule::postfix
                } else {Rule::bitor_expr};
                caches.insert((rule, raw.start, raw.end), Cached {
                    first: None, replay: Vec::new(), sequence: false, contextual: true,
                    deferred: Some(std::sync::Arc::from(recipe)),
                });
            }
        }
        caches
    }
}
impl Builder for RecipeAst<'_> {
    fn supports_shared_inner_prefix(&self,r:usize)->bool{(r==Rule::list_inner as usize || r==Rule::subscript_item as usize) && inner_enabled()}
    fn supports_shared_group_prefix(&self,a:usize,b:usize)->bool{a==Rule::comp_for_arg as usize && b==Rule::expression_list as usize && grouping_enabled()}
    fn supports_shared_prefix(&self, rule: usize) -> bool { rule == Rule::comp_for_arg as usize }
    fn observes_rule(&self, rule: usize) -> bool {
        rule == Rule::postfix as usize || rule == Rule::pratt_scalar_atom as usize || rule == Rule::bitor_expr as usize || rule == Rule::unary as usize && self.active()
    }
    fn supports_pratt_at(&self, rule: usize, start: usize, source: &str) -> bool { rule == Rule::bitor_expr as usize || rule == Rule::postfix as usize && postfix_enabled() && members::candidate(start,source) }
    fn suppress_pratt_captures(&self, rule:usize)->bool {rule==Rule::postfix as usize}
    fn opaque_rule(&self, rule: usize) -> bool { rule != Rule::postfix as usize && rule != Rule::pratt_scalar_atom as usize && rule != Rule::bitor_expr as usize && rule != Rule::unary as usize }
    fn checkpoint(&self) -> usize { self.arena.checkpoint() }
    fn rollback(&mut self, mark: usize) { self.arena.rollback(mark); }
    fn begin(&mut self, rule: usize, start: usize) {
        if rule==Rule::postfix as usize {if postfix_enabled() && members::candidate(start,self.source) {self.arena.push(self.postfix);} return;}
        if rule==Rule::pratt_scalar_atom as usize {return;}
        if rule == Rule::bitor_expr as usize {
            let activates = !self.active(); if activates { self.activity(true); }
            self.arena.push(if activates { self.owner } else { self.nested });
        } else if rule == Rule::unary as usize { self.arena.push(self.unary); }
    }
    fn finish(&mut self, rule: usize, raw: grameo::Span, _: &str) {
        if rule==Rule::pratt_scalar_atom as usize {
            if self.arena.values().last()==Some(&self.postfix) {
                let id=self.arena.alloc(Node::Expr(Box::new(Recipe::Postfix(Box::new(PostfixRecipe{head:raw,head_span:Some(raw),steps:Vec::new(),span:Some(raw)}))),Meta{raw,anchor:None,layer:0,boundary:raw.end}));
                self.arena.push(id);
            }
            return;
        }
        if rule==Rule::postfix as usize {
            if !postfix_enabled() || !members::candidate(raw.start,self.source) {return;}
            let(mut first,mut count)=(None,0);
            while let Some(id)=self.arena.pop(){if matches!(self.arena.get(id),Node::Postfix){break;}first=Some(id);count+=1;}
            let Some(id)=first.filter(|_|count==1) else {self.arena.push(self.gap);return;};
            if !matches!(self.arena.get(id),Node::Expr(recipe,meta) if matches!(&**recipe,Recipe::Postfix(p) if !p.steps.is_empty()) && meta.raw.start==raw.start && meta.raw.end<=raw.end && self.source[meta.raw.end..raw.end].trim().is_empty()) {self.arena.push(self.gap);return;}
            self.update(id,Some(raw),raw.end);self.arena.push(id);return;
        }
        if rule == Rule::unary as usize {
            while let Some(id) = self.arena.pop() { if matches!(self.arena.get(id), Node::Unary) { break; } }
            let id = self.arena.alloc(Node::Atom(Meta { raw, anchor: None, layer: 0, boundary: raw.end }));
            self.arena.push(id); return;
        }
        let (mut first, mut count, mut activates) = (None, 0, false);
        while let Some(id) = self.arena.pop() {
            if let Node::Owner(enabled) = self.arena.get(id) { activates = *enabled; break; }
            first = Some(id); count += 1;
        }
        if activates { self.activity(false); }
        let Some(id) = first.filter(|_| count == 1) else { self.arena.push(self.gap); return; };
        if !self.arena.get(id).meta().is_some_and(|meta| meta.raw.start == raw.start && meta.raw.end <= raw.end && self.source[meta.raw.end..raw.end].trim().is_empty()) {
            self.arena.push(self.gap); return;
        }
        self.update(id, Some(raw), raw.end); self.arena.push(id);
    }
    fn accepts_pratt(&self, _: usize, raw: grameo::Span) -> bool {
        self.arena.values().last().is_some_and(|id| self.arena.get(*id).meta().is_some_and(|meta| meta.raw.start == raw.start && meta.raw.end <= raw.end && self.source[meta.raw.end..raw.end].trim().is_empty()))
    }
    fn try_pratt_boundary(&mut self, _: usize, end: usize) -> Result<(), grameo::builder::BuildError> {
        if let Some(id) = self.arena.values().last().copied().filter(|id| self.arena.get(*id).meta().is_some()) { self.update(id, None, end); }
        Ok(())
    }
    fn reduce(&mut self, _: usize, fixity: grameo::pratt::Fixity, raw_op: grameo::Span) {
        if fixity==grameo::pratt::Fixity::Postfix {
            let Some(id)=self.arena.values().last().copied() else {return;};
            if !matches!(self.arena.get(id),Node::Expr(recipe,meta) if matches!(&**recipe,Recipe::Postfix(_)) && meta.raw.end<=raw_op.start) {return;}
            self.arena.edit(id,|node|{
                let Node::Expr(recipe,meta)=node else {unreachable!()};let Recipe::Postfix(p)=&mut **recipe else {unreachable!()};
                let undo=Edit::Postfix(*meta,p.span);p.steps.push(raw_op);let raw=grameo::Span{start:p.head.start,end:raw_op.end};p.span=Some(raw);*meta=Meta{raw,anchor:None,layer:7,boundary:raw.end};undo
            },|node,undo|{
                let(Node::Expr(recipe,meta),Edit::Postfix(old,old_span))=(node,undo) else {unreachable!()};let Recipe::Postfix(p)=&mut **recipe else {unreachable!()};p.steps.pop();*meta=old;p.span=old_span;
            });return;
        }
        if fixity != grameo::pratt::Fixity::InfixLeft { return; }
        let Some((op, layer)) = arithmetic::operator(&self.source[raw_op.start..raw_op.end]) else { return; };
        let values = self.arena.values(); if values.len() < 2 { return; }
        let (left, right) = (values[values.len()-2], values[values.len()-1]);
        let (Some(lm), Some(rm)) = (self.arena.get(left).meta(), self.arena.get(right).meta()) else { return; };
        if lm.raw.end > raw_op.start || rm.raw.start < raw_op.end { return; }
        self.arena.pop(); self.arena.pop();
        let id = self.arena.binary_with_metadata(left, right, |left, right| {
            let (mut left, lm, left_inline) = left.into_recipe();
            let (mut right, rm, right_inline) = right.into_recipe();
            let inverse = Inverse { left: lm, right: rm, left_span: left.span(), right_span: right.span(), left_inline, right_inline };
            if lm.layer == layer { left.set_span(None); }
            if rm.layer == layer { right.set_span(None); }
            let raw = grameo::Span { start: lm.raw.start, end: rm.boundary.max(rm.raw.end) };
            (Node::Expr(Box::new(Recipe::Binary { op, left, right, span: Some(raw) }),
                Meta { raw, anchor: None, layer, boundary: raw.end }), inverse)
        }, |node, undo| {
            let Node::Expr(recipe, _) = node else { unreachable!() };
            let Recipe::Binary { mut left, mut right, .. } = *recipe else { unreachable!() };
            left.set_span(undo.left_span); right.set_span(undo.right_span);
            (if undo.left_inline { Node::Atom(undo.left) } else { Node::Expr(left, undo.left) },
                if undo.right_inline { Node::Atom(undo.right) } else { Node::Expr(right, undo.right) })
        });
        self.arena.push(id);
    }
    fn committed(&mut self, _: &str, _: usize) {
        #[cfg(test)] if self.retain_journal { return; }
        self.arena.commit();
    }
}
pub(super) struct PostfixRecipe {head:grameo::Span,head_span:Option<grameo::Span>,steps:Vec<grameo::Span>,span:Option<grameo::Span>}
pub(super) enum Recipe {
    Postfix(Box<PostfixRecipe>),
    Atom { raw: grameo::Span, span: Option<grameo::Span> },
    Binary { op: BinOp, left: Box<Recipe>, right: Box<Recipe>, span: Option<grameo::Span> },
}
impl Recipe {
    fn span(&self) -> Option<grameo::Span> { match self { Self::Atom { span, .. } | Self::Binary { span, .. } => *span,Self::Postfix(p)=>p.span } }
    fn set_span(&mut self,new:Option<grameo::Span>){match self{Self::Atom{span,..}|Self::Binary{span,..}=>*span=new,Self::Postfix(p)=>{p.span=new;if p.steps.is_empty(){p.head_span=new;}}}}
    pub(super) fn resolve(&self, walker: &mut PyWalker, root: &Pair<Rule>) -> Result<Expression, String> {
        if let Self::Postfix(p)=self {record_consumed();return p.resolve(walker,root);}
        let mut atoms = HashMap::new(); let mut pending = vec![root.clone()];
        while let Some(pair) = pending.pop() {
            if pair.as_rule() == Rule::unary { let raw = pair.as_span(); atoms.insert((raw.start(), raw.end()), pair); }
            else { pending.extend(pair.into_inner()); }
        }
        let raw = root.as_span(); atoms.entry((raw.start(), raw.end())).or_insert_with(|| root.clone());
        self.resolve_atoms(walker, &mut atoms)
    }
    fn resolve_atoms<'i>(&self, walker: &mut PyWalker, atoms: &mut HashMap<(usize, usize), Pair<'i, Rule>>) -> Result<Expression, String> {
        let mut expression = match self {
            Self::Postfix(_)=>unreachable!("owned postfix consumed at original boundary"),
            Self::Atom { raw, .. } => {
                let pair = atoms.remove(&(raw.start, raw.end)).ok_or_else(|| "deferred operand lacks original capture".to_owned())?;
                walk_expression(walker, pair)?
            }
            Self::Binary { op, left, right, .. } => {
                let left = left.resolve_atoms(walker, atoms)?;
                let right = right.resolve_atoms(walker, atoms)?;
                match op {
                    BinOp::Add => lower_python_owned_additive(walker, "+", left, right),
                    BinOp::Sub => lower_python_owned_additive(walker, "-", left, right),
                    BinOp::Mul => lower_python_owned_multiplicative(walker, "*", left, right),
                    BinOp::Div => lower_python_owned_multiplicative(walker, "/", left, right),
                    BinOp::IDiv => lower_python_owned_multiplicative(walker, "//", left, right),
                    BinOp::Mod => lower_python_owned_multiplicative(walker, "%", left, right),
                    BinOp::MatMul => lower_python_owned_multiplicative(walker, "@", left, right),
                    BinOp::BitOr | BinOp::BitXor | BinOp::BitAnd => lower_python_owned_bitwise(walker, *op, left, right),
                    BinOp::Shl | BinOp::Shr => Expression::new(ExprKind::Binary { op: *op, left: Box::new(left), right: Box::new(right) }),
                    _ => unreachable!("certified arithmetic operator"),
                }
            }
        };
        expression.span = self.span().map_or_else(Span::default, super::span);
        Ok(expression)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn snapshot(module: Result<Module, String>) -> String { format!("{module:?}") }

    #[test]
    fn shared_call_prefix_keeps_rich_arguments_scopes_and_errors() {
        for source in [
            "x=f(g(h(1)),name=obj.m(2),*args,**kwargs)\n",
            "x=sum(f(i) for i in values if g(i))\n",
            "x=f((i for i in values),name=g(2))\n",
            "def outer(x):\n    def inner(y=f(x)):\n        nonlocal x\n        x=g(y)\n        return x\n    return inner\n",
            "class C:\n    __slots__=('x',)\n    def f(self,y=g(2)): return self.x+h(y)\n",
            "x=f(a+)\n", "x=f(x for)\n", "x=f(name=)\n",
        ] {
            assert_eq!(snapshot(crate::parse_with_peg(source)), snapshot(crate::parse_with_direct_ast(source)), "{source}");
        }
        for depth in [8,12] {
            let source=format!("x={}1{}\n", "f(".repeat(depth), ")".repeat(depth));
            assert_eq!(snapshot(crate::parse_with_peg(&source)), snapshot(crate::parse_with_direct_ast(&source)), "depth {depth}");
        }
    }
    #[test]
    fn transparent_atoms_do_not_publish_and_rollback_preserves_inline_boundaries() {
        assert_eq!(std::mem::size_of::<Recipe>(), 48);
        assert_eq!(std::mem::size_of::<Node>(), 64);
        for source in ["x", "f(g(x))", "obj.method(x)[y]", "f'{g(x)}'", "(x)"] {
            let mut builder = RecipeAst::new(source);
            crate::native_grammar::Parser.build_captures(Rule::expression, source, &mut builder).unwrap();
            assert!(builder.finish().keys().all(|(rule,_,_)|*rule==Rule::postfix), "{source}");
        }
        for source in ["x + y", "f(x) + obj[y]", "(x+y)*z"] {
            let mut builder = RecipeAst::new(source);
            crate::native_grammar::Parser.build_captures(Rule::expression, source, &mut builder).unwrap();
            assert!(!builder.finish().is_empty(), "{source}");
        }
        let source = "x + y";
        let mut builder = RecipeAst::new(source);
        builder.retain_journal = true;
        builder.begin(Rule::bitor_expr as usize, 0);
        builder.begin(Rule::unary as usize, 0);
        Builder::finish(&mut builder, Rule::unary as usize, grameo::Span { start: 0, end: 1 }, source);
        builder.try_pratt_boundary(Rule::bitor_expr as usize, 2).unwrap();
        let mark = builder.checkpoint();
        let left = *builder.arena.values().last().unwrap();
        builder.begin(Rule::unary as usize, 4);
        Builder::finish(&mut builder, Rule::unary as usize, grameo::Span { start: 4, end: 5 }, source);
        builder.reduce(Rule::bitor_expr as usize, grameo::pratt::Fixity::InfixLeft, grameo::Span { start: 2, end: 3 });
        builder.rollback(mark);
        assert!(matches!(builder.arena.get(left), Node::Atom(meta) if meta.boundary == 2 && meta.raw.end == 1));
    }
    #[test]
    fn rich_captured_operands_keep_order_scopes_precedence_and_spans() {
        for source in [
            "x=f(1+2)*obj.m(a)[b]+g(c)\n",
            "x=(f(a)+b)*(c+d)\ny=a**b**c+d\nz=-a**b+c\n",
            "x=a|b^c&d<<e+f*g//h\n",
            "x=[f(i)+i for i in values if g(i)]\ny=sum(x*x for x in values)\n",
            "x=2\ndef f(a=g(x+1),*,b=2):\n    global x\n    x=a+b\n    return obj.m(x)[0]+b\n",
            "def outer(x):\n    def f(y=x+1):\n        nonlocal x\n        x=x+y\n        return x\n    return f\n",
            "@decorate(factory(1+2))\nclass C[T](Base,metaclass=Meta,k=f(2+3)):\n    @property\n    def value(self): return self.x+f(self.y)\n",
            "async def f(x):\n    return await g(x+1)+2\n",
            "def f(x):\n    yield g(x)+1\n    yield from values\n",
            "x=f'{f(a+1):>{g(b)}}'\ny=b'abc'*3\n",
            "x=(a +\n b) * c # trivia\ny=a+\\\n b\n",
            "x=f(a+)\n", "x=a***b\n", "class C(\n",
        ] {
            assert_eq!(snapshot(crate::parse_with_peg(source)), snapshot(crate::parse_with_direct_ast(source)), "{source}");
        }
    }
    #[test]
    fn recipe_rollback_preserves_scope_frames_for_nested_budget_failures() {
        for source in ["f(g(1+2), [x+y for x in xs])*obj.m(a)[b]", "f'{g(a+1)}'+x", "f(g(1+),obj.x)"] {
            for depth in 0..16 { for work in 0..256 {
                let mut builder = RecipeAst::new(source); builder.retain_journal = true;
                let mark = builder.checkpoint();
                let result = grameo::engine::build_stackfree_compact_capture_program(
                    &crate::native_grammar::Parser, "expression", source,
                    grameo::MatchOptions { max_rule_depth: depth, max_steps: work }, &mut builder);
                if result.is_err() {
                    assert_eq!(builder.checkpoint(), mark, "{source} depth={depth} work={work}");
                    assert!(builder.arena.values().is_empty()); assert!(!builder.active());
                    assert!(builder.finish().is_empty());
                }
            }}
        }
    }
    #[test]
    fn successful_nested_recipe_ownership_restores_repeatedly() {
        for source in ["a+b*c|g(x)", "f(g(1+2),[x+y for x in xs])", "f'{g(x+1)}'+z"] {
            let mut builder = RecipeAst::new(source); builder.retain_journal = true;
            let mark = builder.checkpoint();
            for _ in 0..3 {
                crate::native_grammar::Parser.build_captures(Rule::expression, source, &mut builder).unwrap();
                builder.rollback(mark); assert_eq!(builder.checkpoint(), mark);
                assert!(builder.arena.values().is_empty()); assert!(!builder.active());
            }
        }
    }
}

impl PostfixRecipe {
    fn resolve(&self,walker:&mut PyWalker,root:&Pair<Rule>)->Result<Expression,String>{
        let source=root.as_str();let base=root.as_span().start();let text=&source[self.head.start-base..self.head.end-base];
        let rule=match text {"True"=>Rule::true_kw,"False"=>Rule::false_kw,"None"=>Rule::none_kw,_=>Rule::identifier};
        let kind=leaf(rule,text).ok_or("owned postfix head")?;
        let mut expression=Expression::with_span(kind,self.head_span.map_or_else(Span::default,super::span));
        // Retained steps already describe the raw call chain. Lower them once at
        // this live-scope boundary instead of building and dismantling a raw AST.
        if self.steps.iter().any(|raw| source[raw.start-base..raw.end-base].trim().starts_with('(')) {
            for raw in self.steps.iter().copied() {
                let text=source[raw.start-base..raw.end-base].trim();
                if let Some(field)=text.strip_prefix('.') {
                    expression=lower_python_owned_member(walker,expression,field.trim().to_owned());
                } else {
                    let mut args=members::scalar_call_args(source,raw,base).ok_or("owned scalar arguments")?;
                    for arg in &mut args {
                        let value=std::mem::replace(&mut arg.value,Expression::null());
                        arg.value=normalize_contextual_expression(walker,value);
                    }
                    expression=if args.is_empty(){lower_python_owned_empty_call(walker,expression)}else{lower_python_owned_call_args(walker,expression,args)};
                }
            }
            let mut expression=desugar_member_reads(walker,expression);
            expression.span=self.span.map_or_else(Span::default,super::span);
            return Ok(expression);
        }
        for (i,raw) in self.steps.iter().copied().enumerate(){
            if i!=0 {expression.span=Span::default();}
            let text=source[raw.start-base..raw.end-base].trim();
            let kind=if let Some(field)=text.strip_prefix('.') {ExprKind::Member{object:Box::new(expression),field:field.trim().to_owned(),null_safe:false}}
            else {let args=members::scalar_call_args(source,raw,base).ok_or("owned scalar arguments")?;ExprKind::Call{callee:Box::new(expression),args,optional:false}};
            expression=Expression::with_span(kind,super::span(grameo::Span{start:self.head.start,end:raw.end}));
        }
        expression.span=self.span.map_or_else(Span::default,super::span);
        Ok(members::normalize_owned_calls(walker,expression))
    }
}
#[cfg(test)] thread_local! {
    static POSTFIX_ENABLED:std::cell::Cell<bool>=const {std::cell::Cell::new(true)};
    static PUBLIC_POSTFIX_STATS:std::cell::Cell<(usize,usize,usize)>=const {std::cell::Cell::new((0,0,0))};
}
fn postfix_enabled()->bool {#[cfg(test)] {return POSTFIX_ENABLED.with(|v|v.get());} #[cfg(not(test))] {true}}
fn record_published(){#[cfg(test)] PUBLIC_POSTFIX_STATS.with(|v|{let(a,b,c)=v.get();v.set((a+1,b,c));});}
fn record_consumed(){#[cfg(test)] PUBLIC_POSTFIX_STATS.with(|v|{let(a,b,c)=v.get();v.set((a,b+1,c));});}
#[cfg(test)] pub(in super::super) fn record_public_captures(c:usize){PUBLIC_POSTFIX_STATS.with(|v|{let(a,b,_)=v.get();v.set((a,b,c));});}

#[cfg(test)]
mod registered_postfix_tests {
    use super::*;
    fn snapshot(module: Result<Module, String>) -> String { format!("{module:?}") }

    #[test]
    fn registered_scalar_calls_preserve_scopes_directives_spans_and_errors() {
        for source in [
            "x=f(a)\n", "x=f(a,b,)().field\n", "x=obj.method(a,2).field\n",
            "x=f ( a ,2,)\n", "x=f(\"x\" ,y)\n", "x=f(True,False,None,NotImplemented)\n",
            "x=f(0xff,1e-2,2.5)\n", "x=f('é',\"a\\n\")\n",
            "print()\nprint(a)\nx=len(a)\ny=isinstance(a,int)\nz=list(a)\n",
            "class C:\n    def f(self,a): return self.g(a).x\n    def g(self,a): return a\nx=C(a)\n",
            "from collections import deque\nx=deque(a)\n",
            "def f(len,a):\n    return len(a)\n",
            "def outer(x):\n    def inner(a):\n        nonlocal x\n        x=f(a)\n        return x\n    return inner\n",
            "x=f(a=1)\n", "x=f(*args,**kwargs)\n", "x=f(g(a))\n", "x=f(i for i in a)\n",
            "x=f(a[1:])\n", "x=f(-2)\n", "x=f(f'{a}')\n", "x=f(b'x')\n",
            "x=f('a' 'b')\n", "x=f(a, # comment\n b)\n", "x=f(a,,)\n", "x=f(a+)\n",
            "x=super().f(a)\n", "x=globals()\n", "x=f(a).g(b).h()\n",
        ] {
            assert_eq!(snapshot(crate::parse_with_peg(source)),snapshot(crate::parse_with_direct_ast(source)),"{source}");
        }
    }
    #[test]
    fn public_parser_publishes_and_consumes_registered_postfixes() {
        assert_eq!(std::mem::size_of::<Recipe>(),48);assert_eq!(std::mem::size_of::<Node>(),64);
        for source in ["x=f(a).g(2)\n","def f(a):\n    return obj.method(a).field\n"] {
            POSTFIX_ENABLED.with(|v|v.set(false));PUBLIC_POSTFIX_STATS.with(|v|v.set((0,0,0)));
            let control=snapshot(crate::parse(source));let before=PUBLIC_POSTFIX_STATS.with(|v|v.get());
            POSTFIX_ENABLED.with(|v|v.set(true));PUBLIC_POSTFIX_STATS.with(|v|v.set((0,0,0)));
            let candidate=snapshot(crate::parse(source));let after=PUBLIC_POSTFIX_STATS.with(|v|v.get());
            assert_eq!(control,candidate,"{source}");assert!(after.0>0 && after.1>0);assert!(after.2<before.2);
            println!("public registered capture census {before:?} -> {after:?}");
        }
    }
    #[test]
    #[ignore = "explicit full corpus gate; pass PY_FRONTEND_CORPUS manifest"]
    fn public_registered_corpus_matches_control_and_peg() {
        let manifest=std::fs::read_to_string(std::env::var("PY_FRONTEND_CORPUS").expect("corpus manifest")).unwrap();
        let(mut accepts,mut rejects,mut published,mut consumed,mut before_captures,mut after_captures)=(0,0,0,0,0,0);
        for path in manifest.lines().filter(|line|!line.is_empty()) {
            let source=std::fs::read_to_string(path).unwrap();
            POSTFIX_ENABLED.with(|v|v.set(false));PUBLIC_POSTFIX_STATS.with(|v|v.set((0,0,0)));
            let control=crate::parse(&source);let before=PUBLIC_POSTFIX_STATS.with(|v|v.get());
            POSTFIX_ENABLED.with(|v|v.set(true));PUBLIC_POSTFIX_STATS.with(|v|v.set((0,0,0)));
            let candidate=crate::parse(&source);let after=PUBLIC_POSTFIX_STATS.with(|v|v.get());
            if candidate.is_ok(){accepts+=1;}else{rejects+=1;}
            let expected=snapshot(control);assert_eq!(expected,snapshot(candidate),"public candidate {path}");
            assert_eq!(expected,snapshot(crate::parse_with_peg(&source)),"PEG {path}");
            published+=after.0;consumed+=after.1;before_captures+=before.2;after_captures+=after.2;
        }
        println!("corpus accepts={accepts} rejects={rejects} published={published} consumed={consumed} captures={before_captures}->{after_captures}");
        assert!(published>0 && consumed>0);
    }
    #[test]
    #[ignore = "bounded private frontend CPU and actual route census"]
    fn public_registered_postfix_frontend_cpu() {
        unsafe extern "C" { fn clock() -> u64; }
        let cpu=|| unsafe { clock() };
        let manifest=std::fs::read_to_string(std::env::var("PY_FRONTEND_CPU_MANIFEST").unwrap()).unwrap();
        let iterations:usize=std::env::var("PY_FRONTEND_CPU_ITERS").unwrap_or("5".into()).parse().unwrap();
        for path in manifest.lines().filter(|p|!p.is_empty()) {
            let source=std::fs::read_to_string(path).unwrap();
            for enabled in [false,true] { POSTFIX_ENABLED.with(|v|v.set(enabled));drop(crate::parse(&source).unwrap()); }
            for order in [[false,true],[true,false]] {
                for enabled in order {
                    POSTFIX_ENABLED.with(|v|v.set(enabled));PUBLIC_POSTFIX_STATS.with(|v|v.set((0,0,0)));
                    let start=cpu();
                    for _ in 0..iterations {let module=crate::parse(std::hint::black_box(&source)).unwrap();std::hint::black_box(&module);drop(module);}
                    let ticks=cpu()-start;let stats=PUBLIC_POSTFIX_STATS.with(|v|v.get());
                    println!("CPU path={path:?} order={order:?} enabled={enabled} iterations={iterations} process_ticks={ticks} published={} consumed={} captures={}",stats.0,stats.1,stats.2);
                }
            }
        }
    }
    #[test]
    fn registered_postfix_work_limits_restore_arena() {
        for source in ["f(a)","f(a,2).g(\"x\")","f(a,,)","f(a,","a.b"] {
            for steps in 1..128 {
                let mut builder=RecipeAst::new(source);
                let options=grameo::MatchOptions { max_steps:steps,max_rule_depth:256 };
                let result=grameo::engine::build_stackfree_compact_capture_program(&crate::native_grammar::Parser,"expression",source,options,&mut builder);
                if result.is_err(){assert!(builder.arena.values().is_empty(),"{source} {steps}");assert!(!builder.active());}
            }
        }
    }
}

#[cfg(test)] thread_local!{static GROUP_ENABLED:std::cell::Cell<bool>=const{std::cell::Cell::new(true)};static GROUP_ADMISSIONS:std::cell::Cell<usize>=const{std::cell::Cell::new(0)};}
fn grouping_enabled()->bool{#[cfg(test)]{GROUP_ADMISSIONS.with(|v|v.set(v.get()+1));return GROUP_ENABLED.with(|v|v.get());}#[cfg(not(test))]{true}}
#[cfg(test)]pub(in super::super)fn grouping_control(enabled:bool){GROUP_ENABLED.with(|v|v.set(enabled));GROUP_ADMISSIONS.with(|v|v.set(0));}
#[cfg(test)]pub(in super::super)fn grouping_admissions()->usize{GROUP_ADMISSIONS.with(|v|v.get())}

#[cfg(test)]thread_local!{static INNER_ENABLED:std::cell::Cell<bool>=const{std::cell::Cell::new(true)};static INNER_ADMISSIONS:std::cell::Cell<usize>=const{std::cell::Cell::new(0)};}
fn inner_enabled()->bool{#[cfg(test)]{INNER_ADMISSIONS.with(|v|v.set(v.get()+1));return INNER_ENABLED.with(|v|v.get());}#[cfg(not(test))]{true}}
#[cfg(test)]pub(in super::super)fn inner_control(enabled:bool){INNER_ENABLED.with(|v|v.set(enabled));INNER_ADMISSIONS.with(|v|v.set(0));}
#[cfg(test)]pub(in super::super)fn inner_admissions()->usize{INNER_ADMISSIONS.with(|v|v.get())}

// Test-only public RecipeAst admission probes; absent from normal library builds.
#[cfg(test)] pub(in super::super) fn v51_reset_public_capture_stats() {
    PUBLIC_POSTFIX_STATS.with(|v| v.set((0,0,0)));
}
#[cfg(test)] pub(in super::super) fn v51_public_capture_count() -> usize {
    PUBLIC_POSTFIX_STATS.with(|v| v.get().2)
}
