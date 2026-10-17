//! Deferred expression-family consumer: matching builds recipes, never common AST.
use super::*;

#[derive(Clone, Copy)]
struct Meta {
    raw: grameo::Span,
    anchor: Option<grameo::Span>,
    layer: u8,
    boundary: usize,
}
enum Node {
    Scope(usize),
    Owner(bool),
    Unary,
    Gap,
    Expr(Box<Recipe>, Meta),
}
struct Inverse {
    left: Meta,
    right: Meta,
    left_span: Option<grameo::Span>,
    right_span: Option<grameo::Span>,
}
enum Edit {
    Scope(usize),
    Expr(Meta, Option<grameo::Span>),
}
pub(in super::super) struct RecipeAst<'i> {
    source: &'i str,
    #[cfg(test)]
    retain_journal: bool,
    arena: EditArena<Node, Inverse, Edit>,
    scope: Id,
    owner: Id,
    nested: Id,
    unary: Id,
    gap: Id,
}
impl<'i> RecipeAst<'i> {
    pub(in super::super) fn new(source: &'i str) -> Self {
        let mut arena = EditArena::default();
        let scope = arena.alloc(Node::Scope(0));
        let owner = arena.alloc(Node::Owner(true));
        let nested = arena.alloc(Node::Owner(false));
        let unary = arena.alloc(Node::Unary);
        let gap = arena.alloc(Node::Gap);
        arena.commit();
        Self {
            source,
            #[cfg(test)]
            retain_journal: false,
            arena,
            scope,
            owner,
            nested,
            unary,
            gap,
        }
    }
    fn active(&self) -> bool {
        matches!(self.arena.get(self.scope),Node::Scope(n) if *n!=0)
    }
    fn activity(&mut self, enabled: bool) {
        self.arena.edit(
            self.scope,
            |node| {
                let Node::Scope(n) = node else { unreachable!() };
                let old = Edit::Scope(*n);
                *n = usize::from(enabled);
                old
            },
            |node, undo| {
                let (Node::Scope(n), Edit::Scope(old)) = (node, undo) else {
                    unreachable!()
                };
                *n = old;
            },
        );
    }
    pub(in super::super) fn into_caches(self) -> body::Caches {
        let mut caches = body::Caches::default();
        for node in self.arena.into_live_nodes() {
            if let Node::Expr(
                recipe,
                Meta {
                    anchor: Some(raw), ..
                },
            ) = node
            {
                // Identity recipes retain the original capture subtree. Lower
                // it once in the live source walker instead of publishing and
                // consuming an identity-only contextual cache.
                if matches!(recipe.as_ref(), Recipe::Atom { .. }) { continue; }
                caches.expressions.insert(
                    (raw.start, raw.end),
                    Cached::Context {
                        first: Some(Box::new(context::Recipe::Deferred(recipe))),
                        raw,
                        facts: 0,
                        rule: Rule::logical_expr,
                    },
                );
            }
        }
        caches
    }
    fn update(&mut self, id: Id, raw: Option<grameo::Span>, end: usize) {
        self.arena.edit(
            id,
            |node| {
                let Node::Expr(recipe, meta) = node else {
                    unreachable!()
                };
                let old = Edit::Expr(*meta, recipe.span());
                if let Some(raw) = raw {
                    meta.anchor = Some(raw);
                    recipe.set_span(Some(raw));
                } else if meta.layer == 0 {
                    recipe.set_span(Some(grameo::Span {
                        start: meta.raw.start,
                        end,
                    }));
                }
                meta.boundary = end;
                old
            },
            |node, undo| {
                let (Node::Expr(recipe, meta), Edit::Expr(old, old_span)) = (node, undo) else {
                    unreachable!()
                };
                *meta = old;
                recipe.set_span(old_span);
            },
        );
    }
}
impl Builder for RecipeAst<'_> {
    fn empty_success_is_inert(&self,rule:usize)->bool{rule==Rule::op_nl as usize||rule==Rule::eat_nl as usize}
    fn empty_suffix_is_inert(&self,rule:usize)->bool{rule==Rule::conditional_expression as usize||rule==Rule::assignment_expression as usize}
    fn supports_choice_heads(&self)->bool{true}
    fn supports_unobserved_failure_prefixes(&self) -> bool { true }
    fn observes_rule(&self, rule: usize) -> bool {
        rule == Rule::logical_expr as usize || rule == Rule::unary as usize && self.active()
    }
    fn supports_pratt_at(&self, rule: usize, _: usize, _: &str) -> bool {
        rule == Rule::logical_expr as usize
    }
    fn opaque_rule(&self, rule: usize) -> bool {
        rule != Rule::logical_expr as usize && rule != Rule::unary as usize
    }
    fn checkpoint(&self) -> usize {
        self.arena.checkpoint()
    }
    fn rollback(&mut self, mark: usize) {
        self.arena.rollback(mark);
    }
    fn begin(&mut self, rule: usize, _: usize) {
        if rule == Rule::logical_expr as usize {
            let activates = !self.active();
            if activates {
                self.activity(true);
            }
            self.arena
                .push(if activates { self.owner } else { self.nested });
        } else if rule == Rule::unary as usize {
            self.arena.push(self.unary);
        }
    }
    fn finish(&mut self, rule: usize, raw: grameo::Span, _: &str) {
        if rule == Rule::unary as usize {
            while let Some(id) = self.arena.pop() {
                if matches!(self.arena.get(id), Node::Unary) {
                    break;
                }
            }
            let id = self.arena.alloc(Node::Expr(
                Box::new(Recipe::Atom {
                    raw,
                    span: Some(raw),
                }),
                Meta {
                    raw,
                    anchor: None,
                    layer: 0,
                    boundary: raw.end,
                },
            ));
            self.arena.push(id);
            return;
        }
        let (mut first, mut count, mut activates) = (None, 0, false);
        while let Some(id) = self.arena.pop() {
            if let Node::Owner(enabled) = self.arena.get(id) {
                activates = *enabled;
                break;
            }
            first = Some(id);
            count += 1;
        }
        if activates {
            self.activity(false);
        }
        let Some(id) = first.filter(|_| count == 1) else {
            self.arena.push(self.gap);
            return;
        };
        if !matches!(self.arena.get(id),Node::Expr(_,meta) if meta.raw.start==raw.start &&meta.raw.end<=raw.end &&self.source[meta.raw.end..raw.end].trim().is_empty())
        {
            self.arena.push(self.gap);
            return;
        }
        self.update(id, Some(raw), raw.end);
        self.arena.push(id);
    }
    fn accepts_pratt(&self, _: usize, raw: grameo::Span) -> bool {
        self.arena.values().last().is_some_and(|id|matches!(self.arena.get(*id),Node::Expr(_,meta) if meta.raw.start==raw.start &&meta.raw.end<=raw.end &&self.source[meta.raw.end..raw.end].trim().is_empty()))
    }
    fn try_pratt_boundary(
        &mut self,
        _: usize,
        end: usize,
    ) -> Result<(), grameo::builder::BuildError> {
        if let Some(id) = self
            .arena
            .values()
            .last()
            .copied()
            .filter(|id| matches!(self.arena.get(*id), Node::Expr(..)))
        {
            self.update(id, None, end);
        }
        Ok(())
    }
    fn reduce(&mut self, _: usize, fixity: grameo::pratt::Fixity, operator_span: grameo::Span) {
        if !matches!(
            fixity,
            grameo::pratt::Fixity::InfixLeft | grameo::pratt::Fixity::InfixRight
        ) {
            return;
        }
        let text = &self.source[operator_span.start..operator_span.end];
        let selected = operator(text)
            .map(|(op, layer)| (op, layer + 2))
            .or_else(|| {
                compound::operator(text).map(|op| {
                    (
                        op,
                        if matches!(
                            op,
                            BinOp::NullCoalesce
                                | BinOp::Or
                                | BinOp::And
                                | BinOp::BitOr
                                | BinOp::BitXor
                                | BinOp::BitAnd
                        ) {
                            1
                        } else {
                            2
                        },
                    )
                })
            });
        let Some((op, layer)) = selected else {
            return;
        };
        let values = self.arena.values();
        if values.len() < 2 {
            return;
        }
        let (left, right) = (values[values.len() - 2], values[values.len() - 1]);
        let (Node::Expr(_, lm), Node::Expr(_, rm)) = (self.arena.get(left), self.arena.get(right))
        else {
            return;
        };
        if lm.raw.end > operator_span.start || rm.raw.start < operator_span.end {
            return;
        }
        self.arena.pop();
        self.arena.pop();
        let id = self.arena.binary_with_metadata(
            left,
            right,
            |left, right| {
                let (Node::Expr(mut left, lm), Node::Expr(mut right, rm)) = (left, right) else {
                    unreachable!()
                };
                let (left_span, right_span) = (left.span(), right.span());
                if lm.layer > layer {
                    left.set_span(Some(grameo::Span {
                        start: lm.raw.start,
                        end: lm.boundary,
                    }));
                }
                if rm.layer > layer || op == BinOp::Pow && rm.layer == layer {
                    right.set_span(Some(grameo::Span {
                        start: rm.raw.start,
                        end: rm.boundary,
                    }));
                }
                (
                    Node::Expr(
                        Box::new(Recipe::Binary {
                            op,
                            left,
                            right,
                            span: None,
                        }),
                        Meta {
                            raw: grameo::Span {
                                start: lm.raw.start,
                                end: rm.raw.end,
                            },
                            anchor: None,
                            layer,
                            boundary: rm.boundary,
                        },
                    ),
                    Inverse {
                        left: lm,
                        right: rm,
                        left_span,
                        right_span,
                    },
                )
            },
            |node, undo| {
                let Node::Expr(recipe, _) = node else {
                    unreachable!()
                };
                let Recipe::Binary {
                    mut left,
                    mut right,
                    ..
                } = *recipe
                else {
                    unreachable!()
                };
                left.set_span(undo.left_span);
                right.set_span(undo.right_span);
                (Node::Expr(left, undo.left), Node::Expr(right, undo.right))
            },
        );
        self.arena.push(id);
    }
    fn committed(&mut self, _: &str, _: usize) {
        #[cfg(test)]
        if self.retain_journal {
            return;
        }
        self.arena.commit();
    }
}

// Byte coordinates remain neutral until committed live-scope cache consumption.
pub(in super::super) enum Recipe {
    Atom {
        raw: grameo::Span,
        span: Option<grameo::Span>,
    },
    Binary {
        op: BinOp,
        left: Box<Recipe>,
        right: Box<Recipe>,
        span: Option<grameo::Span>,
    },
}
impl Recipe {
    fn span(&self) -> Option<grameo::Span> {
        match self {
            Self::Atom { span, .. } | Self::Binary { span, .. } => *span,
        }
    }
    fn set_span(&mut self, new: Option<grameo::Span>) {
        match self {
            Self::Atom { span, .. } | Self::Binary { span, .. } => *span = new,
        }
    }
    pub(super) fn output_span(&self) -> Span {
        self.span().map_or_else(Span::default, super::span)
    }
    pub(super) fn resolve(self, walker: &mut JsWalker, root: &Pair<Rule>) -> Expression {
        #[cfg(private_tdz_profile)] let _private_timer=PrivateConstructionTimer::new(4);
        #[cfg(private_tdz_profile)] PRIVATE_CONSTRUCTION.with(|x|x.borrow_mut()[5]+=1);
        let mut atoms = HashMap::new();
        let mut pending = vec![root.clone()];
        while let Some(pair) = pending.pop() {
            #[cfg(private_tdz_profile)] PRIVATE_CONSTRUCTION.with(|x|x.borrow_mut()[6]+=1);
            if pair.as_rule() == Rule::unary {
                let raw = pair.as_span();
                #[cfg(private_tdz_profile)] PRIVATE_CONSTRUCTION.with(|x|x.borrow_mut()[7]+=1);
                atoms.insert((raw.start(), raw.end()), pair);
            } else {
                pending.extend(pair.into_inner());
            }
        }
        // Identity owners share source anchors with collapsed primary/call pairs.
        // Reuse that supplied capture directly; never rematch the source.
        let raw = root.as_span();
        atoms
            .entry((raw.start(), raw.end()))
            .or_insert_with(|| root.clone());
        self.resolve_atoms(walker, &mut atoms, false)
    }
    fn resolve_atoms<'i>(
        self,
        walker: &mut JsWalker,
        atoms: &mut HashMap<(usize, usize), Pair<'i, Rule>>,
        operand: bool,
    ) -> Expression {
        match self {
            Self::Atom { raw, span: out } => {
                let pair = atoms
                    .remove(&(raw.start, raw.end))
                    .expect("deferred operand retains original capture");
                let kind = walk_expr_kind(
                    walker,
                    collapse_passthrough_expression(pair).expect("complete deferred operand"),
                )
                .expect("certified deferred operand");
                {
                    let expr =
                        Expression::with_span(kind, out.map_or_else(Span::default, super::span));
                    if operand {
                        normalize_optional_chain(walker, expr)
                    } else {
                        expr
                    }
                }
            }
            Self::Binary {
                op,
                left,
                right,
                span: out,
            } => {
                let expr = normalize_typeof_strict_eq(Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(left.resolve_atoms(walker, atoms, true)),
                    right: Box::new(right.resolve_atoms(walker, atoms, true)),
                }));
                Expression::with_span(expr.kind, out.map_or_else(Span::default, super::span))
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn identity_cache_elision_keeps_binary_direct_ast_ownership() {
        let source = "const x=f(obj?.value); const y=a+b*c;";
        let mut builder = RecipeAst::new(source);
        let tree = crate::native_grammar::Parser.build_captures(Rule::program, source, &mut builder).unwrap();
        let caches = builder.into_caches();
        assert_eq!(caches.expressions.len(), 1, "identity-only roots use their retained captures");
        assert!(caches.expressions.values().all(|cache| matches!(cache,
            Cached::Context { first: Some(recipe), .. }
                if matches!(recipe.as_ref(), context::Recipe::Deferred(binary)
                    if matches!(binary.as_ref(), Recipe::Binary { .. }))
        )), "binary roots still lower through direct AST recipes");
        assert!(tree.capture_count() > 0, "identity subtrees retain original source capture context");
        assert_eq!(format!("{:?}", parse(source)), format!("{:?}", parse_direct_recipes(source)));
    }
    #[test]
    fn rich_recipe_roots_preserve_live_scope_and_directives() {
        for source in [
            "'use strict';function f(a=g(1+2)){let x=a;return ()=>x+g({x},obj.m(x)[0]);}",
            "class C extends B{constructor(x){super(x);this.x=x;}m(){return this.x+super.m();}}",
            "const a=(f())['x']+obj?.foo(a).bar; obj.foo(...xs);obj.foo({a:1});obj.foo`tag`;",
            "const a=typeof f({})===\"number\"||z;const b=(a+b)*(c+d);const c=a**b**c;",
            "const a=obj[Symbol.iterator]();Array.isArray(obj[0]);obj.concat(a,b);",
            "function f(){'use strict';let a=1;try{g(()=>a+2);}catch(e){return e;}finally{a++;}}",
            "const α=f(β+2) <\n g({x:γ});const bad=f(1+);",
        ] {
            assert_eq!(
                format!("{:?}", parse(source)),
                format!("{:?}", parse_direct_recipes(source)),
                "{source}"
            );
        }
    }
    #[test]
    fn recipe_journal_restores_all_small_nested_budgets() {
        for source in [
            "f(g(1+2),()=>h(x))*obj.m(a)[b] < c",
            "f({x:a+b})===typeof g(c)",
            "f(g(1+),obj.x)",
        ] {
            for depth in 0..20 {
                for work in 0..384 {
                    let mut builder = RecipeAst::new(source);
                    builder.retain_journal = true;
                    let mark = builder.checkpoint();
                    let result = grameo::engine::build_stackfree_compact_capture_program(
                        &crate::native_grammar::Parser,
                        "expression",
                        source,
                        grameo::MatchOptions {
                            max_rule_depth: depth,
                            max_steps: work,
                        },
                        &mut builder,
                    );
                    if result.is_err() {
                        assert_eq!(
                            builder.checkpoint(),
                            mark,
                            "{source} depth={depth} work={work}"
                        );
                        assert!(builder.arena.values().is_empty());
                        assert!(!builder.active());
                        assert!(builder.into_caches().expressions.is_empty());
                    }
                }
            }
        }
    }
    #[test]
    fn successful_recipe_inverses_restore_repeatedly() {
        for source in [
            "a+b*c<g(x)&&f(y)",
            "f(g(1+2),()=>h(x))",
            "typeof f({})===\"number\"",
        ] {
            let mut builder = RecipeAst::new(source);
            builder.retain_journal = true;
            let mark = builder.checkpoint();
            for _ in 0..3 {
                crate::native_grammar::Parser
                    .build_captures(Rule::expression, source, &mut builder)
                    .unwrap();
                builder.rollback(mark);
                assert_eq!(builder.checkpoint(), mark);
                assert!(builder.arena.values().is_empty());
                assert!(!builder.active());
            }
        }
    }
}
