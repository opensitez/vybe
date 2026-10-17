//! Explicit owned statement/body reductions; unsupported contexts keep captures.
use super::*;
pub(super) const MODEL: u8 = 1;
pub(super) const LEXICAL: u8 = 2;
#[derive(Clone, Copy)]
pub(super) struct BodyMeta {
    pub raw: grameo::Span,
    pub facts: u8,
}
#[derive(Clone, Copy)]
pub(super) struct StatementMeta {
    pub raw: grameo::Span,
    pub facts: u8,
    pub(super) expression: Option<Metadata>,
}
#[derive(Default)]
pub(in super::super) struct Caches {
    pub expressions: HashMap<(usize, usize), Cached>,
    pub statements: HashMap<(usize, usize), CachedStatement>,
    pub bodies: HashMap<(usize, usize), CachedBody>,
}
pub(super) fn observes(rule: usize) -> bool {
    super::region::body_observes(rule)
}
pub(super) fn complete<const S: bool, const B: bool>(
    ast: &IslandAst<'_, S, B>,
    rule: usize,
    raw: grameo::Span,
) -> bool {
    ast.arena
        .values()
        .last()
        .is_some_and(|id| match ast.arena.get(*id) {
            Value::Statement(_, meta) => {
                meta.raw == raw
                    && (rule == Rule::expression_statement as usize
                        || rule == Rule::return_statement as usize
                        || rule == Rule::empty_statement as usize
                        || rule == Rule::debugger_statement as usize)
            }
            Value::Body(_, anchor) => anchor.raw == raw && rule == Rule::function_body as usize,
            _ => false,
        })
}
pub(super) fn finish<const S: bool, const B: bool>(
    ast: &mut IslandAst<'_, S, B>,
    rule: usize,
    raw: grameo::Span,
    source: &str,
) {
    if rule == Rule::expression as usize {
        if !ast.complete(raw) {
            return;
        }
        let id = *ast.arena.values().last().unwrap();
        ast.arena.edit(
            id,
            |value| {
                let Value::Ready(expression, meta) = value else {
                    unreachable!()
                };
                let undo = EditUndo::Boundary(meta.boundary, expression.span);
                meta.boundary = raw.end;
                expression.span = span(raw);
                undo
            },
            |value, undo| {
                let (Value::Ready(expression, meta), EditUndo::Boundary(boundary, original)) =
                    (value, undo)
                else {
                    unreachable!()
                };
                meta.boundary = boundary;
                expression.span = original;
            },
        );
    } else if rule == Rule::expression_statement as usize || rule == Rule::return_statement as usize
    {
        let is_return = rule == Rule::return_statement as usize;
        let candidate = ast.arena.values().last().copied().filter(|id| {
            let Value::Ready(_, meta) = ast.arena.get(*id) else {
                return false;
            };
            if meta.anchor.is_none() || meta.raw.start < raw.start || meta.boundary > raw.end {
                return false;
            }
            let prefix = source[raw.start..meta.raw.start].trim();
            let tail = source[meta.boundary..raw.end].trim();
            (if is_return {
                prefix == "return"
            } else {
                prefix.is_empty()
            }) && matches!(tail, "" | ";")
        });
        if let Some(child) = candidate {
            ast.arena.pop();
            let id = ast.arena.unary(
                child,
                |value| {
                    let Value::Ready(expression, original) = value else {
                        unreachable!()
                    };
                    let kind = if is_return {
                        StmtKind::Return(Some(expression))
                    } else {
                        StmtKind::Expr(expression)
                    };
                    Value::Statement(
                        Box::new(Statement::with_span(kind, span(raw))),
                        StatementMeta {
                            raw,
                            expression: Some(original),
                            facts: original.facts,
                        },
                    )
                },
                split_statement,
            );
            ast.arena.push(id);
        } else if is_return
            && source[raw.start..raw.end]
                .strip_prefix("return")
                .is_some_and(|tail| matches!(tail.trim(), "" | ";"))
        {
            let id = ast.arena.alloc(Value::Statement(
                Box::new(Statement::with_span(StmtKind::Return(None), span(raw))),
                StatementMeta {
                    raw,
                    expression: None,
                    facts: 0,
                },
            ));
            ast.arena.push(id);
        }
    } else if rule == Rule::empty_statement as usize || rule == Rule::debugger_statement as usize {
        let id = ast.arena.alloc(Value::Statement(
            Box::new(Statement::with_span(StmtKind::Empty, span(raw))),
            StatementMeta {
                raw,
                expression: None,
                facts: 0,
            },
        ));
        ast.arena.push(id);
    } else if rule == Rule::statement as usize {
        let represented = ast.arena.values().last().is_some_and(|id| {
            matches!(ast.arena.get(*id), Value::Statement(_, meta)
                if meta.raw.start == raw.start && meta.raw.end <= raw.end
                    && source[meta.raw.end..raw.end].trim().is_empty())
        });
        if !represented {
            ast.arena.push(ast.gap.unwrap());
        }
    }
}
fn split_statement(value: Value) -> Value {
    let Value::Statement(statement, meta) = value else {
        unreachable!()
    };
    let expression = match statement.kind {
        StmtKind::Expr(expression) | StmtKind::Return(Some(expression)) => expression,
        _ => unreachable!(),
    };
    Value::Ready(expression, meta.expression.unwrap())
}
pub(super) fn recapture<'i>(rule: Rule, raw: grameo::Span, source: &'i str) -> Pair<'i, Rule> {
    let tree = grameo::engine::parse_program_at(
        &super::super::super::native_grammar::Parser,
        rule.name(),
        source,
        raw.start,
        grameo::MatchOptions::default(),
    )
    .expect("previously certified owned region");
    let pair = tree
        .into_typed_pairs::<Rule>()
        .next()
        .expect("owned rule capture");
    assert_eq!(
        (pair.as_span().start(), pair.as_span().end()),
        (raw.start, raw.end),
        "recapture keeps absolute anchor"
    );
    pair
}
fn statement_rule(statement: &Statement) -> Option<Rule> {
    Some(match &statement.kind {
        StmtKind::Expr(_) => Rule::expression_statement,
        StmtKind::Return(_) => Rule::return_statement,
        StmtKind::Empty => return None,
        StmtKind::Block(_) => Rule::block_statement,
        StmtKind::VarDecl { .. } => Rule::variable_declaration,
        StmtKind::FunctionDecl { is_async, .. } => {
            if *is_async {
                Rule::async_function_declaration
            } else {
                Rule::function_declaration
            }
        }
        StmtKind::If { .. } => Rule::if_statement,
        StmtKind::While { .. } => Rule::while_statement,
        StmtKind::DoWhile { .. } => Rule::do_while_statement,
        StmtKind::Throw { .. } => Rule::throw_statement,
        StmtKind::Break(_) => Rule::break_statement,
        StmtKind::Continue(_) => Rule::continue_statement,
        StmtKind::Try { .. } => Rule::try_statement,
        _ => unreachable!("certified statement kind"),
    })
}
pub(in super::super) struct CachedStatement {
    first: Option<Box<Statement>>,
    meta: StatementMeta,
    rule: Option<Rule>,
}
impl CachedStatement {
    pub(super) fn new(statement: Statement, meta: StatementMeta) -> Self {
        Self::new_box(Box::new(statement), meta)
    }
    pub(super) fn new_box(statement: Box<Statement>, meta: StatementMeta) -> Self {
        let rule = statement_rule(&statement);
        Self {
            first: Some(statement),
            meta,
            rule,
        }
    }
    pub(super) fn has_lexical_binding(&self) -> bool {
        self.meta.facts & LEXICAL != 0
    }
    pub(super) fn take(&mut self, source: &str, names: &mut bool) -> Statement {
        *names |= self.meta.facts & MODEL != 0;
        if let Some(first) = self.first.take() {
            return *first;
        }
        let Some(rule) = self.rule else {
            return Statement::with_span(StmtKind::Empty, span(self.meta.raw));
        };
        // No recipe, AST clone, traversal or second parse is paid on first use.
        // Only repeated contextual inspection reconstructs the certified region.
        #[cfg(private_tdz_profile)] PRIVATE_CONSTRUCTION.with(|x|x.borrow_mut()[11]+=1);
        let pair = recapture(rule, self.meta.raw, source);
        super::super::walk_statement(&mut JsWalker::default(), pair)
            .expect("certified statement construction")
    }
}
pub(in super::super) struct CachedBody {
    first: Option<Vec<Statement>>,
    meta: BodyMeta,
}
impl CachedBody {
    pub(super) fn new(statements: Vec<Statement>, meta: BodyMeta) -> Self {
        Self {
            first: Some(statements),
            meta,
        }
    }
    pub(super) fn has_lexical_binding(&self) -> bool {
        self.meta.facts & LEXICAL != 0
    }
    pub(super) fn take(&mut self, source: &str, names: &mut bool) -> Vec<Statement> {
        *names |= self.meta.facts & MODEL != 0;
        if let Some(first) = self.first.take() {
            return first;
        }
        super::super::walk_body(
            &mut JsWalker::default(),
            recapture(Rule::function_body, self.meta.raw, source),
        )
        .expect("certified body construction")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn complete_bodies_prune_only_owned_regions_and_keep_mixed_fallback() {
        for (source, count) in [
            ("function f(){'use strict';return x+2;}", 0),
            ("function f(){\n 'use strict';\n return x+2;\n}", 0),
            ("class C { m(x=2){return x+3;} }", 1),
            ("function f(){return 1+2; const x=3;return x+4;}", 0),
            ("function f(){function g(){return 2+3;}return 4+5;}", 0),
        ] {
            let _lines = vybe_ast::line_index::LineIndex::install(source);
            let mut builder = BodyAst::new(source);
            let tree = super::super::super::super::native_grammar::Parser
                .build_captures(Rule::program, source, &mut builder)
                .unwrap();
            let captures = tree.capture_count();
            let caches = builder.into_caches();
            assert_eq!(caches.bodies.len(), count, "{source}");
            assert!(captures > 0, "contextual anchors survive");
            assert_eq!(
                format!("{:?}", super::super::super::parse(source).unwrap()),
                format!(
                    "{:?}",
                    super::super::super::parse_direct_bodies(source).unwrap()
                )
            );
        }
    }
    #[test]
    fn owned_statement_and_list_inverses_preserve_payload_on_repeated_rollback() {
        let source = "{return 2+3;}";
        let _lines = vybe_ast::line_index::LineIndex::install(source);
        let mut builder = BodyAst::new(source);
        builder.begin(Rule::function_body as usize, 0);
        let raw = grameo::Span { start: 8, end: 11 };
        let expression = Expression::with_span(
            ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(Expression::int(2)),
                right: Box::new(Expression::int(3)),
            },
            span(raw),
        );
        let expected = format!("{expression:?}");
        let id = builder.arena.alloc(Value::Ready(
            expression,
            Metadata {
                raw,
                anchor: Some(raw),
                layer: 1,
                facts: 0,
                rule: Rule::additive,
                boundary: raw.end,
            },
        ));
        builder.arena.push(id);
        let mark = builder.checkpoint();
        builder.begin(Rule::expression as usize, raw.start);
        // The expression frame begins before matching its owned operand.
        let frame = builder.arena.pop().unwrap();
        builder.arena.pop();
        builder.arena.push(frame);
        builder.arena.push(id);
        builder.finish(Rule::expression as usize, raw, source);
        builder.finish(
            Rule::return_statement as usize,
            grameo::Span { start: 1, end: 12 },
            source,
        );
        builder.finish(
            Rule::statement as usize,
            grameo::Span { start: 1, end: 12 },
            source,
        );
        builder.finish(
            Rule::function_body as usize,
            grameo::Span { start: 0, end: 13 },
            source,
        );
        assert!(
            matches!(builder.arena.get(*builder.arena.values().last().unwrap()), Value::Body(statements, _) if statements.len() == 1)
        );
        builder.rollback(mark);
        builder.rollback(mark);
        let Value::Ready(actual, meta) = builder.arena.get(id) else {
            panic!("expression restored")
        };
        assert_eq!(format!("{actual:?}"), expected);
        assert_eq!(meta.anchor, Some(raw));
        assert_eq!(builder.arena.values().last(), Some(&id));
    }
    #[test]
    fn failed_matching_cannot_publish_owned_body_or_statement() {
        let source = "function f(){return 2+3;";
        let _lines = vybe_ast::line_index::LineIndex::install(source);
        let mut builder = BodyAst::new(source);
        assert!(
            super::super::super::super::native_grammar::Parser
                .build_captures(Rule::program, source, &mut builder)
                .is_err()
        );
        let caches = builder.into_caches();
        assert!(caches.bodies.is_empty());
        assert!(caches.statements.is_empty());
        assert!(caches.expressions.is_empty());
    }
}
#[cfg(test)]
mod replay_tests {
    use super::*;
    #[test]
    fn body_moves_once_replays_exactly_and_rejects_another_source_identity() {
        let source = "{'use strict';return document+2;}";
        let _lines = vybe_ast::line_index::LineIndex::install(source);
        let mut builder = BodyAst::new(source);
        super::super::super::super::native_grammar::Parser
            .build_captures(Rule::function_body, source, &mut builder)
            .unwrap();
        let mut caches = builder.into_caches();
        let cache = caches.bodies.get(&(0, source.len())).unwrap();
        let owned_pointer = cache.first.as_ref().unwrap().as_ptr();
        let mut walker = JsWalker::default();
        walker.direct_body_mode = true;
        walker.direct_input = Some((source.as_ptr() as usize, source.len()));
        walker.direct_bodies = std::mem::take(&mut caches.bodies);
        let other = source.to_owned();
        let other_pair = JsParser::parse(Rule::function_body, &other)
            .unwrap()
            .next()
            .unwrap();
        assert!(super::super::take_body_for_pair(&mut walker, &other_pair).is_none());
        let pair = JsParser::parse(Rule::function_body, source)
            .unwrap()
            .next()
            .unwrap();
        let first = super::super::take_body_for_pair(&mut walker, &pair).unwrap();
        assert_eq!(
            first.as_ptr(),
            owned_pointer,
            "body buffer is moved, not copied"
        );
        assert!(walker.needs_model_expr_walk);
        walker.needs_model_expr_walk = false;
        let repeated = super::super::take_body_for_pair(&mut walker, &pair).unwrap();
        assert!(walker.needs_model_expr_walk);
        assert_eq!(format!("{first:?}"), format!("{repeated:?}"));
        let expected = super::super::super::walk_body(&mut JsWalker::default(), pair).unwrap();
        assert_eq!(format!("{first:?}"), format!("{expected:?}"));
    }
}
#[cfg(test)]
mod absolute_replay_tests {
    use super::*;
    #[test]
    fn replay_at_nonzero_unicode_prefix_preserves_function_body_coordinates_and_facts() {
        let source = "// café\nclass C {m(x){\nlet y=2;return document+x;\n}}";
        let _lines = vybe_ast::line_index::LineIndex::install(source);
        let mut builder = BodyAst::new(source);
        super::super::super::super::native_grammar::Parser
            .build_captures(Rule::program, source, &mut builder)
            .unwrap();
        let mut caches = builder.into_caches();
        let (&(start, end), cache) = caches
            .bodies
            .iter_mut()
            .next()
            .expect("method body stays available to contextual class construction");
        assert!(start > 0);
        let first_ptr = cache.first.as_ref().unwrap().as_ptr();
        let expected = super::super::super::walk_body(
            &mut JsWalker::default(),
            recapture(Rule::function_body, grameo::Span { start, end }, source),
        )
        .unwrap();
        for index in 0..3 {
            let mut model = false;
            let actual = cache.take(source, &mut model);
            if index == 0 {
                assert_eq!(actual.as_ptr(), first_ptr);
            }
            assert!(model);
            assert!(cache.has_lexical_binding());
            assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
        }
    }
}
