//! Rollback-safe owned JavaScript regions with contextual PEG fallback.
use super::*;
pub(super) mod recipes;
mod body;
mod compound;
mod chain;
mod container;
mod region;
mod context;
mod prefix;
pub(super) use body::{CachedBody, CachedStatement};
use grameo::{
    arena::{Id, EditArena},
    builder::Builder,
};

#[derive(Clone, Copy)]
struct Metadata {
    raw: grameo::Span,
    anchor: Option<grameo::Span>,
    layer: u8,
    facts: u8,
    rule: Rule,
    boundary: usize,
}
enum Value {
    Scope(usize),
    Marker(bool),
    Ready(Expression, Metadata),
    Context(Box<context::Recipe>, Metadata),
    Statement(Box<Statement>, body::StatementMeta),
    Body(Vec<Statement>, body::BodyMeta),
    RegionStart(usize),
    Name(String, grameo::Span),
    Declarator(VarDeclarator, region::HeaderMeta),
    Param(Param, region::HeaderMeta),
    Params(Vec<Param>, grameo::Span, Vec<region::HeaderMeta>),
    Catch(CatchClause, region::CatchMeta),
    Finally(Vec<Statement>, region::FinallyMeta),
    ArrayItem(Expression, Metadata),
    ArrayElision(grameo::Span),
    Property(grameo::Span),
    Link(chain::Link, grameo::Span, u8),
    Gap,
}
// Composition inverses keep full operand metadata. Hot span and activity
// edits use a smaller payload while sharing the arena's single undo journal.
enum Undo {
    Binary(Metadata, Metadata, Span, Span),
    ContextBinary(Metadata, Metadata, Span, Span),
    Prefix(Metadata, Span),
    Region(usize, Vec<region::ChildMeta>),
    LinkIndex(Metadata),
    LinkArgs(std::rc::Rc<Vec<Metadata>>),
    Chain(Metadata, Span, Vec<chain::Step>),
    Array(Vec<container::Route>),
}
enum EditUndo {
    Anchor(Option<grameo::Span>),
    Rule(Rule),
    Property(grameo::Span),
    Boundary(usize, Span),
    Scope(usize),
    ArrayItem(u8),
}
#[derive(Default)]
pub(super) struct IslandAst<'i, const STRINGS: bool, const BODIES: bool = false> {
    source: &'i str,
    #[cfg(test)]
    retain_test_journal: bool,
    arena: EditArena<Value, Undo, EditUndo>,
    marker: Option<Id>,
    ignored_marker: Option<Id>,
    scope: Option<Id>,
    gap: Option<Id>,
    frames: Vec<Option<Id>>,
    ignored_call_frame: Option<Id>,
}
pub(super) type Ast<'i> = IslandAst<'i, false>;
pub(super) type StringAst<'i> = IslandAst<'i, true>;
pub(super) type BodyAst<'i> = IslandAst<'i, true, true>;
impl<'i, const STRINGS: bool, const BODIES: bool> IslandAst<'i, STRINGS, BODIES> {
    pub(super) fn new(source: &'i str) -> Self {
        let mut ast = Self {
            source,
            ..Default::default()
        };
        ast.marker = Some(ast.arena.alloc(Value::Marker(true)));
        ast.ignored_marker = Some(ast.arena.alloc(Value::Marker(false)));
        ast.scope = Some(ast.arena.alloc(Value::Scope(0)));
        if BODIES {
            ast.gap = Some(ast.arena.alloc(Value::Gap));
            // Frames are immutable stack tokens. Reuse one per observed rule;
            // allocating a full Value slot for every entry wastes arena work.
            // Create tokens before commit so rollback can never invalidate IDs.
            ast.ignored_call_frame=Some(ast.arena.alloc(Value::RegionStart(region::IGNORED_CALL)));
            ast.frames = (0..=Rule::EOI as usize).map(|rule| {
                (region::framed(rule) || rule==Rule::unary as usize).then(|| ast.arena.alloc(Value::RegionStart(rule)))
            }).collect();
        }
        ast.arena.commit();
        ast
    }
    pub(super) fn into_expressions(self) -> HashMap<(usize, usize), Cached> {
        self.into_caches().expressions
    }
    pub(super) fn into_caches(self) -> body::Caches {
        let mut caches = body::Caches::default();
        for value in self.arena.into_live_nodes() {
            match value {
                Value::Ready(expression, Metadata {raw, anchor: Some(anchor), facts, rule, ..}) => {
                    caches.expressions.insert((anchor.start, anchor.end), Cached::new_at(expression, raw, anchor, facts, rule));
                }
                Value::Context(recipe,meta) if recipe.is_binary() => {
                    if let Some(anchor)=meta.anchor {caches.expressions.insert((anchor.start,anchor.end),Cached::Context{first:Some(recipe),raw:anchor,facts:meta.facts,rule:meta.rule});}
                }
                Value::Statement(statement, meta) => {
                    caches.statements.insert((meta.raw.start, meta.raw.end), body::CachedStatement::new_box(statement, meta));
                }
                Value::Body(statements, meta) => {
                    caches.bodies.insert((meta.raw.start, meta.raw.end), body::CachedBody::new(statements, meta));
                }
                Value::Declarator(decl, meta) => { region::cache_header(decl.init, meta, &mut caches); }
                Value::Param(param, meta) => { region::cache_header(param.default, meta, &mut caches); }
                Value::Catch(catch, meta) => {
                    caches.statements.insert((meta.block.raw.start, meta.block.raw.end), body::CachedStatement::new(Statement::with_span(StmtKind::Block(catch.body), span(meta.block.raw)), meta.block));
                }
                Value::Finally(statements, meta) => {
                    caches.statements.insert((meta.block.raw.start, meta.block.raw.end), body::CachedStatement::new(Statement::with_span(StmtKind::Block(statements), span(meta.block.raw)), meta.block));
                }
                Value::Params(params, _, metas) => {
                    for (param, meta) in params.into_iter().zip(metas) { region::cache_header(param.default, meta, &mut caches); }
                }
                Value::Link(link,_,_)=>chain::publish(link,&mut caches),
                Value::ArrayItem(expr,meta)=>container::publish(expr,meta,&mut caches),
                _ => {}
            }
        }
        caches
    }
    fn active(&self) -> bool {
        matches!(self.arena.get(self.scope.unwrap()), Value::Scope(depth) if *depth > 0)
    }
    fn change_scope(&mut self, entering: bool) {
        self.arena.edit(
            self.scope.unwrap(),
            |value| {
                let Value::Scope(depth) = value else {
                    unreachable!()
                };
                let undo = EditUndo::Scope(*depth);
                if entering {
                    *depth += 1;
                } else {
                    *depth -= 1;
                }
                undo
            },
            |value, undo| {
                let (Value::Scope(depth), EditUndo::Scope(original)) = (value, undo) else {
                    unreachable!()
                };
                *depth = original;
            },
        );
    }
    fn complete(&self, span: grameo::Span) -> bool {
        self.arena.values().last().is_some_and(|id| {
            matches!(self.arena.get(*id), Value::Ready(_, meta) | Value::Context(_,meta)
            if meta.raw.start == span.start && meta.raw.end <= span.end
                && self.source[meta.raw.end..span.end].trim().is_empty())
        })
    }
}
fn same_span(a:Span,b:Span)->bool {
    (a.start_line,a.start_col,a.end_line,a.end_col)==(b.start_line,b.start_col,b.end_line,b.end_col)
}
fn span(raw: grameo::Span) -> Span {
    vybe_ast::line_index::span_0based(raw.start, raw.end)
        .expect("frontend installs source line index")
}
fn numeric(raw: &str) -> Option<ExprKind> {
    let stripped = normalized_numeric_text(raw);
    let value = if stripped.starts_with("0x") || stripped.starts_with("0X") {
        Literal::Int(i64::from_str_radix(&stripped[2..], 16).ok()?)
    } else if stripped.starts_with("0o") || stripped.starts_with("0O") {
        Literal::Int(i64::from_str_radix(&stripped[2..], 8).ok()?)
    } else if stripped.starts_with("0b") || stripped.starts_with("0B") {
        Literal::Int(i64::from_str_radix(&stripped[2..], 2).ok()?)
    } else if stripped.contains('.') || stripped.contains('e') || stripped.contains('E') {
        Literal::Float(stripped.parse().ok()?)
    } else {
        Literal::Int(stripped.parse().unwrap_or(0))
    };
    Some(ExprKind::Lit(value))
}
fn operator(raw: &str) -> Option<(BinOp, u8)> {
    Some(match raw.trim() {
        "+" => (BinOp::Add, 1),
        "-" => (BinOp::Sub, 1),
        "*" => (BinOp::Mul, 2),
        "/" => (BinOp::Div, 2),
        "%" => (BinOp::Mod, 2),
        "**" => (BinOp::Pow, 3),
        _ => return None,
    })
}
fn split_binary(value: Value, previous: Undo) -> (Value, Value) {
    let Value::Ready(
        Expression {
            kind:
                ExprKind::Binary {
                    mut left,
                    mut right,
                    ..
                },
            ..
        },
        _,
    ) = value
    else {
        unreachable!()
    };
    let Undo::Binary(lm, rm, ls, rs) = previous else {
        unreachable!()
    };
    left.span = ls;
    right.span = rs;
    (Value::Ready(*left, lm), Value::Ready(*right, rm))
}
fn quoted_end(start: usize, source: &str) -> Option<usize> {
    let rule = match source.as_bytes().get(start) {
        Some(b'\'') => Rule::single_quoted_string,
        Some(b'"') => Rule::double_quoted_string,
        _ => return None,
    };
    let fragment = grameo::program::Program::quoted_fragment(
        &super::super::native_grammar::Parser, rule as usize,
    )?;
    grameo::quoted::scan(fragment, source, start)
}
fn owned_candidate_at<const STRINGS: bool, const COMPOUND: bool>(start: usize, source: &str) -> Option<bool> {
        let bytes = source.as_bytes();
        let mut end = start;
        let mut has_operator = false;
        loop {
            let begin = end;
            if STRINGS && let Some(quoted) = quoted_end(end, source) {
                end = quoted;
            } else if bytes.get(end).is_some_and(u8::is_ascii_digit) {
                while bytes
                    .get(end)
                    .is_some_and(|b| b.is_ascii_alphanumeric() || matches!(b, b'_' | b'.'))
                {
                    end += 1;
                }
            } else if bytes
                .get(end)
                .is_some_and(|b| b.is_ascii_alphabetic() || matches!(b, b'_' | b'$'))
            {
                while bytes
                    .get(end)
                    .is_some_and(|b| b.is_ascii_alphanumeric() || matches!(b, b'_' | b'$'))
                {
                    end += 1;
                }
                if matches!(
                    &source[begin..end],
                    "globalThis"
                        | "this"
                        | "super"
                        | "true"
                        | "false"
                        | "null"
                        | "undefined"
                        | "new"
                        | "await"
                        | "yield"
                        | "typeof"
                        | "void"
                        | "delete"
                ) {
                    return None;
                }
            } else {
                return None;
            }
            while bytes.get(end).is_some_and(u8::is_ascii_whitespace) {
                // op_nl is explicit PEG trivia, outside the bound operator
                // rules. Preserve its semantics through contextual captures.
                if bytes[end] == b'\n' {
                    return None;
                }
                end += 1;
            }
            if matches!(
                bytes.get(end),
                None | Some(b';' | b',' | b')' | b']' | b'}' | b'?' | b':')
            ) {
                return Some(has_operator);
            }
            if COMPOUND && (bytes.get(end).is_some_and(|byte|matches!(byte,b'='|b'!'|b'<'|b'>'|b'&'|b'|'|b'^'))
                || ["instanceof", "in"].iter().any(|word|source[end..].strip_prefix(word).is_some_and(|tail|tail.as_bytes().first().is_none_or(|byte| !byte.is_ascii_alphanumeric() && !matches!(byte,b'_'|b'$'))))) {return Some(has_operator);}
            let n = match bytes.get(end) {
                Some(b'*') if bytes.get(end + 1) == Some(&b'*') => 2,
                Some(b'+' | b'-' | b'*' | b'/' | b'%') => 1,
                _ => return None,
            };
            has_operator = true;
            end += n;
            while bytes.get(end).is_some_and(u8::is_ascii_whitespace) {
                // op_nl is explicit PEG trivia, outside the bound operator
                // rules. Preserve its semantics through contextual captures.
                if bytes[end] == b'\n' {
                    return None;
                }
                end += 1;
            }
        }
}
impl<const STRINGS: bool, const BODIES: bool> Builder for IslandAst<'_, STRINGS, BODIES> {
    fn observes_rule(&self, rule: usize) -> bool {
        (BODIES && (rule==Rule::unary as usize || region::observes(rule) && (rule!=Rule::primary as usize || region::primary_enabled(self)) && (rule!=Rule::ident_or_keyword as usize || chain::property_expected(self)) || body::observes(rule) || (rule == Rule::ident_name as usize && region::needs_name(self))))
            || rule == Rule::additive as usize
            || ((rule == Rule::numeric_literal as usize || rule == Rule::ident_name as usize
                || (STRINGS && rule == Rule::string_literal as usize))
                && self.active() && !(BODIES && rule==Rule::ident_name as usize && chain::property_leaf(self)))
    }
    fn supports_pratt_at(&self, rule: usize, start: usize, source: &str) -> bool {
        // Plain owned leaves use PEG matching and prune captures on completion.
        (BODIES && rule==Rule::logical_expr as usize) || rule == Rule::additive as usize && (BODIES || owned_candidate_at::<STRINGS, BODIES>(start, source) == Some(true))
    }
    fn try_pratt_boundary(
        &mut self,
        _: usize,
        end: usize,
    ) -> Result<(), grameo::builder::BuildError> {
        if BODIES && context::boundary(self,end) {return Ok(());}
        let Some(id) = self.arena.values().last().copied() else {
            return Ok(());
        };
        if !matches!(self.arena.get(id), Value::Ready(_, _)) {
            return Ok(());
        }
        self.arena.edit(
            id,
            |value| {
                let Value::Ready(expression, meta) = value else {
                    unreachable!()
                };
                let undo = EditUndo::Boundary(meta.boundary, expression.span);
                meta.boundary = end;
                if meta.layer == 0 {
                    expression.span = span(grameo::Span {
                        start: meta.raw.start,
                        end,
                    });
                }
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
        Ok(())
    }
    fn suppress_pratt_captures(&self, rule: usize) -> bool {
        !BODIES && rule == Rule::additive as usize
    }
    fn accepts_pratt(&self, _: usize, raw: grameo::Span) -> bool {
        self.complete(raw)
    }
    fn checkpoint(&self) -> usize {
        self.arena.checkpoint()
    }
    fn rollback(&mut self, mark: usize) {
        self.arena.rollback(mark);
    }
    fn opaque_rule(&self, rule: usize) -> bool {
        rule != Rule::additive as usize && !(BODIES && rule==Rule::logical_expr as usize)
    }
    fn begin(&mut self, rule: usize, offset: usize) {
        if BODIES && rule==Rule::unary as usize {self.arena.push(self.frames[rule].unwrap());return;}
        if BODIES && region::framed(rule) && rule!=Rule::logical_expr as usize {
            let frame=if rule==Rule::call_expression as usize && !region::call_candidate(offset,self.source) {
                self.ignored_call_frame.unwrap()
            } else {self.frames[rule].unwrap()};
            self.arena.push(frame);
            return;
        }
        if rule == Rule::additive as usize || BODIES && rule==Rule::logical_expr as usize {
            let enabled = BODIES || owned_candidate_at::<STRINGS, BODIES>(offset, self.source).is_some();
            if !enabled && !self.active() { return; }
            if enabled {
                self.change_scope(true);
            }
            self.arena.push(if enabled {
                self.marker.unwrap()
            } else {
                self.ignored_marker.unwrap()
            });
        }
    }
    fn finish(&mut self, rule: usize, raw: grameo::Span, source: &str) {
        let owner=if BODIES && rule==Rule::logical_expr as usize{Rule::logical_expr}else{Rule::additive};
        if BODIES && rule==Rule::unary as usize {context::atom(self,raw,source);return;}
        if BODIES && region::observes(rule) && rule!=Rule::logical_expr as usize {
            region::finish(self, rule, raw, source);
            return;
        }
        if BODIES && body::observes(rule) {
            body::finish(self, rule, raw, source);
            return;
        }
        if BODIES && rule == Rule::ident_name as usize && (!self.active() || region::needs_name(self) && !self.arena.values().last().is_some_and(|id|matches!(self.arena.get(*id),Value::RegionStart(r) if *r==Rule::primary as usize))) {
            let id = self.arena.alloc(Value::Name(source[raw.start..raw.end].into(), raw));
            self.arena.push(id);
            return;
        }
        if rule != Rule::additive as usize && !(BODIES && rule==Rule::logical_expr as usize) {
            if !self.active() {
                return;
            }
            let text = &source[raw.start..raw.end];
            let kind = if rule == Rule::numeric_literal as usize {
                numeric(text)
            } else if STRINGS && rule == Rule::string_literal as usize {
                Some(ExprKind::Lit(Literal::Str(unquote(text))))
            } else if rule == Rule::ident_name as usize
                && !matches!(
                    text,
                    "globalThis" | "this" | "super" | "true" | "false" | "null" | "undefined"
                )
            {
                Some(ExprKind::Ident(text.into()))
            } else {
                None
            };
            if let Some(kind) = kind {
                let facts = if matches!(&kind, ExprKind::Ident(name) if matches!(name.as_str(), "Promise" | "document")) {body::MODEL} else {0};
                let id = self.arena.alloc(Value::Ready(
                    Expression::with_span(kind, span(raw)),
                    Metadata {
                        raw,
                        anchor: None,
                        layer: 0,
                        facts,
                        rule:Rule::additive,
                        boundary: raw.end,
                    },
                ));
                self.arena.push(id);
            }
            return;
        }
        if !self.active() {
            if BODIES && self.complete(raw) {
                // A newly owned primary/call can complete an otherwise inactive
                // arithmetic wrapper. Certify that exact outer operand span;
                // comparison folds and repeat reads need its trivia boundary.
                let id=*self.arena.values().last().unwrap();
                self.arena.edit(id,|value| {
                    let Value::Ready(_,meta)=value else {unreachable!()};
                    let undo=EditUndo::Rule(meta.rule);meta.rule=owner;undo
                },|value,undo| {
                    let (Value::Ready(_,meta),EditUndo::Rule(rule))=(value,undo) else {unreachable!()};
                    meta.rule=rule;
                });
                if !matches!(self.arena.get(id),Value::Ready(_,meta) if meta.anchor==Some(raw)) {
                self.arena.edit(id,|value| {
                    let Value::Ready(_,meta)=value else {unreachable!()};
                    EditUndo::Anchor(meta.anchor.replace(raw))
                },|value,undo| {
                    let (Value::Ready(_,meta),EditUndo::Anchor(anchor))=(value,undo) else {unreachable!()};
                    meta.anchor=anchor;
                });
                }
                let target=span(raw);
                if !matches!(self.arena.get(id),Value::Ready(expr,meta) if meta.boundary==raw.end && same_span(expr.span,target)) {
                self.arena.edit(id,|value| {
                    let Value::Ready(expr,meta)=value else {unreachable!()};
                    let undo=EditUndo::Boundary(meta.boundary,expr.span);
                    meta.boundary=raw.end;expr.span=target;undo
                },|value,undo| {
                    let (Value::Ready(expr,meta),EditUndo::Boundary(boundary,original))=(value,undo) else {unreachable!()};
                    meta.boundary=boundary;expr.span=original;
                });
                }
            }
            return;
        }
        let mut first = None;
        let mut count = 0;
        let mut enabled = false;
        while let Some(id) = self.arena.pop() {
            if let Value::Marker(selected) = self.arena.get(id) {
                enabled = *selected;
                break;
            }
            first = Some(id);
            count += 1;
        }
        if !enabled {
            return;
        }
        self.change_scope(false);
        if count != 1 {
            return;
        }
        let id = first.unwrap();
        if BODIES && context::finish(self,id,raw,source,owner) {self.arena.push(id);return;}
        if !matches!(self.arena.get(id), Value::Ready(_, meta) if meta.raw.start == raw.start
            && meta.raw.end <= raw.end && source[meta.raw.end..raw.end].trim().is_empty())
        {
            return;
        }
        if owner==Rule::logical_expr {
            self.arena.edit(id,|value|{let Value::Ready(expr,meta)=value else{unreachable!()};let undo=EditUndo::Boundary(meta.boundary,expr.span);meta.boundary=raw.end;expr.span=span(raw);undo},|value,undo|{let(Value::Ready(expr,meta),EditUndo::Boundary(boundary,old))=(value,undo)else{unreachable!()};meta.boundary=boundary;expr.span=old;});
        }
        if !matches!(self.arena.get(id),Value::Ready(_,meta) if meta.rule==owner) {
            self.arena.edit(id,|value| {let Value::Ready(_,meta)=value else{unreachable!()};let undo=EditUndo::Rule(meta.rule);meta.rule=owner;undo},|value,undo| {let (Value::Ready(_,meta),EditUndo::Rule(rule))=(value,undo) else{unreachable!()};meta.rule=rule;});
        }
        self.arena.edit(
            id,
            |value| {
                let Value::Ready(_, meta) = value else {
                    unreachable!()
                };
                EditUndo::Anchor(meta.anchor.replace(raw))
            },
            |value, undo| {
                let (Value::Ready(_, meta), EditUndo::Anchor(anchor)) = (value, undo) else {
                    unreachable!()
                };
                meta.anchor = anchor;
            },
        );
        self.arena.push(id);
    }
    fn prune_capture_children(&self, rule: usize, raw: grameo::Span) -> bool {
        (rule == Rule::unary as usize && BODIES && prefix::complete(self,rule,raw))
            || ((rule == Rule::additive as usize || BODIES && rule==Rule::logical_expr as usize) && self.complete(raw)
            && self.arena.values().last().is_some_and(|id| matches!(self.arena.get(*id),Value::Ready(_,meta) if meta.anchor==Some(raw))))
            || (BODIES && (body::complete(self, rule, raw) || region::complete(self, rule, raw)))
    }
    fn reduce(&mut self, _: usize, fixity: grameo::pratt::Fixity, op_span: grameo::Span) {
        if !matches!(
            fixity,
            grameo::pratt::Fixity::InfixLeft | grameo::pratt::Fixity::InfixRight
        ) {
            return;
        }
        let text=&self.source[op_span.start..op_span.end];
        let selected=operator(text).map(|(op,layer)|(op,if BODIES{layer+2}else{layer})).or_else(||if BODIES{compound::operator(text).map(|op|(op,if matches!(op,BinOp::NullCoalesce|BinOp::Or|BinOp::And|BinOp::BitOr|BinOp::BitXor|BinOp::BitAnd){1}else{2}))}else{None});
        let Some((op, layer)) = selected else {
            return;
        };
        if BODIES && context::reduce(self,op,layer,op_span) {return;}
        let values = self.arena.values();
        if values.len() < 2 {
            return;
        }
        let (left, right) = (values[values.len() - 2], values[values.len() - 1]);
        let (Value::Ready(_, lm), Value::Ready(_, rm)) =
            (self.arena.get(left), self.arena.get(right))
        else {
            return;
        };
        if lm.raw.end > op_span.start
            || rm.raw.start < op_span.end
            || self.source[lm.raw.end..rm.raw.start].trim()
                != self.source[op_span.start..op_span.end].trim()
        {
            return;
        }
        self.arena.pop();
        self.arena.pop();
        let id = self.arena.binary_with_metadata(
            left,
            right,
            |left, right| {
                let (Value::Ready(mut left, lm), Value::Ready(mut right, rm)) = (left, right)
                else {
                    unreachable!()
                };
                let (ls, rs) = (left.span, right.span);
                if lm.layer > layer {
                    left.span = span(grameo::Span {
                        start: lm.raw.start,
                        end: lm.boundary,
                    });
                }
                if rm.layer > layer || (op == BinOp::Pow && rm.layer == layer) {
                    right.span = span(grameo::Span {
                        start: rm.raw.start,
                        end: rm.boundary,
                    });
                }
                let raw = grameo::Span {
                    start: lm.raw.start,
                    end: rm.raw.end,
                };
                (
                    Value::Ready(
                        Expression {
                            kind: ExprKind::Binary {
                                op,
                                left: Box::new(left),
                                right: Box::new(right),
                            },
                            span: Span::default(),
                        },
                        Metadata {
                            raw,
                            anchor: None,
                            layer,
                            facts: lm.facts | rm.facts,
                            rule:Rule::additive,
                            boundary: rm.boundary,
                        },
                    ),
                    Undo::Binary(lm, rm, ls, rs),
                )
            },
            split_binary,
        );
        self.arena.push(id);
    }
    fn committed(&mut self, _: &str, _: usize) {
        #[cfg(test)]
        if self.retain_test_journal {return;}
        self.arena.commit();
    }
}

pub(super) enum Cached {
    Context {first:Option<Box<context::Recipe>>,raw:grameo::Span,facts:u8,rule:Rule},
    Ident { first: Option<String>, raw: grameo::Span, span: Span },
    String { first: Option<String>, raw: grameo::Span, span: Span },
    Number { first: Option<Literal>, raw: grameo::Span, span: Span },
    Tree { first: Option<Box<Expression>>, raw:grameo::Span, span:Span, facts:u8, rule:Rule },
}
impl Cached {
    #[cfg(test)]
    fn new(expression: Expression, raw: grameo::Span) -> Self {Self::new_at(expression,raw,raw,0,Rule::additive)}
    fn new_at(expression: Expression, raw: grameo::Span, anchor:grameo::Span, facts:u8, rule:Rule) -> Self {
        match expression {
            Expression { kind: ExprKind::Ident(name), span } =>
                Self::Ident { first: Some(name), raw, span },
            Expression { kind: ExprKind::Lit(Literal::Str(value)), span } =>
                Self::String { first: Some(value), raw, span },
            Expression { kind: ExprKind::Lit(value), span } =>
                Self::Number { first: Some(value), raw, span },
            expression => {let span=expression.span;Self::Tree {first:Some(Box::new(expression)),raw:anchor,span,facts,rule}},
        }
    }
    #[cfg(test)]
    fn take(&mut self, source: &str) -> Expression {
        self.take_and_record(source, &mut false)
    }
    fn take_and_record(&mut self, source: &str, model_names: &mut bool) -> Expression {
        // Move the owned leaf on first use. Repeated inspections reconstruct
        // only that leaf from its certified source span, preserving the AST
        // span that matching attached after consuming boundary trivia.
        match self {
            Self::Context{first,raw,facts,rule}=>{
                let pair=body::recapture(*rule,*raw,source);
                let mut walker=JsWalker::default();
                *model_names|=*facts&body::MODEL!=0;
                return if let Some(recipe)=first.take(){recipe.resolve(&mut walker,&pair)} else {walk_expression(&mut walker,pair).expect("certified contextual expression")};
            },
            Self::Ident { first, raw, span } => {
                let name = first.take().unwrap_or_else(|| source[raw.start..raw.end].into());
                *model_names |= matches!(name.as_str(), "Promise" | "document");
                return Expression::with_span(ExprKind::Ident(name), *span);
            },
            Self::String { first, raw, span } => return Expression::with_span(
                ExprKind::Lit(Literal::Str(first.take().unwrap_or_else(|| unquote(&source[raw.start..raw.end])))), *span,
            ),
            Self::Number { first, raw, span } => return Expression::with_span(
                ExprKind::Lit(first.take().unwrap_or_else(|| {
                    let Some(ExprKind::Lit(value)) = numeric(&source[raw.start..raw.end]) else {
                        unreachable!("certified owned numeric leaf")
                    };
                    value
                })), *span,
            ),
            Self::Tree {first,raw,span,facts,rule} => {
                *model_names |= *facts & body::MODEL !=0;
                if let Some(first)=first.take() {return *first;}
                let pair=body::recapture(*rule,*raw,source);
                let kind=walk_expr_kind(&mut JsWalker::default(),pair).expect("certified arithmetic construction");
                Expression::with_span(kind,*span)
            }
        }
    }

}
pub(super) fn take_for_pair(walker: &mut JsWalker, pair: &Pair<Rule>) -> Option<Expression> {
    if pair.as_rule() != Rule::additive && !compound::is_rule(pair.as_rule() as usize)
        && !matches!(pair.as_rule(), Rule::primary | Rule::call_expression | Rule::array_literal | Rule::function_expression | Rule::async_function_expression | Rule::unary) {
        return None;
    }
    let raw = pair.as_span();
    let input = raw.get_input();
    if walker.direct_input != Some((input.as_ptr() as usize, input.len())) {
        return None;
    }
    let key=(raw.start(),raw.end());
    if matches!(walker.direct_expressions.get(&key),Some(Cached::Context{..})) {
        let Cached::Context{mut first,raw,facts,rule}=walker.direct_expressions.remove(&key).unwrap() else{unreachable!()};
        walker.needs_model_expr_walk|=facts&body::MODEL!=0;
        let result=if let Some(recipe)=first.take() {recipe.resolve(walker,pair)} else {
            let captured=body::recapture(rule,raw,input);
            walk_expression(walker,captured).expect("certified contextual expression")
        };
        walker.direct_expressions.insert(key,Cached::Context{first:None,raw,facts,rule});
        return Some(result);
    }
    let cached = walker.direct_expressions.get_mut(&key)?;
    Some(cached.take_and_record(input, &mut walker.needs_model_expr_walk))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn direct_arithmetic_preserves_full_ast_spans_and_repeated_inspection() {
        for source in [
            "const x = a + b * 3 - 4 / 2;",
            "const x = 2 ** 3 ** 2 + 1;",
            "const x = 1 + 2 + 3;",
            "const x = a + b ;",
            "const x = 0x10 + 1e-3 * 2;",
            "class C extends Base { [a + 1]() { return b * 2; } }",
            "const x = obj?.value; x += 2; const y = globalThis;",
        ] {
            assert_eq!(
                format!("{:?}", parse(source).unwrap()),
                format!("{:?}", parse_direct_source(source).unwrap()),
                "{source}"
            );
        }
        let source = "a + b * 2";
        let _index = vybe_ast::line_index::LineIndex::install(source);
        let mut builder = Ast::new(source);
        let tree = super::super::super::native_grammar::Parser
            .build_captures(Rule::additive, source, &mut builder)
            .unwrap();
        let mut expressions = builder.into_expressions();
        assert_eq!(
            expressions.len(),
            1,
            "the arithmetic island actually constructs a value"
        );
        let value = expressions.values_mut().next().unwrap();
        assert_eq!(format!("{:?}", value.take(source)), format!("{:?}", value.take(source)));
        let pair = tree.into_typed_pairs::<Rule>().next().unwrap();
        assert!(
            pair.into_inner().next().is_none(),
            "completed arithmetic children are pruned"
        );
    }

    #[test]
    fn plain_leaves_construct_owned_ast_without_pratt_frames() {
        for source in ["answer", "42", "0x10", "1.5"] {
            let _index = vybe_ast::line_index::LineIndex::install(source);
            let mut builder = Ast::new(source);
            assert!(!builder.supports_pratt_at(Rule::additive as usize, 0, source));
            let tree = super::super::super::native_grammar::Parser
                .build_captures(Rule::additive, source, &mut builder).unwrap();
            let mut values = builder.into_expressions();
            assert_eq!(values.len(), 1, "{source}");
            let original = JsParser::parse(Rule::additive, source).unwrap().next().unwrap();
            let mut walker = JsWalker::default();
            let expected = walk_expression(&mut walker, original).unwrap();
            assert_eq!(format!("{:?}", values.values_mut().next().unwrap().take(source)), format!("{expected:?}"));
            assert!(tree.into_typed_pairs::<Rule>().next().unwrap().into_inner().next().is_none());
        }
    }

    #[test]
    fn inactive_parent_avoids_markers_and_preserves_nested_owned_islands() {
        let source = "f(a + b, x);";
        let _index = vybe_ast::line_index::LineIndex::install(source);
        let mut builder = Ast::new(source);
        let initial = builder.checkpoint();
        builder.begin(Rule::additive as usize, 0);
        assert_eq!(builder.checkpoint(), initial);
        builder.finish(Rule::additive as usize, grameo::Span { start: 0, end: source.len() }, source);
        assert_eq!(builder.checkpoint(), initial);
        super::super::super::native_grammar::Parser
            .build_captures(Rule::program, source, &mut builder).unwrap();
        assert!(!builder.active());
        let expressions = builder.into_expressions();
        assert_eq!(expressions.len(), 2);
        assert!(expressions.contains_key(&(2, 7)));
        assert!(expressions.contains_key(&(9, 10)));
        assert_eq!(format!("{:?}", parse(source).unwrap()), format!("{:?}", parse_direct_source(source).unwrap()));
    }

    #[test]
    fn failed_source_restores_owned_values_and_checkpoint() {
        let source = "a + b )";
        let _index = vybe_ast::line_index::LineIndex::install(source);
        let mut builder = Ast::new(source);
        let initial = builder.checkpoint();
        assert!(
            super::super::super::native_grammar::Parser
                .build_captures(Rule::program, source, &mut builder)
                .is_err()
        );
        assert_eq!(builder.checkpoint(), initial);
        assert!(builder.arena.values().is_empty());
        builder.rollback(initial);
        assert!(builder.into_expressions().is_empty());
    }

    #[test]
    fn expression_cache_belongs_to_its_source_allocation() {
        let source = "a + b";
        let _index = vybe_ast::line_index::LineIndex::install(source);
        let mut builder = Ast::new(source);
        let tree = super::super::super::native_grammar::Parser
            .build_captures(Rule::additive, source, &mut builder)
            .unwrap();
        let mut walker = JsWalker::default();
        walker.direct_input = Some((source.as_ptr() as usize, source.len()));
        walker.direct_expressions = builder.into_expressions();
        let secondary_source = source.to_owned();
        let secondary = JsParser::parse(Rule::additive, &secondary_source)
            .unwrap()
            .next()
            .unwrap();
        assert!(take_for_pair(&mut walker, &secondary).is_none());
        let original = tree.into_typed_pairs::<Rule>().next().unwrap();
        assert!(take_for_pair(&mut walker, &original).is_some());
    }

    #[test]
    fn mixed_operator_trivia_and_rejection_match_peg() {
        for first in ["+", "-", "*", "/", "%", "**"] {
            for second in ["+", "-", "*", "/", "%", "**"] {
                for source in [
                    format!("const x = a {first} b {second} c;"),
                    format!("const x = a {first}\n b {second} c;"),
                    format!("const x = a /* comment */ {first} (b {second} c);"),
                    format!("const x = a {first} b {second};"),
                ] {
                    let peg = parse(&source).map(|m| format!("{m:?}"));
                    let direct = parse_direct_source(&source).map(|m| format!("{m:?}"));
                    assert_eq!(direct, peg, "{source}");
                }
            }
        }
    }

    #[test]
    fn mode_checkpoints_interleave_with_arena_and_repeat() {
        let source = "a + b";
        let mut builder = Ast::new(source);
        let initial = builder.checkpoint();
        builder.change_scope(true);
        let first = builder.checkpoint();
        builder.change_scope(true);
        let second = builder.checkpoint();
        builder.change_scope(false);
        for _ in 0..2 {
            builder.rollback(second);
            assert_eq!(
                match builder.arena.get(builder.scope.unwrap()) {
                    Value::Scope(d) => *d,
                    _ => unreachable!(),
                },
                2
            );
            builder.change_scope(false);
        }
        builder.rollback(first);
        assert_eq!(
            match builder.arena.get(builder.scope.unwrap()) {
                Value::Scope(d) => *d,
                _ => unreachable!(),
            },
            1
        );
        builder.arena.push(builder.marker.unwrap());
        let stack_mark = builder.checkpoint();
        builder.change_scope(false);
        builder.arena.pop();
        builder.rollback(stack_mark);
        assert_eq!(
            match builder.arena.get(builder.scope.unwrap()) {
                Value::Scope(d) => *d,
                _ => unreachable!(),
            },
            1
        );
        assert_eq!(builder.arena.values(), &[builder.marker.unwrap()]);
        builder.rollback(initial);
        assert_eq!(
            match builder.arena.get(builder.scope.unwrap()) {
                Value::Scope(d) => *d,
                _ => unreachable!(),
            },
            0
        );
        assert!(builder.arena.values().is_empty());
        assert_eq!(builder.checkpoint(), initial);
    }

    #[test]
    fn mode_cursor_restores_all_mixed_journal_boundaries() {
        let mut builder = Ast::new("a + b");
        let depth = |builder: &Ast<'_>| match builder.arena.get(builder.scope.unwrap()) {
            Value::Scope(depth) => *depth, _ => unreachable!(),
        };
        let mut states = Vec::new();
        for step in 0..128 {
            states.push((builder.checkpoint(), depth(&builder), builder.arena.values().to_vec()));
            match step % 8 {
                0 | 2 => builder.change_scope(true),
                1 => builder.arena.push(builder.marker.unwrap()),
                3 | 7 => { builder.arena.pop().unwrap(); },
                4 | 6 => builder.change_scope(false),
                5 => builder.arena.push(builder.ignored_marker.unwrap()),
                _ => unreachable!(),
            }
        }
        for (mark, expected_depth, values) in states.into_iter().rev() {
            for _ in 0..2 {
                builder.rollback(mark);
                assert_eq!(depth(&builder), expected_depth);
                assert_eq!(builder.arena.values(), values.as_slice());
                assert_eq!(builder.checkpoint(), mark);
            }
        }
        assert_eq!(builder.checkpoint(), 0);
    }

    #[test]
    fn repeated_speculative_reductions_restore_exact_children() {
        let source = "a + b";
        let _index = vybe_ast::line_index::LineIndex::install(source);
        let mut builder = Ast::new(source);
        builder.begin(Rule::additive as usize, 0);
        builder.finish(
            Rule::ident_name as usize,
            grameo::Span { start: 0, end: 1 },
            source,
        );
        builder.finish(
            Rule::ident_name as usize,
            grameo::Span { start: 4, end: 5 },
            source,
        );
        let before = builder.checkpoint();
        let stack = builder.arena.values().to_vec();
        for _ in 0..2 {
            builder.reduce(
                Rule::additive as usize,
                grameo::pratt::Fixity::InfixLeft,
                grameo::Span { start: 2, end: 3 },
            );
            builder.rollback(before);
            assert_eq!(builder.arena.values(), stack.as_slice());
            for (id, name) in stack.iter().skip(1).zip(["a", "b"]) {
                assert!(
                    matches!(builder.arena.get(*id), Value::Ready(Expression { kind: ExprKind::Ident(value), .. }, _) if value == name)
                );
            }
        }
    }
}

#[cfg(test)]
mod suppression_tests {
    use super::*;
    #[test]
    fn unowned_bigint_island_recaptures_contextual_peg_tree() {
        for source in ["1n + 2n", "0x10n + 2n", "1_000n + 2n"] {
            let _lines = vybe_ast::line_index::LineIndex::install(source);
            let mut builder = Ast::new(source);
            assert!(builder.supports_pratt_at(Rule::additive as usize, 0, source));
            assert!(builder.suppress_pratt_captures(Rule::additive as usize));
            let tree = super::super::super::native_grammar::Parser
                .build_captures(Rule::additive, source, &mut builder).unwrap();
            assert!(builder.into_expressions().is_empty());
            let pair = tree.into_typed_pairs::<Rule>().next().unwrap();
            assert!(pair.clone().into_inner().next().is_some(), "fallback retains captures");
            let original = JsParser::parse(Rule::additive, source).unwrap().next().unwrap();
            let mut actual = JsWalker::default();
            let mut expected = JsWalker::default();
            assert_eq!(format!("{:?}", walk_expression(&mut actual, pair).unwrap()),
                format!("{:?}", walk_expression(&mut expected, original).unwrap()), "{source}");
        }
    }
}

#[cfg(test)]
mod leaf_move_tests {
    use super::*;
    #[test]
    fn identifier_first_take_moves_storage_and_replays_exact_span() {
        let source = "  answer\t";
        let expected = Expression::with_span(ExprKind::Ident("answer".into()), Span {
            start_line: 3, start_col: 4, end_line: 3, end_col: 11,
        });
        let mut cached = Cached::new(expected.clone(), grameo::Span { start: 2, end: 8 });
        let Cached::Ident { first: Some(name), .. } = &cached else { unreachable!() };
        let pointer = name.as_ptr();
        let first = cached.take(source);
        let ExprKind::Ident(name) = &first.kind else { unreachable!() };
        assert_eq!(name.as_ptr(), pointer, "first binding moves the existing String");
        assert!(matches!(&cached, Cached::Ident { first: None, .. }));
        for actual in [first, cached.take(source), cached.take(source)] {
            assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
        }
    }
    #[test]
    fn numeric_first_take_moves_value_and_replays_decoding_and_span() {
        for source in ["1_234.5", "0x10", "0b101", "0o71", "1e-3", "42"] {
            let expected = Expression::with_span(numeric(source).unwrap(), Span {
                start_line: 1, start_col: 2, end_line: 1, end_col: 30,
            });
            let mut cached = Cached::new(expected.clone(), grameo::Span { start: 0, end: source.len() });
            let first = cached.take(source);
            assert!(matches!(&cached, Cached::Number { first: None, .. }));
            for actual in [first, cached.take(source), cached.take(source)] {
                assert_eq!(format!("{actual:?}"), format!("{expected:?}"), "{source}");
            }
        }
    }
}

#[cfg(test)]
mod owned_string_tests {
    use super::*;
    #[test]
    fn owned_strings_preserve_arithmetic_context_escapes_and_spans() {
        for source in ["const x = 'hello';", "const x = \"café\";", "const x = 'a' + 'b';",
            "const x = 'a' + 2 * 3;", "const x = f('hello', 'a' + 'b');",
            "const x = 'a\\'b';", "const x = \"\\u{1f600}\";", "const x = 'a\rb';",
            "const x = 'a'.length;", "const x = 'a' /*c*/ + 'b';",
            "const key = 'hello'; class C { [key]() { return 'world'; } }"] {
            assert_eq!(format!("{:?}", parse(source)), format!("{:?}", parse_direct_strings(source)), "{source}");
        }
    }
    #[test]
    fn string_first_binding_moves_storage_and_repeated_binding_decodes_raw_token() {
        let source = "'a\\'b'";
        let _lines = vybe_ast::line_index::LineIndex::install(source);
        let mut builder = StringAst::new(source);
        super::super::super::native_grammar::Parser.build_captures(Rule::additive, source, &mut builder).unwrap();
        let mut values = builder.into_expressions();
        let cached = values.values_mut().next().expect("standalone string is owned");
        let Cached::String { first: Some(text), .. } = &cached else { unreachable!() };
        let pointer = text.as_ptr();
        let first = cached.take(source);
        let ExprKind::Lit(Literal::Str(value)) = &first.kind else { unreachable!() };
        assert_eq!(value.as_ptr(), pointer);
        assert_eq!(format!("{first:?}"), format!("{:?}", cached.take(source)));
    }
}

#[cfg(test)]
mod model_fact_tests {
    use super::*;
    #[test]
    fn transferred_model_names_preserve_first_and_repeated_bindings() {
        for (source, expected) in [
            ("document", true), ("Promise", true), ("1 + document", true),
            ("Promise + 1", true), ("answer + 1", false),
            ("'document' + 1", false), ("'Promise' + document", true),
        ] {
            let _lines = vybe_ast::line_index::LineIndex::install(source);
            let mut builder = StringAst::new(source);
            super::super::super::native_grammar::Parser
                .build_captures(Rule::additive, source, &mut builder).unwrap();
            let mut cache = builder.into_expressions();
            let cached = cache.values_mut().next().unwrap();
            let mut first_fact = false;
            let first = cached.take_and_record(source, &mut first_fact);
            let mut repeated_fact = false;
            let repeated = cached.take_and_record(source, &mut repeated_fact);
            assert_eq!(first_fact, expected, "{source}");
            assert_eq!(repeated_fact, expected, "{source}");
            assert_eq!(format!("{first:?}"), format!("{repeated:?}"), "{source}");
        }
    }
}

#[cfg(test)]
mod scope_policy_tests {
    use super::*;
    #[test]
    fn direct_paths_preserve_scope_shapes_prologues_and_module_policies() {
        for source in [
            "'use strict'; let x=1; function f(){return x;}",
            "var x=1; function f(){'use strict'; return function(){return x;};}",
            "function outer(a){let x=2; return ()=>{let x=3; return a+x;};}",
            "class C extends Base { constructor(){super();} m(){'use strict'; return this;} }",
            "var F=function inner(a){return inner;}; var f=(x)=>x+1;",
            "function Foo(){return 1;} function foo(){return 2;} var X=1; var x=2;",
            "var global=1; var nonlocal=2; globalThis.value=global+nonlocal;",
            "for(let document=0;true;){(()=>2)();break;}",
        ] {
            let peg = parse(source).unwrap();
            for direct in [parse_direct_source(source).unwrap(), parse_direct_strings(source).unwrap()] {
                assert_eq!(format!("{peg:?}"), format!("{direct:?}"), "{source}");
                assert_eq!(direct.directives.receiver_binding, Some(ReceiverBinding::UniversalParameter));
                assert_eq!(direct.directives.method_receiver, Some(MethodReceiver::Prototype));
                assert_eq!(direct.directives.reference_binding, None);
                assert_eq!(direct.directives.variable_fold(), None);
                assert_eq!(direct.directives.callable_fold(), None);
            }
        }
    }
}

pub(super) fn take_statement_for_pair(walker: &mut JsWalker, pair: &Pair<Rule>) -> Option<Statement> {
    if !walker.direct_body_mode { return None; }
    let raw = pair.as_span();
    let input = raw.get_input();
    if walker.direct_input != Some((input.as_ptr() as usize, input.len())) { return None; }
    let cached = walker.direct_statements.get_mut(&(raw.start(), raw.end()))?;
    let owned = cached.take(input, &mut walker.needs_model_expr_walk);
    walker.has_lexical_binding |= cached.has_lexical_binding();
    Some(owned)
}
pub(super) fn take_body_for_pair(walker: &mut JsWalker, pair: &Pair<Rule>) -> Option<Vec<Statement>> {
    if !walker.direct_body_mode { return None; }
    let raw = pair.as_span();
    let input = raw.get_input();
    if walker.direct_input != Some((input.as_ptr() as usize, input.len())) { return None; }
    let cached = walker.direct_bodies.get_mut(&(raw.start(), raw.end()))?;
    let owned = cached.take(input, &mut walker.needs_model_expr_walk);
    walker.has_lexical_binding |= cached.has_lexical_binding();
    Some(owned)
}
