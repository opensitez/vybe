//! Opt-in, reversible construction at context-independent PEG boundaries.
//! Contextual expressions, spreads and comprehensions retain their captures.
use super::*;
mod arithmetic;
pub(super) mod recipes;
mod assignment;
mod subscript;
mod members;
use grameo::{
    arena::{Id, MetadataArena},
    builder::Builder,
};

#[derive(Clone, Copy)]
struct Metadata {
    raw: grameo::Span,
    anchor: Option<grameo::Span>,
    anchor_rule: Rule,
    sequence: bool,
    contextual: bool,
    layer: u8,
    boundary: usize,
}
#[derive(Clone, Copy)]
enum Replay {
    Member { raw: grameo::Span, span: Span },
    Leaf {
        rule: Rule,
        raw: grameo::Span,
        span: Span,
    },
    Binary {
        op: BinOp,
        span: Span,
    },
    List {
        count: usize,
        span: Span,
        tuple: bool,
    },
}
impl Replay {
    fn span(&mut self) -> &mut Span {
        match self {
            Self::Leaf { span, .. } | Self::List { span, .. } | Self::Binary { span, .. } | Self::Member {span,..} => span,
        }
    }
}
enum Value {
    Depth(usize),
    Frame(Rule),
    Statement {
        statement: Statement,
        raw: grameo::Span,
        rule: Rule,
        child: Option<Metadata>,
        replay: Vec<Replay>,
    },
    Ready {
        expression: Expression,
        metadata: Metadata,
        replay: Vec<Replay>,
    },
    Assignment {
        statement: Statement,
        raw: grameo::Span,
        rule: Rule,
        modes: Vec<assignment::Mode>,
        replay: Vec<Replay>,
    },
}
enum Undo {
    Depth(usize),
    Span(Metadata, Span, Span),
    Anchor(Metadata),
    Member(Metadata,Span),
    List(Vec<(Metadata, usize)>),
    Binary(Metadata, Metadata, usize, Span, Span),
    Assignment(Vec<(Metadata, usize)>),
}
pub(super) struct Cached {
    deferred: Option<std::sync::Arc<recipes::Recipe>>,
    first: Option<Expression>,
    replay: Vec<Replay>,
    sequence: bool,
    contextual: bool,
}
impl Cached {
    fn take_expression(&mut self, source: &str, start: usize) -> Expression {
        if let Some(first) = self.first.take() {
            return first;
        }
        replay_expression(&self.replay, source, start)
    }
    fn take(&mut self, source: &str, start: usize) -> ExprKind {
        self.take_expression(source, start).kind
    }
}
fn replay_expression(replay: &[Replay], source: &str, start: usize) -> Expression {
    let mut stack = Vec::new();
    for event in replay {
        match *event {
            Replay::Member {raw,span} => {
                let object=stack.pop().expect("member object");
                stack.push(Expression::with_span(ExprKind::Member {object:Box::new(object),field:source[raw.start-start..raw.end-start].trim().strip_prefix('.').unwrap().trim().to_owned(),null_safe:false},span));
            }
            Replay::Leaf { rule, raw, span } => stack.push(Expression::with_span(
                leaf(rule, &source[raw.start - start..raw.end - start]).expect("certified literal"),
                span,
            )),
            Replay::Binary { op, span } => {
                let right = stack.pop().expect("binary right operand");
                let left = stack.pop().expect("binary left operand");
                stack.push(Expression::with_span(
                    ExprKind::Binary {
                        op,
                        left: Box::new(left),
                        right: Box::new(right),
                    },
                    span,
                ));
            }
            Replay::List { count, span, tuple } => {
                let children = stack.split_off(stack.len() - count);
                stack.push(Expression::with_span(
                    if tuple {
                        ExprKind::Tuple(children)
                    } else {
                        array(children)
                    },
                    span,
                ));
            }
        }
    }
    stack.pop().expect("certified expression")
}
pub(super) struct CachedStatement {
    first: Option<Statement>,
    raw: grameo::Span,
    rule: Rule,
    replay: Vec<Replay>,
    sequence: bool,
    contextual: bool,
    assignment_modes: Vec<assignment::Mode>,
}
impl CachedStatement {
    fn take(&mut self, source: &str) -> Statement {
        if let Some(first) = self.first.take() {
            return first;
        }
        let kind = if matches!(self.rule, Rule::expr_or_assign_stmt | Rule::expr_or_assign_inline) {
            let ExprKind::Tuple(mut operands) = replay_expression(&self.replay, source, self.raw.start).kind else { unreachable!() };
            let value = operands.pop().unwrap();
            StmtKind::Assign { targets: operands, value, by_ref: false }
        } else if matches!(self.rule, Rule::return_stmt | Rule::return_inline) {
            let value = if self.replay.is_empty() {
                None
            } else {
                Some(replay_expression(&self.replay, source, self.raw.start))
            };
            StmtKind::Return(value)
        } else {
            simple_statement(self.rule, source).expect("certified statement")
        };
        Statement::with_span(kind, span(self.raw))
    }
}
fn simple_statement(rule: Rule, source: &str) -> Option<StmtKind> {
    Some(match rule {
        Rule::pass_stmt => StmtKind::Empty,
        Rule::break_stmt | Rule::break_inline => StmtKind::Break(BreakTarget::Implicit),
        Rule::continue_stmt | Rule::continue_inline => StmtKind::Continue(ContinueTarget::Implicit),
        Rule::global_stmt | Rule::global_inline | Rule::nonlocal_stmt | Rule::nonlocal_inline => {
            // Explicit line continuations retain the contextual capture adapter.
            if source.contains('\\') {
                return None;
            }
            let global = matches!(rule, Rule::global_stmt | Rule::global_inline);
            let names = source[if global { 6 } else { 8 }..]
                .split(['\n', '\r', ';', '#'])
                .next()
                .unwrap()
                .split(',')
                .map(|name| name.trim().to_owned())
                .collect();
            StmtKind::ScopeDecl {
                kind: if global {
                    ScopeDeclKind::Global
                } else {
                    ScopeDeclKind::Nonlocal
                },
                names,
            }
        }
        _ => return None,
    })
}
fn split_return(value: Value) -> Value {
    let Value::Statement {
        statement:
            Statement {
                kind: StmtKind::Return(Some(expression)),
                ..
            },
        child: Some(metadata),
        replay,
        ..
    } = value
    else {
        unreachable!()
    };
    Value::Ready {
        expression,
        metadata,
        replay,
    }
}
fn array(children: Vec<Expression>) -> ExprKind {
    ExprKind::Array(
        children
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
            .collect(),
    )
}
fn leaf(rule: Rule, text: &str) -> Option<ExprKind> {
    match rule {
        Rule::identifier => Some(if text == "NotImplemented" {
            ExprKind::Call {
                callee: Box::new(Expression::ident("__py_notimplemented")),
                args: Vec::new(),
                optional: false,
            }
        } else {
            ExprKind::Ident(text.to_owned())
        }),
        Rule::numeric_literal => parse_number(text).ok(),
        Rule::string_literal => Some(if is_bytes_prefix(text) {
            parse_bytes_literal(text)
        } else {
            ExprKind::Lit(Literal::Str(parse_python_string(text)))
        }),
        Rule::true_kw => Some(ExprKind::Lit(Literal::Bool(true))),
        Rule::false_kw => Some(ExprKind::Lit(Literal::Bool(false))),
        Rule::none_kw => Some(ExprKind::Lit(Literal::Null)),
        _ => None,
    }
}
fn span(raw: grameo::Span) -> Span {
    let (start_line, start_col) = line_col_at(raw.start);
    let (end_line, end_col) = line_col_at(raw.end);
    Span {
        start_line,
        start_col,
        end_line,
        end_col,
    }
}

fn split_list(value: Value, undo: Undo) -> Vec<Value> {
    let (
        Value::Ready {
            expression: Expression { kind, .. },
            mut replay,
            ..
        },
        Undo::List(metadata),
    ) = (value, undo)
    else {
        unreachable!()
    };
    replay.pop(); // The list event follows all of its children.
    let mut events = replay.into_iter();
    let elements = match kind {
        ExprKind::Array(elements) => elements.into_iter().map(|element| element.value).collect(),
        ExprKind::Tuple(elements) => elements,
        _ => unreachable!(),
    };
    elements
        .into_iter()
        .zip(metadata)
        .map(|(element, (metadata, count))| Value::Ready {
            expression: element,
            metadata,
            replay: events.by_ref().take(count).collect(),
        })
        .collect()
}
pub(super) struct Ast<'i> {
    source: Option<&'i str>,
    // Selection, begin and finish ask about the same immutable source region.
    // This pure lexical hint needs no rollback journal or capture allocation.
    arithmetic_hint: std::cell::Cell<(usize, bool)>,
    pratt_frame: Id,
    bitwise_frame: Id,
    bitwise_hint: std::cell::Cell<(usize, bool)>,
    arena: MetadataArena<Value, Undo>,
    depth: Id,
    expression_frame: Id,
    list_frame: Id,
    tuple_frame: Id,
    return_frame: Id,
    return_inline_frame: Id,
    assignment_frame: Id,
    assignment_inline_frame: Id,
    member_hint: std::cell::Cell<(usize,bool)>,
    member_frame: Id,
    member_operator_frame: Id,
    subscript_frame: Id,
    subscript_enter_frame: Id,
    statements: HashMap<(usize, usize), CachedStatement>,
    expressions: HashMap<(Rule, usize, usize), Cached>,
}
impl<'i> Default for Ast<'i> {
    fn default() -> Self {
        let mut arena = MetadataArena::default();
        let depth = arena.alloc(Value::Depth(0));
        let pratt_frame = arena.alloc(Value::Frame(Rule::additive));
        let bitwise_frame = arena.alloc(Value::Frame(Rule::bitor_expr));
        // Immutable sentinels participate only in stack push/pop, never owned
        // AST composition. Reusing them preserves rollback without allocating
        // a node for each speculative expression boundary.
        let expression_frame = arena.alloc(Value::Frame(Rule::expression));
        let list_frame = arena.alloc(Value::Frame(Rule::list_inner));
        let tuple_frame = arena.alloc(Value::Frame(Rule::expression_list));
        let return_frame = arena.alloc(Value::Frame(Rule::return_stmt));
        let return_inline_frame = arena.alloc(Value::Frame(Rule::return_inline));
        let assignment_frame = arena.alloc(Value::Frame(Rule::expr_or_assign_stmt));
        let assignment_inline_frame = arena.alloc(Value::Frame(Rule::expr_or_assign_inline));
        let member_frame=arena.alloc(Value::Frame(Rule::postfix));
        let member_operator_frame=arena.alloc(Value::Frame(Rule::owned_member_op));
        let subscript_frame=arena.alloc(Value::Frame(Rule::subscript));
        let subscript_enter_frame=arena.alloc(Value::Frame(Rule::subscript));
        arena.commit();
        Self {
            source: None,
            arithmetic_hint: std::cell::Cell::new((usize::MAX, false)),
            pratt_frame,
            bitwise_frame,
            bitwise_hint: std::cell::Cell::new((usize::MAX,false)),
            arena,
            depth,
            expression_frame,
            list_frame,
            tuple_frame,
            return_frame,
            return_inline_frame,
            assignment_frame,
            assignment_inline_frame,
            member_hint: std::cell::Cell::new((usize::MAX,false)),
            member_frame,
            member_operator_frame,
            subscript_frame,
            subscript_enter_frame,
            statements: HashMap::new(),
            expressions: HashMap::new(),
        }
    }
}
impl<'i> Ast<'i> {
    fn arithmetic_candidate(&self, start: usize) -> bool {
        let Some(source) = self.source else { return false };
        let (previous, result) = self.arithmetic_hint.get();
        if previous == start { return result; }
        let result = arithmetic::candidate(start, source);
        self.arithmetic_hint.set((start, result));
        result
    }
    pub(super) fn new(source: &'i str) -> Self {
        Self {
            source: Some(source),
            ..Self::default()
        }
    }
    pub(super) fn take_statements(&mut self) -> HashMap<(usize, usize), CachedStatement> {
        std::mem::take(&mut self.statements)
    }
    pub(super) fn finish(self) -> HashMap<(Rule, usize, usize), Cached> {
        self.expressions
    }
    fn region_candidate(&self, rule: Rule, start: usize) -> bool {
        if rule == Rule::postfix {
            let Some(source)=self.source else {return false;};
            let (previous,result)=self.member_hint.get();
            if previous==start {return result;}
            let result=members::eager_candidate(start,source);
            self.member_hint.set((start,result));
            return result;
        }
        if rule != Rule::bitor_expr { return self.arithmetic_candidate(start); }
        let Some(source) = self.source else { return false; };
        let (old, result) = self.bitwise_hint.get();
        if old == start { return result; }
        let result = arithmetic::bitwise_candidate(start, source);
        self.bitwise_hint.set((start, result));
        result
    }
    fn active(&self) -> bool {
        matches!(self.arena.get(self.depth), Value::Depth(depth) if *depth > 0)
    }
    fn depth(&mut self, entering: bool) {
        self.adjust_depth(self.depth, entering);
    }
    fn adjust_depth(&mut self, id: Id, entering: bool) {
        self.arena.edit(
            id,
            |value| {
                let Value::Depth(depth) = value else {
                    unreachable!()
                };
                let original = *depth;
                if entering {
                    *depth += 1;
                } else {
                    *depth -= 1;
                }
                Undo::Depth(original)
            },
            |value, undo| {
                let (Value::Depth(depth), Undo::Depth(original)) = (value, undo) else {
                    unreachable!()
                };
                *depth = original;
            },
        );
    }
    fn frame_children(&mut self, rule: Rule) -> Vec<Id> {
        let mut children = Vec::new();
        while let Some(id) = self.arena.pop() {
            if matches!(self.arena.get(id), Value::Frame(found) if *found == rule) {
                children.reverse();
                return children;
            }
            children.push(id);
        }
        unreachable!("semantic frame must balance")
    }
    fn single_child(&mut self, rule: Rule) -> Option<Id> {
        let mut child = None;
        let mut multiple = false;
        while let Some(id) = self.arena.pop() {
            if matches!(self.arena.get(id), Value::Frame(found) if *found == rule) {
                return if multiple { None } else { child };
            }
            if child.is_some() {
                multiple = true;
            } else {
                child = Some(id);
            }
        }
        unreachable!("semantic frame must balance")
    }
    fn wrap(&mut self, id: Id, raw: grameo::Span) {
        // A literal expression commonly has its leaf's exact span. Reuse that
        // metadata without another line-index lookup or undo record.
        if matches!(self.arena.get(id), Value::Ready { metadata, .. } if metadata.raw == raw) {
            self.arena.push(id);
            return;
        }
        self.arena.edit(
            id,
            |value| {
                let Value::Ready {
                    expression,
                    metadata,
                    replay,
                } = value
                else {
                    unreachable!()
                };
                let original = Undo::Span(
                    *metadata,
                    expression.span,
                    *replay.last_mut().unwrap().span(),
                );
                metadata.raw = raw;
                expression.span = span(raw);
                *replay.last_mut().unwrap().span() = expression.span;
                original
            },
            |value, undo| {
                let (
                    Value::Ready {
                        expression,
                        metadata,
                        replay,
                    },
                    Undo::Span(old, old_span, event_span),
                ) = (value, undo)
                else {
                    unreachable!()
                };
                *metadata = old;
                expression.span = old_span;
                *replay.last_mut().unwrap().span() = event_span;
            },
        );
        self.arena.push(id);
    }
}
impl Builder for Ast<'_> {
    fn observes_rule(&self, rule: usize) -> bool {
        const LIST: usize = Rule::list_inner as usize;
        const TUPLE: usize = Rule::expression_list as usize;
        const EXPR: usize = Rule::expression as usize;
        const NUMBER: usize = Rule::numeric_literal as usize;
        const STRING: usize = Rule::string_literal as usize;
        const TRUE: usize = Rule::true_kw as usize;
        const FALSE: usize = Rule::false_kw as usize;
        const NONE: usize = Rule::none_kw as usize;
        matches!(
            rule,
            LIST | TUPLE | EXPR | NUMBER | STRING | TRUE | FALSE | NONE
        ) || rule == Rule::postfix as usize || rule == Rule::owned_member_op as usize || rule == Rule::additive as usize || rule == Rule::bitor_expr as usize
            || (rule == Rule::identifier as usize && self.active())
            || Rule::from_capture(grameo::CaptureRule::User(rule)).is_some_and(|rule| {
                matches!(
                    rule,
                    Rule::subscript | Rule::return_stmt
                        | Rule::expr_or_assign_stmt
                        | Rule::expr_or_assign_inline
                        | Rule::return_inline
                        | Rule::pass_stmt
                        | Rule::break_stmt
                        | Rule::break_inline
                        | Rule::continue_stmt
                        | Rule::continue_inline
                        | Rule::global_stmt
                        | Rule::global_inline
                        | Rule::nonlocal_stmt
                        | Rule::nonlocal_inline
                )
            })
    }
    fn supports_pratt_at(&self, rule: usize, start: usize, _: &str) -> bool {
        (rule == Rule::postfix as usize && self.region_candidate(Rule::postfix,start)) || (rule == Rule::additive as usize && self.arithmetic_candidate(start)) || (rule == Rule::bitor_expr as usize && self.region_candidate(Rule::bitor_expr,start))
    }
    fn suppress_pratt_captures(&self, rule: usize) -> bool {
        matches!(Rule::from_capture(grameo::CaptureRule::User(rule)), Some(Rule::additive | Rule::bitor_expr | Rule::postfix))
    }
    fn accepts_pratt(&self, rule: usize, raw: grameo::Span) -> bool {
        arithmetic::complete(self, raw, Rule::from_capture(grameo::CaptureRule::User(rule)).unwrap())
    }
    fn try_pratt_boundary(
        &mut self,
        _: usize,
        end: usize,
    ) -> Result<(), grameo::builder::BuildError> {
        arithmetic::boundary(self, end);
        Ok(())
    }
    fn reduce(&mut self, _: usize, fixity: grameo::pratt::Fixity, op: grameo::Span) {
        if fixity==grameo::pratt::Fixity::Postfix {members::reduce(self,op);return;}
        arithmetic::reduce(self, fixity, op);
    }
    fn opaque_rule(&self, rule: usize) -> bool {
        // Identifier's descendants are only the pure keyword predicate and
        // character classes. Keep its outer capture/finish hook, while letting
        // the certified atomic fragment match the complete lexical token.
        rule == Rule::numeric_literal as usize || rule == Rule::string_literal as usize
            || rule == Rule::identifier as usize
    }
    fn checkpoint(&self) -> usize {
        self.arena.checkpoint()
    }
    fn rollback(&mut self, mark: usize) {
        self.arena.rollback(mark);
    }
    fn begin(&mut self, rule: usize, start: usize) {
        let Some(rule) = Rule::from_capture(grameo::CaptureRule::User(rule)) else {
            return;
        };
        if rule==Rule::owned_member_op {self.arena.push(self.member_operator_frame);return;}
        if rule==Rule::postfix {arithmetic::begin_rule(self,rule,start);return;}
        if rule==Rule::subscript {
            let frame=if self.active() {self.subscript_frame} else {self.depth(true);self.subscript_enter_frame};
            self.arena.push(frame);return;
        }
        if rule == Rule::additive {
            arithmetic::begin(self, start);
            return;
        }
        if rule == Rule::bitor_expr { arithmetic::begin_rule(self, rule, start); return; }
        if matches!(rule, Rule::expr_or_assign_stmt | Rule::expr_or_assign_inline) {
            self.depth(true);
            self.arena.push(if rule == Rule::expr_or_assign_stmt { self.assignment_frame } else { self.assignment_inline_frame });
            return;
        }
        if matches!(
            rule,
            Rule::list_inner | Rule::expression_list | Rule::return_stmt | Rule::return_inline
        ) {
            self.depth(true);
        }
        if self.active()
            && matches!(
                rule,
                Rule::list_inner
                    | Rule::expression_list
                    | Rule::expression
                    | Rule::return_stmt
                    | Rule::return_inline
            )
        {
            let id = if rule == Rule::expression {
                self.expression_frame
            } else if rule == Rule::return_stmt {
                self.return_frame
            } else if rule == Rule::return_inline {
                self.return_inline_frame
            } else if rule == Rule::expression_list {
                self.tuple_frame
            } else {
                self.list_frame
            };
            self.arena.push(id);
        }
    }
    fn finish(&mut self, rule: usize, raw: grameo::Span, source: &str) {
        let Some(rule) = Rule::from_capture(grameo::CaptureRule::User(rule)) else {
            return;
        };
        if rule==Rule::owned_member_op {assert_eq!(self.arena.pop(),Some(self.member_operator_frame));return;}
        if rule==Rule::postfix {arithmetic::finish_rule(self,rule,raw,source);return;}
        if rule==Rule::identifier && self.arena.values().last()==Some(&self.member_operator_frame) {return;}
        if rule==Rule::subscript {subscript::finish(self,raw,source);return;}
        if rule == Rule::additive {
            arithmetic::finish(self, raw, source);
            return;
        }
        if rule == Rule::bitor_expr { arithmetic::finish_rule(self, rule, raw, source); return; }
        if matches!(rule, Rule::expr_or_assign_stmt | Rule::expr_or_assign_inline) {
            assignment::finish(self, rule, raw, source);
            return;
        }
        if let Some(kind) = simple_statement(rule, &source[raw.start..raw.end]) {
            let id = self.arena.alloc(Value::Statement {
                statement: Statement::with_span(kind, span(raw)),
                raw,
                rule,
                child: None,
                replay: Vec::new(),
            });
            self.arena.push(id);
            return;
        }
        if !self.active() {
            return;
        }
        if matches!(rule, Rule::return_stmt | Rule::return_inline) {
            let children = self.frame_children(rule);
            self.depth(false);
            if children.len() == 1 {
                let id = children[0];
                let Value::Ready { metadata, .. } = self.arena.get(id) else {
                    return;
                };
                if source[raw.start..metadata.raw.start].trim() != "return"
                    || !source[metadata.raw.end..raw.end]
                        .trim()
                        .trim_matches(';')
                        .is_empty()
                {
                    return;
                }
                let id = self.arena.unary(
                    id,
                    |value| {
                        let Value::Ready {
                            expression,
                            metadata,
                            replay,
                        } = value
                        else {
                            unreachable!()
                        };
                        Value::Statement {
                            statement: Statement::with_span(
                                StmtKind::Return(Some(expression)),
                                span(raw),
                            ),
                            raw,
                            rule,
                            child: Some(metadata),
                            replay,
                        }
                    },
                    split_return,
                );
                self.arena.push(id);
            } else if children.is_empty()
                && source[raw.start..raw.end]
                    .trim()
                    .trim_end_matches(';')
                    .trim()
                    == "return"
            {
                let id = self.arena.alloc(Value::Statement {
                    statement: Statement::with_span(StmtKind::Return(None), span(raw)),
                    raw,
                    rule,
                    child: None,
                    replay: Vec::new(),
                });
                self.arena.push(id);
            }
            return;
        }
        if let Some(kind) = leaf(rule, &source[raw.start..raw.end]) {
            let leaf_span = span(raw);
            let id = self.arena.alloc(Value::Ready {
                expression: Expression::with_span(kind, leaf_span),
                metadata: Metadata {
                    raw,
                    anchor: None,
                    anchor_rule: rule,
                    sequence: false,
                    contextual: rule == Rule::identifier,
                    layer: 0,
                    boundary: raw.end,
                },
                replay: vec![Replay::Leaf {
                    rule,
                    raw,
                    span: leaf_span,
                }],
            });
            self.arena.push(id);
            return;
        }
        if !matches!(
            rule,
            Rule::list_inner
                | Rule::expression_list
                | Rule::expression
                | Rule::return_stmt
                | Rule::return_inline
        ) {
            return;
        }
        if rule == Rule::expression {
            // A completed exact-span child needs no span/context rewrite.
            // Remove only its immutable wrapper frame, retaining the child
            // with one undo entry instead of pop/pop/push.
            let values = self.arena.values();
            if values.len() >= 2
                && values[values.len() - 2] == self.expression_frame
                && matches!(self.arena.get(values[values.len() - 1]), Value::Ready { metadata, .. } if metadata.raw == raw)
            {
                self.arena.remove_value_below_top();
                return;
            }
            let Some(id) = self.single_child(rule) else {
                return;
            };
            let Value::Ready {
                metadata,
                expression,
                ..
            } = self.arena.get(id)
            else {
                return;
            };
            if metadata.raw.start < raw.start || metadata.raw.end > raw.end {
                return;
            }
            let before = source[raw.start..metadata.raw.start].trim();
            let after = source[metadata.raw.end..raw.end].trim();
            let simple = before.is_empty() && after.is_empty();
            let list = rule == Rule::expression
                && matches!(expression.kind, ExprKind::Array(_))
                && before == "["
                && after == "]";
            let grouping = !before.is_empty()
                && before.bytes().all(|byte| byte == b'(')
                && after.bytes().all(|byte| byte == b')')
                && before.len() == after.len();
            // Contextual grouping has an additional postfix normalization
            // boundary. Keep its outer capture until that boundary is owned.
            // Its completed inner expression-list can still be deferred.
            if grouping && metadata.contextual {
                return;
            }
            if simple || list || grouping {
                self.wrap(id, raw);
            }
            return;
        }
        let children = self.frame_children(rule);
        if matches!(rule, Rule::list_inner | Rule::expression_list) {
            self.depth(false);
            let mut cursor = raw.start;
            let mut valid = true;
            for (index, id) in children.iter().enumerate() {
                let Value::Ready { metadata, .. } = self.arena.get(*id) else {
                    valid = false;
                    break;
                };
                if metadata.raw.start < cursor || metadata.raw.end > raw.end
                {
                    valid = false;
                    break;
                }
                let gap = source[cursor..metadata.raw.start].trim();
                if if index == 0 {
                    !gap.is_empty()
                } else {
                    gap != ","
                } {
                    valid = false;
                    break;
                }
                cursor = metadata.raw.end;
            }
            let tail = source[cursor..raw.end].trim();
            if !valid || (!tail.is_empty() && tail != ",") {
                return;
            }
            // One expression without a comma preserves its child span.
            if rule == Rule::expression_list && children.len() == 1 && tail.is_empty() {
                let id = children[0];
                self.arena.edit(
                    id,
                    |value| {
                        let Value::Ready { metadata, .. } = value else {
                            unreachable!()
                        };
                        let previous = *metadata;
                        metadata.anchor = Some(raw);
                        metadata.anchor_rule = rule;
                        metadata.sequence = false;
                        Undo::Anchor(previous)
                    },
                    |value, undo| {
                        let (Value::Ready { metadata, .. }, Undo::Anchor(previous)) = (value, undo)
                        else {
                            unreachable!()
                        };
                        *metadata = previous;
                    },
                );
                self.arena.push(id);
                return;
            }
            let tuple = rule == Rule::expression_list;
            let id = self.arena.list_with_metadata(
                children,
                |values| {
                    let mut expressions = Vec::with_capacity(values.len());
                    let mut metadata = Vec::with_capacity(values.len());
                    let mut replay = Vec::new();
                    let mut contextual = false;
                    for value in values {
                        let Value::Ready {
                            expression,
                            metadata: child,
                            replay: events,
                        } = value
                        else {
                            unreachable!()
                        };
                        metadata.push((child, events.len()));
                        contextual |= child.contextual;
                        expressions.push(expression);
                        replay.extend(events);
                    }
                    replay.push(Replay::List {
                        count: expressions.len(),
                        span: span(raw),
                        tuple,
                    });
                    (
                        Value::Ready {
                            expression: Expression::with_span(
                                if tuple {
                                    ExprKind::Tuple(expressions)
                                } else {
                                    array(expressions)
                                },
                                span(raw),
                            ),
                            metadata: Metadata {
                                raw,
                                anchor: Some(raw),
                                anchor_rule: rule,
                                sequence: tuple,
                                contextual,
                                layer: 0,
                                boundary: raw.end,
                            },
                            replay,
                        },
                        Undo::List(metadata),
                    )
                },
                split_list,
            );
            self.arena.push(id);
        }
    }
    fn prune_capture_children(&self, rule: usize, raw: grameo::Span) -> bool {
        self.arena
            .values()
            .last()
            .is_some_and(|id| match self.arena.get(*id) {
                Value::Ready { metadata, .. } => {
                    (rule == Rule::additive as usize
                        || rule == Rule::bitor_expr as usize
                        || rule == Rule::postfix as usize
                        || rule == Rule::subscript as usize
                        || rule == Rule::list_inner as usize
                        || rule == Rule::expression_list as usize)
                        && metadata.anchor == Some(raw)
                        && metadata.anchor_rule as usize == rule
                }
                Value::Statement {
                    raw: found,
                    rule: found_rule,
                    ..
                } => *found == raw && *found_rule as usize == rule,
                Value::Assignment { raw: found, rule: found_rule, .. } => *found == raw && *found_rule as usize == rule,
                _ => false,
            })
    }
    fn committed(&mut self, _: &str, _: usize) {
        for value in std::mem::take(&mut self.arena).into_live_nodes() {
            match value {
                Value::Ready {
                    expression,
                    metadata:
                        Metadata {
                            anchor: Some(raw),
                            anchor_rule,
                            sequence,
                            contextual,
                            ..
                        },
                    replay,
                } => {
                    self.expressions.insert(
                        (anchor_rule, raw.start, raw.end),
                        Cached {
                            deferred: None,
                            first: Some(expression),
                            replay,
                            sequence,
                            contextual,
                        },
                    );
                }
                Value::Statement {
                    statement,
                    raw,
                    rule,
                    replay,
                    child,
                } => {
                    self.statements.insert(
                        (raw.start, raw.end),
                        CachedStatement {
                            first: Some(statement),
                            raw,
                            rule,
                            replay,
                            sequence: child.is_some_and(|metadata| metadata.sequence),
                            contextual: child.is_some_and(|metadata| metadata.contextual),
                            assignment_modes: Vec::new(),
                        },
                    );
                }
                Value::Assignment { statement, raw, rule, modes, replay } => {
                    self.statements.insert((raw.start, raw.end), CachedStatement {
                        first: Some(statement), raw, rule, replay,
                        sequence: false, contextual: true, assignment_modes: modes,
                    });
                }
                _ => {}
            }
        }
    }
}
// Preserve the established postfix/member normalization on the owned operands.
// A tuple made by expression_list is a sequence root: only its individual
// expression children passed through postfix. A scalar root passed through it
// as a whole. Each outer walk_expression reinstated its original source span.
fn normalize_expression_list(
    walker: &mut PyWalker,
    mut expression: Expression,
    sequence: bool,
) -> Expression {
    fn operand(walker: &mut PyWalker, expression: Expression) -> Expression {
        let span = expression.span;
        let mut normalized = desugar_member_reads(walker, expression);
        normalized.span = span;
        normalized
    }
    if sequence {
        let ExprKind::Tuple(items) = &mut expression.kind else {
            unreachable!()
        };
        for item in items {
            let original = std::mem::replace(item, Expression::null());
            *item = operand(walker, original);
        }
        expression
    } else {
        operand(walker, expression)
    }
}

// These certified trees contain only lexical leaves, arithmetic and owned
// containers. Normalize child expressions in source order before the enclosing
// postfix boundary, just as the contextual capture walker did.
fn normalize_contextual_expression(walker: &mut PyWalker, expression: Expression) -> Expression {
    let original_span = expression.span;
    if matches!(expression.kind,ExprKind::Member{..}) {return members::normalize(walker,expression);}
    let kind = match expression.kind {
        ExprKind::Array(mut elements) => {
            for element in &mut elements {
                let value = std::mem::replace(&mut element.value, Expression::null());
                element.value = normalize_contextual_expression(walker, value);
            }
            ExprKind::Array(elements)
        }
        ExprKind::Tuple(mut items) => {
            for item in &mut items {
                let value = std::mem::replace(item, Expression::null());
                *item = normalize_contextual_expression(walker, value);
            }
            ExprKind::Tuple(items)
        }
        kind => return arithmetic::normalize(walker, Expression::with_span(kind, original_span)),
    };
    let mut normalized = desugar_member_reads(walker, Expression::with_span(kind, original_span));
    normalized.span = original_span;
    normalized
}
fn normalize_contextual_list(
    walker: &mut PyWalker,
    mut expression: Expression,
    sequence: bool,
    list_inner: bool,
) -> Expression {
    if sequence {
        let ExprKind::Tuple(items) = &mut expression.kind else { unreachable!() };
        for item in items {
            let value = std::mem::replace(item, Expression::null());
            *item = normalize_contextual_expression(walker, value);
        }
        expression
    } else if list_inner {
        let ExprKind::Array(elements) = &mut expression.kind else { unreachable!() };
        for element in elements {
            let value = std::mem::replace(&mut element.value, Expression::null());
            element.value = normalize_contextual_expression(walker, value);
        }
        expression
    } else {
        normalize_contextual_expression(walker, expression)
    }
}
pub(super) fn take_statement_for_pair(
    walker: &mut PyWalker,
    pair: &Pair<Rule>,
) -> Option<Result<Statement, String>> {
    let input = walker.direct_input?;
    let raw = pair.as_span();
    if pair.as_str().as_ptr() as usize != input.0 + raw.start() {
        return None;
    }
    let cached = walker
        .direct_statements
        .get_mut(&(raw.start(), raw.end()))?;
    let sequence = cached.sequence;
    let contextual = cached.contextual;
    let assignment_modes = cached.assignment_modes.clone();
    let mut statement = cached.take(pair.as_str());
    if !assignment_modes.is_empty() {
        return Some(assignment::normalize(walker, statement, &assignment_modes));
    }
    if let StmtKind::Return(value) = &mut statement.kind {
        if let Some(expression) = value.take() {
            *value = Some(if contextual {
                normalize_contextual_list(walker, expression, sequence, false)
            } else {
                normalize_expression_list(walker, expression, sequence)
            });
        }
    }
    Some(Ok(statement))
}
pub(super) fn take_expression_for_pair(
    walker: &mut PyWalker,
    pair: &Pair<Rule>,
) -> Option<Result<Expression, String>> {
    let input = walker.direct_input?;
    let raw = pair.as_span();
    if pair.as_str().as_ptr() as usize != input.0 + raw.start() {
        return None;
    }
    let cached = walker
        .direct_lists
        .get_mut(&(pair.as_rule(), raw.start(), raw.end()))?;
    if let Some(recipe) = cached.deferred.clone() {
        return Some(recipe.resolve(walker, pair));
    }
    let sequence = cached.sequence;
    let contextual = cached.contextual;
    let expression = cached.take_expression(pair.as_str(), raw.start());
    Some(Ok(if contextual && matches!(pair.as_rule(), Rule::additive | Rule::bitor_expr) {
        arithmetic::normalize(walker, expression)
    } else if contextual {
        normalize_contextual_list(walker, expression, sequence, pair.as_rule() == Rule::list_inner)
    } else if pair.as_rule() == Rule::expression_list {
        normalize_expression_list(walker, expression, sequence)
    } else {
        expression
    }))
}
pub(super) fn take_for_pair(walker: &mut PyWalker, pair: &Pair<Rule>) -> Option<Result<ExprKind, String>> {
    take_expression_for_pair(walker, pair).map(|expression| expression.map(|expression| expression.kind))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn expression(source: &str, direct: bool) -> (Expression, usize) {
        set_line_index(source);
        let mut walker = PyWalker::default();
        let (pairs, count) = if direct {
            let mut builder = Ast::new(source);
            let tree = crate::native_grammar::Parser
                .build_captures(Rule::expression, source, &mut builder)
                .unwrap();
            walker.direct_lists = builder.finish();
            walker.direct_input = Some((source.as_ptr() as usize, source.len()));
            let count = walker.direct_lists.len();
            (tree.into_typed_pairs(), count)
        } else {
            (PythonParser::parse(Rule::expression, source).unwrap(), 0)
        };
        (
            walk_expression(&mut walker, pairs.into_iter().next().unwrap()).unwrap(),
            count,
        )
    }

    fn scoped_module_snapshot(mut module: Result<Module,String>) -> String {
        if let Ok(module)=&mut module {
            for statement in &mut module.body {
                if statement.span.start_line==0
                    && let StmtKind::Assign{targets,value,..}=&mut statement.kind
                    && targets.len()==1
                    && matches!(&targets[0].kind,ExprKind::Ident(name) if name=="__py_sys_modules")
                    && let ExprKind::Object(entries)=&mut value.kind
                {entries.sort_by_key(|entry|format!("{entry:?}"));}
            }
        }
        format!("{module:?}")
    }

    #[test]
    fn owned_members_preserve_ordered_scope_and_postfix_boundaries() {
        for source in ["obj.field","obj.a.b","obj . field . name","a.b+c.d","a.b ** 2","True.x","None.x","f().field","obj.field()","obj.field[0]","obj['x'].field","(obj).field"] {
            assert_eq!(format!("{:?}",expression(source,false).0),format!("{:?}",expression(source,true).0),"{source}");
        }
        for source in [
            "class C:\n    pass\nx=C.__name__\ny=C.__mro__\n",
            "import json\nx=json.decoder.JSONDecoder\n",
            "import sys as _sys\nx=(_sys.maxsize*2+1)\n",
            "import re as _re\nx=_re.VERBOSE\n",
            "class C:\n    def f(self):\n        return self.field\n",
            "def gen(a):\n    yield a.field\n",
            "a=bytearray(b'xy')\nx=a.hex\n",
            "from enum import IntFlag\nclass F(IntFlag):\n    A=1\nx=F.A.value\n",
        ] {assert_eq!(scoped_module_snapshot(super::super::parse(source)),scoped_module_snapshot(super::super::parse_with_direct_literals(source)),"{source}");}
        let source="C.__name__";
        set_line_index(source);
        let mut ast=Ast::new(source);
        let pair=crate::native_grammar::Parser.build_captures(Rule::expression,source,&mut ast).unwrap().into_typed_pairs().into_iter().next().unwrap();
        let mut walker=PyWalker::default();
        walker.direct_lists=ast.finish();
        walker.direct_input=Some((source.as_ptr() as usize,source.len()));
        for defined in [true,false,true] {
            if defined {walker.py_defined_classes.insert("C".into());} else {walker.py_defined_classes.remove("C");}
            let owned=walk_expression(&mut walker,pair.clone()).unwrap();
            let mut reference=PyWalker::default();
            if defined {reference.py_defined_classes.insert("C".into());}
            let reference=walk_expression(&mut reference,PythonParser::parse(Rule::expression,source).unwrap().into_iter().next().unwrap()).unwrap();
            assert_eq!(format!("{owned:?}"),format!("{reference:?}"));
        }
        let source="f._state";
        set_line_index(source);
        let mut ast=Ast::new(source);
        let pair=crate::native_grammar::Parser.build_captures(Rule::expression,source,&mut ast).unwrap().into_typed_pairs().into_iter().next().unwrap();
        let mut walker=PyWalker::default();walker.py_deleted_names.insert("f".into());
        walker.direct_lists=ast.finish();walker.direct_input=Some((source.as_ptr() as usize,source.len()));
        let owned=walk_expression(&mut walker,pair).unwrap();
        let mut reference=PyWalker::default();reference.py_deleted_names.insert("f".into());
        let reference=walk_expression(&mut reference,PythonParser::parse(Rule::expression,source).unwrap().into_iter().next().unwrap()).unwrap();
        assert_eq!(format!("{owned:?}"),format!("{reference:?}"));
        let (_,owned)=expression("obj.field",true);assert!(owned>0);
        for source in ["obj.field","obj.a.b","obj . field","obj.field()","obj.class"] {
            set_line_index(source);
            for limit in 1..128 {
                for options in [grameo::MatchOptions{max_steps:limit,..Default::default()},grameo::MatchOptions{max_rule_depth:limit,..Default::default()}] {
                    let mut ast=Ast::new(source);
                    let result=grameo::engine::build_compact_capture_program(&crate::native_grammar::Parser,"expression",source,options,&mut ast);
                    if result.is_err() {assert!(ast.arena.values().is_empty(),"{source} limit={limit}");assert!(!ast.active());}
                }
            }
        }
    }

    #[test]
    fn pratt_selection_requires_an_operator_and_keeps_plain_owned_names() {
        for source in ["x", "12.5", "value,", "value if flag else other"] {
            assert!(!arithmetic::candidate(0, source), "{source}");
        }
        for source in ["x+y", "a*b", "a/b+c", "x+1 if flag else y"] {
            assert!(arithmetic::candidate(0, source), "{source}");
        }
        for source in ["x=1\ny=x\nz=y+1\n", "int_alias=int\nx=int_alias('2')\n", "def f(x):\n    return x\n", "value=[x,y,2]\n", "value=f(x,y=2)\n"] {
            assert_eq!(format!("{:?}", super::super::parse(source)), format!("{:?}", super::super::parse_with_direct_literals(source)), "{source}");
        }
    }

    #[test]
    fn atom_only_matching_keeps_peg_limits_predicates_trivia_and_rollback() {
        use grameo::{MatchOptions, ParseErrorKind};
        struct CaptureOnly;
        impl Builder for CaptureOnly {
            fn checkpoint(&self) -> usize { 0 }
            fn rollback(&mut self, _: usize) {}
            fn begin(&mut self, _: usize, _: usize) {}
            fn finish(&mut self, _: usize, _: grameo::Span, _: &str) {}
        }
        let parser = crate::native_grammar::Parser;
        for source in ["x", "12.5", "value ", "True", "None", "lambda", "x y"] {
            set_line_index(source);
            for steps in 1..96 {
                let options = MatchOptions { max_steps: steps, ..Default::default() };
                let reference = grameo::engine::build_compact_capture_program(&parser, "expression", source, options, &mut CaptureOnly).map(|_| ());
                let mut builder = Ast::new(source);
                let owned = grameo::engine::build_compact_capture_program(&parser, "expression", source, options, &mut builder).map(|_| ());
                assert_eq!(reference, owned, "{source} steps={steps}");
                if owned.is_err() { assert!(builder.arena.values().is_empty()); assert!(!builder.active()); }
            }
            for depth in 1..32 {
                let options = MatchOptions { max_rule_depth: depth, ..Default::default() };
                let reference = grameo::engine::build_compact_capture_program(&parser, "expression", source, options, &mut CaptureOnly).map(|_| ());
                let mut builder = Ast::new(source);
                let owned = grameo::engine::build_compact_capture_program(&parser, "expression", source, options, &mut builder).map(|_| ());
                assert_eq!(reference, owned, "{source} depth={depth}");
            }
        }
        for source in ["x+y", "x+1 ", "x+"] {
            for options in [MatchOptions { max_steps: 1, ..Default::default() }, MatchOptions { max_rule_depth: 1, ..Default::default() }] {
                let mut builder = Ast::new(source);
                let error = grameo::engine::build_compact_capture_program(&parser, "expression", source, options, &mut builder).unwrap_err();
                assert!(matches!(error.kind, ParseErrorKind::WorkLimit | ParseErrorKind::DepthLimit));
                assert!(builder.arena.values().is_empty());
                assert!(!builder.active());
            }
        }
    }

    #[test]
    fn exact_expression_wrapper_keeps_child_and_rolls_back_frame() {
        let source = "1";
        set_line_index(source);
        let mut builder = Ast::new(source);
        builder.depth(true);
        builder.arena.push(builder.expression_frame);
        Builder::finish(&mut builder, Rule::numeric_literal as usize, grameo::Span { start: 0, end: 1 }, source);
        let child = *builder.arena.values().last().unwrap();
        let before = builder.arena.values().to_vec();
        let checkpoint = builder.arena.checkpoint();
        Builder::finish(&mut builder, Rule::expression as usize, grameo::Span { start: 0, end: 1 }, source);
        assert_eq!(builder.arena.values(), &[child]);
        builder.arena.rollback(checkpoint);
        assert_eq!(builder.arena.values(), before);
        Builder::finish(&mut builder, Rule::expression as usize, grameo::Span { start: 0, end: 1 }, source);
        assert_eq!(builder.arena.values(), &[child]);
    }

    #[test]
    fn nested_region_hint_cache_preserves_pure_selection_across_rollback() {
        for source in ["a+b*(c+d)", "f(a+1, b*c)", "1 # comment\n", "α+1", "[a+b,c*d]"] {
            let mut builder = Ast::new(source);
            let checkpoint = builder.arena.checkpoint();
            for start in 0..=source.len() {
                assert_eq!(builder.arithmetic_candidate(start), arithmetic::candidate(start, source));
            }
            builder.arena.rollback(checkpoint);
            for start in (0..=source.len()).rev() {
                assert_eq!(builder.arithmetic_candidate(start), arithmetic::candidate(start, source));
            }
            assert!(builder.arena.values().is_empty());
        }
    }

    #[test]
    fn pratt_arithmetic_owns_operands_and_keeps_ordered_python_lowering() {
        for source in [
            "1+2*3",
            "a+b*c",
            "a*b+c*d",
            "a-b-c",
            "a/b//c%d",
            "x@y+z",
            "1 + 2 * 3 ",
            "a + b + c",
            "a * b * c + d",
            "x+1 if flag else y+2",
            "f(a+1)",
            "(a+1)*2",
            "[a+1,2]",
            "[[a+1],2]",
            "a+[1,2]",
            "a**b+1",
            "-a+1",
            "obj.value+1",
            "items[0]+1",
            "f()+1",
            "NotImplemented",
            "True+1",
            "None+1",
            "1/0",
            "1//0",
            "1%0",
            "1j+2",
            "1j*2",
            "a+=1",
            "a +",
            "a *",
            "a + # comment\n 1",
        ] {
            set_line_index(source);
            let baseline = PythonParser::parse(Rule::expression, source);
            let mut builder = Ast::new(source);
            let direct = crate::native_grammar::Parser.build_captures(
                Rule::expression,
                source,
                &mut builder,
            );
            match (baseline, direct) {
                (Ok(baseline), Ok(direct)) => {
                    let mut bw = PyWalker::default();
                    let expected = walk_expression(&mut bw, baseline.into_iter().next().unwrap());
                    let mut dw = PyWalker::default();
                    dw.direct_lists = builder.finish();
                    dw.direct_input = Some((source.as_ptr() as usize, source.len()));
                    let actual = walk_expression(
                        &mut dw,
                        direct.into_typed_pairs().into_iter().next().unwrap(),
                    );
                    assert_eq!(format!("{expected:?}"), format!("{actual:?}"), "{source}");
                }
                (Err(expected), Err(actual)) => assert_eq!(expected, actual, "{source}"),
                _ => panic!("acceptance mismatch: {source}"),
            }
        }
        let (_, count) = expression("a+b*c", true);
        assert_eq!(count, 1, "one completed owned arithmetic anchor");
    }

    #[test]
    fn bitwise_pratt_owns_scope_neutral_operands_with_ordered_lowering() {
        for source in [
            "a|b^c&d<<2+e*3", "a<<b<<c", "a>>b+c", "a|b|c", "a&b&c",
            "True|False", "1|2^3&4", "a| (b+c)", "-a|b", "a**2|b",
            "f()|x", "obj.field|x", "[a|b,c^d]", "a|b if flag else c&d",
            "a||b", "a&&b", "a|=b", "a<<", "a|#comment\n b",
        ] {
            set_line_index(source);
            let expected = PythonParser::parse(Rule::expression, source);
            let mut builder = Ast::new(source);
            let actual = crate::native_grammar::Parser.build_captures(Rule::expression, source, &mut builder);
            match (expected, actual) {
                (Ok(expected), Ok(actual)) => {
                    let mut ew = PyWalker::default();
                    let expected = walk_expression(&mut ew, expected.into_iter().next().unwrap());
                    let mut aw = PyWalker::default();
                    aw.direct_input = Some((source.as_ptr() as usize, source.len()));
                    aw.direct_lists = builder.finish();
                    let actual = walk_expression(&mut aw, actual.into_typed_pairs().into_iter().next().unwrap());
                    assert_eq!(format!("{expected:?}"),format!("{actual:?}"),"{source}");
                }
                (Err(expected),Err(actual)) => assert_eq!(expected,actual,"{source}"),
                _ => panic!("acceptance mismatch {source}"),
            }
        }
        assert_eq!(expression("a|b^c&d",true).1,1,"one owned bitwise anchor");
    }

    #[test]
    fn bitwise_scope_and_budget_rollback_are_preserved() {
        for source in [
            "from enum import IntFlag\nclass F(IntFlag):\n    A=1\n    B=2\na=F.A\nb=F.B\nc=a|b\n",
            "a={'x':1}\nb={'y':2}\nc=a|b\n",
            "a={1,2}\nb={2,3}\nc=a^b\n",
            "from collections import Counter\na=Counter({'x':2})\nb=Counter({'x':1})\nc=a|b\n",
            "def outer(x):\n    def inner(y):\n        nonlocal x\n        x=x|y\n        return x&y\n    return inner\n",
            "def gen(x):\n    yield x|2\n",
        ] {
            assert_eq!(format!("{:?}",super::super::parse(source)),format!("{:?}",super::super::parse_with_direct_literals(source)),"{source}");
        }
        let parser=crate::native_grammar::Parser;
        for source in ["a|b^c&d<<2+e*3", "a | b ", "a<<", "a|b||c", "[a|b,c&d]"] {
            set_line_index(source);
            for limit in 1..96 {
                for options in [grameo::MatchOptions { max_steps:limit,..Default::default() },grameo::MatchOptions { max_rule_depth:limit,..Default::default() }] {
                    let mut ast=Ast::new(source);
                    let result=grameo::engine::build_compact_capture_program(&parser,"expression",source,options,&mut ast);
                    if result.is_err() { assert!(ast.arena.values().is_empty(),"{source} limit={limit}"); assert!(!ast.active()); }
                }
            }
            let mut ast=Ast::new(source);
            let checkpoint=ast.checkpoint();
            for start in 0..=source.len() { assert_eq!(ast.region_candidate(Rule::bitor_expr,start),arithmetic::bitwise_candidate(start,source)); }
            ast.rollback(checkpoint);
            assert!(ast.arena.values().is_empty());
        }
    }

    #[test]
    fn scalar_subscripts_defer_index_protocol_and_preserve_legacy_slices() {
        for source in [
            "x=a[0]\n", "x=a[i]\n", "x=a[i+1]\n", "x=a[i|j]\n", "x=a['key']\n",
            "x=a[1,]\n", "x=a[1,2]\n", "x=a[(i,j)]\n", "x=a[[1,2]]\n",
            "x=a[-1]\n", "x=a[1:2]\n", "x=a[':']\n", "x=a[f()]\n",
            "a={'x':1}\ni='x'\nx=a[i]\n", "from typing import List\nx=List[int]\n",
            "class C:\n    def __getitem__(self,index):\n        return index\na=C()\nx=a[2]\n",
            "a=[1,2]\na[0]=3\nx=a[0]\n", "def gen(a,i):\n    yield a[i]\n",
        ] {assert_eq!(format!("{:?}",super::super::parse(source)),format!("{:?}",super::super::parse_with_direct_literals(source)),"{source}");}
        let source="a[i+1]";set_line_index(source);
        let mut ast=Ast::new(source);
        crate::native_grammar::Parser.build_captures(Rule::expression,source,&mut ast).unwrap();
        let cache=ast.finish();
        assert!(cache.keys().any(|(rule,_,_)|*rule==Rule::subscript));
        for source in ["a[i+1]","a[0]","a[1:2]","a[i,]","a[]"] {
            set_line_index(source);
            for limit in 1..96 {
                for options in [grameo::MatchOptions {max_steps:limit,..Default::default()},grameo::MatchOptions {max_rule_depth:limit,..Default::default()}] {
                    let mut ast=Ast::new(source);
                    let result=grameo::engine::build_compact_capture_program(&crate::native_grammar::Parser,"expression",source,options,&mut ast);
                    if result.is_err() {assert!(ast.arena.values().is_empty(),"{source} limit={limit}");assert!(!ast.active());}
                }
            }
        }
    }

    #[test]
    fn subscript_capture_coverage() {
        struct WithoutSubscripts<'a,'i>(&'a mut Ast<'i>);
        impl Builder for WithoutSubscripts<'_,'_> {
            fn observes_rule(&self,r:usize)->bool { !matches!(Rule::from_capture(grameo::CaptureRule::User(r)),Some(Rule::subscript)) && self.0.observes_rule(r) }
            fn checkpoint(&self)->usize {self.0.checkpoint()}
            fn rollback(&mut self,m:usize){self.0.rollback(m)}
            fn begin(&mut self,r:usize,o:usize){self.0.begin(r,o)}
            fn finish(&mut self,r:usize,s:grameo::Span,t:&str){self.0.finish(r,s,t)}
            fn committed(&mut self,s:&str,o:usize){self.0.committed(s,o)}
            fn supports_pratt_at(&self,r:usize,o:usize,s:&str)->bool{self.0.supports_pratt_at(r,o,s)}
            fn accepts_pratt(&self,r:usize,s:grameo::Span)->bool{self.0.accepts_pratt(r,s)}
            fn suppress_pratt_captures(&self,r:usize)->bool{self.0.suppress_pratt_captures(r)}
            fn opaque_rule(&self,r:usize)->bool{self.0.opaque_rule(r)}
            fn prune_capture_children(&self,r:usize,s:grameo::Span)->bool{self.0.prune_capture_children(r,s)}
            fn try_pratt_boundary(&mut self,r:usize,o:usize)->Result<(),grameo::builder::BuildError>{self.0.try_pratt_boundary(r,o)}
            fn reduce(&mut self,r:usize,f:grameo::pratt::Fixity,s:grameo::Span){self.0.reduce(r,f,s)}
        }
        fn count(pair:Pair<'_,Rule>,args:&mut usize)->usize {
            if pair.as_rule()==Rule::subscript {*args+=1;}
            1+pair.into_inner().map(|pair|count(pair,args)).sum::<usize>()
        }
        for (label,path) in [
            ("list_index","tests/python/data_structures/list_index_access.py"),
            ("linecache","tests/python/python_linecache_cache_ops/test_linecache_cache_entry_structure.py"),
            ("getitem","tests/python/oop_inheritance_extended/dunder_getitem.py"),
        ] {
            let source=preprocess_indentation(&std::fs::read_to_string(path).unwrap());
            set_line_index(&source);
            let mut before=Ast::new(&source);
            let tree=crate::native_grammar::Parser.build_captures(Rule::program,&source,&mut WithoutSubscripts(&mut before)).unwrap();
            let mut matched=0;
            let before_nodes=tree.into_typed_pairs().into_iter().map(|pair|count(pair,&mut matched)).sum::<usize>();
            let mut after=Ast::new(&source);
            let tree=crate::native_grammar::Parser.build_captures(Rule::program,&source,&mut after).unwrap();
            let after_nodes=tree.into_typed_pairs().into_iter().map(|pair|count(pair,&mut 0)).sum::<usize>();
            let owned=after.finish().keys().filter(|(rule,_,_)|*rule==Rule::subscript).count();
            println!("{label} matched_subscripts={matched} owned_subscripts={owned} before_capture_nodes={before_nodes} after_capture_nodes={after_nodes}");
            assert!(after_nodes<=before_nodes);
        }
    }

    #[test]
    fn member_capture_coverage() {
        struct WithoutMembers<'a,'i>(&'a mut Ast<'i>);
        impl Builder for WithoutMembers<'_,'_> {
            fn observes_rule(&self,r:usize)->bool { !matches!(Rule::from_capture(grameo::CaptureRule::User(r)),Some(Rule::postfix)) && self.0.observes_rule(r) }
            fn checkpoint(&self)->usize {self.0.checkpoint()}
            fn rollback(&mut self,m:usize){self.0.rollback(m)}
            fn begin(&mut self,r:usize,o:usize){self.0.begin(r,o)}
            fn finish(&mut self,r:usize,s:grameo::Span,t:&str){self.0.finish(r,s,t)}
            fn committed(&mut self,s:&str,o:usize){self.0.committed(s,o)}
            fn supports_pratt_at(&self,r:usize,o:usize,s:&str)->bool{r!=Rule::postfix as usize && self.0.supports_pratt_at(r,o,s)}
            fn accepts_pratt(&self,r:usize,s:grameo::Span)->bool{self.0.accepts_pratt(r,s)}
            fn suppress_pratt_captures(&self,r:usize)->bool{self.0.suppress_pratt_captures(r)}
            fn opaque_rule(&self,r:usize)->bool{self.0.opaque_rule(r)}
            fn prune_capture_children(&self,r:usize,s:grameo::Span)->bool{self.0.prune_capture_children(r,s)}
            fn try_pratt_boundary(&mut self,r:usize,o:usize)->Result<(),grameo::builder::BuildError>{self.0.try_pratt_boundary(r,o)}
            fn reduce(&mut self,r:usize,f:grameo::pratt::Fixity,s:grameo::Span){self.0.reduce(r,f,s)}
        }
        fn count(pair:Pair<'_,Rule>,args:&mut usize)->usize {
            if pair.as_rule()==Rule::postfix_chain && pair.clone().into_inner().next().is_some_and(|p|p.as_rule()==Rule::identifier) {*args+=1;}
            1+pair.into_inner().map(|pair|count(pair,args)).sum::<usize>()
        }
        for (label,path) in [
            ("numbers","tests/python/python_numbers_abstract_hierarchy/test_numbers_subclass_integral_implementation.py"),
            ("property","tests/python/python_property_advanced/test_property_getter.py"),
            ("pydoc","tests/python/python_pydoc_documentation_renderer/test_pydoc_render_doc_text_mode.py"),
        ] {
            let source=preprocess_indentation(&std::fs::read_to_string(path).unwrap());
            set_line_index(&source);
            let mut before=Ast::new(&source);
            let tree=crate::native_grammar::Parser.build_captures(Rule::program,&source,&mut WithoutMembers(&mut before)).unwrap();
            let mut matched=0;
            let before_nodes=tree.into_typed_pairs().into_iter().map(|pair|count(pair,&mut matched)).sum::<usize>();
            let mut after=Ast::new(&source);
            let tree=crate::native_grammar::Parser.build_captures(Rule::program,&source,&mut after).unwrap();
            let after_nodes=tree.into_typed_pairs().into_iter().map(|pair|count(pair,&mut 0)).sum::<usize>();
            let owned=after.expressions.values().map(|cached|cached.replay.iter().filter(|r|matches!(r,Replay::Member{..})).count()).sum::<usize>()
                + after.statements.values().map(|cached|cached.replay.iter().filter(|r|matches!(r,Replay::Member{..})).count()).sum::<usize>();
            let before_statements=before.statements.len();
            let after_statements=after.statements.len();
            let before_returns=before.statements.values().filter(|c|matches!(c.rule,Rule::return_stmt | Rule::return_inline)).count();
            let after_returns=after.statements.values().filter(|c|matches!(c.rule,Rule::return_stmt | Rule::return_inline)).count();
            println!("{label} matched_member_operations={matched} owned_member_operations={owned} before_capture_nodes={before_nodes} after_capture_nodes={after_nodes} before_owned_statements={before_statements} after_owned_statements={after_statements} before_owned_returns={before_returns} after_owned_returns={after_returns}");
            assert!(after_nodes<=before_nodes);
        }
    }

    #[test]
    fn literal_lists_keep_values_and_all_expression_spans() {
        for source in [
            "[]",
            "[1,2,3]",
            "[True,False,None]",
            "[1,(2,3),[4,(5,)]]",
            "(1,2,3)",
            "(1,)",
            "((1,2),[3,4])",
            "[1.5,1e2,0xff,1j]",
            "['a', r'\\n', b'abc']",
            "[1,[2,'x'],[]]",
            "[ 1 , 2 , ]",
            "[[[[[1]]]]]",
            "['a' 'b', 1]",
            "[-1,2]",
            "[x+1,[2,3]]",
            "[f([1,2]),[3]]",
            "[*items,[1]]",
            "[x for x in [1,2]]",
        ] {
            let (baseline, _) = expression(source, false);
            let (direct, _) = expression(source, true);
            assert_eq!(format!("{baseline:?}"), format!("{direct:?}"), "{source}");
        }
        assert_eq!(expression("[1,[2,'x'],[]]", true).1, 1);
    }

    #[test]
    fn contextual_containers_and_returns_defer_ordered_scope_lowering() {
        for source in [
            "[a,b,c]", "[a+b,c*d,[x,y+2]]", "[a,1j,b'abc',(1,2)]",
            "(a,b+1)", "[(a+b),c]", "([a+b])", "[(a,b),c]",
        ] {
            let (expected, _) = expression(source, false);
            let (actual, _) = expression(source, true);
            assert_eq!(format!("{expected:?}"), format!("{actual:?}"), "{source}");
        }
        assert_eq!(expression("[a+b,c*d,[x,y+2]]", true).1, 1,
            "one complete owned list replaces descendant arithmetic captures");
        for source in [
            "x='outer'\ndef f(x,y):\n    return x+y\n",
            "def f(x,y):\n    return [x+y,x,[y*2]]\n",
            "def f(x,y):\n    return x,y+1,[x*2]\n",
            "def outer():\n    x=1\n    def inner():\n        nonlocal x\n        x=x+1\n        return x+2\n    return inner\n",
            "def gen(x):\n    yield x\n    return [x+1,x]\n",
        ] {
            let expected = super::super::parse(source);
            let actual = super::super::parse_with_direct_literals(source);
            assert_eq!(format!("{expected:?}"), format!("{actual:?}"), "{source}");
        }
        let source = "return [x+1,y,[z*2]]\n";
        let prepared = preprocess_indentation(source);
        set_line_index(&prepared);
        let mut builder = Ast::new(&prepared);
        let tree = crate::native_grammar::Parser.build_captures(Rule::program, &prepared, &mut builder).unwrap();
        let mut walker = PyWalker::default();
        walker.direct_input = Some((prepared.as_ptr() as usize, prepared.len()));
        walker.direct_statements = builder.take_statements();
        walker.direct_lists = builder.finish();
        assert_eq!(walker.direct_statements.len(), 1, "whole contextual return is owned");
        fn find(pair: Pair<'_, Rule>) -> Option<Pair<'_, Rule>> {
            if pair.as_rule() == Rule::return_stmt { return Some(pair); }
            pair.into_inner().find_map(find)
        }
        let pair = tree.into_typed_pairs().into_iter().find_map(find).unwrap();
        let first = take_statement_for_pair(&mut walker, &pair).unwrap().unwrap();
        let repeated = take_statement_for_pair(&mut walker, &pair).unwrap().unwrap();
        assert_eq!(format!("{first:?}"), format!("{repeated:?}"));
    }

    #[test]
    fn owned_name_assignments_keep_binding_facts_chains_and_scope() {
        for source in [
            "x = 'operator **= text'\n",
            "class C:\n    pattern = 'operator |= text'\n",
            "x=1\ny=x+2\nz=[x,y,y*2]\n",
            "a=b=2\nc=a+b\n",
            "a,b=1,2\nc=a+b\n",
            "s='a'\nt=s\nu=t+'b'\n",
            "int_alias=int\nx=int_alias('2')\n",
            "x=1\ndef f(x):\n    y=x+2\n    return [y,x]\n",
            "x=1\ndef f():\n    global x\n    x=x+1\n    return x\n",
            "def outer():\n    x=1\n    def inner():\n        nonlocal x\n        x=x+1\n        return x\n    return inner\n",
            "class C:\n    field=2\n    def value(self):\n        x=3\n        return x\n",
            "x: int=2\na=[1]\na[0]=3\nx+=4\n",
        ] {
            let expected = super::super::parse(source);
            let actual = super::super::parse_with_direct_literals(source);
            assert_eq!(format!("{expected:?}"), format!("{actual:?}"), "{source}");
        }
        let source = "x=1\ny=x+2\na=b=3\nc,d=x,y\n";
        let prepared = preprocess_indentation(source);
        set_line_index(&prepared);
        let mut builder = Ast::new(&prepared);
        let tree = crate::native_grammar::Parser.build_captures(Rule::program, &prepared, &mut builder).unwrap();
        let mut walker = PyWalker::default();
        walker.direct_input = Some((prepared.as_ptr() as usize, prepared.len()));
        walker.direct_statements = builder.take_statements();
        walker.direct_lists = builder.finish();
        assert_eq!(walker.direct_statements.len(), 4, "complete assignment statements are owned");
        fn find<'i>(pair: Pair<'i, Rule>, out: &mut Vec<Pair<'i, Rule>>) {
            if pair.as_rule() == Rule::expr_or_assign_stmt { out.push(pair); return; }
            for pair in pair.into_inner() { find(pair, out); }
        }
        let mut pairs = Vec::new();
        for pair in tree.into_typed_pairs() { find(pair, &mut pairs); }
        for pair in pairs {
            let first = take_statement_for_pair(&mut walker, &pair).unwrap().unwrap();
            let repeated = take_statement_for_pair(&mut walker, &pair).unwrap().unwrap();
            assert_eq!(format!("{first:?}"), format!("{repeated:?}"));
        }
    }

    #[test]
    fn repeated_capture_reads_replay_without_an_ast_snapshot() {
        let source = "[1,[2,'x'],True]";
        set_line_index(source);
        let mut builder = Ast::default();
        crate::native_grammar::Parser
            .build_captures(Rule::expression, source, &mut builder)
            .unwrap();
        let mut cache = builder.finish();
        let ((_, start, end), cached) = cache.iter_mut().next().unwrap();
        let text = &source[*start..*end];
        let first = cached.take(text, *start);
        assert!(cached.first.is_none());
        let second = cached.take(text, *start);
        let third = cached.take(text, *start);
        assert_eq!(format!("{first:?}"), format!("{second:?}"));
        assert_eq!(format!("{second:?}"), format!("{third:?}"));
    }

    #[test]
    fn failed_program_restores_owned_values_and_can_retry() {
        let invalid = "x = [1,[2,'x']]\ndef f(\n";
        let prepared = preprocess_indentation(invalid);
        set_line_index(&prepared);
        let mut builder = Ast::default();
        let initial = builder.checkpoint();
        assert!(
            crate::native_grammar::Parser
                .build_captures(Rule::program, &prepared, &mut builder)
                .is_err()
        );
        assert_eq!(builder.checkpoint(), initial);
        assert!(!builder.active());
        assert!(builder.arena.values().is_empty());
        let source = "[1,2]";
        set_line_index(source);
        crate::native_grammar::Parser
            .build_captures(Rule::expression, source, &mut builder)
            .unwrap();
        assert_eq!(builder.finish().len(), 1);
    }
    #[test]
    fn complete_returns_and_scope_statements_move_owned_common_ast() {
        let source = "global items, other\ndef f():\n    return [1,(2,3),[4]]\npass\ndef g():\n    return x + 1\n";
        let prepared = preprocess_indentation(source);
        set_line_index(&prepared);
        let mut builder = Ast::default();
        crate::native_grammar::Parser
            .build_captures(Rule::program, &prepared, &mut builder)
            .unwrap();
        let mut statements = builder.take_statements();
        assert_eq!(
            statements.len(),
            3,
            "global, complete return, and pass own their AST"
        );
        for ((start, end), cached) in &mut statements {
            let first = cached.take(&prepared[*start..*end]);
            assert!(cached.first.is_none());
            let second = cached.take(&prepared[*start..*end]);
            assert_eq!(format!("{first:?}"), format!("{second:?}"));
        }
    }

    #[test]
    fn module_lists_preserve_comments_scopes_and_directives() {
        fn canonical(module: &mut Module) {
            for stmt in &mut module.body {
                if let StmtKind::Assign { targets, value, .. } = &mut stmt.kind
                    && targets.len() == 1
                    && matches!(&targets[0].kind, ExprKind::Ident(name) if name == "__py_sys_modules")
                    && let ExprKind::Object(entries) = &mut value.kind
                {
                    entries.sort_by_key(|entry| format!("{entry:?}"));
                }
                stmt.walk_exprs_mut(&mut |expr| {
                    if let ExprKind::Call { callee, args, .. } = &mut expr.kind
                        && matches!(&callee.kind, ExprKind::Ident(name) if name == "__py_copy_stamp_fields")
                        && let Some(arg) = args.get_mut(1)
                        && let ExprKind::Array(elements) = &mut arg.value.kind {
                        elements.sort_by_key(|element| format!("{:?}", element.value));
                    }
                });
            }
        }
        for source in [
            "x = [1, # comment\n 2]\n",
            "x=(1,2,3)\ndef f():\n    return 1,2,3\n",
            "def f():\n    global x,y\n    return [1,(2,3)]\n",
            "def f(): global x,y; return (1,2)\n",
            "def f():\n    global x, \\\n        y\n    return 1\n",
            "def f():\n    return\n",
            "for x in [1,2]:\n    if x: break\n    else: continue\npass\n",
            "x = [1, [2, # nested comment\n 3]]\n",
            "x=0\ndef outer():\n    global x\n    items=[1,[2,3]]\n    def inner():\n        nonlocal items\n        items=[4,5]\n        return items\n    return inner\n",
            "class C:\n    def method(self):\n        global values\n        values=[1,2]\n        return values\n",
            "x=[1,2]\ndef f(\n",
        ] {
            match (parse(source), parse_with_direct_literals(source)) {
                (Ok(mut baseline), Ok(mut direct)) => {
                    canonical(&mut baseline);
                    canonical(&mut direct);
                    assert_eq!(format!("{baseline:?}"), format!("{direct:?}"), "{source}");
                }
                (Err(baseline), Err(direct)) => assert_eq!(baseline, direct, "{source}"),
                _ => panic!("acceptance mismatch: {source}"),
            }
        }
    }    #[test]
    fn opaque_identifier_lexer_preserves_scope_keywords_trivia_and_rollback() {
        for input in [
            "classify=1\nTrue_=2\nassert_=3\nx=classify+True_+assert_\n",
            "x=1\n# line\nif x:\n    x=x+1\n",
            "def f(async_, global_, return_):\n    return async_+global_+return_\n",
            "class C:\n    __namex__=1\nx=C.__name__\n",
            "from collections import Counter as Counter_\nx=Counter_()\n",
            "x=1\ndel x\ny=x\n",
            "x = name \\\n + other\n", "x = Falsex\n", "x=True_\n", "x = class\n",
        ] {
            assert_eq!(scoped_module_snapshot(super::super::parse(input)),scoped_module_snapshot(super::super::parse_with_direct_literals(input)),"{input}");
        }
        for input in ["name", "classify", "class", "True_", "__name__", "a.b", "a+b", "a + b ", "a+", "assert name\n", "# comment\nx=name\n"] {
            set_line_index(input);
            for limit in 1..128 {
                for options in [grameo::MatchOptions{max_steps:limit,..Default::default()},grameo::MatchOptions{max_rule_depth:limit,..Default::default()}] {
                    let mut ast=Ast::new(input);
                    let entry=if input.contains('\n') {"program"} else {"expression"};
                    let result=grameo::engine::build_compact_capture_program(&crate::native_grammar::Parser,entry,input,options,&mut ast);
                    if result.is_err() {assert!(ast.arena.values().is_empty(),"{input} limit={limit}");assert!(!ast.active());}
                }
            }
        }
    }

}
