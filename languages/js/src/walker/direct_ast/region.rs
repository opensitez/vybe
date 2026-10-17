//! Owned grammar regions. A frame accepts only its exact, complete children;
//! unsupported children publish a gap and retain their contextual PEG captures.
use super::*;

#[derive(Clone, Copy)]
pub(super) struct HeaderMeta {
    pub expression: Option<Metadata>,
}
pub(super) fn cache_header(
    expression: Option<Expression>,
    meta: HeaderMeta,
    caches: &mut body::Caches,
) {
    if let (
        Some(expression),
        Some(Metadata {
            raw,
            anchor: Some(anchor),
            facts,
            rule,
            ..
        }),
    ) = (expression, meta.expression)
    {
        caches.expressions.insert(
            (anchor.start, anchor.end),
            Cached::new_at(expression, raw, anchor, facts, rule),
        );
    }
}
#[derive(Clone, Copy)]
pub(super) struct CatchMeta {
    pub block: body::StatementMeta,
}
#[derive(Clone, Copy)]
pub(super) struct FinallyMeta {
    pub block: body::StatementMeta,
}
pub(super) enum ChildMeta {
    Expression(Metadata),
    Statement(body::StatementMeta),
    Body(body::BodyMeta),
    Name(grameo::Span),
    Declarator(HeaderMeta),
    Param(HeaderMeta),
    Params(grameo::Span, Vec<HeaderMeta>),
    Catch(CatchMeta),
    Finally(FinallyMeta),
}
enum Payload {
    Expression(Expression),
    Statement(Statement),
    Body(Vec<Statement>),
    Name(String),
    Declarator(VarDeclarator),
    Param(Param),
    Params(Vec<Param>),
    Catch(CatchClause),
    Finally(Vec<Statement>),
}
impl ChildMeta {
    fn facts(&self) -> u8 {
        match self {
            Self::Expression(meta) => meta.facts,
            Self::Statement(meta) => meta.facts,
            Self::Body(meta) => meta.facts,
            Self::Declarator(meta) | Self::Param(meta) => {
                meta.expression.map_or(0, |meta| meta.facts)
            }
            Self::Params(_, metas) => metas.iter().fold(0, |facts, meta| {
                facts | meta.expression.map_or(0, |meta| meta.facts)
            }),
            Self::Catch(meta) => meta.block.facts,
            Self::Finally(meta) => meta.block.facts,
            Self::Name(_) => 0,
        }
    }
    fn of(value: &Value) -> Option<Self> {
        Some(match value {
            Value::Ready(_, meta) => Self::Expression(*meta),
            Value::Statement(_, meta) => Self::Statement(*meta),
            Value::Body(_, raw) => Self::Body(*raw),
            Value::Name(_, raw) => Self::Name(*raw),
            Value::Declarator(_, raw) => Self::Declarator(*raw),
            Value::Param(_, raw) => Self::Param(*raw),
            Value::Params(_, raw, metas) => Self::Params(*raw, metas.clone()),
            Value::Catch(_, meta) => Self::Catch(*meta),
            Value::Finally(_, meta) => Self::Finally(*meta),
            _ => return None,
        })
    }
    fn restore(self, payload: Payload) -> Value {
        match (self, payload) {
            (Self::Expression(meta), Payload::Expression(value)) => Value::Ready(value, meta),
            (Self::Statement(meta), Payload::Statement(value)) => {
                Value::Statement(Box::new(value), meta)
            }
            (Self::Body(raw), Payload::Body(value)) => Value::Body(value, raw),
            (Self::Name(raw), Payload::Name(value)) => Value::Name(value, raw),
            (Self::Declarator(raw), Payload::Declarator(value)) => Value::Declarator(value, raw),
            (Self::Param(raw), Payload::Param(value)) => Value::Param(value, raw),
            (Self::Params(raw, metas), Payload::Params(value)) => Value::Params(value, raw, metas),
            (Self::Catch(meta), Payload::Catch(value)) => Value::Catch(value, meta),
            (Self::Finally(meta), Payload::Finally(value)) => Value::Finally(value, meta),
            _ => unreachable!("region inverse preserves child kind"),
        }
    }
}
fn payload(value: Value) -> Payload {
    match value {
        Value::Ready(value, _) => Payload::Expression(value),
        Value::Statement(value, _) => Payload::Statement(*value),
        Value::Body(value, _) => Payload::Body(value),
        Value::Name(value, _) => Payload::Name(value),
        Value::Declarator(value, _) => Payload::Declarator(value),
        Value::Param(value, _) => Payload::Param(value),
        Value::Params(value, _, _) => Payload::Params(value),
        Value::Catch(value, _) => Payload::Catch(value),
        Value::Finally(value, _) => Payload::Finally(value),
        _ => unreachable!("validated region child"),
    }
}
fn expression(value: Payload) -> Expression {
    let Payload::Expression(value) = value else {
        unreachable!()
    };
    value
}
fn statement(value: Payload) -> Statement {
    let Payload::Statement(value) = value else {
        unreachable!()
    };
    value
}
fn name(value: Payload) -> String {
    let Payload::Name(value) = value else {
        unreachable!()
    };
    value
}

// Only call callee primaries need a new name reducer. Scalar arithmetic
// already observes its leaves. Unsupported member/unary/container primaries
// stay captured, avoiding owned leaf construction that cannot complete them.
pub(super) const IGNORED_CALL: usize=Rule::EOI as usize+1;
pub(super) fn primary_enabled<const S:bool,const B:bool>(ast:&IslandAst<'_,S,B>)->bool {
    ast.arena.values().last().is_some_and(|id|matches!(ast.arena.get(*id),Value::RegionStart(rule) if *rule==Rule::call_expression as usize))
}
pub(super) fn call_candidate(start:usize,source:&str)->bool {
    let bytes=source.as_bytes();let mut end=start;
    if !bytes.get(end).is_some_and(|b|b.is_ascii_alphabetic()||matches!(b,b'_'|b'$')||*b>=128) {return false;}
    // A permissive eligibility check, not a lexer: grammar ownership still
    // certifies the complete identifier, source range, and ordinary call.
    while bytes.get(end).is_some_and(|b|b.is_ascii_alphanumeric()||matches!(b,b'_'|b'$')||*b>=128) {end+=1;}
    let keyword=matches!(&source[start..end],"this"|"super"|"true"|"false"|"null"|"undefined");
    while bytes.get(end).is_some_and(u8::is_ascii_whitespace) {end+=1;}
    matches!(bytes.get(end),Some(b'('|b'.'|b'[')) || keyword
}
fn frame_matches(rule:usize,stored:usize)->bool {
    stored==rule || rule==Rule::call_expression as usize && stored==IGNORED_CALL
}
/// Names are semantic children only at a declaration/binding/jump boundary.
/// Observing every identifier would allocate discarded name values throughout
/// unsupported calls, properties and contextual expressions.
pub(super) fn needs_name<const S: bool, const B: bool>(ast: &IslandAst<'_, S, B>) -> bool {
    ast.arena.values().last().is_some_and(|id| matches!(ast.arena.get(*id), Value::RegionStart(rule)
        if *rule == Rule::primary as usize || *rule == Rule::binding_pattern as usize || *rule == Rule::function_declaration as usize
            || *rule == Rule::async_function_declaration as usize || function_expression(*rule) || *rule == Rule::break_statement as usize
            || *rule == Rule::continue_statement as usize))
}
const RULE_FLAGS: [u8; Rule::EOI as usize + 1] = {
    let mut flags = [0; Rule::EOI as usize + 1];
    let frames = [
        Rule::function_body,
        Rule::block_statement,
        Rule::if_statement,
        Rule::while_statement,
        Rule::do_while_statement,
        Rule::variable_declaration,
        Rule::var_declarator,
        Rule::binding_pattern,
        Rule::function_declaration,
        Rule::async_function_declaration,
        Rule::function_expression,
        Rule::async_function_expression,
        Rule::param,
        Rule::param_list,
        Rule::expression,
        Rule::assignment_expression,
        Rule::throw_statement,
        Rule::break_statement,
        Rule::continue_statement,
        Rule::try_statement,
        Rule::catch_clause,
        Rule::finally_clause,
        Rule::comparison,
        Rule::logical_expr,
        Rule::primary,
        Rule::call_expression,
        Rule::call_chain,
        Rule::array_literal,
        Rule::array_element,
    ];
    let statements = [
        Rule::block_statement,
        Rule::if_statement,
        Rule::while_statement,
        Rule::do_while_statement,
        Rule::variable_declaration,
        Rule::function_declaration,
        Rule::async_function_declaration,
        Rule::throw_statement,
        Rule::break_statement,
        Rule::continue_statement,
        Rule::try_statement,
    ];
    let bodies = [
        Rule::expression,
        Rule::expression_statement,
        Rule::return_statement,
        Rule::statement,
        Rule::function_body,
        Rule::empty_statement,
        Rule::debugger_statement,
    ];
    let mut index = 0;
    while index < frames.len() {
        flags[frames[index] as usize] |= 1;
        index += 1;
    }
    index = 0;
    while index < statements.len() {
        flags[statements[index] as usize] |= 2;
        index += 1;
    }
    index = 0;
    while index < bodies.len() {
        flags[bodies[index] as usize] |= 4;
        index += 1;
    }
    flags[Rule::generator_marker as usize] |= 8;
    flags[Rule::ident_or_keyword as usize] |= 8;
    flags[Rule::array_elision as usize] |= 8;
    flags
};
fn flagged(rule: usize, flag: u8) -> bool {
    RULE_FLAGS.get(rule).is_some_and(|flags| flags & flag != 0)
}
fn function_expression(rule: usize) -> bool {
    rule == Rule::function_expression as usize || rule == Rule::async_function_expression as usize
}
pub(super) fn framed(rule: usize) -> bool {
    flagged(rule, 1)
}
pub(super) fn observes(rule: usize) -> bool {
    flagged(rule, 1 | 8)
}
pub(super) fn body_observes(rule: usize) -> bool {
    flagged(rule, 4)
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
            Value::Statement(_, meta) => meta.raw == raw && matches_rule_statement(rule),
            Value::Body(_, anchor) => anchor.raw == raw && rule == Rule::function_body as usize,
            // Header components stay captured until a complete declaration is owned.
            Value::Ready(_, meta) => {
                (compound::is_rule(rule) || rule == Rule::primary as usize || rule == Rule::call_expression as usize || rule == Rule::array_literal as usize || function_expression(rule)) && meta.rule as usize == rule && meta.anchor == Some(raw)
            }
            _ => false,
        })
}
fn matches_rule_statement(rule: usize) -> bool {
    flagged(rule, 2)
}
fn valid<const S: bool, const B: bool>(
    ast: &IslandAst<'_, S, B>,
    rule: usize,
    ids: &[Id],
    raw: grameo::Span,
    source: &str,
) -> bool {
    let get = |index: usize| ids.get(index).map(|id| ast.arena.get(*id));
    let children = (ids.len(), get(0), get(1), get(2));
    if rule == Rule::primary as usize {
        return (ids.is_empty() && matches!(source[raw.start..raw.end].trim(),"this"|"super"|"true"|"false"|"null"|"undefined")) || matches!(children,(1,Some(Value::Name(name,anchor)),_,_)
            if *anchor==raw && !matches!(name.as_str(),"globalThis"|"this"|"super"|"true"|"false"|"null"|"undefined"));
    }
    if rule == Rule::call_expression as usize {return chain::valid(ast,ids,raw,source);}
    if rule == Rule::expression as usize || rule == Rule::assignment_expression as usize {
        return matches!(children, (1,Some(Value::Ready(_,meta)),_,_) if meta.raw.start == raw.start
            && meta.raw.end <= raw.end && source[meta.raw.end..raw.end].trim().is_empty());
    }
    if rule == Rule::function_body as usize || rule == Rule::block_statement as usize {
        return ids
            .iter()
            .all(|id| matches!(ast.arena.get(*id), Value::Statement(_, _)));
    }
    if rule == Rule::binding_pattern as usize {
        return matches!(children, (1,Some(Value::Name(_,anchor)),_,_) if *anchor == raw);
    }
    if rule == Rule::var_declarator as usize {
        return matches!(
            children,
            (1, Some(Value::Name(_, _)), _, _)
                | (2, Some(Value::Name(_, _)), Some(Value::Ready(_, _)), _)
        ) && (ids.len() == 2 || !source[raw.start..raw.end].contains('='));
    }
    if rule == Rule::variable_declaration as usize {
        let is_const = source[raw.start..raw.end].starts_with("const");
        return !ids.is_empty() && ids.iter().all(|id| matches!(ast.arena.get(*id), Value::Declarator(decl, _) if !is_const || decl.init.is_some()));
    }
    if rule == Rule::param as usize {
        return matches!(
            children,
            (1, Some(Value::Name(_, _)), _, _)
                | (2, Some(Value::Name(_, _)), Some(Value::Ready(_, _)), _)
        );
    }
    if rule == Rule::param_list as usize {
        return ids
            .iter()
            .all(|id| matches!(ast.arena.get(*id), Value::Param(_, _)));
    }
    if rule == Rule::function_declaration as usize
        || rule == Rule::async_function_declaration as usize
    {
        return matches!(children, (2,Some(Value::Name(_, _)),Some(Value::Body(_, _)),_)
            | (3,Some(Value::Name(_, _)),Some(Value::Params(_, _, _)),Some(Value::Body(_, _))))
            // Preserve scope-aware discovery for every positive yield spelling.
            && !source[raw.start..raw.end].contains("yield");
    }
    if function_expression(rule) {
        return matches!(children, (1,Some(Value::Body(_, _)),_,_)
            | (2,Some(Value::Params(_, _, _)),Some(Value::Body(_, _)),_))
            && !source[raw.start..raw.end].contains("yield");
    }
    if rule == Rule::if_statement as usize {
        return matches!(
            children,
            (2, Some(Value::Ready(_, _)), Some(Value::Statement(_, _)), _)
                | (
                    3,
                    Some(Value::Ready(_, _)),
                    Some(Value::Statement(_, _)),
                    Some(Value::Statement(_, _))
                )
        );
    }
    if rule == Rule::while_statement as usize {
        return matches!(
            children,
            (2, Some(Value::Ready(_, _)), Some(Value::Statement(_, _)), _)
        );
    }
    if rule == Rule::do_while_statement as usize {
        return matches!(
            children,
            (2, Some(Value::Statement(_, _)), Some(Value::Ready(_, _)), _)
        );
    }
    if rule == Rule::throw_statement as usize {
        return matches!(children, (1, Some(Value::Ready(_, _)), _, _));
    }
    if rule == Rule::break_statement as usize || rule == Rule::continue_statement as usize {
        return matches!(children, (0, _, _, _) | (1, Some(Value::Name(_, _)), _, _));
    }
    if rule == Rule::catch_clause as usize {
        return matches!(children, (1,Some(Value::Statement(stmt, _)),_,_) | (2,Some(Value::Name(_, _)),Some(Value::Statement(stmt, _)),_) if matches!(stmt.kind, StmtKind::Block(_)));
    }
    if rule == Rule::finally_clause as usize {
        return matches!(children, (1,Some(Value::Statement(stmt, _)),_,_) if matches!(stmt.kind, StmtKind::Block(_)));
    }
    if rule == Rule::try_statement as usize {
        return matches!(children, (1,Some(Value::Statement(stmt, _)),_,_)
            | (2,Some(Value::Statement(stmt, _)),Some(Value::Catch(_, _)),_)
            | (2,Some(Value::Statement(stmt, _)),Some(Value::Finally(_, _)),_)
            | (3,Some(Value::Statement(stmt, _)),Some(Value::Catch(_, _)),Some(Value::Finally(_, _))) if matches!(stmt.kind, StmtKind::Block(_)));
    }
    false
}
pub(super) fn finish<const S: bool, const B: bool>(
    ast: &mut IslandAst<'_, S, B>,
    rule: usize,
    raw: grameo::Span,
    source: &str,
) {
    if matches!(rule,r if r==Rule::array_literal as usize || r==Rule::array_element as usize || r==Rule::array_elision as usize) {container::finish(ast,rule,raw,source);return;}
    if rule==Rule::ident_or_keyword as usize {chain::finish_property(ast,raw);return;}
    if rule==Rule::call_chain as usize {chain::finish_link(ast,raw,source);return;}
    if compound::is_rule(rule) {
        compound::finish(ast, rule, raw, source);
        return;
    }
    if rule == Rule::primary as usize || rule == Rule::call_expression as usize {
        let values=ast.arena.values();
        if values.len()>=2 && matches!(ast.arena.get(values[values.len()-2]),Value::RegionStart(r) if frame_matches(rule,*r))
            && matches!(ast.arena.get(*values.last().unwrap()),Value::Ready(_,meta) if meta.raw.start==raw.start && meta.raw.end<=raw.end && source[meta.raw.end..raw.end].trim().is_empty()) {
            let child=*values.last().unwrap();
            ast.arena.remove_value_below_top().unwrap();
            if rule == Rule::primary as usize {
                if !matches!(ast.arena.get(child),Value::Ready(_,meta) if meta.anchor==Some(raw)) {
                    ast.arena.edit(child,|value| {let Value::Ready(_,meta)=value else{unreachable!()};EditUndo::Anchor(meta.anchor.replace(raw))},|value,undo| {let (Value::Ready(_,meta),EditUndo::Anchor(anchor))=(value,undo) else{unreachable!()};meta.anchor=anchor;});
                }
                let target=span(raw);
                if !matches!(ast.arena.get(child),Value::Ready(expr,meta) if meta.boundary==raw.end && same_span(expr.span,target)) {
                    ast.arena.edit(child,|value| {let Value::Ready(expr,meta)=value else{unreachable!()};let undo=EditUndo::Boundary(meta.boundary,expr.span);meta.boundary=raw.end;expr.span=target;undo},|value,undo| {let (Value::Ready(expr,meta),EditUndo::Boundary(boundary,previous))=(value,undo) else{unreachable!()};meta.boundary=boundary;expr.span=previous;});
                }
            }
            return;
        }
    }
    if rule == Rule::generator_marker as usize {
        ast.arena.push(ast.gap.unwrap());
        return;
    }
    if rule == Rule::expression as usize
        || rule == Rule::assignment_expression as usize
        || rule == Rule::binding_pattern as usize
    {
        // Exact identity wrappers retain the child on the stack. A single
        // reversible removal avoids popping and re-pushing the child ID.
        let values=ast.arena.values();
        if values.len()>=2 && matches!(ast.arena.get(values[values.len()-2]),Value::RegionStart(r) if frame_matches(rule,*r)) {
            let id=*values.last().unwrap();
            if valid(ast,rule,&[id],raw,source) {
                ast.arena.remove_value_below_top().unwrap();
                if rule!=Rule::binding_pattern as usize {
                    let target=span(raw);
                    if !matches!(ast.arena.get(id),Value::Ready(expr,meta) if meta.boundary==raw.end && same_span(expr.span,target)) {
                        ast.arena.edit(id,|value| {
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
        }
        let mut single = None;
        let mut count = 0;
        while let Some(id) = ast.arena.pop() {
            if matches!(ast.arena.get(id), Value::RegionStart(r) if frame_matches(rule,*r)) {
                break;
            }
            single = Some(id);
            count += 1;
        }
        if count != 1 {
            ast.arena.push(ast.gap.unwrap());
            return;
        }
        let id = single.unwrap();
        if !valid(ast, rule, &[id], raw, source) {
            ast.arena.push(ast.gap.unwrap());
            return;
        }
        // These wrappers retain the same owned operand. Its output span is changed
        // journaledly; lexical/raw spans remain available for exact replay.
        if rule == Rule::expression as usize || rule == Rule::assignment_expression as usize {
            let target=span(raw);
            if matches!(ast.arena.get(id),Value::Ready(expression,meta) if meta.boundary==raw.end && same_span(expression.span,target)) {
                ast.arena.push(id);return;
            }
            ast.arena.edit(
                id,
                |value| {
                    let Value::Ready(expression, meta) = value else {
                        unreachable!()
                    };
                    let undo = EditUndo::Boundary(meta.boundary, expression.span);
                    expression.span = target;
                    meta.boundary = raw.end;
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
            ast.arena.push(id);
            return;
        }
        ast.arena.push(id);
        return;
    }
    let mut ids = Vec::new();
    while let Some(id) = ast.arena.pop() {
        if matches!(ast.arena.get(id), Value::RegionStart(r) if frame_matches(rule,*r)) {
            break;
        }
        ids.push(id);
    }
    ids.reverse();
    if !valid(ast, rule, &ids, raw, source) {
        ast.arena.push(ast.gap.unwrap());
        return;
    }
    if rule==Rule::call_expression as usize {
        let id=ast.arena.list_with_metadata(ids,|values|chain::construct(values,raw,source),chain::split);
        ast.arena.push(id);return;
    }
    let id = ast.arena.list_with_metadata(
        ids,
        |values| {
            let metadata: Vec<_> = values
                .iter()
                .map(|value| ChildMeta::of(value).unwrap())
                .collect();
            let value = construct(
                rule,
                values.into_iter().map(payload),
                raw,
                source,
                &metadata,
            );
            (value, Undo::Region(rule, metadata))
        },
        split,
    );
    ast.arena.push(id);
}
fn construct(
    rule: usize,
    values: impl Iterator<Item = Payload>,
    raw: grameo::Span,
    source: &str,
    metadata: &[ChildMeta],
) -> Value {
    let facts = metadata.iter().fold(0, |facts, meta| facts | meta.facts())
        | if rule == Rule::variable_declaration as usize
            && !source[raw.start..raw.end].starts_with("var")
        {
            body::LEXICAL
        } else {
            0
        };
    if rule == Rule::primary as usize {
        let kind=match source[raw.start..raw.end].trim() {
            "this"=>ExprKind::This,"super"=>ExprKind::Super,
            "true"=>ExprKind::Lit(Literal::Bool(true)),"false"=>ExprKind::Lit(Literal::Bool(false)),
            "null"=>ExprKind::Lit(Literal::Null),"undefined"=>ExprKind::Lit(Literal::Undefined),
            _=>ExprKind::Ident(name(values.into_iter().next().unwrap())),
        };
        let facts=if matches!(&kind,ExprKind::Ident(name) if matches!(name.as_str(),"Promise"|"document")) {body::MODEL} else {0};
        return Value::Ready(Expression::with_span(kind,span(raw)),Metadata {
            raw,anchor:Some(raw),layer:0,facts,rule:Rule::primary,boundary:raw.end,
        });
    }
    if rule == Rule::call_expression as usize {
        let mut values=values.into_iter();
        let callee=expression(values.next().unwrap());
        let args=values.map(expression).map(Argument::positional).collect();
        return Value::Ready(Expression::with_span(ExprKind::Call {
            callee:Box::new(callee),args,optional:false,
        },span(raw)),Metadata {raw,anchor:Some(raw),layer:0,facts,rule:Rule::call_expression,boundary:raw.end});
    }
    if function_expression(rule) {
        let mut values=values;
        let (params, body)=match values.next().unwrap() {
            Payload::Body(body)=>(Vec::new(),body),
            Payload::Params(params)=>{
                let Payload::Body(body)=values.next().unwrap() else {unreachable!()};
                (params,body)
            }
            _=>unreachable!(),
        };
        let function=wrap_generator_if_needed(String::new(),params,body,
            rule==Rule::async_function_expression as usize,false);
        return Value::Ready(Expression::with_span(ExprKind::FunctionExpr(Box::new(Statement::new(function))),span(raw)),
            Metadata{raw,anchor:Some(raw),layer:0,facts,
                rule:if rule==Rule::async_function_expression as usize {Rule::async_function_expression}else{Rule::function_expression},boundary:raw.end});
    }
    if rule == Rule::function_body as usize {
        return Value::Body(
            values.into_iter().map(statement).collect(),
            body::BodyMeta { raw, facts },
        );
    }
    let header = HeaderMeta {
        expression: metadata.iter().find_map(|meta| {
            if let ChildMeta::Expression(meta) = meta {
                Some(*meta)
            } else {
                None
            }
        }),
    };
    let mut values = values;
    if rule == Rule::var_declarator as usize {
        return Value::Declarator(
            VarDeclarator {
                pattern: BindingPattern::Ident(name(values.next().unwrap())),
                type_hint: None,
                init: values.next().map(expression),
                array_bounds: None,
                with_events: false,
            },
            header,
        );
    }
    if rule == Rule::param as usize {
        return Value::Param(
            Param {
                name: name(values.next().unwrap()),
                type_hint: None,
                default: values.next().map(expression),
                pass_by: PassBy::Value,
                is_rest: source[raw.start..raw.end].starts_with("..."),
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            },
            header,
        );
    }
    if rule == Rule::param_list as usize {
        return Value::Params(
            values
                .map(|value| {
                    let Payload::Param(value) = value else {
                        unreachable!()
                    };
                    value
                })
                .collect(),
            raw,
            metadata
                .iter()
                .map(|meta| {
                    let ChildMeta::Param(meta) = meta else {
                        unreachable!()
                    };
                    *meta
                })
                .collect(),
        );
    }
    if rule == Rule::catch_clause as usize {
        let (var_name, block) = match values.next().unwrap() {
            Payload::Name(name) => (Some(name), statement(values.next().unwrap())),
            Payload::Statement(block) => (None, block),
            _ => unreachable!(),
        };
        let StmtKind::Block(body) = block.kind else {
            unreachable!()
        };
        let ChildMeta::Statement(block) = metadata.last().unwrap() else {
            unreachable!()
        };
        return Value::Catch(
            CatchClause {
                types: Vec::new(),
                var_name,
                stack_var: None,
                body,
                when_clause: None,
            },
            CatchMeta { block: *block },
        );
    }
    if rule == Rule::finally_clause as usize {
        let StmtKind::Block(body) = statement(values.next().unwrap()).kind else {
            unreachable!()
        };
        let ChildMeta::Statement(block) = metadata[0] else {
            unreachable!()
        };
        return Value::Finally(body, FinallyMeta { block });
    }
    let kind = if rule == Rule::block_statement as usize {
        StmtKind::Block(values.map(statement).collect())
    } else if rule == Rule::variable_declaration as usize {
        let prefix = &source[raw.start..raw.end];
        let kind = if prefix.starts_with("var") {
            VarDeclKind::FunctionScoped
        } else if prefix.starts_with("const") {
            VarDeclKind::Const
        } else {
            VarDeclKind::Let
        };
        StmtKind::VarDecl {
            kind,
            declarations: values
                .map(|value| {
                    let Payload::Declarator(value) = value else {
                        unreachable!()
                    };
                    value
                })
                .collect(),
        }
    } else if rule == Rule::function_declaration as usize
        || rule == Rule::async_function_declaration as usize
    {
        let name = name(values.next().unwrap());
        let next = values.next().unwrap();
        let (params, body) = match next {
            Payload::Body(body) => (Vec::new(), body),
            Payload::Params(params) => {
                let Payload::Body(body) = values.next().unwrap() else {
                    unreachable!()
                };
                (params, body)
            }
            _ => unreachable!(),
        };
        wrap_generator_if_needed(
            name,
            params,
            body,
            rule == Rule::async_function_declaration as usize,
            false,
        )
    } else if rule == Rule::if_statement as usize {
        StmtKind::If {
            cond: expression(values.next().unwrap()),
            then_body: vec![statement(values.next().unwrap())],
            elifs: Vec::new(),
            else_body: values.next().map(|value| vec![statement(value)]),
        }
    } else if rule == Rule::while_statement as usize {
        StmtKind::While {
            cond: expression(values.next().unwrap()),
            body: vec![statement(values.next().unwrap())],
            else_body: None,
        }
    } else if rule == Rule::do_while_statement as usize {
        StmtKind::DoWhile {
            body: vec![statement(values.next().unwrap())],
            cond: expression(values.next().unwrap()),
            until: false,
        }
    } else if rule == Rule::try_statement as usize {
        let StmtKind::Block(body) = statement(values.next().unwrap()).kind else {
            unreachable!()
        };
        let mut catches = Vec::new();
        let mut finally = None;
        for value in values {
            match value {
                Payload::Catch(catch) => catches.push(catch),
                Payload::Finally(body) => finally = Some(body),
                _ => unreachable!(),
            }
        }
        StmtKind::Try {
            body,
            catches,
            finally,
            else_body: None,
        }
    } else if rule == Rule::throw_statement as usize {
        StmtKind::Throw {
            expr: Some(expression(values.next().unwrap())),
            cause: None,
        }
    } else if rule == Rule::break_statement as usize {
        StmtKind::Break(
            values
                .next()
                .map(|value| BreakTarget::Label(name(value)))
                .unwrap_or(BreakTarget::Implicit),
        )
    } else if rule == Rule::continue_statement as usize {
        StmtKind::Continue(
            values
                .next()
                .map(|value| ContinueTarget::Label(name(value)))
                .unwrap_or(ContinueTarget::Implicit),
        )
    } else {
        unreachable!()
    };
    Value::Statement(
        Box::new(Statement::with_span(kind, span(raw))),
        body::StatementMeta {
            raw,
            expression: None,
            facts,
        },
    )
}
fn split(value: Value, undo: Undo) -> Vec<Value> {
    let Undo::Region(rule, metadata) = undo else {
        unreachable!()
    };
    let payloads = match payload(value) {
        Payload::Expression(expression) => {
            if rule==Rule::primary as usize {
                if let ExprKind::Ident(name)=expression.kind {vec![Payload::Name(name)]} else {Vec::new()}
            } else if function_expression(rule) {
                let ExprKind::FunctionExpr(function)=expression.kind else {unreachable!()};
                let StmtKind::FunctionDecl{params,body,..}=function.kind else {unreachable!()};
                let mut values=Vec::new();
                if metadata.len()==2 {values.push(Payload::Params(params));}
                values.push(Payload::Body(body));values
            } else {
                let ExprKind::Call {callee,args,..}=expression.kind else {unreachable!()};
                let mut values=vec![Payload::Expression(*callee)];
                values.extend(args.into_iter().map(|arg|Payload::Expression(arg.value)));
                values
            }
        }
        Payload::Body(body) => body.into_iter().map(Payload::Statement).collect(),
        Payload::Declarator(decl) => {
            let BindingPattern::Ident(name) = decl.pattern else {
                unreachable!()
            };
            let mut values = vec![Payload::Name(name)];
            values.extend(decl.init.map(Payload::Expression));
            values
        }
        Payload::Param(param) => {
            let mut values = vec![Payload::Name(param.name)];
            values.extend(param.default.map(Payload::Expression));
            values
        }
        Payload::Params(params) => params.into_iter().map(Payload::Param).collect(),
        Payload::Catch(catch) => {
            let mut values = Vec::new();
            values.extend(catch.var_name.map(Payload::Name));
            let ChildMeta::Statement(block) = metadata.last().unwrap() else {
                unreachable!()
            };
            values.push(Payload::Statement(Statement::with_span(
                StmtKind::Block(catch.body),
                span(block.raw),
            )));
            values
        }
        Payload::Finally(body) => {
            let ChildMeta::Statement(block) = metadata[0] else {
                unreachable!()
            };
            vec![Payload::Statement(Statement::with_span(
                StmtKind::Block(body),
                span(block.raw),
            ))]
        }
        Payload::Statement(stmt) => match stmt.kind {
            StmtKind::Block(body) => body.into_iter().map(Payload::Statement).collect(),
            StmtKind::VarDecl { declarations, .. } => {
                declarations.into_iter().map(Payload::Declarator).collect()
            }
            StmtKind::FunctionDecl {
                name, params, body, ..
            } => {
                let mut values = vec![Payload::Name(name)];
                if metadata.len() == 3 {
                    values.push(Payload::Params(params));
                }
                values.push(Payload::Body(body));
                values
            }
            StmtKind::If {
                cond,
                then_body,
                else_body,
                ..
            } => {
                let mut values = vec![Payload::Expression(cond)];
                values.extend(then_body.into_iter().map(Payload::Statement));
                values.extend(else_body.into_iter().flatten().map(Payload::Statement));
                values
            }
            StmtKind::While { cond, body, .. } => {
                let mut values = vec![Payload::Expression(cond)];
                values.extend(body.into_iter().map(Payload::Statement));
                values
            }
            StmtKind::DoWhile { cond, body, .. } => {
                let mut values: Vec<_> = body.into_iter().map(Payload::Statement).collect();
                values.push(Payload::Expression(cond));
                values
            }
            StmtKind::Try {
                body,
                catches,
                finally,
                ..
            } => {
                let ChildMeta::Statement(block) = metadata[0] else {
                    unreachable!()
                };
                let mut values = vec![Payload::Statement(Statement::with_span(
                    StmtKind::Block(body),
                    span(block.raw),
                ))];
                values.extend(catches.into_iter().map(Payload::Catch));
                values.extend(finally.map(Payload::Finally));
                values
            }
            StmtKind::Throw { expr, .. } => expr.into_iter().map(Payload::Expression).collect(),
            StmtKind::Break(BreakTarget::Label(name))
            | StmtKind::Continue(ContinueTarget::Label(name)) => vec![Payload::Name(name)],
            StmtKind::Break(BreakTarget::Implicit)
            | StmtKind::Continue(ContinueTarget::Implicit) => Vec::new(),
            _ => unreachable!("owned rule {rule}"),
        },
        _ => unreachable!("owned region result"),
    };
    assert_eq!(metadata.len(), payloads.len());
    metadata
        .into_iter()
        .zip(payloads)
        .map(|(meta, value)| meta.restore(value))
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn collapsed_identity_frames_restore_nested_children_and_exact_spans() {
        for source in ["f(g(1),2) < h(3) && obj.x", "[f(1),...g(2)]", "function(a=1){return f(a)&&g();}"] {
            let _lines=vybe_ast::line_index::LineIndex::install(source);
            let mut builder=BodyAst::new(source);
            builder.retain_test_journal=true;
            let checkpoint=builder.checkpoint();
            for _ in 0..3 {
                super::super::super::super::native_grammar::Parser
                    .build_captures(Rule::expression,source,&mut builder).unwrap();
                builder.rollback(checkpoint);
                assert_eq!(builder.checkpoint(),checkpoint);
                assert!(builder.arena.values().is_empty());
            }
            super::super::super::super::native_grammar::Parser
                .build_captures(Rule::expression,source,&mut builder).unwrap();
            let mut caches=builder.into_caches();
            let first=caches.expressions.get_mut(&(0,source.len())).unwrap_or_else(||panic!("restored owned wrapper root {source}")).take(source);
            let replay=caches.expressions.get_mut(&(0,source.len())).unwrap().take(source);
            assert_eq!(format!("{first:?}"),format!("{replay:?}"),"{source}");
        }
    }
    #[test]
    fn anonymous_function_expressions_move_scopes_and_preserve_fallback() {
        for source in [
            "const f=function(a,b=2,...rest){let x=a+2;return x;};",
            "const f=async function(x=3){if(x)return x+2;return 0;};",
            "const f=function(){return function(a){return a+1;};};",
            "f(function(x){return x+2;}); const a=[function(){return 1;}];",
            "const f=function Named(x){return Named(x+1);};",
            "const f=function Promise(){return document;};",
            "const f=function*(){yield 2;}; const g=async function*(){yield 3;};",
            "const f=function({x}){return x;}; const g=function(){return 'yield';};",
            "const f=function(){let document=2;return document;};",
            "const f=function(){'use strict';return 2;};",
        ] {
            assert_eq!(format!("{:?}",super::super::super::parse(source)),
                format!("{:?}",super::super::super::parse_direct_bodies(source)),"{source}");
        }
        let source="function(a=2,...rest){return a+3;}";
        let _lines=vybe_ast::line_index::LineIndex::install(source);
        let mut builder=BodyAst::new(source);
        let checkpoint=builder.checkpoint();
        for _ in 0..3 {
            let tree=super::super::super::super::native_grammar::Parser
                .build_captures(Rule::function_expression,source,&mut builder).unwrap();
            assert_eq!(tree.capture_count(),1,"complete function owns its entire scope");
            builder.rollback(checkpoint);
        }
        super::super::super::super::native_grammar::Parser
            .build_captures(Rule::function_expression,source,&mut builder).unwrap();
        let mut caches=builder.into_caches();
        let cache=caches.expressions.get_mut(&(0,source.len())).expect("owned function root");
        let first=cache.take(source);
        assert!(matches!(first.kind,ExprKind::FunctionExpr(_)));
        assert_eq!(format!("{first:?}"),format!("{:?}",cache.take(source)));
    }
    #[test]
    fn ordinary_identifier_calls_own_arguments_and_preserve_contextual_fallback() {
        for source in [
            "function f(a){let x=g(a,2+3);if(h(x))return j(k(4));return empty();}",
            "const x=f(1,2,); f(); f(g(1),h(2));",
            "function π(x){return x+1;} const y=π(2); const z=ƒ(3);",
            "const x=f(1) < 2; const y=g() instanceof C; class C extends B {m(){return f() > 3;}}",
            "const x=f(Promise,document); const y=obj.method(f(2));",
            "f(...xs); obj?.method(f(1)); new C(f(2)); f(1)(2);",
            "Array.isArray(f(2)); Intl.DateTimeFormat.supportedLocalesOf(f(2));",
            "const x=f(globalThis); typeof f(2) === \"number\";",
        ] {
            assert_eq!(format!("{:?}",super::super::super::parse(source)),
                format!("{:?}",super::super::super::parse_direct_bodies(source)),"{source}");
        }
        let source="f(1) ";
        let _lines=vybe_ast::line_index::LineIndex::install(source);
        let mut builder=BodyAst::new(source);
        let checkpoint=builder.checkpoint();
        for _ in 0..3 {
            super::super::super::super::native_grammar::Parser
                .build_captures(Rule::additive,source,&mut builder).unwrap();
            builder.rollback(checkpoint);
        }
        super::super::super::super::native_grammar::Parser
            .build_captures(Rule::additive,source,&mut builder).unwrap();
        let mut caches=builder.into_caches();
        let cache=caches.expressions.get_mut(&(0,source.len())).expect("promoted call span includes trivia");
        let first=cache.take(source);
        assert_eq!(format!("{:?}",first.span),format!("{:?}",span(grameo::Span{start:0,end:source.len()})));
        assert_eq!(format!("{first:?}"),format!("{:?}",cache.take(source)));
        let source="f(1,g(2))";
        let _lines=vybe_ast::line_index::LineIndex::install(source);
        let mut builder=BodyAst::new(source);
        let checkpoint=builder.checkpoint();
        for _ in 0..3 {
            let tree=super::super::super::super::native_grammar::Parser
                .build_captures(Rule::call_expression,source,&mut builder).unwrap();
            assert!(tree.capture_count()<10,"complete call prunes children");
            builder.rollback(checkpoint);
        }
        super::super::super::super::native_grammar::Parser
            .build_captures(Rule::call_expression,source,&mut builder).unwrap();
        let mut caches=builder.into_caches();
        let cache=caches.expressions.get_mut(&(0,source.len())).expect("owned call root");
        let first=cache.take(source);
        assert!(matches!(first.kind,ExprKind::Call{..}));
        assert_eq!(format!("{first:?}"),format!("{:?}",cache.take(source)));
    }
    #[test]
    fn scope_and_control_regions_construct_owned_common_ast() {
        for source in [
            "function f(a,b=2,...rest){let x=a+2,y=4;if(x){const z=3;return z;}else return y;}",
            "async function f(x=3){while(x){let y=x+2;break;}do {continue;}while(x);throw x;}",
            "function f(){function g(x){return x+3;}return 4;}",
            "let x=1,y=2; {const z=3; if(x) return z; else throw y;}",
        ] {
            let _lines = vybe_ast::line_index::LineIndex::install(source);
            let mut builder = BodyAst::new(source);
            let tree = super::super::super::super::native_grammar::Parser
                .build_captures(Rule::program, source, &mut builder)
                .unwrap();
            let caches = builder.into_caches();
            assert!(
                !caches.statements.is_empty(),
                "owned root statement: {source}"
            );
            assert!(
                tree.capture_count()
                    < super::super::super::super::native_grammar::Parser
                        .parse(Rule::program, source)
                        .unwrap()
                        .capture_count(),
                "complete statements prune contextual children: {source}"
            );
            assert_eq!(
                format!("{:?}", super::super::super::parse(source)),
                format!("{:?}", super::super::super::parse_direct_bodies(source)),
                "{source}"
            );
        }
    }
    #[test]
    fn incomplete_headers_publish_consumed_expression_anchors_for_fallback() {
        for source in [
            "function f(a=2){return call(a);}",
            "function f(a=2,{x}){return a+3;}",
            "const a=2,b;",
            "const a=2,b=call();",
            "let {x=2}=call();",
            "function* f(a=2){yield a+3;}",
            "class C {m(a=2){let x=3;return a+x;}}",
            "function f(a=2){'yield';return a+3;}",
            "try {let x=2;} catch(e) {return 3;} finally {const y=4;}",
            "async function f(){try {await g();}catch(e){let x=2;return x;}}",
        ] {
            assert_eq!(
                format!("{:?}", super::super::super::parse(source)),
                format!("{:?}", super::super::super::parse_direct_bodies(source)),
                "{source}"
            );
        }
    }
    #[test]
    fn owned_region_list_inverse_restores_headers_and_scopes_repeatedly() {
        let _lines = vybe_ast::line_index::LineIndex::install("let x=2;");
        let mut builder = BodyAst::new("let x=2;");
        let source = "let x=2;";
        let raw = grameo::Span {
            start: 0,
            end: source.len(),
        };
        builder.begin(Rule::variable_declaration as usize, 0);
        let decl = VarDeclarator {
            pattern: BindingPattern::Ident("x".into()),
            type_hint: None,
            init: Some(Expression::int(2)),
            array_bounds: None,
            with_events: false,
        };
        let id = builder
            .arena
            .alloc(Value::Declarator(decl, HeaderMeta { expression: None }));
        builder.arena.push(id);
        let mark = builder.checkpoint();
        let expected = format!(
            "{:?}",
            match builder.arena.get(id) {
                Value::Declarator(d, _) => d,
                _ => unreachable!(),
            }
        );
        for _ in 0..2 {
            builder.finish(Rule::variable_declaration as usize, raw, source);
            assert!(matches!(
                builder.arena.get(*builder.arena.values().last().unwrap()),
                Value::Statement(_, _)
            ));
            builder.rollback(mark);
            assert_eq!(
                format!(
                    "{:?}",
                    match builder.arena.get(id) {
                        Value::Declarator(d, _) => d,
                        _ => unreachable!(),
                    }
                ),
                expected
            );
        }
    }
}
#[cfg(test)]
mod replay_tests {
    use super::*;
    #[test]
    fn complete_scope_regions_move_once_and_replay_full_shapes_and_spans() {
        for source in [
            "function f(a,b=2,...rest){let x=a+2,y=4;if(x){const z=3;return z;}else return y;}",
            "async function f(x=3){while(x){let y=x+2;break;}do {continue;}while(x);throw x;}",
            "try {let x=2;} catch(e) {return document+3;} finally {const y=4;}",
        ] {
            let _lines = vybe_ast::line_index::LineIndex::install(source);
            let mut builder = BodyAst::new(source);
            let tree = super::super::super::super::native_grammar::Parser
                .build_captures(Rule::program, source, &mut builder)
                .unwrap();
            let mut caches = builder.into_caches();
            let pair = tree
                .into_typed_pairs::<Rule>()
                .next()
                .unwrap()
                .into_inner()
                .next()
                .unwrap();
            let raw = pair.as_span();
            let cached = caches
                .statements
                .get_mut(&(raw.start(), raw.end()))
                .expect("complete root is owned");
            let original = JsParser::parse(Rule::program, source)
                .unwrap()
                .next()
                .unwrap()
                .into_inner()
                .next()
                .unwrap();
            let expected =
                super::super::super::walk_statement(&mut JsWalker::default(), original).unwrap();
            for _ in 0..3 {
                let mut names = false;
                let actual = cached.take(source, &mut names);
                assert_eq!(format!("{actual:?}"), format!("{expected:?}"), "{source}");
            }
        }
    }
}
#[cfg(test)]
mod rollback_tests {
    use super::*;
    struct Speculative<'i>(BodyAst<'i>);
    impl Builder for Speculative<'_> {
        fn observes_rule(&self, r: usize) -> bool {
            self.0.observes_rule(r)
        }
        fn supports_pratt_at(&self, r: usize, o: usize, s: &str) -> bool {
            self.0.supports_pratt_at(r, o, s)
        }
        fn try_pratt_boundary(
            &mut self,
            r: usize,
            e: usize,
        ) -> Result<(), grameo::builder::BuildError> {
            self.0.try_pratt_boundary(r, e)
        }
        fn suppress_pratt_captures(&self, r: usize) -> bool {
            self.0.suppress_pratt_captures(r)
        }
        fn accepts_pratt(&self, r: usize, s: grameo::Span) -> bool {
            self.0.accepts_pratt(r, s)
        }
        fn checkpoint(&self) -> usize {
            self.0.checkpoint()
        }
        fn rollback(&mut self, m: usize) {
            self.0.rollback(m)
        }
        fn opaque_rule(&self, r: usize) -> bool {
            self.0.opaque_rule(r)
        }
        fn begin(&mut self, r: usize, o: usize) {
            self.0.begin(r, o)
        }
        fn finish(&mut self, r: usize, s: grameo::Span, i: &str) {
            self.0.finish(r, s, i)
        }
        fn prune_capture_children(&self, r: usize, s: grameo::Span) -> bool {
            self.0.prune_capture_children(r, s)
        }
        fn reduce(&mut self, r: usize, f: grameo::pratt::Fixity, s: grameo::Span) {
            self.0.reduce(r, f, s)
        }
        // Keep the complete inverse chain for this enclosing speculative parse.
        fn committed(&mut self, _: &str, _: usize) {}
    }
    #[test]
    fn entire_nested_declaration_control_and_catch_inverses_restore_repeatedly() {
        for source in [
            "function f(a,b=2,...rest){let x=a+2,y=4;if(x){const z=3;return z;}else return y;}",
            "async function f(x=3){while(x){let y=x+2;break;}do {continue;}while(x);throw x;}",
            "try {let x=2;} catch(e) {return document+3;} finally {const y=4;}",
            "function f(){try {throw 2;} catch {return 3;}finally {4+5;}}",
            "function f(a,b){if(a<b&&b>2)return a+b;while(a<4)break;}",
            "function f(a){const x=-a;return !x;}const y=~f(2);",
            "function f(a){return typeof a;}const y=void [1,2];",
            "function f(a){return typeof a===\"string\";}const y=!-~a;",
        ] {
            let _lines = vybe_ast::line_index::LineIndex::install(source);
            let mut builder = Speculative(BodyAst::new(source));
            let start = builder.checkpoint();
            for _ in 0..3 {
                super::super::super::super::native_grammar::Parser
                    .build_captures(Rule::program, source, &mut builder)
                    .unwrap();
                assert!(matches!(
                    builder
                        .0
                        .arena
                        .get(*builder.0.arena.values().last().unwrap()),
                    Value::Statement(_, _)
                ));
                builder.rollback(start);
                builder.rollback(start);
                assert_eq!(builder.checkpoint(), start);
                assert!(builder.0.arena.values().is_empty());
                assert!(!builder.0.active());
            }
            assert!(builder.0.into_caches().statements.is_empty());
        }
    }
}
