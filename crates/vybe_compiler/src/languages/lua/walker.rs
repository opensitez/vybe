use super::{LuaParser, Rule};
use crate::ast::*;
use pest::Parser;
use pest::iterators::Pair;

fn to_span(pair: &Pair<Rule>) -> Span {
    let (start_line, start_col) = pair.as_span().start_pos().line_col();
    let (end_line, end_col) = pair.as_span().end_pos().line_col();
    Span {
        start_line: start_line as u32,
        start_col: start_col as u32,
        end_line: end_line as u32,
        end_col: end_col as u32,
    }
}

pub fn parse(source: &str) -> Result<Module, String> {
    let pairs = LuaParser::parse(Rule::chunk, source).map_err(|e| format!("Parse error: {e}"))?;
    let mut body = Vec::new();
    for pair in pairs {
        if pair.as_rule() == Rule::chunk {
            for inner in pair.into_inner() {
                if inner.as_rule() == Rule::block {
                    for stmt in inner.into_inner() {
                        body.push(walk_statement(stmt)?);
                    }
                }
            }
        }
    }
    let mut module = Module {
        name: "main".to_string(),
        language: Lang::Lua,
        body,
        imports: Vec::new(),
    };
    super::normalize::normalize_module(&mut module);
    Ok(module)
}

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    // Pest may still emit a `statement` wrapper depending on call site.
    let pair = if pair.as_rule() == Rule::statement {
        pair.into_inner().next().ok_or("empty statement wrapper")?
    } else {
        pair
    };
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::local_function => walk_local_function(pair)?,
        Rule::local_var => walk_local_var(pair)?,
        Rule::function_decl => walk_function_decl(pair)?,
        Rule::for_numeric => walk_for_numeric(pair)?,
        Rule::for_generic => walk_for_generic(pair)?,
        Rule::do_statement => walk_do_statement(pair)?,
        Rule::if_statement => walk_if_statement(pair)?,
        Rule::while_statement => walk_while_statement(pair)?,
        Rule::repeat_statement => walk_repeat_statement(pair)?,
        Rule::return_statement => walk_return_statement(pair)?,
        Rule::break_statement => StmtKind::Break(BreakTarget::Implicit),
        Rule::assign_stmt => walk_assign_stmt(pair)?,
        Rule::expr => StmtKind::Expr(walk_expression(pair)?),
        other => return Err(format!("Unhandled statement rule: {other:?}")),
    };
    Ok(Statement::with_span(kind, span))
}

fn walk_do_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let body = pair
        .into_inner()
        .find(|p| matches!(p.as_rule(), Rule::body_block | Rule::block))
        .map(walk_block)
        .transpose()?
        .unwrap_or_default();
    Ok(StmtKind::Block(body))
}

/// Lua `do`/`for`/`while` bodies are separate scopes; wrap so `local` bindings
/// inside the body do not shadow loop control variables during `cond`/`update`.
fn lua_scoped_body(body: Vec<Statement>) -> Vec<Statement> {
    if body.is_empty() {
        body
    } else {
        vec![Statement::new(StmtKind::Block(body))]
    }
}

fn walk_for_numeric(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    let mut iter = inner.into_iter();
    iter.next(); // for
    let var = iter
        .find(|p| p.as_rule() == Rule::name)
        .ok_or("missing numeric for variable")?
        .as_str()
        .to_string();
    let mut exprs = Vec::new();
    let mut body_pair = None;
    for p in iter {
        match p.as_rule() {
            Rule::expr => exprs.push(p),
            Rule::body_block | Rule::block => body_pair = Some(p),
            _ => {}
        }
    }
    let start = walk_expression(exprs.remove(0))?;
    let limit = walk_expression(exprs.remove(0))?;
    let step = if let Some(step_pair) = exprs.pop() {
        walk_expression(step_pair)?
    } else {
        Expression::new(ExprKind::Lit(Literal::Int(1)))
    };
    let body = body_pair.map(walk_block).transpose()?.unwrap_or_default();

    Ok(super::normalize::build_numeric_for(
        var, start, limit, step, body,
    ))
}

fn walk_for_generic(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let inner = pair.into_inner();
    let mut names = Vec::new();
    let mut expl = Vec::new();
    let mut body = Vec::new();
    for p in inner {
        match p.as_rule() {
            Rule::for_varlist => {
                names = p
                    .into_inner()
                    .filter(|c| c.as_rule() == Rule::name)
                    .map(|c| c.as_str().to_string())
                    .collect();
            }
            Rule::expr => expl.push(walk_expression(p)?),
            Rule::body_block | Rule::block => {
                body = walk_block(p)?;
            }
            _ => {}
        }
    }
    if expl.is_empty() {
        return Err("missing generic for iterator expression".to_string());
    }
    let iter = if expl.len() == 1 {
        expl.into_iter().next().unwrap()
    } else {
        Expression::new(ExprKind::Sequence(expl))
    };
    Ok(StmtKind::ForIn {
        var: names.first().cloned().unwrap_or_default(),
        key: names.get(1).cloned(),
        iter,
        body: lua_scoped_body(body),
        of: true,
        else_body: None,
        is_async: false,
    })
}

fn walk_local_function(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    inner.next(); // local
    inner.next(); // function
    let name = inner
        .next()
        .filter(|p| p.as_rule() == Rule::name)
        .ok_or("missing local function name")?
        .as_str()
        .to_string();
    let (params, body) = walk_function_parts(inner.collect())?;
    Ok(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name),
            type_hint: None,
            init: Some(Expression::new(ExprKind::Lambda {
                params,
                body: LambdaBody::Block(body),
                is_async: false,
                captures: Vec::new(),
            })),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    })
}

fn walk_function_parts(parts: Vec<Pair<Rule>>) -> Result<(Vec<Param>, Vec<Statement>), String> {
    let mut params = Vec::new();
    let mut body = Vec::new();
    for p in parts {
        match p.as_rule() {
            Rule::param_list => params = walk_param_list(p)?,
            Rule::func_body | Rule::body_block | Rule::block => {
                body = walk_block(p)?;
            }
            _ => {}
        }
    }
    Ok((params, body))
}

fn walk_param_list(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ELLIPSIS => params.push(Param {
                name: "...".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: true,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }),
            Rule::name => params.push(Param {
                name: p.as_str().to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }),
            Rule::param => {
                let child = p.into_inner().next().ok_or("empty param")?;
                match child.as_rule() {
                    Rule::ELLIPSIS => params.push(Param {
                        name: "...".to_string(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: true,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    }),
                    Rule::name => params.push(Param {
                        name: child.as_str().to_string(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    }),
                    other => return Err(format!("unexpected param child: {other:?}")),
                }
            }
            Rule::comma | Rule::lparen | Rule::rparen => {}
            other => return Err(format!("unexpected param_list child: {other:?}")),
        }
    }
    Ok(params)
}

fn walk_local_var(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    inner.next(); // local
    let mut declarations = Vec::new();
    let mut values = Vec::new();
    let mut in_values = false;
    for p in inner {
        match p.as_rule() {
            Rule::name => {
                if in_values {
                    return Err("unexpected name in local initializer".into());
                }
                declarations.push(VarDeclarator {
                    pattern: BindingPattern::Ident(p.as_str().to_string()),
                    type_hint: None,
                    init: None,
                    array_bounds: None,
                    with_events: false,
                });
            }
            Rule::assign => in_values = true,
            Rule::expr => values.push(walk_expression(p)?),
            _ => {}
        }
    }
    for (i, decl) in declarations.iter_mut().enumerate() {
        if let Some(val) = values.get(i) {
            decl.init = Some(val.clone());
        }
    }
    Ok(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Let,
    })
}

fn walk_function_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    inner.next(); // function
    let func_name = inner
        .next()
        .ok_or("missing function name")?
        .into_inner()
        .collect::<Vec<_>>();
    let (mut params, body) = walk_function_parts(inner.collect())?;
    if func_name.len() == 1 {
        let global = func_name[0].as_str().to_string();
        return Ok(StmtKind::FunctionDecl {
            name: global,
            params,
            body,
            modifiers: Modifiers::default(),
            is_async: false,
            is_generator: false,
            is_sub: false,
            handles: Vec::new(),
            return_type: None,
        });
    }

    let (target, colon_method) = walk_qualified_func_name(&func_name)?;
    if colon_method {
        lua_ensure_self_param(&mut params);
    }

    let lambda = Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    });

    if colon_method {
        if let Some((table, field)) = split_method_table_field(&target) {
            return Ok(StmtKind::Block(vec![
                walk_preserve_data_field(table, &field),
                Statement::new(StmtKind::Assign {
                    targets: vec![target],
                    value: lambda,
                }),
            ]));
        }
    }

    Ok(StmtKind::Assign {
        targets: vec![target],
        value: lambda,
    })
}

fn split_method_table_field(target: &Expression) -> Option<(Expression, String)> {
    let ExprKind::Index { object, index, .. } = &target.kind else {
        return None;
    };
    let ExprKind::Lit(Literal::Str(field)) = &index.kind else {
        return None;
    };
    Some((object.as_ref().clone(), field.clone()))
}

/// JS-style: keep data fields when `function T:method()` would clobber `T.method`.
fn walk_preserve_data_field(table: Expression, field: &str) -> Statement {
    let field_key = Expression::new(ExprKind::Lit(Literal::Str(field.to_string())));
    let slot_val = Expression::new(ExprKind::Index {
        object: Box::new(table.clone()),
        index: Box::new(field_key.clone()),
        null_safe: false,
    });
    let data_key = Expression::new(ExprKind::Lit(Literal::Str("__lua_data".to_string())));
    let data_table = Expression::new(ExprKind::Index {
        object: Box::new(table.clone()),
        index: Box::new(data_key.clone()),
        null_safe: false,
    });
    let data_slot = Expression::new(ExprKind::Index {
        object: Box::new(data_table.clone()),
        index: Box::new(field_key),
        null_safe: false,
    });
    let is_function = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(Expression::new(ExprKind::Unary {
            op: UnaryOp::Typeof,
            expr: Box::new(slot_val.clone()),
        })),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
            "function".to_string(),
        )))),
    });
    let cond = Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(is_function),
        })),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(slot_val.clone()),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
        })),
    });
    Statement::new(StmtKind::If {
        cond,
        then_body: vec![
            Statement::new(StmtKind::If {
                cond: Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(data_table.clone()),
                    right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
                }),
                then_body: vec![Statement::new(StmtKind::Assign {
                    targets: vec![Expression::new(ExprKind::Index {
                        object: Box::new(table.clone()),
                        index: Box::new(data_key),
                        null_safe: false,
                    })],
                    value: Expression::new(ExprKind::Array(Vec::new())),
                })],
                elifs: Vec::new(),
                else_body: None,
            }),
            Statement::new(StmtKind::Assign {
                targets: vec![data_slot],
                value: slot_val,
            }),
        ],
        elifs: Vec::new(),
        else_body: None,
    })
}

/// `name`, `dot`/`colon`, `name`, … → lvalue `a.b.c` for `function a.b.c()`.
fn walk_qualified_func_name(parts: &[Pair<Rule>]) -> Result<(Expression, bool), String> {
    let mut iter = parts.iter();
    let base = iter
        .next()
        .ok_or("missing qualified function base name")?
        .as_str();
    let mut target = Expression::new(ExprKind::Ident(base.to_string()));
    let mut colon_method = false;
    while let Some(suffix) = iter.next() {
        if suffix.as_str().trim_start().starts_with(':') {
            colon_method = true;
        }
        let field = suffix
            .clone()
            .into_inner()
            .find(|p| p.as_rule() == Rule::name)
            .ok_or("missing qualified function field")?
            .as_str()
            .to_string();
        target = Expression::new(ExprKind::Index {
            object: Box::new(target),
            index: Box::new(Expression::new(ExprKind::Lit(Literal::Str(field)))),
            null_safe: false,
        });
    }
    Ok((target, colon_method))
}

fn lua_ensure_self_param(params: &mut Vec<Param>) {
    if params.first().is_some_and(|p| p.name == "self") {
        return;
    }
    params.insert(
        0,
        Param {
            name: "self".to_string(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        },
    );
}

fn walk_block(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();
    for p in pair.into_inner() {
        let stmt_pair = match p.as_rule() {
            Rule::stmt_not_else | Rule::stmt_not_end => {
                p.into_inner().next().ok_or("empty guarded statement")?
            }
            Rule::statement => p,
            other => return Err(format!("unexpected block item: {other:?}")),
        };
        body.push(walk_statement(stmt_pair)?);
    }
    Ok(body)
}

fn walk_if_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let inner = pair.into_inner();
    let mut cond = None;
    let mut then_body = Vec::new();
    let mut elifs = Vec::new();
    let mut else_body = None;
    let mut pending_elif_cond = None;
    for p in inner {
        match p.as_rule() {
            Rule::expr => {
                let e = walk_expression(p)?;
                if cond.is_none() {
                    cond = Some(e);
                } else {
                    pending_elif_cond = Some(e);
                }
            }
            Rule::then_block | Rule::else_block | Rule::block => {
                let stmts = walk_block(p)?;
                if then_body.is_empty() && pending_elif_cond.is_none() {
                    then_body = stmts;
                } else if let Some(elif_cond) = pending_elif_cond.take() {
                    elifs.push((elif_cond, stmts));
                } else {
                    else_body = Some(stmts);
                }
            }
            _ => {}
        }
    }
    Ok(StmtKind::If {
        cond: cond.ok_or("missing if condition")?,
        then_body,
        elifs,
        else_body,
    })
}

fn walk_while_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let inner = pair.into_inner();
    let mut cond = None;
    let mut body = Vec::new();
    for p in inner {
        match p.as_rule() {
            Rule::expr => cond = Some(walk_expression(p)?),
            Rule::body_block | Rule::block => {
                body = walk_block(p)?;
            }
            _ => {}
        }
    }
    Ok(StmtKind::While {
        cond: cond.ok_or("missing while condition")?,
        body: lua_scoped_body(body),
        else_body: None,
    })
}

fn walk_repeat_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let inner = pair.into_inner();
    let mut body = Vec::new();
    let mut cond = None;
    for p in inner {
        match p.as_rule() {
            Rule::body_block | Rule::block => {
                body = walk_block(p)?;
            }
            Rule::expr => cond = Some(walk_expression(p)?),
            _ => {}
        }
    }
    Ok(StmtKind::DoWhile {
        body: lua_scoped_body(body),
        cond: cond.ok_or("missing until condition")?,
        until: true,
    })
}

fn walk_return_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut values = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::expr {
            values.push(walk_expression(p)?);
        }
    }
    if values.len() > 1 {
        let elems = values
            .into_iter()
            .map(|v| ArrayElement {
                key: None,
                value: v,
                spread: false,
                by_ref: false,
            })
            .collect();
        Ok(StmtKind::Return(Some(Expression::new(ExprKind::Array(
            elems,
        )))))
    } else {
        Ok(StmtKind::Return(values.into_iter().next()))
    }
}

fn walk_assign_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let inner = pair.into_inner();
    let mut targets = Vec::new();
    let mut values = Vec::new();
    let mut seen_assign = false;
    for p in inner {
        match p.as_rule() {
            Rule::assign_lhs => {
                for lhs in p.into_inner() {
                    if lhs.as_rule() == Rule::postfix {
                        targets.push(walk_expression(lhs)?);
                    }
                }
            }
            Rule::assign_rhs => {
                for rhs in p.into_inner() {
                    if rhs.as_rule() == Rule::expr {
                        values.push(walk_expression(rhs)?);
                    }
                }
            }
            Rule::assign => seen_assign = true,
            Rule::postfix if !seen_assign => targets.push(walk_expression(p)?),
            Rule::expr if seen_assign => values.push(walk_expression(p)?),
            _ => {}
        }
    }
    let value = if values.len() == 1 {
        values.remove(0)
    } else {
        let elems = values
            .into_iter()
            .map(|v| ArrayElement {
                key: None,
                value: v,
                spread: false,
                by_ref: false,
            })
            .collect();
        Expression::new(ExprKind::Array(elems))
    };
    Ok(StmtKind::Assign { targets, value })
}

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let kind = walk_expr_kind(pair)?;
    Ok(Expression::with_span(kind, span))
}

fn walk_expr_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    match pair.as_rule() {
        Rule::call_expression => return walk_call_expression(pair),
        Rule::primary => return walk_primary(pair),
        Rule::postfix => {
            let inner = pair.into_inner().next().ok_or("empty postfix")?;
            return walk_call_expression(inner);
        }
        _ => {}
    }

    let rule = pair.as_rule();
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expression(inner.remove(0)).map(|e| e.kind);
    }
    match rule {
        Rule::or_expr => walk_binary_chain(inner, |_| BinOp::Or),
        Rule::and_expr => walk_binary_chain(inner, |_| BinOp::And),
        Rule::compare_expr => walk_compare_chain(inner),
        Rule::bor_expr => walk_binary_chain_with_ops(inner),
        Rule::bxor_expr => walk_binary_chain_with_ops(inner),
        Rule::band_expr => walk_binary_chain_with_ops(inner),
        Rule::shift_expr => walk_binary_chain_with_ops(inner),
        Rule::concat_expr => walk_binary_chain(inner, |_| BinOp::Concat),
        Rule::additive | Rule::multiplicative => walk_binary_chain_with_ops(inner),
        Rule::pow_expr => walk_pow_expr(inner),
        Rule::unary => walk_unary_from_inner(inner),
        Rule::expr => walk_expression(inner.remove(0)).map(|e| e.kind),
        other => Err(format!("Unhandled expression rule: {other:?}")),
    }
}

fn walk_binary_chain(
    mut items: Vec<Pair<Rule>>,
    op_fn: impl Fn(&str) -> BinOp,
) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    for item in items {
        if is_lua_expr_rule(item.as_rule()) {
            let right = walk_expression(item)?;
            left = Expression::new(ExprKind::Binary {
                op: op_fn(""),
                left: Box::new(left),
                right: Box::new(right),
            });
        }
    }
    Ok(left.kind)
}

fn walk_compare_chain(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    let mut i = 0;
    while i < items.len() {
        if items[i].as_rule() == Rule::compare_op {
            let op = parse_binop(items[i].as_str().trim())?;
            i += 1;
            if i < items.len() {
                let right = walk_expression(items[i].clone())?;
                i += 1;
                left = Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(left),
                    right: Box::new(right),
                });
            }
        } else {
            i += 1;
        }
    }
    Ok(left.kind)
}

/// Lua `^` is right-associative: `2 ^ 3 ^ 2` → `2 ^ (3 ^ 2)`.
fn walk_pow_expr(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let mut operands = Vec::new();
    operands.push(walk_expression(items.remove(0))?);
    let mut i = 1;
    while i < items.len() {
        let p = &items[i];
        if is_lua_op_rule(p.as_rule()) || p.as_rule() == Rule::CARET {
            i += 1;
            if i < items.len() {
                operands.push(walk_expression(items[i].clone())?);
                i += 1;
            }
        } else if is_lua_expr_rule(p.as_rule()) {
            operands.push(walk_expression(items[i].clone())?);
            i += 1;
        } else {
            i += 1;
        }
    }
    let mut acc = operands
        .pop()
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    while let Some(left) = operands.pop() {
        acc = Expression::new(ExprKind::Binary {
            op: BinOp::Pow,
            left: Box::new(left),
            right: Box::new(acc),
        });
    }
    Ok(acc.kind)
}

fn walk_binary_chain_with_ops(mut items: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let mut left = walk_expression(items.remove(0))?;
    let mut i = 0;
    while i < items.len() {
        let p = &items[i];
        if is_lua_op_rule(p.as_rule()) {
            let op = parse_binop(p.as_str().trim())?;
            i += 1;
            if i < items.len() {
                let right = walk_expression(items[i].clone())?;
                i += 1;
                left = Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(left),
                    right: Box::new(right),
                });
            }
        } else {
            i += 1;
        }
    }
    Ok(left.kind)
}

fn is_lua_expr_rule(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::or_expr
            | Rule::and_expr
            | Rule::compare_expr
            | Rule::concat_expr
            | Rule::shift_expr
            | Rule::band_expr
            | Rule::bxor_expr
            | Rule::bor_expr
            | Rule::additive
            | Rule::multiplicative
            | Rule::pow_expr
            | Rule::unary
            | Rule::postfix
            | Rule::call_expression
            | Rule::primary
            | Rule::expr
    )
}

fn is_lua_op_rule(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::additive_op
            | Rule::mul_op
            | Rule::pow_op
            | Rule::shift_op
            | Rule::compare_op
            | Rule::PLUS
            | Rule::MINUS
            | Rule::STAR
            | Rule::SLASH
            | Rule::DOUBLESLASH
            | Rule::PERCENT
            | Rule::CARET
            | Rule::CONCAT
            | Rule::AMP
            | Rule::PIPE
            | Rule::TILDE
            | Rule::LSHIFT
            | Rule::RSHIFT
            | Rule::EQ
            | Rule::NE
            | Rule::LT
            | Rule::GT
            | Rule::LE
            | Rule::GE
    )
}

fn walk_unary_from_inner(mut inner: Vec<Pair<Rule>>) -> Result<ExprKind, String> {
    let first = inner.remove(0);
    if first.as_rule() == Rule::unop {
        let op_str = first.as_str();
        let operand = walk_expression(inner.remove(0))?;
        if op_str == "#" {
            return Ok(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident("__lua_len".to_string()))),
                args: vec![Argument::positional(operand)],
                optional: false,
            });
        }
        let op = parse_unop(op_str)?;
        return Ok(ExprKind::Unary {
            op,
            expr: Box::new(operand),
        });
    }
    walk_expression(first).map(|e| e.kind)
}

/// Walk `primary ~ call_chain*` — same shape as JS `walk_call_chain`.
fn walk_call_expression(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("empty call_expression")?;
    let mut expr = walk_expression(first)?;

    for chain in inner {
        if chain.as_rule() != Rule::call_chain {
            continue;
        }
        let chain_src = chain.as_str().trim_start();
        let chain_inner: Vec<Pair<Rule>> = chain.into_inner().collect();

        if chain_src.starts_with('(') {
            let mut args = Vec::new();
            for arg in chain_inner {
                if arg.as_rule() == Rule::expr {
                    args.push(Argument::positional(walk_expression(arg)?));
                }
            }
            expr = Expression::new(ExprKind::Call {
                callee: Box::new(expr),
                args,
                optional: false,
            });
        } else if chain_src.starts_with('.') {
            let field = chain_inner
                .iter()
                .find(|p| p.as_rule() == Rule::name)
                .ok_or("missing member name")?
                .as_str()
                .to_string();
            expr = Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field,
                null_safe: false,
            });
        } else if chain_src.starts_with(':') {
            let field = chain_inner
                .iter()
                .find(|p| p.as_rule() == Rule::name)
                .ok_or("missing method name")?
                .as_str()
                .to_string();
            let mut args = Vec::new();
            for arg in chain_inner {
                if arg.as_rule() == Rule::expr {
                    args.push(Argument::positional(walk_expression(arg)?));
                }
            }
            let receiver = expr.clone();
            args.insert(0, Argument::positional(receiver));
            expr = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field,
                    null_safe: true,
                })),
                args,
                optional: false,
            });
        } else if chain_src.starts_with('[') {
            let index = chain_inner
                .iter()
                .find(|p| p.as_rule() == Rule::expr)
                .map(|p| walk_expression(p.clone()))
                .transpose()?
                .ok_or("missing index")?;
            expr = Expression::new(ExprKind::Index {
                object: Box::new(expr),
                index: Box::new(index),
                null_safe: false,
            });
        }
    }
    Ok(expr.kind)
}

fn unescape_lua_string(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let bytes = s.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'\\' && i + 1 < bytes.len() {
            i += 1;
            match bytes[i] {
                b'n' => {
                    out.push('\n');
                    i += 1;
                }
                b't' => {
                    out.push('\t');
                    i += 1;
                }
                b'r' => {
                    out.push('\r');
                    i += 1;
                }
                b'a' => {
                    out.push('\x07');
                    i += 1;
                }
                b'b' => {
                    out.push('\x08');
                    i += 1;
                }
                b'f' => {
                    out.push('\x0C');
                    i += 1;
                }
                b'v' => {
                    out.push('\x0B');
                    i += 1;
                }
                b'\\' => {
                    out.push('\\');
                    i += 1;
                }
                b'\'' => {
                    out.push('\'');
                    i += 1;
                }
                b'"' => {
                    out.push('"');
                    i += 1;
                }
                b'\n' => {
                    out.push('\n');
                    i += 1;
                }
                b'x' => {
                    // \xNN — two hex digits
                    if i + 2 < bytes.len() {
                        let hex = &s[i + 1..i + 3];
                        if let Ok(n) = u8::from_str_radix(hex, 16) {
                            out.push(n as char);
                        }
                        i += 3;
                    } else {
                        out.push('x');
                        i += 1;
                    }
                }
                b'u' => {
                    // \u{NNNN}
                    if i + 1 < bytes.len() && bytes[i + 1] == b'{' {
                        if let Some(end) = s[i + 2..].find('}') {
                            let hex = &s[i + 2..i + 2 + end];
                            if let Ok(n) = u32::from_str_radix(hex, 16) {
                                if let Some(c) = char::from_u32(n) {
                                    out.push(c);
                                }
                            }
                            i += 2 + end + 1;
                        } else {
                            out.push('u');
                            i += 1;
                        }
                    } else {
                        out.push('u');
                        i += 1;
                    }
                }
                b'z' => {
                    // \z — skip following whitespace
                    i += 1;
                    while i < bytes.len()
                        && (bytes[i] == b' '
                            || bytes[i] == b'\t'
                            || bytes[i] == b'\n'
                            || bytes[i] == b'\r')
                    {
                        i += 1;
                    }
                }
                d if d.is_ascii_digit() => {
                    // \DDD — up to 3 decimal digits
                    let start = i;
                    let mut count = 0;
                    while i < bytes.len() && count < 3 && bytes[i].is_ascii_digit() {
                        i += 1;
                        count += 1;
                    }
                    if let Ok(n) = s[start..i].parse::<u8>() {
                        out.push(n as char);
                    }
                }
                other => {
                    out.push(other as char);
                    i += 1;
                }
            }
        } else {
            out.push(bytes[i] as char);
            i += 1;
        }
    }
    out
}

fn strip_long_brackets(raw: &str) -> &str {
    // raw is the inner text produced by one of long_string_0/eq1/eq2/eq3 rules.
    // The outer brackets are already stripped by pest since they're silent rules
    // for long_string_0, but for eq1/eq2/eq3 the body is an atomic rule that
    // includes the content. We receive the full text including opening/closing brackets.
    // Strip [=*[ prefix and ]=*] suffix.
    let bytes = raw.as_bytes();
    if bytes.len() < 4 || bytes[0] != b'[' {
        return raw;
    }
    let mut eq_count = 0;
    let mut i = 1;
    while i < bytes.len() && bytes[i] == b'=' {
        eq_count += 1;
        i += 1;
    }
    if i >= bytes.len() || bytes[i] != b'[' {
        return raw;
    }
    let content_start = i + 1;
    // find matching close: ]=*]
    let suffix_len = eq_count + 2; // ]===...=]
    if raw.len() < content_start + suffix_len {
        return raw;
    }
    let content_end = raw.len() - suffix_len;
    // Skip optional first newline per Lua spec
    let content = &raw[content_start..content_end];
    if content.starts_with('\n') {
        &content[1..]
    } else {
        content
    }
}

fn walk_primary(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("empty primary")?;
    match first.as_rule() {
        Rule::KW_TRUE => Ok(ExprKind::Lit(Literal::Bool(true))),
        Rule::KW_FALSE => Ok(ExprKind::Lit(Literal::Bool(false))),
        Rule::KW_NIL => Ok(ExprKind::Lit(Literal::Null)),
        Rule::number => {
            let raw = first.as_str();
            // Remove underscore separators for parsing
            let clean: String = raw.chars().filter(|&c| c != '_').collect();
            if clean.starts_with("0b") || clean.starts_with("0B") {
                let v = i64::from_str_radix(&clean[2..], 2).unwrap_or(0);
                return Ok(ExprKind::Lit(Literal::Int(v)));
            }
            if clean.starts_with("0x") || clean.starts_with("0X") {
                let v = i64::from_str_radix(&clean[2..], 16).unwrap_or(0);
                return Ok(ExprKind::Lit(Literal::Int(v)));
            }
            let is_float = clean.contains('.') || clean.contains('e') || clean.contains('E');
            if is_float {
                Ok(ExprKind::Lit(Literal::Float(clean.parse().unwrap_or(0.0))))
            } else {
                Ok(ExprKind::Lit(Literal::Int(clean.parse().unwrap_or(0))))
            }
        }
        Rule::string => {
            let raw = first.as_str();
            let content = if raw.len() >= 2
                && ((raw.starts_with('"') && raw.ends_with('"'))
                    || (raw.starts_with('\'') && raw.ends_with('\'')))
            {
                &raw[1..raw.len() - 1]
            } else {
                raw
            };
            Ok(ExprKind::Lit(Literal::Str(unescape_lua_string(content))))
        }
        Rule::long_string => {
            // long_string contains one of long_string_0/eq1/eq2/eq3 as child
            let child = first.into_inner().next();
            let content = if let Some(child) = child {
                let raw = child.as_str();
                // The atomic body of long_string_0 is already content-only (body_0).
                // For eq1/eq2/eq3 the rule matched the whole [=..=] including brackets.
                // strip_long_brackets handles both.
                match child.as_rule() {
                    Rule::long_string_0 => {
                        // body_0 is atomic child inside long_string_0
                        let body = child.into_inner().next().map(|p| p.as_str()).unwrap_or("");
                        if body.starts_with('\n') {
                            body[1..].to_string()
                        } else {
                            body.to_string()
                        }
                    }
                    _ => strip_long_brackets(raw).to_string(),
                }
            } else {
                String::new()
            };
            Ok(ExprKind::Lit(Literal::Str(content)))
        }
        Rule::name => Ok(ExprKind::Ident(first.as_str().to_string())),
        Rule::table_constructor => walk_table_constructor(first),
        Rule::func_expr => walk_func_expr(first),
        Rule::ELLIPSIS => Ok(ExprKind::Spread(Box::new(Expression::null()))),
        Rule::lparen => {
            let expr = inner
                .find(|p| p.as_rule() == Rule::expr)
                .ok_or("empty parentheses")?;
            walk_expression(expr).map(|e| e.kind)
        }
        other => Err(format!("Unhandled primary: {other:?}")),
    }
}

fn walk_table_constructor(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut elements = Vec::new();
    for field in pair.into_inner() {
        if field.as_rule() == Rule::table_field {
            walk_table_field(field, &mut elements)?;
        }
    }
    Ok(ExprKind::Array(elements))
}

fn walk_table_field(field: Pair<Rule>, out: &mut Vec<ArrayElement>) -> Result<(), String> {
    let inner: Vec<Pair<Rule>> = field.into_inner().collect();
    if inner.is_empty() {
        return Ok(());
    }

    if inner.first().map(|p| p.as_rule()) == Some(Rule::ELLIPSIS) {
        let spread_expr = inner
            .iter()
            .find(|p| p.as_rule() == Rule::expr)
            .ok_or("missing spread expression")?;
        out.push(ArrayElement {
            key: None,
            value: walk_expression(spread_expr.clone())?,
            spread: true,
            by_ref: false,
        });
        return Ok(());
    }

    if inner.first().map(|p| p.as_rule()) == Some(Rule::lbracket) {
        let exprs: Vec<_> = inner
            .iter()
            .filter(|p| p.as_rule() == Rule::expr)
            .cloned()
            .collect();
        out.push(ArrayElement {
            key: Some(walk_expression(exprs[0].clone())?),
            value: walk_expression(
                exprs
                    .get(1)
                    .cloned()
                    .ok_or("missing bracketed table field value")?,
            )?,
            spread: false,
            by_ref: false,
        });
        return Ok(());
    }

    if inner.first().map(|p| p.as_rule()) == Some(Rule::name) {
        let key_name = inner[0].as_str().to_string();
        let value_expr = inner
            .iter()
            .find(|p| p.as_rule() == Rule::expr)
            .ok_or("missing named table field value")?;
        out.push(ArrayElement {
            key: Some(Expression::new(ExprKind::Lit(Literal::Str(key_name)))),
            value: walk_expression(value_expr.clone())?,
            spread: false,
            by_ref: false,
        });
        return Ok(());
    }

    if let Some(expr_pair) = inner.iter().find(|p| p.as_rule() == Rule::expr) {
        out.push(ArrayElement {
            key: None,
            value: walk_expression(expr_pair.clone())?,
            spread: false,
            by_ref: false,
        });
        return Ok(());
    }

    Err("unhandled table field shape".into())
}

fn walk_func_expr(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let (params, body) = walk_function_parts(pair.into_inner().collect())?;
    Ok(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    })
}

fn parse_binop(s: &str) -> Result<BinOp, String> {
    Ok(match s {
        "and" => BinOp::And,
        "or" => BinOp::Or,
        "<" => BinOp::Lt,
        ">" => BinOp::Gt,
        "<=" => BinOp::LtEq,
        ">=" => BinOp::GtEq,
        "~=" => BinOp::NotEq,
        "==" => BinOp::Eq,
        ".." => BinOp::Concat,
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "//" => BinOp::FloorDiv,
        "%" => BinOp::Mod,
        "^" => BinOp::Pow,
        "&" => BinOp::BitAnd,
        "|" => BinOp::BitOr,
        "~" => BinOp::BitXor,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        _ => return Err(format!("unknown binop: {s}")),
    })
}

fn parse_unop(s: &str) -> Result<UnaryOp, String> {
    Ok(match s {
        "not" => UnaryOp::Not,
        "-" => UnaryOp::Neg,
        "~" => UnaryOp::BitNot,
        _ => return Err(format!("unknown unop: {s}")),
    })
}
