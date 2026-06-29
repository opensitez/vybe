//! Lua → JS-shaped AST so the **shared compiler** lowers control flow and
//! expressions through `crates/vybe_compiler/src/emitter/` (`loops`, `expressions`,
//! `strings`, `collections`, `io`, …) — same path as JavaScript.

use crate::ast::*;

#[derive(Clone, Copy, PartialEq, Eq)]
enum LuaExprCtx {
    Read,
    Write,
}

pub fn normalize_module(module: &mut Module) {
    let mut body = Vec::new();
    for stmt in std::mem::take(&mut module.body) {
        for stmt in desugar_multi_assign(stmt) {
            body.extend(desugar_multi_var_decl(stmt));
        }
    }
    module.body = body;
    flatten_numeric_loop_blocks(&mut module.body);
    for stmt in &mut module.body {
        normalize_stmt(&mut stmt.kind);
    }
    flatten_numeric_loop_blocks(&mut module.body);
    for stmt in &mut module.body {
        maybe_wrap_numeric_loop_iife(&mut stmt.kind);
    }
    for stmt in &mut module.body {
        split_iife_assigns_in_stmt(&mut stmt.kind);
    }
}

fn flatten_numeric_loop_blocks(body: &mut Vec<Statement>) {
    let mut out = Vec::with_capacity(body.len());
    for stmt in std::mem::take(body) {
        if let StmtKind::Block(stmts) = stmt.kind {
            if is_numeric_loop_prelude_block(&stmts) {
                out.extend(stmts);
                continue;
            }
            out.push(Statement::new(StmtKind::Block(stmts)));
        } else {
            out.push(stmt);
        }
    }
    *body = out;
}

fn is_numeric_loop_prelude_block(stmts: &[Statement]) -> bool {
    if stmts.len() < 2 {
        return false;
    }
    if !matches!(stmts.last().map(|s| &s.kind), Some(StmtKind::While { .. })) {
        return false;
    }
    stmts[..stmts.len() - 1]
        .iter()
        .all(|s| matches!(s.kind, StmtKind::VarDecl { .. }))
}

fn fresh_lua_temp(prefix: &str) -> String {
    use std::sync::atomic::{AtomicUsize, Ordering};
    static NEXT: AtomicUsize = AtomicUsize::new(0);
    let n = NEXT.fetch_add(1, Ordering::Relaxed);
    format!("{prefix}_{n}")
}

fn is_lua_iife_call(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call {
            callee,
            args,
            optional: false,
        } if args.is_empty()
            && matches!(callee.kind, ExprKind::Lambda { .. })
    )
}

fn desugar_multi_assign(stmt: Statement) -> Vec<Statement> {
    let StmtKind::Assign { targets, value } = stmt.kind else {
        return vec![stmt];
    };
    if targets.len() <= 1 {
        return vec![Statement::with_span(
            StmtKind::Assign { targets, value },
            stmt.span,
        )];
    }
    let values = match value.kind.clone() {
        ExprKind::Sequence(values) => values,
        ExprKind::Call {
            callee,
            args,
            optional,
        } => {
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                let is_unpack = field == "unpack" && is_ident(object, "table");
                let is_modf = field == "modf" && is_ident(object, "math");
                if (is_unpack || is_modf) && !args.is_empty() {
                    let src = if is_modf {
                        Expression::new(ExprKind::Call {
                            callee: callee.clone(),
                            args: args.clone(),
                            optional,
                        })
                    } else {
                        args[0].value.clone()
                    };
                    let start = args
                        .get(1)
                        .map(|a| a.value.clone())
                        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Int(1))));
                    let mut out = Vec::with_capacity(targets.len());
                    for i in 0..targets.len() {
                        let idx = if i == 0 {
                            start.clone()
                        } else {
                            Expression::new(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(start.clone()),
                                right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(
                                    i as i64,
                                )))),
                            })
                        };
                        out.push(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Ident(
                                "rawget".to_string(),
                            ))),
                            args: vec![
                                Argument::positional(src.clone()),
                                Argument::positional(idx),
                            ],
                            optional: false,
                        }));
                    }
                    out
                } else {
                    return vec![Statement::with_span(
                        StmtKind::Assign { targets, value },
                        stmt.span,
                    )];
                }
            } else {
                return vec![Statement::with_span(
                    StmtKind::Assign { targets, value },
                    stmt.span,
                )];
            }
        }
        _ => {
            return vec![Statement::with_span(
                StmtKind::Assign { targets, value },
                stmt.span,
            )];
        }
    };
    if values.len() != targets.len() {
        return vec![Statement::with_span(
            StmtKind::Assign {
                targets,
                value: Expression::with_span(ExprKind::Sequence(values), value.span),
            },
            stmt.span,
        )];
    }

    let temps: Vec<String> = (0..targets.len())
        .map(|i| format!("__lua_multi_{i}"))
        .collect();
    let mut out = Vec::new();
    out.push(Statement::with_span(
        StmtKind::VarDecl {
            declarations: temps
                .iter()
                .zip(values.iter())
                .map(|(name, init)| VarDeclarator {
                    pattern: BindingPattern::Ident(name.clone()),
                    type_hint: None,
                    init: Some(init.clone()),
                    array_bounds: None,
                    with_events: false,
                })
                .collect(),
            kind: VarDeclKind::Let,
        },
        stmt.span.clone(),
    ));
    for (target, temp) in targets.iter().zip(temps.iter()) {
        out.push(Statement::with_span(
            StmtKind::Assign {
                targets: vec![target.clone()],
                value: Expression::new(ExprKind::Ident(temp.clone())),
            },
            stmt.span.clone(),
        ));
    }
    out
}

fn desugar_multi_var_decl(stmt: Statement) -> Vec<Statement> {
    let StmtKind::VarDecl { declarations, kind } = stmt.kind else {
        return vec![stmt];
    };
    if declarations.len() <= 1 {
        return vec![Statement::with_span(
            StmtKind::VarDecl { declarations, kind },
            stmt.span,
        )];
    }
    let Some(first_init) = declarations.first().and_then(|d| d.init.clone()) else {
        return vec![Statement::with_span(
            StmtKind::VarDecl { declarations, kind },
            stmt.span,
        )];
    };
    let first_init_expr = first_init.clone();
    let ExprKind::Call { callee, args, .. } = first_init.kind else {
        return vec![Statement::with_span(
            StmtKind::VarDecl { declarations, kind },
            stmt.span,
        )];
    };
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return vec![Statement::with_span(
            StmtKind::VarDecl { declarations, kind },
            stmt.span,
        )];
    };
    let is_unpack = field == "unpack" && is_ident(object, "table");
    let is_modf = field == "modf" && is_ident(object, "math");
    if (!is_unpack && !is_modf) || args.is_empty() {
        return vec![Statement::with_span(
            StmtKind::VarDecl { declarations, kind },
            stmt.span,
        )];
    }
    if declarations.iter().skip(1).any(|d| d.init.is_some()) {
        return vec![Statement::with_span(
            StmtKind::VarDecl { declarations, kind },
            stmt.span,
        )];
    }
    let src = if is_modf {
        first_init_expr
    } else {
        args[0].value.clone()
    };
    let start = args
        .get(1)
        .map(|a| a.value.clone())
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Int(1))));
    let rewritten = declarations
        .into_iter()
        .enumerate()
        .map(|(i, mut decl)| {
            let idx = if i == 0 {
                start.clone()
            } else {
                Expression::new(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(start.clone()),
                    right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(i as i64)))),
                })
            };
            decl.init = Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident("rawget".to_string()))),
                args: vec![Argument::positional(src.clone()), Argument::positional(idx)],
                optional: false,
            }));
            decl
        })
        .collect();
    vec![Statement::with_span(
        StmtKind::VarDecl {
            declarations: rewritten,
            kind,
        },
        stmt.span,
    )]
}

fn normalize_stmt(kind: &mut StmtKind) {
    if let Some(replacement) = try_desugar_loop(kind) {
        *kind = replacement;
    }
    match kind {
        StmtKind::Block(body) => {
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init_stmt) = init {
                normalize_stmt(&mut init_stmt.kind);
            }
            if let Some(c) = cond {
                normalize_expr(c);
            }
            if let Some(u) = update {
                normalize_expr(u);
            }
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_expr(iter);
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
            if let Some(eb) = else_body {
                for s in eb.iter_mut() {
                    normalize_stmt(&mut s.kind);
                }
            }
        }
        StmtKind::Expr(expr) => normalize_expr(expr),
        StmtKind::Assign { targets, value } => {
            if targets.len() == 1 {
                if let Some(expr) = try_desugar_lua_index_assign(targets[0].clone(), value) {
                    *kind = StmtKind::Expr(Expression::new(expr));
                    return;
                }
            }
            for t in targets.iter_mut() {
                normalize_expr_ctx(t, LuaExprCtx::Write);
            }
            normalize_expr(value);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations.iter_mut() {
                if let Some(init) = &mut d.init {
                    normalize_expr(init);
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } => {
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            normalize_expr(cond);
            for s in then_body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
            for (c, body) in elifs.iter_mut() {
                normalize_expr(c);
                for s in body.iter_mut() {
                    normalize_stmt(&mut s.kind);
                }
            }
            if let Some(body) = else_body {
                for s in body.iter_mut() {
                    normalize_stmt(&mut s.kind);
                }
            }
        }
        StmtKind::While { cond, body, .. } => {
            normalize_expr(cond);
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
            normalize_expr(cond);
        }
        StmtKind::Return(expr) => {
            if let Some(e) = expr {
                normalize_expr(e);
            }
        }
        _ => {}
    }
}

fn split_iife_assign_in_block(stmts: &mut [Statement]) {
    if stmts.len() <= 1 {
        return;
    }
    for stmt in stmts.iter_mut() {
        if let StmtKind::Assign { targets, value } = &mut stmt.kind {
            if targets.len() == 1 && is_lua_iife_call(value) {
                let temp = fresh_lua_temp("__lua_iife");
                let targets = std::mem::take(targets);
                let value = std::mem::replace(value, Expression::new(ExprKind::Lit(Literal::Null)));
                stmt.kind = StmtKind::Block(vec![
                    Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(temp.clone()),
                            type_hint: None,
                            init: Some(value),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    Statement::new(StmtKind::Assign {
                        targets,
                        value: Expression::new(ExprKind::Ident(temp)),
                    }),
                ]);
            }
        }
        split_iife_assigns_in_stmt(&mut stmt.kind);
    }
}

fn split_iife_assigns_in_stmt(kind: &mut StmtKind) {
    match kind {
        StmtKind::Block(stmts) => {
            split_iife_assign_in_block(stmts);
        }
        StmtKind::While { body, .. } | StmtKind::DoWhile { body, .. } => {
            split_iife_assign_in_block(body);
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            split_iife_assign_in_block(then_body);
            for (_, body) in elifs.iter_mut() {
                split_iife_assign_in_block(body);
            }
            if let Some(body) = else_body.as_mut() {
                split_iife_assign_in_block(body);
            }
        }
        StmtKind::For { body, .. } | StmtKind::ForIn { body, .. } => {
            split_iife_assign_in_block(body);
        }
        StmtKind::FunctionDecl { body, .. } => {
            for s in body.iter_mut() {
                split_iife_assigns_in_stmt(&mut s.kind);
            }
        }
        _ => {}
    }
}

fn maybe_wrap_numeric_loop_iife(kind: &mut StmtKind) {
    let StmtKind::While { body, .. } = kind else {
        return;
    };
    let Some((user_body, index_var, bind_prefix)) = numeric_loop_user_body(body) else {
        return;
    };
    let wrap_start = bind_prefix;
    let wrap_end = user_body.len();
    if wrap_end <= wrap_start {
        return;
    }
    if !stmt_list_contains_function_value(&user_body[wrap_start..wrap_end]) {
        return;
    }
    let wrapped = Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: index_var.clone(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Block(user_body[wrap_start..wrap_end].to_vec()),
            is_async: false,
            captures: Vec::new(),
        })),
        args: vec![Argument::positional(Expression::new(ExprKind::Ident(
            index_var,
        )))],
        optional: false,
    })));
    user_body.splice(wrap_start..wrap_end, [wrapped]);
}

/// Numeric `for` lowers to `while` with `[Block(user), ctrl += step]`.
fn numeric_loop_user_body(body: &mut [Statement]) -> Option<(&mut Vec<Statement>, String, usize)> {
    if body.len() != 2 {
        return None;
    }
    let ctrl_ok = numeric_loop_ctrl_increment(&body[1]);
    if !ctrl_ok {
        return None;
    }
    let StmtKind::Block(inner) = &mut body[0].kind else {
        return None;
    };
    let bind_prefix = numeric_loop_bind_prefix(inner)?;
    Some((inner, bind_prefix.0, bind_prefix.1))
}

fn numeric_loop_ctrl_increment(stmt: &Statement) -> bool {
    let StmtKind::Assign { targets, value } = &stmt.kind else {
        return false;
    };
    let Some(target) = targets.first() else {
        return false;
    };
    let ExprKind::Ident(name) = &target.kind else {
        return false;
    };
    if !name.starts_with("__lua_ctrl_") {
        return false;
    }
    matches!(
        &value.kind,
        ExprKind::Binary {
            op: BinOp::Add,
            left,
            ..
        } if matches!(&left.kind, ExprKind::Ident(l) if l == name)
    )
}

fn numeric_loop_bind_prefix(inner: &[Statement]) -> Option<(String, usize)> {
    let first = inner.first()?;
    if let StmtKind::VarDecl { declarations, .. } = &first.kind {
        if declarations.len() == 1 {
            if let BindingPattern::Ident(name) = &declarations[0].pattern {
                if let Some(ExprKind::Ident(ctrl)) = declarations[0].init.as_ref().map(|e| &e.kind)
                {
                    let expected = format!("__lua_ctrl_{name}");
                    if *ctrl == expected {
                        return Some((name.clone(), 1));
                    }
                }
            }
        }
    }
    None
}

fn stmt_list_contains_function_value(stmts: &[Statement]) -> bool {
    stmts.iter().any(stmt_contains_function_value)
}

fn stmt_contains_function_value(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Expr(expr) => expr_contains_function_value(expr),
        StmtKind::Assign { targets, value } => {
            targets.iter().any(expr_contains_function_value) || expr_contains_function_value(value)
        }
        StmtKind::VarDecl { declarations, .. } => declarations
            .iter()
            .filter_map(|d| d.init.as_ref())
            .any(expr_contains_function_value),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            expr_contains_function_value(cond)
                || stmt_list_contains_function_value(then_body)
                || elifs.iter().any(|(c, b)| {
                    expr_contains_function_value(c) || stmt_list_contains_function_value(b)
                })
                || else_body
                    .as_ref()
                    .is_some_and(|b| stmt_list_contains_function_value(b))
        }
        StmtKind::Block(stmts) => stmt_list_contains_function_value(stmts),
        _ => false,
    }
}

fn expr_contains_function_value(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_) => true,
        ExprKind::Binary { left, right, .. } => {
            expr_contains_function_value(left) || expr_contains_function_value(right)
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Delete(expr)
        | ExprKind::Void(expr)
        | ExprKind::Spread(expr)
        | ExprKind::RefLoad(expr) => expr_contains_function_value(expr),
        ExprKind::Yield(Some(expr)) => expr_contains_function_value(expr),
        ExprKind::Yield(None) => false,
        ExprKind::Ternary { cond, then, else_ } => {
            expr_contains_function_value(cond)
                || expr_contains_function_value(then)
                || expr_contains_function_value(else_)
        }
        ExprKind::Member { object, .. } => expr_contains_function_value(object),
        ExprKind::Index { object, index, .. } => {
            expr_contains_function_value(object) || expr_contains_function_value(index)
        }
        ExprKind::Call { callee, args, .. } => {
            expr_contains_function_value(callee)
                || args.iter().any(|a| expr_contains_function_value(&a.value))
        }
        ExprKind::New { class, args } => {
            expr_contains_function_value(class)
                || args.iter().any(|a| expr_contains_function_value(&a.value))
        }
        ExprKind::Assign { target, value } => {
            expr_contains_function_value(target) || expr_contains_function_value(value)
        }
        ExprKind::NullCoalesce { left, right } => {
            expr_contains_function_value(left) || expr_contains_function_value(right)
        }
        ExprKind::Array(items) => items.iter().any(|e| expr_contains_function_value(&e.value)),
        ExprKind::Tuple(items) | ExprKind::Set(items) => {
            items.iter().any(expr_contains_function_value)
        }
        _ => false,
    }
}

/// Lower Lua `ipairs` / `pairs` / generic `for` to `For` / `ForIn` / `While` so
/// `compiler/mod.rs` + `emitter/loops.rs` handle iteration (no lua emitter fork).
fn try_desugar_loop(kind: &StmtKind) -> Option<StmtKind> {
    match kind {
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            else_body,
            is_async,
            ..
        } => {
            if else_body.is_some() || *is_async {
                return None;
            }
            if let ExprKind::Call { callee, args, .. } = &iter.kind {
                if is_ident(callee, "ipairs") && args.len() == 1 {
                    return Some(desugar_ipairs(
                        var.clone(),
                        key.clone(),
                        args[0].value.clone(),
                        body.clone(),
                    ));
                }
                if is_ident(callee, "pairs") && args.len() == 1 {
                    return Some(desugar_pairs(
                        var.clone(),
                        key.clone(),
                        args[0].value.clone(),
                        body.clone(),
                    ));
                }
            }
            Some(desugar_generic_for(
                var.clone(),
                key.clone(),
                iter.clone(),
                body.clone(),
            ))
        }
        _ => None,
    }
}

fn desugar_ipairs(
    var: String,
    key: Option<String>,
    table: Expression,
    body: Vec<Statement>,
) -> StmtKind {
    let index_var = if key.is_some() && var != "_" {
        var.clone()
    } else {
        "__lua_i".to_string()
    };
    let mut prelude = Vec::new();
    if let Some(value_name) = ipairs_value_name(&var, &key) {
        prelude.push(Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Ident(value_name))],
            value: lua_table_index(&table, &index_var),
        }));
    }
    let mut loop_body = prelude;
    loop_body.extend(unwrap_lua_body(body));
    build_numeric_for(
        index_var,
        Expression::new(ExprKind::Lit(Literal::Float(1.0))),
        lua_len_call(&table),
        Expression::new(ExprKind::Lit(Literal::Float(1.0))),
        loop_body,
    )
}

fn desugar_pairs(
    var: String,
    key: Option<String>,
    table: Expression,
    body: Vec<Statement>,
) -> StmtKind {
    let of = key.is_some();
    StmtKind::ForIn {
        var,
        key,
        iter: table,
        body,
        of,
        else_body: None,
        is_async: false,
    }
}

fn desugar_generic_for(
    var: String,
    key: Option<String>,
    iter: Expression,
    body: Vec<Statement>,
) -> StmtKind {
    let loop_vars = loop_var_names(&var, &key);
    let (f_expr, s_expr, ctrl_init) = parse_explist(iter);
    let f_slot = "__lua_gen_f";
    let s_slot = "__lua_gen_s";
    let ctrl_slot = "__lua_gen_ctrl";

    let call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident(f_slot.to_string()))),
        args: vec![
            Argument::positional(Expression::new(ExprKind::Ident(s_slot.to_string()))),
            Argument::positional(Expression::new(ExprKind::Ident(ctrl_slot.to_string()))),
        ],
        optional: false,
    });

    let assign_targets: Vec<Expression> = loop_vars
        .iter()
        .map(|name| Expression::new(ExprKind::Ident(name.clone())))
        .collect();

    let mut while_body = vec![
        Statement::new(StmtKind::Assign {
            targets: assign_targets,
            value: call,
        }),
        Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Ident(ctrl_slot.to_string()))],
            value: Expression::new(ExprKind::Ident(loop_vars[0].clone())),
        }),
        Statement::new(StmtKind::If {
            cond: Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(Expression::new(ExprKind::Ident(ctrl_slot.to_string()))),
                right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
            }),
            then_body: vec![Statement::new(StmtKind::Break(BreakTarget::Implicit))],
            elifs: Vec::new(),
            else_body: None,
        }),
    ];
    while_body.extend(unwrap_lua_body(body));

    StmtKind::Block(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![
                VarDeclarator {
                    pattern: BindingPattern::Ident(f_slot.to_string()),
                    type_hint: None,
                    init: Some(f_expr),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(s_slot.to_string()),
                    type_hint: None,
                    init: Some(s_expr),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(ctrl_slot.to_string()),
                    type_hint: None,
                    init: Some(ctrl_init),
                    array_bounds: None,
                    with_events: false,
                },
            ],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::While {
            cond: Expression::new(ExprKind::Lit(Literal::Bool(true))),
            body: while_body,
            else_body: None,
        }),
    ])
}

pub(crate) fn build_numeric_for(
    index_var: String,
    start: Expression,
    limit: Expression,
    step: Expression,
    body: Vec<Statement>,
) -> StmtKind {
    if expr_is_zero(&step) {
        return StmtKind::Throw {
            expr: Some(Expression::new(ExprKind::Lit(Literal::Str(
                "'step' argument is zero".to_string(),
            )))),
            cause: None,
        };
    }
    let ctrl_var = format!("__lua_ctrl_{index_var}");
    let limit_temp = format!("__lua_for_limit_{index_var}");
    let step_temp = format!("__lua_for_step_{index_var}");
    let ctrl_expr = Expression::new(ExprKind::Ident(ctrl_var.clone()));
    let index_expr = Expression::new(ExprKind::Ident(index_var.clone()));
    let compare_op = if lua_step_is_negative(&step) {
        BinOp::GtEq
    } else {
        BinOp::LtEq
    };
    let cond = Expression::new(ExprKind::Binary {
        op: compare_op,
        left: Box::new(ctrl_expr.clone()),
        right: Box::new(Expression::new(ExprKind::Ident(limit_temp.clone()))),
    });
    // Bind the user-visible loop variable each iteration; keep the hidden control
    // variable separate so `local i = …` in the body cannot break the increment.
    let bind_loop_var = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(index_var.clone()),
            type_hint: None,
            init: Some(ctrl_expr.clone()),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });
    let mut scoped_body = vec![bind_loop_var];
    scoped_body.extend(body);
    let increment = Statement::new(StmtKind::Assign {
        targets: vec![ctrl_expr.clone()],
        value: Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(ctrl_expr),
            right: Box::new(Expression::new(ExprKind::Ident(step_temp.clone()))),
        }),
    });
    let while_body = vec![Statement::new(StmtKind::Block(scoped_body)), increment];

    StmtKind::Block(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(ctrl_var.clone()),
                type_hint: None,
                init: Some(start),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(limit_temp.clone()),
                type_hint: None,
                init: Some(limit),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(step_temp.clone()),
                type_hint: None,
                init: Some(step),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::While {
            cond,
            body: while_body,
            else_body: None,
        }),
    ])
}

fn lua_scoped_body(body: Vec<Statement>) -> Vec<Statement> {
    if body.is_empty() {
        body
    } else {
        vec![Statement::new(StmtKind::Block(body))]
    }
}

fn unwrap_lua_body(body: Vec<Statement>) -> Vec<Statement> {
    if body.len() == 1 {
        if let StmtKind::Block(inner) = &body[0].kind {
            if !inner.is_empty() {
                return inner.clone();
            }
        }
    }
    body
}

fn loop_var_names(var: &str, key: &Option<String>) -> Vec<String> {
    let mut names = vec![var.to_string()];
    if let Some(k) = key {
        names.push(k.clone());
    }
    names
}

fn ipairs_value_name(var: &str, key: &Option<String>) -> Option<String> {
    match key {
        Some(name) if name != "_" => Some(name.clone()),
        Some(_) => None,
        None if var != "_" => Some(var.to_string()),
        None => None,
    }
}

fn parse_explist(iter: Expression) -> (Expression, Expression, Expression) {
    let nil = Expression::new(ExprKind::Lit(Literal::Null));
    match iter.kind {
        ExprKind::Sequence(mut parts) => {
            let f = parts.remove(0);
            let s = parts.first().cloned().unwrap_or_else(|| nil.clone());
            let ctrl = parts.get(1).cloned().unwrap_or_else(|| nil.clone());
            (f, s, ctrl)
        }
        other => (Expression::new(other), nil.clone(), nil),
    }
}

fn lua_len_call(table: &Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("__lua_len".to_string()))),
        args: vec![Argument::positional(table.clone())],
        optional: false,
    })
}

fn lua_table_index(table: &Expression, index_name: &str) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(table.clone()),
        index: Box::new(Expression::new(ExprKind::Ident(index_name.to_string()))),
        null_safe: false,
    })
}

fn lit_bool(v: bool) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Bool(v)))
}

fn lit_float(n: f64) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Float(n)))
}

fn lit_str(s: &str) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Str(s.to_string())))
}

fn call_ident(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident(name.to_string()))),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn lua_index(object: Expression, index: Expression) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(index),
        null_safe: false,
    })
}

fn lua_return(expr: Expression) -> Statement {
    Statement::new(StmtKind::Return(Some(expr)))
}

/// Immediately-invoked lambda — short-circuit `and` / `or` and Lua comparisons.
fn lua_iife(stmts: Vec<Statement>) -> ExprKind {
    ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: vec![],
            body: LambdaBody::Block(stmts),
            is_async: false,
            captures: vec![],
        })),
        args: vec![],
        optional: false,
    }
}

/// Lua falsy: `nil` and `false` only (not JS truthiness).
fn lua_is_falsy(expr: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(expr.clone()),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
        })),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(expr),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(false)))),
        })),
    })
}

fn lua_is_truthy(expr: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(expr.clone()),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
        })),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(expr),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(false)))),
        })),
    })
}

fn lua_typeof(expr: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::Typeof,
        expr: Box::new(expr),
    })
}

fn lua_type_is(expr: Expression, ty: &str) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(lua_typeof(expr)),
        right: Box::new(lit_str(ty)),
    })
}

fn lua_tonumber(expr: Expression) -> Expression {
    call_ident("tonumber", vec![expr])
}

fn desugar_lua_and(left: Expression, right: Expression) -> ExprKind {
    ExprKind::Ternary {
        cond: Box::new(lua_is_falsy(left.clone())),
        then: Box::new(left),
        else_: Box::new(right),
    }
}

fn desugar_lua_or(left: Expression, right: Expression) -> ExprKind {
    ExprKind::Ternary {
        cond: Box::new(lua_is_truthy(left.clone())),
        then: Box::new(left),
        else_: Box::new(right),
    }
}

fn lua_to_number(expr: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("Number".to_string()))),
        args: vec![Argument::positional(expr)],
        optional: false,
    })
}

fn desugar_lua_add(left: Expression, right: Expression) -> ExprKind {
    let needs_numeric_coercion = matches!(left.kind, ExprKind::Lit(Literal::Str(_)))
        || matches!(right.kind, ExprKind::Lit(Literal::Str(_)));
    if needs_numeric_coercion {
        return ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(lua_to_number(left)),
            right: Box::new(lua_to_number(right)),
        };
    }
    desugar_lua_mm_binop(left, right, "__add", BinOp::Add, false)
}

fn desugar_lua_mod(left: Expression, right: Expression) -> ExprKind {
    let a = "__lua_mod_a";
    let b = "__lua_mod_b";
    let f = "__lua_mod_f";
    let quotient = Expression::new(ExprKind::Binary {
        op: BinOp::Div,
        left: Box::new(Expression::new(ExprKind::Ident(a.to_string()))),
        right: Box::new(Expression::new(ExprKind::Ident(b.to_string()))),
    });
    let floored = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::new(ExprKind::Ident("math".to_string()))),
            field: "floor".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(quotient)],
        optional: false,
    });
    let fallback = Expression::new(ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(Expression::new(ExprKind::Ident(a.to_string()))),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Mul,
            left: Box::new(floored),
            right: Box::new(Expression::new(ExprKind::Ident(b.to_string()))),
        })),
    });
    lua_iife(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![
                VarDeclarator {
                    pattern: BindingPattern::Ident(a.to_string()),
                    type_hint: None,
                    init: Some(left),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(b.to_string()),
                    type_hint: None,
                    init: Some(right),
                    array_bounds: None,
                    with_events: false,
                },
            ],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::If {
            cond: lua_type_is_function(lua_mm_lookup(a, "__mod")),
            then_body: vec![
                Statement::new(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(f.to_string()),
                        type_hint: None,
                        init: Some(lua_mm_lookup(a, "__mod")),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Let,
                }),
                lua_return(lua_mm_call(
                    Expression::new(ExprKind::Ident(f.to_string())),
                    Expression::new(ExprKind::Ident(a.to_string())),
                    Expression::new(ExprKind::Ident(b.to_string())),
                )),
            ],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: lua_type_is_function(lua_mm_lookup(b, "__mod")),
            then_body: vec![
                Statement::new(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(f.to_string()),
                        type_hint: None,
                        init: Some(lua_mm_lookup(b, "__mod")),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Let,
                }),
                lua_return(lua_mm_call(
                    Expression::new(ExprKind::Ident(f.to_string())),
                    Expression::new(ExprKind::Ident(b.to_string())),
                    Expression::new(ExprKind::Ident(a.to_string())),
                )),
            ],
            elifs: Vec::new(),
            else_body: None,
        }),
        lua_return(fallback),
    ])
}

fn desugar_lua_eq(left: Expression, right: Expression) -> ExprKind {
    let a = "__lua_eq_a";
    let b = "__lua_eq_b";
    let f = "__lua_eq_f";
    let fallback = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(Expression::new(ExprKind::Ident(a.to_string()))),
        right: Box::new(Expression::new(ExprKind::Ident(b.to_string()))),
    });
    lua_iife(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![
                VarDeclarator {
                    pattern: BindingPattern::Ident(a.to_string()),
                    type_hint: None,
                    init: Some(left),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(b.to_string()),
                    type_hint: None,
                    init: Some(right),
                    array_bounds: None,
                    with_events: false,
                },
            ],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::If {
            cond: lua_same_metatable(a, b),
            then_body: vec![
                Statement::new(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(f.to_string()),
                        type_hint: None,
                        init: Some(lua_mm_lookup(a, "__eq")),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Let,
                }),
                Statement::new(StmtKind::If {
                    cond: lua_type_is_function(Expression::new(ExprKind::Ident(f.to_string()))),
                    then_body: vec![lua_return(lua_mm_call(
                        Expression::new(ExprKind::Ident(f.to_string())),
                        Expression::new(ExprKind::Ident(a.to_string())),
                        Expression::new(ExprKind::Ident(b.to_string())),
                    ))],
                    elifs: Vec::new(),
                    else_body: None,
                }),
            ],
            elifs: Vec::new(),
            else_body: None,
        }),
        lua_return(fallback),
    ])
}

fn desugar_lua_rel(op: BinOp, left: Expression, right: Expression) -> ExprKind {
    match op {
        BinOp::Lt => desugar_lua_mm_binop(left, right, "__lt", BinOp::Lt, false),
        BinOp::LtEq => desugar_lua_mm_binop(left, right, "__le", BinOp::LtEq, false),
        _ => ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        },
    }
}

fn expr_is_lit_str(expr: &Expression, value: &str) -> bool {
    matches!(&expr.kind, ExprKind::Lit(Literal::Str(s)) if s == value)
}

fn desugar_os_date(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let first = iter.next();
    let second = iter.next();
    match (first, second) {
        (None, None) => call_ident("__lua_strftime", vec![lit_str("%c")]).kind,
        (Some(fmt), None) => {
            let format = Expression::new(normalize_expr_kind(fmt.value.kind, LuaExprCtx::Read));
            if expr_is_lit_str(&format, "*t") {
                call_ident("__lua_getdate", vec![]).kind
            } else {
                call_ident("__lua_strftime", vec![format]).kind
            }
        }
        (None, Some(ts)) => {
            let ts = Expression::new(normalize_expr_kind(ts.value.kind, LuaExprCtx::Read));
            call_ident("__lua_strftime", vec![lit_str("%c"), ts]).kind
        }
        (Some(fmt), Some(ts)) => {
            let format = Expression::new(normalize_expr_kind(fmt.value.kind, LuaExprCtx::Read));
            let ts = Expression::new(normalize_expr_kind(ts.value.kind, LuaExprCtx::Read));
            if expr_is_lit_str(&format, "*t") {
                call_ident("__lua_getdate", vec![ts]).kind
            } else {
                call_ident("__lua_strftime", vec![format, ts]).kind
            }
        }
    }
}

fn desugar_os_difftime(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let t1 = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    let t2 = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(t1),
        right: Box::new(t2),
    }
}

fn desugar_os_clock() -> ExprKind {
    let ns = call_ident("__lua_monotonic_now", vec![]);
    ExprKind::Binary {
        op: BinOp::Div,
        left: Box::new(ns),
        right: Box::new(lit_float(1_000_000_000.0)),
    }
}

fn desugar_os_setlocale(_args: Vec<Argument>) -> ExprKind {
    ExprKind::Lit(Literal::Str("C".to_string()))
}

fn lua_type_is_function(expr: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(lua_typeof(expr)),
        right: Box::new(lit_str("function")),
    })
}

fn lua_obj_mt(obj: &str) -> Expression {
    lua_raw_index(
        Expression::new(ExprKind::Ident(obj.to_string())),
        lit_str("__lua_mt"),
    )
}

/// Check if a value is a Lua table (object or array in the VM).
/// Strings/numbers/booleans/nil are NOT tables and have no metatable.
fn lua_is_table_value(obj: &str) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(lua_type_is(
            Expression::new(ExprKind::Ident(obj.to_string())),
            "object",
        )),
        right: Box::new(lua_type_is(
            Expression::new(ExprKind::Ident(obj.to_string())),
            "array",
        )),
    })
}

fn lua_mm_lookup(obj: &str, mm: &str) -> Expression {
    // Only objects/arrays (tables) can have metatables.
    // For strings/numbers/booleans ARRAY_GET returns a character (index 0),
    // which is truthy and causes metamethod dispatch to fire incorrectly.
    let mm_key = lit_str(mm);
    let from_mt = lua_raw_index(lua_obj_mt(obj), mm_key.clone());
    let from_proto = lua_raw_index(
        lua_raw_index(
            Expression::new(ExprKind::Ident(obj.to_string())),
            lit_str("__lua_proto"),
        ),
        mm_key,
    );
    Expression::new(ExprKind::Ternary {
        cond: Box::new(lua_is_table_value(obj)),
        then: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(lua_is_truthy(from_mt.clone())),
            then: Box::new(from_mt),
            else_: Box::new(from_proto),
        })),
        else_: Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
    })
}

fn lua_mm_call(mm_fn: Expression, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(mm_fn),
        args: vec![Argument::positional(left), Argument::positional(right)],
        optional: false,
    })
}

fn lua_mm_call1(mm_fn: Expression, arg: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(mm_fn),
        args: vec![Argument::positional(arg)],
        optional: false,
    })
}

/// Binary op with Lua metamethod fallback (`__add`, `__sub`, …).
/// Uses flat stmts inside one IIFE (not nested callee IIFEs). Unique temp names
/// avoid collisions; `split_iife_assign_in_block` hoists IIFE RHS in multi-stmt blocks.
fn desugar_lua_mm_binop(
    left: Expression,
    right: Expression,
    mm: &str,
    fallback_op: BinOp,
    commutative: bool,
) -> ExprKind {
    let a = fresh_lua_temp("__lua_a");
    let b = fresh_lua_temp("__lua_b");
    let f = fresh_lua_temp("__lua_mm_fn");
    let fallback = Expression::new(ExprKind::Binary {
        op: fallback_op,
        left: Box::new(Expression::new(ExprKind::Ident(a.clone()))),
        right: Box::new(Expression::new(ExprKind::Ident(b.clone()))),
    });
    let mut stmts = vec![Statement::new(StmtKind::VarDecl {
        declarations: vec![
            VarDeclarator {
                pattern: BindingPattern::Ident(a.clone()),
                type_hint: None,
                init: Some(left),
                array_bounds: None,
                with_events: false,
            },
            VarDeclarator {
                pattern: BindingPattern::Ident(b.clone()),
                type_hint: None,
                init: Some(right),
                array_bounds: None,
                with_events: false,
            },
        ],
        kind: VarDeclKind::Let,
    })];
    let left_mm = lua_mm_lookup(&a, mm);
    stmts.push(Statement::new(StmtKind::If {
        cond: lua_is_truthy(left_mm.clone()),
        then_body: vec![
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(f.clone()),
                    type_hint: None,
                    init: Some(left_mm),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            }),
            lua_return(lua_mm_call(
                Expression::new(ExprKind::Ident(f.clone())),
                Expression::new(ExprKind::Ident(a.clone())),
                Expression::new(ExprKind::Ident(b.clone())),
            )),
        ],
        elifs: Vec::new(),
        else_body: None,
    }));
    let right_mm = lua_mm_lookup(&b, mm);
    stmts.push(Statement::new(StmtKind::If {
        cond: lua_is_truthy(right_mm.clone()),
        then_body: vec![
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(f.clone()),
                    type_hint: None,
                    init: Some(right_mm),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            }),
            lua_return(lua_mm_call(
                Expression::new(ExprKind::Ident(f.clone())),
                Expression::new(ExprKind::Ident(b.clone())),
                Expression::new(ExprKind::Ident(a.clone())),
            )),
        ],
        elifs: Vec::new(),
        else_body: None,
    }));
    let _ = commutative;
    stmts.push(lua_return(fallback));
    lua_iife(stmts)
}

fn try_desugar_lua_index_assign(target: Expression, value: &mut Expression) -> Option<ExprKind> {
    match target.kind {
        ExprKind::Member { object, field, .. } => {
            let object = Expression::new(normalize_expr_kind(object.kind, LuaExprCtx::Read));
            if is_lua_profile_namespace(&object) {
                return None;
            }
            normalize_expr(value);
            Some(wrap_lua_proto_set(object, lit_str(&field), value.clone()))
        }
        ExprKind::Index { object, index, .. } => {
            let object = Expression::new(normalize_expr_kind(object.kind, LuaExprCtx::Read));
            if is_lua_profile_namespace(&object) {
                return None;
            }
            normalize_expr(value);
            let index = lua_one_based_index(Expression::new(normalize_expr_kind(
                index.kind,
                LuaExprCtx::Read,
            )));
            Some(wrap_lua_proto_set(object, index, value.clone()))
        }
        _ => None,
    }
}

fn is_lua_global_builtin(name: &str) -> bool {
    matches!(
        name,
        "print"
            | "tonumber"
            | "tostring"
            | "type"
            | "setmetatable"
            | "getmetatable"
            | "rawget"
            | "rawset"
            | "rawequal"
            | "pcall"
            | "xpcall"
            | "error"
            | "assert"
            | "ipairs"
            | "pairs"
            | "next"
            | "select"
            | "unpack"
            | "load"
            | "loadfile"
            | "dofile"
            | "require"
            | "collectgarbage"
            | "__lua_len"
    )
}

fn desugar_lua_unary_mm(expr: Expression, _mm: &str, fallback_op: UnaryOp) -> ExprKind {
    let a = "__lua_un_a";
    let f = "__lua_un_f";
    lua_iife(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(a.to_string()),
                type_hint: None,
                init: Some(expr),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(f.to_string()),
                type_hint: None,
                init: Some(lua_mm_lookup(a, _mm)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::If {
            cond: lua_is_truthy(Expression::new(ExprKind::Ident(f.to_string()))),
            then_body: vec![lua_return(lua_mm_call1(
                Expression::new(ExprKind::Ident(f.to_string())),
                Expression::new(ExprKind::Ident(a.to_string())),
            ))],
            elifs: Vec::new(),
            else_body: None,
        }),
        lua_return(Expression::new(ExprKind::Unary {
            op: fallback_op,
            expr: Box::new(Expression::new(ExprKind::Ident(a.to_string()))),
        })),
    ])
}

fn desugar_lua_len_call(args: Vec<Argument>) -> ExprKind {
    let obj = args
        .into_iter()
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    let v = "__lua_len_obj".to_string();
    let i = "__lua_len_i".to_string();
    let f = "__lua_len_fn".to_string();
    let cur = Expression::new(ExprKind::Index {
        object: Box::new(Expression::new(ExprKind::Ident(v.clone()))),
        index: Box::new(Expression::new(ExprKind::Ident(i.clone()))),
        null_safe: false,
    });
    let mut stmts = vec![Statement::new(StmtKind::VarDecl {
        declarations: vec![
            VarDeclarator {
                pattern: BindingPattern::Ident(v.clone()),
                type_hint: None,
                init: Some(obj),
                array_bounds: None,
                with_events: false,
            },
            VarDeclarator {
                pattern: BindingPattern::Ident(i.clone()),
                type_hint: None,
                init: Some(lit_float(0.0)),
                array_bounds: None,
                with_events: false,
            },
            VarDeclarator {
                pattern: BindingPattern::Ident(f.clone()),
                type_hint: None,
                init: Some(lua_mm_lookup(&v, "__len")),
                array_bounds: None,
                with_events: false,
            },
        ],
        kind: VarDeclKind::Let,
    })];
    stmts.push(Statement::new(StmtKind::If {
        cond: lua_type_is_function(Expression::new(ExprKind::Ident(f.clone()))),
        then_body: vec![lua_return(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Ident(f))),
            args: vec![Argument::positional(Expression::new(ExprKind::Ident(
                v.clone(),
            )))],
            optional: false,
        }))],
        elifs: Vec::new(),
        else_body: None,
    }));
    stmts.push(Statement::new(StmtKind::While {
        cond: Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::NotEq,
                left: Box::new(cur.clone()),
                right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
            })),
            right: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::NotEq,
                left: Box::new(cur),
                right: Box::new(Expression::new(ExprKind::Lit(Literal::Undefined))),
            })),
        }),
        body: vec![Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Ident(i.clone()))],
            value: Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(Expression::new(ExprKind::Ident(i.clone()))),
                right: Box::new(lit_float(1.0)),
            }),
        })],
        else_body: None,
    }));
    stmts.push(lua_return(Expression::new(ExprKind::Ident(i))));
    lua_iife(stmts)
}

/// Lua `type(v)` — returns one of: "nil", "boolean", "number", "string",
/// "function", "table". The VM's typeof returns JS tags ("object", "array",
/// "function", "number", "string", "boolean", "undefined") so we remap.
fn desugar_lua_type_call(arg: Expression) -> ExprKind {
    let v = "__lua_type_v";
    let t = "__lua_type_t";
    let arg = Expression::new(normalize_expr_kind(arg.kind, LuaExprCtx::Read));
    lua_iife(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![
                VarDeclarator {
                    pattern: BindingPattern::Ident(v.to_string()),
                    type_hint: None,
                    init: Some(arg),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(t.to_string()),
                    type_hint: None,
                    init: Some(lua_typeof(Expression::new(ExprKind::Ident(v.to_string())))),
                    array_bounds: None,
                    with_events: false,
                },
            ],
            kind: VarDeclKind::Let,
        }),
        // nil
        Statement::new(StmtKind::If {
            cond: Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(Expression::new(ExprKind::Ident(v.to_string()))),
                right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
            }),
            then_body: vec![lua_return(lit_str("nil"))],
            elifs: Vec::new(),
            else_body: None,
        }),
        // number (VM returns "number" for I32/I64/F64)
        Statement::new(StmtKind::If {
            cond: Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(Expression::new(ExprKind::Ident(t.to_string()))),
                right: Box::new(lit_str("number")),
            }),
            then_body: vec![lua_return(lit_str("number"))],
            elifs: Vec::new(),
            else_body: None,
        }),
        // string
        Statement::new(StmtKind::If {
            cond: Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(Expression::new(ExprKind::Ident(t.to_string()))),
                right: Box::new(lit_str("string")),
            }),
            then_body: vec![lua_return(lit_str("string"))],
            elifs: Vec::new(),
            else_body: None,
        }),
        // boolean
        Statement::new(StmtKind::If {
            cond: Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(Expression::new(ExprKind::Ident(t.to_string()))),
                right: Box::new(lit_str("boolean")),
            }),
            then_body: vec![lua_return(lit_str("boolean"))],
            elifs: Vec::new(),
            else_body: None,
        }),
        // function
        Statement::new(StmtKind::If {
            cond: Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(Expression::new(ExprKind::Ident(t.to_string()))),
                right: Box::new(lit_str("function")),
            }),
            then_body: vec![lua_return(lit_str("function"))],
            elifs: Vec::new(),
            else_body: None,
        }),
        // everything else (object, array, ...) → table
        lua_return(lit_str("table")),
    ])
}

/// Lua `print(...)` always calls tostring on each argument.
/// This ensures `nil` prints as "nil", tables as "table: 0x...", etc.
fn desugar_lua_print(args: Vec<Argument>) -> ExprKind {
    let wrapped: Vec<Argument> = args
        .into_iter()
        .map(|a| Argument::positional(Expression::new(desugar_lua_tostring(vec![a]))))
        .collect();
    ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("print".to_string()))),
        args: wrapped,
        optional: false,
    }
}

fn desugar_lua_tostring(args: Vec<Argument>) -> ExprKind {
    let obj = args
        .into_iter()
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    let o = "__lua_tostr_o";
    let f = "__lua_tostr_f";
    lua_iife(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![
                VarDeclarator {
                    pattern: BindingPattern::Ident(o.to_string()),
                    type_hint: None,
                    init: Some(obj),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(f.to_string()),
                    type_hint: None,
                    init: Some(lua_mm_lookup(o, "__tostring")),
                    array_bounds: None,
                    with_events: false,
                },
            ],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::If {
            cond: lua_is_truthy(Expression::new(ExprKind::Ident(f.to_string()))),
            then_body: vec![lua_return(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident(f.to_string()))),
                args: vec![Argument::positional(Expression::new(ExprKind::Ident(
                    o.to_string(),
                )))],
                optional: false,
            }))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(Expression::new(ExprKind::Ident(o.to_string()))),
                right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
            }),
            then_body: vec![lua_return(lit_str("nil"))],
            elifs: Vec::new(),
            else_body: None,
        }),
        lua_return(call_ident(
            "tostring",
            vec![Expression::new(ExprKind::Ident(o.to_string()))],
        )),
    ])
}

fn desugar_tonumber(args: Vec<Argument>) -> ExprKind {
    if let Some(first) = args.first() {
        if let ExprKind::Lit(Literal::Str(raw)) = &first.value.kind {
            let s = raw.trim();
            if let Some(second) = args.get(1) {
                let base = match &second.value.kind {
                    ExprKind::Lit(Literal::Int(n)) => Some(*n),
                    ExprKind::Lit(Literal::Float(n)) if n.fract() == 0.0 => Some(*n as i64),
                    _ => None,
                };
                if let Some(base) = base {
                    if (2..=36).contains(&base) {
                        let (neg, digits) = if let Some(rest) = s.strip_prefix('-') {
                            (true, rest)
                        } else if let Some(rest) = s.strip_prefix('+') {
                            (false, rest)
                        } else {
                            (false, s)
                        };
                        if digits.is_empty() {
                            return ExprKind::Lit(Literal::Null);
                        }
                        if let Ok(v) = i64::from_str_radix(digits, base as u32) {
                            let v = if neg { -v } else { v };
                            return ExprKind::Lit(Literal::Int(v));
                        }
                        return ExprKind::Lit(Literal::Null);
                    }
                }
            }
        }
    }

    let mut iter = args.into_iter();
    let value = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    let base = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)));

    let v = "__lua_tonum_v";
    let n = "__lua_tonum_n";
    let b = "__lua_tonum_b";

    let mut stmts = vec![Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(v.to_string()),
            type_hint: None,
            init: Some(value),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    })];

    stmts.push(Statement::new(StmtKind::If {
        cond: lua_type_is(Expression::new(ExprKind::Ident(v.to_string())), "number"),
        then_body: vec![lua_return(Expression::new(ExprKind::Ident(v.to_string())))],
        elifs: Vec::new(),
        else_body: None,
    }));

    stmts.push(Statement::new(StmtKind::If {
        cond: Expression::new(ExprKind::Binary {
            op: BinOp::StrictNotEq,
            left: Box::new(lua_typeof(Expression::new(ExprKind::Ident(v.to_string())))),
            right: Box::new(lit_str("string")),
        }),
        then_body: vec![lua_return(Expression::new(ExprKind::Lit(Literal::Null)))],
        elifs: Vec::new(),
        else_body: None,
    }));

    stmts.push(Statement::new(StmtKind::If {
        cond: Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(Expression::new(ExprKind::Ident(v.to_string()))),
            right: Box::new(lit_str("")),
        }),
        then_body: vec![lua_return(Expression::new(ExprKind::Lit(Literal::Null)))],
        elifs: Vec::new(),
        else_body: None,
    }));

    if let Some(base_expr) = base {
        stmts.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![
                VarDeclarator {
                    pattern: BindingPattern::Ident(b.to_string()),
                    type_hint: None,
                    init: Some(base_expr),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(n.to_string()),
                    type_hint: None,
                    init: Some(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Ident("parseInt".to_string()))),
                        args: vec![
                            Argument::positional(Expression::new(ExprKind::Ident(v.to_string()))),
                            Argument::positional(Expression::new(ExprKind::Ident(b.to_string()))),
                        ],
                        optional: false,
                    })),
                    array_bounds: None,
                    with_events: false,
                },
            ],
            kind: VarDeclKind::Let,
        }));
    } else {
        stmts.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(n.to_string()),
                type_hint: None,
                init: Some(lua_to_number(Expression::new(ExprKind::Ident(
                    v.to_string(),
                )))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }));
    }

    stmts.push(Statement::new(StmtKind::If {
        cond: Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(Expression::new(ExprKind::Ident(n.to_string()))),
            right: Box::new(Expression::new(ExprKind::Ident(n.to_string()))),
        }),
        then_body: vec![lua_return(Expression::new(ExprKind::Lit(Literal::Null)))],
        elifs: Vec::new(),
        else_body: None,
    }));

    stmts.push(lua_return(Expression::new(ExprKind::Ident(n.to_string()))));
    lua_iife(stmts)
}

fn desugar_pcall(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let f = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    let rest: Vec<Argument> = iter
        .map(|mut a| {
            a.value = Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read));
            a
        })
        .collect();
    let call = Expression::new(ExprKind::Call {
        callee: Box::new(f),
        args: rest,
        optional: false,
    });
    lua_iife(vec![Statement::new(StmtKind::Try {
        body: vec![
            Statement::new(StmtKind::Expr(call)),
            lua_return(lit_bool(true)),
        ],
        catches: vec![CatchClause {
            types: vec![],
            var_name: None,
            stack_var: None,
            body: vec![lua_return(lit_bool(false))],
            when_clause: None,
        }],
        else_body: None,
        finally: None,
    })])
}

fn desugar_getmetatable(args: Vec<Argument>) -> ExprKind {
    let obj = args
        .into_iter()
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    let o = "__lua_gm_o";
    let mt = "__lua_gm_mt";
    let hidden = "__lua_gm_hidden";
    lua_iife(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(o.to_string()),
                type_hint: None,
                init: Some(obj),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(mt.to_string()),
                type_hint: None,
                init: Some(lua_obj_mt(o)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::If {
            cond: lua_is_falsy(Expression::new(ExprKind::Ident(mt.to_string()))),
            then_body: vec![lua_return(Expression::new(ExprKind::Lit(Literal::Null)))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(hidden.to_string()),
                type_hint: None,
                init: Some(lua_raw_index(
                    Expression::new(ExprKind::Ident(mt.to_string())),
                    lit_str("__metatable"),
                )),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::If {
            cond: lua_is_truthy(Expression::new(ExprKind::Ident(hidden.to_string()))),
            then_body: vec![lua_return(Expression::new(ExprKind::Ident(
                hidden.to_string(),
            )))],
            elifs: Vec::new(),
            else_body: None,
        }),
        lua_return(Expression::new(ExprKind::Ident(mt.to_string()))),
    ])
}

fn desugar_rawequal(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let left = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    let right = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    ExprKind::Binary {
        op: BinOp::StrictEq,
        left: Box::new(left),
        right: Box::new(right),
    }
}

fn desugar_rawget(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let obj = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    let key = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    let key = lua_one_based_index(key);
    let v = "__lua_rawget_v";
    lua_iife(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(v.to_string()),
                type_hint: None,
                init: Some(Expression::new(ExprKind::Index {
                    object: Box::new(obj),
                    index: Box::new(key),
                    null_safe: false,
                })),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::If {
            cond: Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(Expression::new(ExprKind::Ident(v.to_string()))),
                right: Box::new(Expression::new(ExprKind::Lit(Literal::Undefined))),
            }),
            then_body: vec![lua_return(Expression::new(ExprKind::Lit(Literal::Null)))],
            elifs: Vec::new(),
            else_body: None,
        }),
        lua_return(Expression::new(ExprKind::Ident(v.to_string()))),
    ])
}

fn desugar_rawset(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let obj = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    let key = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    let val = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    ExprKind::Assign {
        target: Box::new(Expression::new(ExprKind::Index {
            object: Box::new(obj),
            index: Box::new(key),
            null_safe: false,
        })),
        value: Box::new(val),
    }
}

fn desugar_math_deg(args: Vec<Argument>) -> ExprKind {
    let x = args
        .into_iter()
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| lit_float(0.0));
    ExprKind::Binary {
        op: BinOp::Div,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Mul,
            left: Box::new(x),
            right: Box::new(lit_float(180.0)),
        })),
        right: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::new(ExprKind::Ident("math".to_string()))),
            field: "pi".to_string(),
            null_safe: false,
        })),
    }
}

fn desugar_math_rad(args: Vec<Argument>) -> ExprKind {
    let x = args
        .into_iter()
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| lit_float(0.0));
    ExprKind::Binary {
        op: BinOp::Div,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Mul,
            left: Box::new(x),
            right: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::Ident("math".to_string()))),
                field: "pi".to_string(),
                null_safe: false,
            })),
        })),
        right: Box::new(lit_float(180.0)),
    }
}

fn desugar_math_modf(args: Vec<Argument>) -> ExprKind {
    let x = args
        .into_iter()
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| lit_float(0.0));
    let slot = "__lua_modf_x";
    let i = "__lua_modf_i";
    lua_iife(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(slot.to_string()),
                type_hint: None,
                init: Some(x),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(i.to_string()),
                type_hint: None,
                init: Some(call_ident(
                    "math.floor",
                    vec![Expression::new(ExprKind::Ident(slot.to_string()))],
                )),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        lua_return(Expression::new(ExprKind::Array(vec![
            ArrayElement {
                key: None,
                value: Expression::new(ExprKind::Ident(i.to_string())),
                spread: false,
                by_ref: false,
            },
            ArrayElement {
                key: None,
                value: Expression::new(ExprKind::Binary {
                    op: BinOp::Sub,
                    left: Box::new(Expression::new(ExprKind::Ident(slot.to_string()))),
                    right: Box::new(Expression::new(ExprKind::Ident(i.to_string()))),
                }),
                spread: false,
                by_ref: false,
            },
        ]))),
    ])
}

fn desugar_math_fmod(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let left = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| lit_float(0.0));
    let right = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| lit_float(0.0));
    ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Or,
            left: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(right.clone()),
                right: Box::new(Expression::new(ExprKind::Lit(Literal::Float(
                    f64::INFINITY,
                )))),
            })),
            right: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(right.clone()),
                right: Box::new(Expression::new(ExprKind::Lit(Literal::Float(
                    f64::NEG_INFINITY,
                )))),
            })),
        })),
        then: Box::new(left.clone()),
        else_: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Mod,
            left: Box::new(left),
            right: Box::new(right),
        })),
    }
}

fn desugar_math_type(args: Vec<Argument>) -> ExprKind {
    if let Some(arg) = args.into_iter().next() {
        match &arg.value.kind {
            ExprKind::Lit(Literal::Int(_)) => {
                return ExprKind::Lit(Literal::Str("integer".to_string()));
            }
            ExprKind::Lit(Literal::Float(_)) => {
                return ExprKind::Lit(Literal::Str("float".to_string()));
            }
            _ => {}
        }
        let x = Expression::new(normalize_expr_kind(arg.value.kind, LuaExprCtx::Read));
        let slot = "__lua_mt_x";
        lua_iife(vec![
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(slot.to_string()),
                    type_hint: None,
                    init: Some(x),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            }),
            Statement::new(StmtKind::If {
                cond: Expression::new(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(lua_typeof(Expression::new(ExprKind::Ident(
                        slot.to_string(),
                    )))),
                    right: Box::new(lit_str("number")),
                }),
                then_body: vec![lua_return(Expression::new(ExprKind::Lit(Literal::Null)))],
                elifs: Vec::new(),
                else_body: None,
            }),
            Statement::new(StmtKind::If {
                cond: Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(Expression::new(ExprKind::Binary {
                        op: BinOp::Sub,
                        left: Box::new(Expression::new(ExprKind::Ident(slot.to_string()))),
                        right: Box::new(call_ident(
                            "math.floor",
                            vec![Expression::new(ExprKind::Ident(slot.to_string()))],
                        )),
                    })),
                    right: Box::new(lit_float(0.0)),
                }),
                then_body: vec![lua_return(lit_str("integer"))],
                elifs: Vec::new(),
                else_body: Some(vec![lua_return(lit_str("float"))]),
            }),
        ])
    } else {
        ExprKind::Lit(Literal::Null)
    }
}

fn desugar_math_ult(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let a = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| lit_float(0.0));
    let b = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| lit_float(0.0));
    ExprKind::Binary {
        op: BinOp::Lt,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::UShr,
            left: Box::new(a),
            right: Box::new(lit_float(0.0)),
        })),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::UShr,
            left: Box::new(b),
            right: Box::new(lit_float(0.0)),
        })),
    }
}

fn desugar_math_randomseed(args: Vec<Argument>) -> ExprKind {
    let seed = args
        .into_iter()
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| lit_float(0.0));
    ExprKind::Assign {
        target: Box::new(Expression::new(ExprKind::Ident(
            "__lua_rng_seed".to_string(),
        ))),
        value: Box::new(seed),
    }
}

fn desugar_math_random(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let a = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)));
    let b = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)));
    let seed_slot = "__lua_rng_seed";
    let r_slot = "__lua_rng_r";
    let next_seed = Expression::new(ExprKind::Binary {
        op: BinOp::Mod,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Mul,
                left: Box::new(Expression::new(ExprKind::Ident(seed_slot.to_string()))),
                right: Box::new(lit_float(1_103_515_245.0)),
            })),
            right: Box::new(lit_float(12_345.0)),
        })),
        right: Box::new(lit_float(2_147_483_647.0)),
    });
    let frac = Expression::new(ExprKind::Binary {
        op: BinOp::Div,
        left: Box::new(Expression::new(ExprKind::Ident(r_slot.to_string()))),
        right: Box::new(lit_float(2_147_483_647.0)),
    });
    let mut stmts = vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(r_slot.to_string()),
                type_hint: None,
                init: Some(next_seed.clone()),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Ident(seed_slot.to_string()))],
            value: Expression::new(ExprKind::Ident(r_slot.to_string())),
        }),
    ];
    let ret = match (a, b) {
        (None, None) => frac,
        (Some(hi), None) => Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::Ident("math".to_string()))),
                field: "floor".to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(lit_float(1.0)),
                right: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Mul,
                    left: Box::new(frac),
                    right: Box::new(hi),
                })),
            }))],
            optional: false,
        }),
        (Some(lo), Some(hi)) => Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::Ident("math".to_string()))),
                field: "floor".to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(lo.clone()),
                right: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Mul,
                    left: Box::new(frac),
                    right: Box::new(Expression::new(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(Expression::new(ExprKind::Binary {
                            op: BinOp::Sub,
                            right: Box::new(hi),
                            left: Box::new(lo.clone()),
                        })),
                        right: Box::new(lit_float(1.0)),
                    })),
                })),
            }))],
            optional: false,
        }),
        (None, Some(_)) => frac,
    };
    stmts.push(lua_return(ret));
    lua_iife(stmts)
}

fn desugar_lua_value_call(callee: Expression, args: Vec<Argument>) -> ExprKind {
    let c = "__lua_call_c";
    let mt = "__lua_call_mt";
    let f = "__lua_call_f";
    let norm_args: Vec<Argument> = args
        .into_iter()
        .map(|mut a| {
            a.value = Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read));
            a
        })
        .collect();
    let direct = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident(c.to_string()))),
        args: norm_args.clone(),
        optional: false,
    });
    lua_iife(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(c.to_string()),
                type_hint: None,
                init: Some(callee),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![
                VarDeclarator {
                    pattern: BindingPattern::Ident(mt.to_string()),
                    type_hint: None,
                    init: Some(Expression::new(ExprKind::Lit(Literal::Null))),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(f.to_string()),
                    type_hint: None,
                    init: Some(Expression::new(ExprKind::Lit(Literal::Null))),
                    array_bounds: None,
                    with_events: false,
                },
            ],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::If {
            cond: lua_is_truthy(lua_raw_index(
                Expression::new(ExprKind::Ident(c.to_string())),
                lit_str("__lua_mt"),
            )),
            then_body: vec![
                Statement::new(StmtKind::Assign {
                    targets: vec![Expression::new(ExprKind::Ident(mt.to_string()))],
                    value: lua_raw_index(
                        Expression::new(ExprKind::Ident(c.to_string())),
                        lit_str("__lua_mt"),
                    ),
                }),
                Statement::new(StmtKind::Assign {
                    targets: vec![Expression::new(ExprKind::Ident(f.to_string()))],
                    value: lua_raw_index(
                        Expression::new(ExprKind::Ident(mt.to_string())),
                        lit_str("__call"),
                    ),
                }),
            ],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: lua_is_truthy(Expression::new(ExprKind::Ident(f.to_string()))),
            then_body: vec![lua_return(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident(f.to_string()))),
                args: std::iter::once(Argument::positional(Expression::new(ExprKind::Ident(
                    c.to_string(),
                ))))
                .chain(norm_args)
                .collect(),
                optional: false,
            }))],
            elifs: Vec::new(),
            else_body: None,
        }),
        lua_return(direct),
    ])
}

fn lua_same_metatable(a: &str, b: &str) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(lua_is_truthy(lua_obj_mt(a))),
            right: Box::new(lua_is_truthy(lua_obj_mt(b))),
        })),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(lua_obj_mt(a)),
            right: Box::new(lua_obj_mt(b)),
        })),
    })
}

fn wrap_lua_proto_set(object: Expression, index: Expression, value: Expression) -> ExprKind {
    let o = "__lua_set_o";
    let k = "__lua_set_k";
    let v = "__lua_set_v";
    let mt = "__lua_set_mt";
    let ni = "__lua_set_newindex";
    let current = lua_raw_index(
        Expression::new(ExprKind::Ident(o.to_string())),
        Expression::new(ExprKind::Ident(k.to_string())),
    );
    lua_iife(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![
                VarDeclarator {
                    pattern: BindingPattern::Ident(o.to_string()),
                    type_hint: None,
                    init: Some(object),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(k.to_string()),
                    type_hint: None,
                    init: Some(index),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(v.to_string()),
                    type_hint: None,
                    init: Some(value),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(mt.to_string()),
                    type_hint: None,
                    init: Some(lua_obj_mt(o)),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(ni.to_string()),
                    type_hint: None,
                    init: Some(Expression::new(ExprKind::Lit(Literal::Null))),
                    array_bounds: None,
                    with_events: false,
                },
            ],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::If {
            cond: lua_is_truthy(current.clone()),
            then_body: vec![
                Statement::new(StmtKind::Assign {
                    targets: vec![lua_raw_index(
                        Expression::new(ExprKind::Ident(o.to_string())),
                        Expression::new(ExprKind::Ident(k.to_string())),
                    )],
                    value: Expression::new(ExprKind::Ident(v.to_string())),
                }),
                lua_return(Expression::new(ExprKind::Ident(v.to_string()))),
            ],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: lua_is_truthy(Expression::new(ExprKind::Ident(mt.to_string()))),
            then_body: vec![Statement::new(StmtKind::Assign {
                targets: vec![Expression::new(ExprKind::Ident(ni.to_string()))],
                value: lua_raw_index(
                    Expression::new(ExprKind::Ident(mt.to_string())),
                    lit_str("__newindex"),
                ),
            })],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: lua_type_is_function(Expression::new(ExprKind::Ident(ni.to_string()))),
            then_body: vec![
                Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Ident(ni.to_string()))),
                    args: vec![
                        Argument::positional(Expression::new(ExprKind::Ident(o.to_string()))),
                        Argument::positional(Expression::new(ExprKind::Ident(k.to_string()))),
                        Argument::positional(Expression::new(ExprKind::Ident(v.to_string()))),
                    ],
                    optional: false,
                }))),
                lua_return(Expression::new(ExprKind::Ident(v.to_string()))),
            ],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: lua_is_truthy(Expression::new(ExprKind::Ident(ni.to_string()))),
            then_body: vec![
                Statement::new(StmtKind::Assign {
                    targets: vec![lua_raw_index(
                        Expression::new(ExprKind::Ident(ni.to_string())),
                        Expression::new(ExprKind::Ident(k.to_string())),
                    )],
                    value: Expression::new(ExprKind::Ident(v.to_string())),
                }),
                lua_return(Expression::new(ExprKind::Ident(v.to_string()))),
            ],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::Assign {
            targets: vec![lua_raw_index(
                Expression::new(ExprKind::Ident(o.to_string())),
                Expression::new(ExprKind::Ident(k.to_string())),
            )],
            value: Expression::new(ExprKind::Ident(v.to_string())),
        }),
        lua_return(Expression::new(ExprKind::Ident(v.to_string()))),
    ])
}

/// Colon method dispatch — own function slot, then `__lua_proto` (flat stmts, no nested IIFE callee).
fn desugar_lua_colon_call(object: Expression, field: String, mut args: Vec<Argument>) -> ExprKind {
    if !args.is_empty() {
        args.remove(0);
    }
    let recv = "__lua_recv";
    let mtd = "__lua_mtd";
    let field_lit = lit_str(&field);
    let own_fn = lua_raw_index(
        Expression::new(ExprKind::Ident(recv.to_string())),
        field_lit.clone(),
    );
    let proto = lua_raw_index(
        Expression::new(ExprKind::Ident(recv.to_string())),
        lit_str("__lua_proto"),
    );
    let proto_fn = lua_raw_index(proto.clone(), field_lit);
    let mut call_args = vec![Argument::positional(Expression::new(ExprKind::Ident(
        recv.to_string(),
    )))];
    call_args.extend(args.into_iter().map(|mut a| {
        a.value = Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read));
        a
    }));
    lua_iife(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![
                VarDeclarator {
                    pattern: BindingPattern::Ident(recv.to_string()),
                    type_hint: None,
                    init: Some(Expression::new(normalize_expr_kind(
                        object.kind,
                        LuaExprCtx::Read,
                    ))),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(mtd.to_string()),
                    type_hint: None,
                    init: Some(Expression::new(ExprKind::Lit(Literal::Null))),
                    array_bounds: None,
                    with_events: false,
                },
            ],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::If {
            cond: lua_type_is_function(own_fn.clone()),
            then_body: vec![Statement::new(StmtKind::Assign {
                targets: vec![Expression::new(ExprKind::Ident(mtd.to_string()))],
                value: own_fn,
            })],
            elifs: vec![(
                Expression::new(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(lua_is_truthy(proto.clone())),
                    right: Box::new(lua_type_is_function(proto_fn.clone())),
                }),
                vec![Statement::new(StmtKind::Assign {
                    targets: vec![Expression::new(ExprKind::Ident(mtd.to_string()))],
                    value: proto_fn,
                })],
            )],
            else_body: None,
        }),
        lua_return(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Ident(mtd.to_string()))),
            args: call_args,
            optional: false,
        })),
    ])
}

fn desugar_lua_colon_call_expr(object: Expression, field: String, args: Vec<Argument>) -> ExprKind {
    desugar_lua_colon_call(object, field, args)
}

fn desugar_setmetatable(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let obj = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    let mt = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    let obj_slot = "__lua_sm_obj";
    let mt_slot = "__lua_sm_mt";
    let idx_slot = "__lua_sm_idx";
    lua_iife(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![
                VarDeclarator {
                    pattern: BindingPattern::Ident(obj_slot.to_string()),
                    type_hint: None,
                    init: Some(obj),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(mt_slot.to_string()),
                    type_hint: None,
                    init: Some(mt),
                    array_bounds: None,
                    with_events: false,
                },
            ],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::If {
            cond: Expression::new(ExprKind::Binary {
                op: BinOp::And,
                left: Box::new(lua_is_truthy(Expression::new(ExprKind::Ident(
                    mt_slot.to_string(),
                )))),
                right: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(lua_key_is_present(mt_slot, "__metatable")),
                    right: Box::new(lua_type_is(
                        lua_raw_index(
                            Expression::new(ExprKind::Ident(mt_slot.to_string())),
                            lit_str("__metatable"),
                        ),
                        "string",
                    )),
                })),
            }),
            then_body: vec![Statement::new(StmtKind::Throw {
                expr: Some(lit_str("cannot change a protected metatable")),
                cause: None,
            })],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::Assign {
            targets: vec![lua_raw_index(
                Expression::new(ExprKind::Ident(obj_slot.to_string())),
                lit_str("__lua_mt"),
            )],
            value: Expression::new(ExprKind::Ident(mt_slot.to_string())),
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(idx_slot.to_string()),
                type_hint: None,
                init: Some(lua_raw_index(
                    Expression::new(ExprKind::Ident(mt_slot.to_string())),
                    lit_str("__index"),
                )),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::If {
            cond: Expression::new(ExprKind::Binary {
                op: BinOp::And,
                left: Box::new(lua_is_truthy(Expression::new(ExprKind::Ident(
                    idx_slot.to_string(),
                )))),
                right: Box::new(Expression::new(ExprKind::Unary {
                    op: UnaryOp::Not,
                    expr: Box::new(lua_type_is_function(Expression::new(ExprKind::Ident(
                        idx_slot.to_string(),
                    )))),
                })),
            }),
            then_body: vec![Statement::new(StmtKind::Assign {
                targets: vec![lua_raw_index(
                    Expression::new(ExprKind::Ident(obj_slot.to_string())),
                    lit_str("__lua_proto"),
                )],
                value: Expression::new(ExprKind::Ident(idx_slot.to_string())),
            })],
            elifs: Vec::new(),
            else_body: None,
        }),
        lua_return(Expression::new(ExprKind::Ident(obj_slot.to_string()))),
    ])
}

fn lua_raw_index(object: Expression, index: Expression) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(index),
        null_safe: false,
    })
}

fn lua_key_is_present(object: &str, key: &str) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(lua_raw_index(
                Expression::new(ExprKind::Ident(object.to_string())),
                lit_str(key),
            )),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
        })),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(lua_raw_index(
                Expression::new(ExprKind::Ident(object.to_string())),
                lit_str(key),
            )),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Undefined))),
        })),
    })
}

/// Lua `__index` / field read — data in `__lua_data`, methods on table (JS prototype split).
fn wrap_lua_proto_get(object: Expression, index: Expression) -> ExprKind {
    let o = "__lua_get_o";
    let k = "__lua_get_k";
    let mt = "__lua_get_mt";
    let idx = "__lua_get_idx";
    let direct = "__lua_get_direct";
    let chain = "__lua_get_chain";
    lua_iife(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![
                VarDeclarator {
                    pattern: BindingPattern::Ident(o.to_string()),
                    type_hint: None,
                    init: Some(object),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(k.to_string()),
                    type_hint: None,
                    init: Some(index),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(direct.to_string()),
                    type_hint: None,
                    init: Some(lua_raw_index(
                        Expression::new(ExprKind::Ident(o.to_string())),
                        Expression::new(ExprKind::Ident(k.to_string())),
                    )),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(mt.to_string()),
                    type_hint: None,
                    init: Some(lua_obj_mt(o)),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(idx.to_string()),
                    type_hint: None,
                    init: Some(Expression::new(ExprKind::Lit(Literal::Null))),
                    array_bounds: None,
                    with_events: false,
                },
            ],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::If {
            cond: lua_type_is(Expression::new(ExprKind::Ident(o.to_string())), "number"),
            then_body: vec![lua_return(Expression::new(ExprKind::Ident(o.to_string())))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: Expression::new(ExprKind::Binary {
                op: BinOp::And,
                left: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(Expression::new(ExprKind::Ident(direct.to_string()))),
                    right: Box::new(Expression::new(ExprKind::Lit(Literal::Undefined))),
                })),
                right: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(Expression::new(ExprKind::Ident(direct.to_string()))),
                    right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
                })),
            }),
            then_body: vec![lua_return(Expression::new(ExprKind::Ident(
                direct.to_string(),
            )))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: lua_is_truthy(Expression::new(ExprKind::Ident(mt.to_string()))),
            then_body: vec![Statement::new(StmtKind::Assign {
                targets: vec![Expression::new(ExprKind::Ident(idx.to_string()))],
                value: lua_raw_index(
                    Expression::new(ExprKind::Ident(mt.to_string())),
                    lit_str("__index"),
                ),
            })],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: lua_type_is_function(Expression::new(ExprKind::Ident(idx.to_string()))),
            then_body: vec![
                Statement::new(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(direct.to_string()),
                        type_hint: None,
                        init: Some(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Ident(idx.to_string()))),
                            args: vec![
                                Argument::positional(Expression::new(ExprKind::Ident(
                                    o.to_string(),
                                ))),
                                Argument::positional(Expression::new(ExprKind::Ident(
                                    k.to_string(),
                                ))),
                            ],
                            optional: false,
                        })),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Let,
                }),
                Statement::new(StmtKind::If {
                    cond: Expression::new(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(Expression::new(ExprKind::Ident(direct.to_string()))),
                        right: Box::new(Expression::new(ExprKind::Lit(Literal::Undefined))),
                    }),
                    then_body: vec![lua_return(Expression::new(ExprKind::Lit(Literal::Null)))],
                    elifs: Vec::new(),
                    else_body: None,
                }),
                lua_return(Expression::new(ExprKind::Ident(direct.to_string()))),
            ],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: lua_is_truthy(Expression::new(ExprKind::Ident(idx.to_string()))),
            then_body: vec![
                Statement::new(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(direct.to_string()),
                        type_hint: None,
                        init: Some(lua_raw_index(
                            Expression::new(ExprKind::Ident(idx.to_string())),
                            Expression::new(ExprKind::Ident(k.to_string())),
                        )),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Let,
                }),
                Statement::new(StmtKind::If {
                    cond: Expression::new(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(Expression::new(ExprKind::Ident(direct.to_string()))),
                        right: Box::new(Expression::new(ExprKind::Lit(Literal::Undefined))),
                    }),
                    then_body: vec![
                        Statement::new(StmtKind::VarDecl {
                            declarations: vec![VarDeclarator {
                                pattern: BindingPattern::Ident(chain.to_string()),
                                type_hint: None,
                                init: Some(lua_raw_index(
                                    Expression::new(ExprKind::Ident(idx.to_string())),
                                    lit_str("__lua_proto"),
                                )),
                                array_bounds: None,
                                with_events: false,
                            }],
                            kind: VarDeclKind::Let,
                        }),
                        Statement::new(StmtKind::While {
                            cond: lua_is_truthy(Expression::new(ExprKind::Ident(
                                chain.to_string(),
                            ))),
                            body: vec![
                                Statement::new(StmtKind::Assign {
                                    targets: vec![Expression::new(ExprKind::Ident(
                                        direct.to_string(),
                                    ))],
                                    value: lua_raw_index(
                                        Expression::new(ExprKind::Ident(chain.to_string())),
                                        Expression::new(ExprKind::Ident(k.to_string())),
                                    ),
                                }),
                                Statement::new(StmtKind::If {
                                    cond: Expression::new(ExprKind::Binary {
                                        op: BinOp::And,
                                        left: Box::new(Expression::new(ExprKind::Binary {
                                            op: BinOp::NotEq,
                                            left: Box::new(Expression::new(ExprKind::Ident(
                                                direct.to_string(),
                                            ))),
                                            right: Box::new(Expression::new(ExprKind::Lit(
                                                Literal::Undefined,
                                            ))),
                                        })),
                                        right: Box::new(Expression::new(ExprKind::Binary {
                                            op: BinOp::NotEq,
                                            left: Box::new(Expression::new(ExprKind::Ident(
                                                direct.to_string(),
                                            ))),
                                            right: Box::new(Expression::new(ExprKind::Lit(
                                                Literal::Null,
                                            ))),
                                        })),
                                    }),
                                    then_body: vec![lua_return(Expression::new(ExprKind::Ident(
                                        direct.to_string(),
                                    )))],
                                    elifs: Vec::new(),
                                    else_body: None,
                                }),
                                Statement::new(StmtKind::Assign {
                                    targets: vec![Expression::new(ExprKind::Ident(
                                        chain.to_string(),
                                    ))],
                                    value: lua_raw_index(
                                        Expression::new(ExprKind::Ident(chain.to_string())),
                                        lit_str("__lua_proto"),
                                    ),
                                }),
                            ],
                            else_body: None,
                        }),
                        lua_return(Expression::new(ExprKind::Lit(Literal::Null))),
                    ],
                    elifs: Vec::new(),
                    else_body: None,
                }),
                lua_return(Expression::new(ExprKind::Ident(direct.to_string()))),
            ],
            elifs: Vec::new(),
            else_body: None,
        }),
        lua_return(Expression::new(ExprKind::Lit(Literal::Null))),
    ])
}

/// `string.byte(s [, i [, j]])` → `s.charCodeAt(i - 1)` (Lua default index 1).
fn desugar_string_byte_call(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let s = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    let index = iter
        .next()
        .map(|a| {
            lua_one_based_index(Expression::new(normalize_expr_kind(
                a.value.kind,
                LuaExprCtx::Read,
            )))
        })
        .unwrap_or_else(|| lit_float(0.0));
    ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(s),
            field: "charCodeAt".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(index)],
        optional: false,
    }
}

/// `string.sub(s, i [, j])` → `s.substring(i - 1 [, j])` — same `invoke:substring`
/// path as JavaScript (`languages/lua/profile` `[value_methods]`).
fn desugar_string_sub_call(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let s = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    let mut invoke_args = Vec::new();
    if let Some(start) = iter.next() {
        invoke_args.push(Argument::positional(lua_one_based_index(Expression::new(
            normalize_expr_kind(start.value.kind, LuaExprCtx::Read),
        ))));
    }
    if let Some(end) = iter.next() {
        // Lua 1-based inclusive `j` matches JS `substring` exclusive end index.
        invoke_args.push(Argument::positional(Expression::new(normalize_expr_kind(
            end.value.kind,
            LuaExprCtx::Read,
        ))));
    }
    ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(s),
            field: "substring".to_string(),
            null_safe: false,
        })),
        args: invoke_args,
        optional: false,
    }
}

fn desugar_table_pack(args: Vec<Argument>) -> ExprKind {
    let mut elements: Vec<ArrayElement> = args
        .iter()
        .map(|a| ArrayElement {
            key: None,
            value: Expression::new(normalize_expr_kind(a.value.clone().kind, LuaExprCtx::Read)),
            spread: false,
            by_ref: false,
        })
        .collect();
    elements.push(ArrayElement {
        key: Some(lit_str("n")),
        value: Expression::new(ExprKind::Lit(Literal::Int(args.len() as i64))),
        spread: false,
        by_ref: false,
    });
    ExprKind::Array(elements)
}

fn desugar_table_unpack(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let src = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Array(Vec::new())));
    if let Some(start) = iter.next() {
        let start_idx = lua_one_based_index(Expression::new(normalize_expr_kind(
            start.value.kind,
            LuaExprCtx::Read,
        )));
        let end_arg = iter
            .next()
            .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)));
        let mut slice_args = vec![Argument::positional(start_idx)];
        if let Some(end) = end_arg {
            slice_args.push(Argument::positional(end));
        }
        return ExprKind::Spread(Box::new(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(src),
                field: "slice".to_string(),
                null_safe: false,
            })),
            args: slice_args,
            optional: false,
        })));
    }
    ExprKind::Spread(Box::new(src))
}

fn desugar_table_insert(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let table = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Array(Vec::new())));
    let second = iter.next();
    let third = iter.next();

    match (second, third) {
        (Some(value), None) => ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(table),
                field: "push".to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(Expression::new(normalize_expr_kind(
                value.value.kind,
                LuaExprCtx::Read,
            )))],
            optional: false,
        },
        (Some(pos), Some(value)) => ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(table),
                field: "splice".to_string(),
                null_safe: false,
            })),
            args: vec![
                Argument::positional(lua_one_based_index(Expression::new(normalize_expr_kind(
                    pos.value.kind,
                    LuaExprCtx::Read,
                )))),
                Argument::positional(lit_float(0.0)),
                Argument::positional(Expression::new(normalize_expr_kind(
                    value.value.kind,
                    LuaExprCtx::Read,
                ))),
            ],
            optional: false,
        },
        _ => ExprKind::Lit(Literal::Null),
    }
}

fn desugar_table_remove(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let table = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Array(Vec::new())));
    if let Some(pos) = iter.next() {
        return ExprKind::Index {
            object: Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(table),
                    field: "splice".to_string(),
                    null_safe: false,
                })),
                args: vec![
                    Argument::positional(lua_one_based_index(Expression::new(
                        normalize_expr_kind(pos.value.kind, LuaExprCtx::Read),
                    ))),
                    Argument::positional(lit_float(1.0)),
                ],
                optional: false,
            })),
            index: Box::new(lit_float(0.0)),
            null_safe: false,
        };
    }
    ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(table),
            field: "pop".to_string(),
            null_safe: false,
        })),
        args: vec![],
        optional: false,
    }
}

fn desugar_table_concat(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let table = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Array(Vec::new())));
    let sep = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| lit_str(""));
    let start = iter
        .next()
        .map(|a| {
            lua_one_based_index(Expression::new(normalize_expr_kind(
                a.value.kind,
                LuaExprCtx::Read,
            )))
        })
        .unwrap_or_else(|| lit_float(0.0));
    let end = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| {
            Expression::new(desugar_lua_len_call(vec![Argument::positional(
                table.clone(),
            )]))
        });
    let slice = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(table),
            field: "slice".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(start), Argument::positional(end)],
        optional: false,
    });
    ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(slice),
            field: "join".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(sep)],
        optional: false,
    }
}

fn desugar_table_sort(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let table = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Array(Vec::new())));
    if let Some(cmp) = iter.next() {
        let cmp_fn = Expression::new(normalize_expr_kind(cmp.value.kind, LuaExprCtx::Read));
        let a = "a".to_string();
        let b = "b".to_string();
        let cmp_ab = Expression::new(ExprKind::Call {
            callee: Box::new(cmp_fn.clone()),
            args: vec![
                Argument::positional(Expression::new(ExprKind::Ident(a.clone()))),
                Argument::positional(Expression::new(ExprKind::Ident(b.clone()))),
            ],
            optional: false,
        });
        let cmp_ba = Expression::new(ExprKind::Call {
            callee: Box::new(cmp_fn),
            args: vec![
                Argument::positional(Expression::new(ExprKind::Ident(b.clone()))),
                Argument::positional(Expression::new(ExprKind::Ident(a.clone()))),
            ],
            optional: false,
        });
        let wrapper = Expression::new(ExprKind::Lambda {
            params: vec![
                Param {
                    name: a,
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                },
                Param {
                    name: b,
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                },
            ],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Ternary {
                cond: Box::new(cmp_ab),
                then: Box::new(lit_float(-1.0)),
                else_: Box::new(Expression::new(ExprKind::Ternary {
                    cond: Box::new(cmp_ba),
                    then: Box::new(lit_float(1.0)),
                    else_: Box::new(lit_float(0.0)),
                })),
            }))),
            is_async: false,
            captures: Vec::new(),
        });
        return ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(table),
                field: "sort".to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(wrapper)],
            optional: false,
        };
    }
    ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(table),
            field: "sort".to_string(),
            null_safe: false,
        })),
        args: vec![],
        optional: false,
    }
}

fn desugar_table_move(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let table = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Array(Vec::new())));
    let from = iter
        .next()
        .map(|a| {
            lua_one_based_index(Expression::new(normalize_expr_kind(
                a.value.kind,
                LuaExprCtx::Read,
            )))
        })
        .unwrap_or_else(|| lit_float(0.0));
    let end_exclusive = iter
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| lit_float(0.0));
    let target = iter
        .next()
        .map(|a| {
            lua_one_based_index(Expression::new(normalize_expr_kind(
                a.value.kind,
                LuaExprCtx::Read,
            )))
        })
        .unwrap_or_else(|| lit_float(0.0));
    let src = "__lua_move_src".to_string();
    let tmp = "__lua_move_tmp".to_string();
    let mut stmts = vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(src.clone()),
                type_hint: None,
                init: Some(table),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(tmp.clone()),
                type_hint: None,
                init: Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::Ident(src.clone()))),
                        field: "slice".to_string(),
                        null_safe: false,
                    })),
                    args: vec![
                        Argument::positional(from),
                        Argument::positional(end_exclusive),
                    ],
                    optional: false,
                })),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::Ident(src.clone()))),
                field: "splice".to_string(),
                null_safe: false,
            })),
            args: vec![
                Argument::positional(target),
                Argument::positional(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::Ident(tmp.clone()))),
                    field: "length".to_string(),
                    null_safe: false,
                })),
                Argument {
                    value: Expression::new(ExprKind::Ident(tmp)),
                    name: None,
                    by_ref: false,
                    spread: true,
                },
            ],
            optional: false,
        }))),
    ];
    stmts.push(lua_return(Expression::new(ExprKind::Ident(src))));
    lua_iife(stmts)
}

fn desugar_next(args: Vec<Argument>) -> ExprKind {
    if args.is_empty() {
        return ExprKind::Lit(Literal::Null);
    }
    let normalized: Vec<Argument> = args
        .into_iter()
        .map(|mut a| {
            a.value = Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read));
            a
        })
        .collect();
    if let Some(first) = normalized.first() {
        if matches!(&first.value.kind, ExprKind::Lit(Literal::Null)) {
            return ExprKind::Lit(Literal::Null);
        }
        if matches!(&first.value.kind, ExprKind::Array(elements) if elements.is_empty()) {
            return ExprKind::Lit(Literal::Null);
        }
    }
    ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("next".to_string()))),
        args: normalized,
        optional: false,
    }
}

fn is_ident(expr: &Expression, name: &str) -> bool {
    matches!(&expr.kind, ExprKind::Ident(n) if n == name)
}

fn expr_is_zero(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(0)) => true,
        ExprKind::Lit(Literal::Float(f)) => *f == 0.0,
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => expr_is_zero(expr),
        _ => false,
    }
}

fn lua_step_is_negative(step: &Expression) -> bool {
    match &step.kind {
        ExprKind::Lit(Literal::Int(n)) => *n < 0,
        ExprKind::Lit(Literal::Float(n)) => *n < 0.0,
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => match &expr.kind {
            ExprKind::Lit(Literal::Int(n)) => *n > 0,
            ExprKind::Lit(Literal::Float(n)) => *n > 0.0,
            _ => false,
        },
        _ => false,
    }
}

fn normalize_expr(expr: &mut Expression) {
    normalize_expr_ctx(expr, LuaExprCtx::Read);
}

fn normalize_expr_ctx(expr: &mut Expression, ctx: LuaExprCtx) {
    expr.kind = normalize_expr_kind(
        std::mem::replace(&mut expr.kind, ExprKind::Lit(Literal::Null)),
        ctx,
    );
}

fn normalize_expr_kind(kind: ExprKind, ctx: LuaExprCtx) -> ExprKind {
    let read = LuaExprCtx::Read;
    let kind = match kind {
        ExprKind::Unary { op, expr } => {
            let expr = Expression::new(normalize_expr_kind(expr.kind, read));
            return match op {
                UnaryOp::Not => lua_is_falsy(expr).kind,
                UnaryOp::Neg => desugar_lua_unary_mm(expr, "__unm", UnaryOp::Neg),
                UnaryOp::BitNot => desugar_lua_unary_mm(expr, "__bnot", UnaryOp::BitNot),
                other => ExprKind::Unary {
                    op: other,
                    expr: Box::new(expr),
                },
            };
        }
        ExprKind::Binary { op, left, right } => {
            let left = Expression::new(normalize_expr_kind(left.kind, read));
            let right = Expression::new(normalize_expr_kind(right.kind, read));
            return match op {
                BinOp::And => desugar_lua_and(left, right),
                BinOp::Or => desugar_lua_or(left, right),
                BinOp::Add => desugar_lua_add(left, right),
                BinOp::Sub => desugar_lua_mm_binop(left, right, "__sub", BinOp::Sub, false),
                BinOp::Mul => desugar_lua_mm_binop(left, right, "__mul", BinOp::Mul, false),
                BinOp::Div => desugar_lua_mm_binop(left, right, "__div", BinOp::Div, false),
                BinOp::IDiv | BinOp::FloorDiv => {
                    desugar_lua_mm_binop(left, right, "__idiv", BinOp::IDiv, false)
                }
                BinOp::Pow => desugar_lua_mm_binop(left, right, "__pow", BinOp::Pow, false),
                BinOp::BitAnd => desugar_lua_mm_binop(left, right, "__band", BinOp::BitAnd, false),
                BinOp::BitOr => desugar_lua_mm_binop(left, right, "__bor", BinOp::BitOr, false),
                BinOp::BitXor => desugar_lua_mm_binop(left, right, "__bxor", BinOp::BitXor, false),
                BinOp::Shl => desugar_lua_mm_binop(left, right, "__shl", BinOp::Shl, false),
                BinOp::Shr => desugar_lua_mm_binop(left, right, "__shr", BinOp::Shr, false),
                BinOp::Concat => {
                    desugar_lua_mm_binop(left, right, "__concat", BinOp::Concat, false)
                }
                BinOp::Mod => desugar_lua_mod(left, right),
                BinOp::Eq => desugar_lua_eq(left, right),
                BinOp::NotEq => ExprKind::Unary {
                    op: UnaryOp::Not,
                    expr: Box::new(Expression::new(desugar_lua_eq(left, right))),
                },
                BinOp::Gt => desugar_lua_rel(BinOp::Lt, right, left),
                BinOp::GtEq => desugar_lua_rel(BinOp::LtEq, right, left),
                BinOp::Lt | BinOp::LtEq => desugar_lua_rel(op, left, right),
                _ => ExprKind::Binary {
                    op,
                    left: Box::new(left),
                    right: Box::new(right),
                },
            };
        }
        ExprKind::Assign { target, value } => ExprKind::Assign {
            target: Box::new(Expression::new(normalize_expr_kind(
                target.kind,
                LuaExprCtx::Write,
            ))),
            value: Box::new(Expression::new(normalize_expr_kind(value.kind, read))),
        },
        ExprKind::Lambda {
            params,
            body,
            is_async,
            captures,
        } => {
            let body = match body {
                LambdaBody::Expr(expr) => LambdaBody::Expr(Box::new(Expression::new(
                    normalize_expr_kind(expr.kind, read),
                ))),
                LambdaBody::Block(mut stmts) => {
                    for s in stmts.iter_mut() {
                        normalize_stmt(&mut s.kind);
                    }
                    LambdaBody::Block(stmts)
                }
            };
            ExprKind::Lambda {
                params,
                body,
                is_async,
                captures,
            }
        }
        ExprKind::Array(elements) => {
            let elements = elements
                .into_iter()
                .map(|mut e| {
                    if let Some(key) = &mut e.key {
                        let normalized_key = Expression::new(normalize_expr_kind(
                            key.kind.clone(),
                            LuaExprCtx::Read,
                        ));
                        *key = lua_one_based_index(normalized_key);
                    }
                    normalize_expr(&mut e.value);
                    let value_kind =
                        std::mem::replace(&mut e.value.kind, ExprKind::Lit(Literal::Null));
                    if let ExprKind::Spread(inner) = value_kind {
                        e.value = *inner;
                        e.spread = true;
                    } else {
                        e.value.kind = value_kind;
                    }
                    e
                })
                .collect();
            ExprKind::Array(elements)
        }
        ExprKind::Call {
            callee,
            args,
            optional,
        } => {
            if let ExprKind::Member {
                object,
                field,
                null_safe: true,
                ..
            } = &callee.kind
            {
                return desugar_lua_colon_call_expr(object.as_ref().clone(), field.clone(), args);
            }
            if let ExprKind::Ident(name) = &callee.kind {
                match name.as_str() {
                    "print" => return desugar_lua_print(args),
                    "setmetatable" => return desugar_setmetatable(args),
                    "getmetatable" => return desugar_getmetatable(args),
                    "rawget" => return desugar_rawget(args),
                    "rawset" => return desugar_rawset(args),
                    "rawequal" => return desugar_rawequal(args),
                    "pcall" => return desugar_pcall(args),
                    "next" => return desugar_next(args),
                    "tostring" => return desugar_lua_tostring(args),
                    "tonumber" => return desugar_tonumber(args),
                    "__lua_len" => return desugar_lua_len_call(args),
                    _ => {}
                }
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if is_ident(object, "os") {
                    match field.as_str() {
                        "date" => return desugar_os_date(args),
                        "difftime" => return desugar_os_difftime(args),
                        "clock" => return desugar_os_clock(),
                        "setlocale" => return desugar_os_setlocale(args),
                        _ => {}
                    }
                }
                if is_ident(object, "math") {
                    match field.as_str() {
                        "deg" => return desugar_math_deg(args),
                        "rad" => return desugar_math_rad(args),
                        "modf" => return desugar_math_modf(args),
                        "fmod" => return desugar_math_fmod(args),
                        "type" => return desugar_math_type(args),
                        "ult" => return desugar_math_ult(args),
                        "randomseed" => return desugar_math_randomseed(args),
                        "random" => return desugar_math_random(args),
                        _ => {}
                    }
                }
                if is_ident(object, "table") {
                    match field.as_str() {
                        "insert" => return desugar_table_insert(args),
                        "remove" => return desugar_table_remove(args),
                        "concat" => return desugar_table_concat(args),
                        "sort" => return desugar_table_sort(args),
                        "move" => return desugar_table_move(args),
                        "pack" => return desugar_table_pack(args),
                        "unpack" => return desugar_table_unpack(args),
                        _ => {}
                    }
                }
                if field == "sub" && is_lua_profile_namespace(object) {
                    return desugar_string_sub_call(args);
                }
                if field == "byte" && is_lua_profile_namespace(object) {
                    return desugar_string_byte_call(args);
                }
            }
            let callee_expr = Expression::new(normalize_expr_kind(callee.kind, read));
            let mut args: Vec<Argument> = args
                .into_iter()
                .map(|mut a| {
                    normalize_expr(&mut a.value);
                    let value_kind =
                        std::mem::replace(&mut a.value.kind, ExprKind::Lit(Literal::Null));
                    if let ExprKind::Spread(inner) = value_kind {
                        a.value = *inner;
                        a.spread = true;
                    } else {
                        a.value.kind = value_kind;
                    }
                    a
                })
                .collect();
            if let ExprKind::Ident(name) = &callee_expr.kind {
                if name == "type" && args.len() == 1 {
                    let arg = args.remove(0).value;
                    return desugar_lua_type_call(arg);
                }
            }
            let callee_is_profile_member = matches!(
                &callee_expr.kind,
                ExprKind::Member { object, .. } if is_lua_profile_namespace(object)
            );
            let callee_is_known_builtin = matches!(
                &callee_expr.kind,
                ExprKind::Ident(name) if is_lua_global_builtin(name) || name == "type"
            );
            if !args.iter().any(|a| a.spread)
                && !callee_is_profile_member
                && !callee_is_known_builtin
            {
                return desugar_lua_value_call(callee_expr, args);
            }
            ExprKind::Call {
                callee: Box::new(callee_expr),
                args,
                optional,
            }
        }
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => {
            let object = Expression::new(normalize_expr_kind(object.kind, read));
            if is_ident(&object, "math") && ctx == read {
                match field.as_str() {
                    "huge" => return ExprKind::Lit(Literal::Float(f64::INFINITY)),
                    "maxinteger" => return ExprKind::Lit(Literal::Int(9_007_199_254_740_991)),
                    "mininteger" => return ExprKind::Lit(Literal::Int(-9_007_199_254_740_991)),
                    _ => {}
                }
            }
            if is_lua_profile_namespace(&object) {
                ExprKind::Member {
                    object: Box::new(object),
                    field,
                    null_safe,
                }
            } else {
                let index = Expression::new(ExprKind::Lit(Literal::Str(field)));
                if ctx == read {
                    return wrap_lua_proto_get(object, index);
                }
                ExprKind::Index {
                    object: Box::new(object),
                    index: Box::new(index),
                    null_safe,
                }
            }
        }
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => {
            let object = Expression::new(normalize_expr_kind(object.kind, read));
            let index = lua_one_based_index(Expression::new(normalize_expr_kind(index.kind, read)));
            if ctx == read {
                return wrap_lua_proto_get(object, index);
            }
            ExprKind::Index {
                object: Box::new(object),
                index: Box::new(index),
                null_safe,
            }
        }
        ExprKind::Sequence(values) => ExprKind::Sequence(
            values
                .into_iter()
                .map(|e| Expression::new(normalize_expr_kind(e.kind, read)))
                .collect(),
        ),
        other => normalize_literal(other),
    };
    kind
}

/// Preserve integer literals for `math.type` (7 vs 7.0).
fn normalize_literal(kind: ExprKind) -> ExprKind {
    match kind {
        ExprKind::Lit(Literal::BigInt(n)) => ExprKind::Lit(Literal::Float(n as f64)),
        other => other,
    }
}

/// Lua tables compile to ECMA arrays/maps (0-based); adjust numeric keys at use sites.
fn lua_one_based_index(index: Expression) -> Expression {
    if matches!(index.kind, ExprKind::Lit(Literal::Str(_))) {
        return index;
    }

    Expression::new(ExprKind::Ternary {
        cond: Box::new(lua_type_is(index.clone(), "number")),
        then: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Sub,
            left: Box::new(index.clone()),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Float(1.0)))),
        })),
        else_: Box::new(index),
    })
}

/// Roots bound in `languages/lua/profile` `[builtins]` (`string.*`, `table.*`, …).
fn is_lua_profile_namespace(object: &Expression) -> bool {
    let ExprKind::Ident(name) = &object.kind else {
        return false;
    };
    matches!(
        name.as_str(),
        "string"
            | "table"
            | "math"
            | "os"
            | "io"
            | "coroutine"
            | "debug"
            | "package"
            | "utf8"
            | "Object"
    )
}
