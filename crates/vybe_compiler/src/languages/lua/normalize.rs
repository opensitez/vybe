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
        body.extend(desugar_multi_assign(stmt));
    }
    module.body = body;
    for stmt in &mut module.body {
        normalize_stmt(&mut stmt.kind);
    }
}

fn desugar_multi_assign(stmt: Statement) -> Vec<Statement> {
    let StmtKind::Assign { targets, value } = stmt.kind else {
        return vec![stmt];
    };
    if targets.len() <= 1 {
        return vec![Statement::with_span(StmtKind::Assign { targets, value }, stmt.span)];
    }
    let ExprKind::Sequence(values) = value.kind else {
        return vec![Statement::with_span(StmtKind::Assign { targets, value }, stmt.span)];
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
        StmtKind::For { update, .. } if for_update_is_zero_step(update.as_ref()) => {
            Some(StmtKind::Throw {
                expr: Some(Expression::new(ExprKind::Lit(Literal::Str(
                    "'step' argument is zero".to_string(),
                )))),
                cause: None,
            })
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

fn build_numeric_for(
    index_var: String,
    start: Expression,
    limit: Expression,
    step: Expression,
    body: Vec<Statement>,
) -> StmtKind {
    let limit_temp = format!("__lua_for_limit_{index_var}");
    let init = Box::new(Statement::new(StmtKind::VarDecl {
        declarations: vec![
            VarDeclarator {
                pattern: BindingPattern::Ident(index_var.clone()),
                type_hint: None,
                init: Some(start),
                array_bounds: None,
                with_events: false,
            },
            VarDeclarator {
                pattern: BindingPattern::Ident(limit_temp.clone()),
                type_hint: None,
                init: Some(limit),
                array_bounds: None,
                with_events: false,
            },
        ],
        kind: VarDeclKind::Let,
    }));

    let index_expr = Expression::new(ExprKind::Ident(index_var));
    let compare_op = if lua_step_is_negative(&step) {
        BinOp::GtEq
    } else {
        BinOp::LtEq
    };
    let cond = Expression::new(ExprKind::Binary {
        op: compare_op,
        left: Box::new(index_expr.clone()),
        right: Box::new(Expression::new(ExprKind::Ident(limit_temp))),
    });
    let update = Expression::new(ExprKind::Assign {
        target: Box::new(index_expr.clone()),
        value: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(index_expr),
            right: Box::new(step),
        })),
    });

    StmtKind::For {
        init: Some(init),
        cond: Some(cond),
        update: Some(update),
        body: lua_scoped_body(body),
    }
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

fn desugar_lua_mod(left: Expression, right: Expression) -> ExprKind {
    let quotient = Expression::new(ExprKind::Binary {
        op: BinOp::Div,
        left: Box::new(left.clone()),
        right: Box::new(right.clone()),
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
    ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(left),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Mul,
            left: Box::new(floored),
            right: Box::new(right),
        })),
    }
}

fn desugar_lua_eq(left: Expression, right: Expression) -> ExprKind {
    ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(left),
        right: Box::new(right),
    }
}

fn desugar_lua_rel(op: BinOp, left: Expression, right: Expression) -> ExprKind {
    ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
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

fn lua_mm_lookup(obj: &str, mm: &str) -> Expression {
    lua_raw_index(lua_obj_mt(obj), lit_str(mm))
}

fn lua_mm_call(mm_fn: Expression, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(mm_fn),
        args: vec![
            Argument::positional(left),
            Argument::positional(right),
        ],
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
fn desugar_lua_mm_binop(
    left: Expression,
    right: Expression,
    mm: &str,
    fallback_op: BinOp,
    commutative: bool,
) -> ExprKind {
    let a = "__lua_a";
    let b = "__lua_b";
    let f = "__lua_mm_fn";
    let fallback = Expression::new(ExprKind::Binary {
        op: fallback_op,
        left: Box::new(Expression::new(ExprKind::Ident(a.to_string()))),
        right: Box::new(Expression::new(ExprKind::Ident(b.to_string()))),
    });
    let mut stmts = vec![Statement::new(StmtKind::VarDecl {
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
    })];
    let left_mm = lua_mm_lookup(a, mm);
    stmts.push(Statement::new(StmtKind::If {
        cond: lua_type_is_function(left_mm.clone()),
        then_body: vec![
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(f.to_string()),
                    type_hint: None,
                    init: Some(left_mm),
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
    }));
    let right_mm = lua_mm_lookup(b, mm);
    stmts.push(Statement::new(StmtKind::If {
        cond: lua_type_is_function(right_mm.clone()),
        then_body: vec![
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(f.to_string()),
                    type_hint: None,
                    init: Some(right_mm),
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
    }));
    if commutative {
        // already handled with swapped args above
    }
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
            Some(wrap_lua_proto_set(
                object,
                lit_str(&field),
                value.clone(),
            ))
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
    ExprKind::Unary {
        op: fallback_op,
        expr: Box::new(expr),
    }
}

fn desugar_lua_len_call(args: Vec<Argument>) -> ExprKind {
    let obj = args
        .into_iter()
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("__lua_len".to_string()))),
        args: vec![Argument::positional(obj)],
        optional: false,
    }
}

fn desugar_lua_tostring(args: Vec<Argument>) -> ExprKind {
    let obj = args
        .into_iter()
        .next()
        .map(|a| Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read)))
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
    call_ident("tostring", vec![obj]).kind
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
            then_body: vec![lua_return(Expression::new(ExprKind::Ident(hidden.to_string())))],
            elifs: Vec::new(),
            else_body: None,
        }),
        lua_return(Expression::new(ExprKind::Ident(mt.to_string()))),
    ])
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
    ExprKind::Index {
        object: Box::new(obj),
        index: Box::new(key),
        null_safe: false,
    }
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
        lua_return(Expression::new(ExprKind::Sequence(vec![
            Expression::new(ExprKind::Ident(i.to_string())),
            Expression::new(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(Expression::new(ExprKind::Ident(slot.to_string()))),
                right: Box::new(Expression::new(ExprKind::Ident(i.to_string()))),
            }),
        ]))),
    ])
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
        target: Box::new(Expression::new(ExprKind::Ident("__lua_rng_seed".to_string()))),
        value: Box::new(seed),
    }
}

fn desugar_math_random(args: Vec<Argument>) -> ExprKind {
    let mut iter = args.into_iter();
    let a = iter.next().map(|a| {
        Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read))
    });
    let b = iter.next().map(|a| {
        Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read))
    });
    let seed_slot = "__lua_rng_seed";
    let r_slot = "__lua_rng_r";
    let next_seed = Expression::new(ExprKind::Binary {
        op: BinOp::UShr,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Mul,
                left: Box::new(Expression::new(ExprKind::Ident(seed_slot.to_string()))),
                right: Box::new(lit_float(1_103_515_245.0)),
            })),
            right: Box::new(lit_float(12_345.0)),
        })),
        right: Box::new(lit_float(0.0)),
    });
    let frac = Expression::new(ExprKind::Binary {
        op: BinOp::Div,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(Expression::new(ExprKind::Ident(r_slot.to_string()))),
            right: Box::new(lit_float(2_147_483_647.0)),
        })),
        right: Box::new(lit_float(2_147_483_648.0)),
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
    let f = "__lua_call_fn";
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
        Statement::new(StmtKind::If {
            cond: lua_type_is_function(Expression::new(ExprKind::Ident(c.to_string()))),
            then_body: vec![lua_return(direct)],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: lua_type_is_function(lua_mm_lookup(c, "__call")),
            then_body: vec![lua_return(Expression::new(ExprKind::Call {
                callee: Box::new(lua_mm_lookup(c, "__call")),
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
        lua_return(Expression::new(ExprKind::Lit(Literal::Null))),
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
    ExprKind::Assign {
        target: Box::new(Expression::new(ExprKind::Index {
            object: Box::new(object),
            index: Box::new(index),
            null_safe: false,
        })),
        value: Box::new(value),
    }
}

/// Colon method dispatch — own function slot, then `__lua_proto` (flat stmts, no nested IIFE callee).
fn desugar_lua_colon_call(
    object: Expression,
    field: String,
    mut args: Vec<Argument>,
) -> ExprKind {
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
    call_args.extend(
        args.into_iter()
            .map(|mut a| {
                a.value = Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read));
                a
            }),
    );
    lua_iife(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![
                VarDeclarator {
                    pattern: BindingPattern::Ident(recv.to_string()),
                    type_hint: None,
                    init: Some(Expression::new(normalize_expr_kind(object.kind, LuaExprCtx::Read))),
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

fn desugar_lua_colon_call_expr(
    object: Expression,
    field: String,
    args: Vec<Argument>,
) -> ExprKind {
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
                right: Box::new(lua_key_is_present(mt_slot, "__metatable")),
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
                Expression::new(ExprKind::Ident(key.to_string())),
            )),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
        })),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(lua_raw_index(
                Expression::new(ExprKind::Ident(object.to_string())),
                Expression::new(ExprKind::Ident(key.to_string())),
            )),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Str("undefined".to_string())))),
        })),
    })
}

/// Lua `__index` / field read — data in `__lua_data`, methods on table (JS prototype split).
fn wrap_lua_proto_get(object: Expression, index: Expression) -> ExprKind {
    ExprKind::Index {
        object: Box::new(object),
        index: Box::new(index),
        null_safe: false,
    }
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
        .map(|a| lua_one_based_index(Expression::new(normalize_expr_kind(a.value.kind, LuaExprCtx::Read))))
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

fn is_ident(expr: &Expression, name: &str) -> bool {
    matches!(&expr.kind, ExprKind::Ident(n) if n == name)
}

fn for_update_is_zero_step(update: Option<&Expression>) -> bool {
    let Some(update) = update else {
        return false;
    };
    let ExprKind::Assign { value, .. } = &update.kind else {
        return false;
    };
    let ExprKind::Binary { op: BinOp::Add, right, .. } = &value.kind else {
        return false;
    };
    expr_is_zero(right)
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
        ExprKind::Lambda { params, body, is_async, captures } => {
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
                        normalize_expr_ctx(key, LuaExprCtx::Write);
                    }
                    normalize_expr(&mut e.value);
                    e
                })
                .collect();
            ExprKind::Array(elements)
        }
        ExprKind::Call { callee, args, optional } => {
            if let ExprKind::Member {
                object,
                field,
                null_safe: true,
                ..
            } = &callee.kind
            {
                return desugar_lua_colon_call_expr(
                    object.as_ref().clone(),
                    field.clone(),
                    args,
                );
            }
            if let ExprKind::Ident(name) = &callee.kind {
                match name.as_str() {
                    "setmetatable" => return desugar_setmetatable(args),
                    "getmetatable" => return desugar_getmetatable(args),
                    "rawget" => return desugar_rawget(args),
                    "rawset" => return desugar_rawset(args),
                    "pcall" => return desugar_pcall(args),
                    "tostring" => return desugar_lua_tostring(args),
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
                        "type" => return desugar_math_type(args),
                        "ult" => return desugar_math_ult(args),
                        "randomseed" => return desugar_math_randomseed(args),
                        "random" => return desugar_math_random(args),
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
            let args: Vec<Argument> = args
                .into_iter()
                .map(|mut a| {
                    normalize_expr(&mut a.value);
                    a
                })
                .collect();
            if let ExprKind::Ident(name) = &callee_expr.kind {
                if name == "type" && args.len() == 1 {
                    return ExprKind::Unary {
                        op: UnaryOp::Typeof,
                        expr: Box::new(args.into_iter().next().unwrap().value),
                    };
                }
            }
            ExprKind::Call {
                callee: Box::new(callee_expr),
                args,
                optional,
            }
        }
        ExprKind::Member { object, field, null_safe } => {
            let object = Expression::new(normalize_expr_kind(object.kind, read));
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
        ExprKind::Index { object, index, null_safe } => {
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
    Expression::new(ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(index),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Float(1.0)))),
    })
}

/// Roots bound in `languages/lua/profile` `[builtins]` (`string.*`, `table.*`, …).
fn is_lua_profile_namespace(object: &Expression) -> bool {
    let ExprKind::Ident(name) = &object.kind else {
        return false;
    };
    matches!(
        name.as_str(),
        "string" | "table" | "math" | "os" | "io" | "coroutine" | "debug" | "package" | "utf8"
            | "Object"
    )
}
