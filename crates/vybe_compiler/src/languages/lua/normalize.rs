//! Lua → JS-shaped AST so the **shared compiler** lowers control flow and
//! expressions through `crates/vybe_compiler/src/emitter/` (`loops`, `expressions`,
//! `strings`, `collections`, `io`, …) — same path as JavaScript.

use crate::ast::*;

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
            for t in targets.iter_mut() {
                normalize_expr(t);
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

fn normalize_expr(expr: &mut Expression) {
    expr.kind = normalize_expr_kind(std::mem::replace(&mut expr.kind, ExprKind::Lit(Literal::Null)));
}

fn normalize_expr_kind(kind: ExprKind) -> ExprKind {
    let kind = match kind {
        ExprKind::Unary { op, expr } => ExprKind::Unary {
            op,
            expr: Box::new(Expression::new(normalize_expr_kind(expr.kind))),
        },
        ExprKind::Binary { op, left, right } => {
            let left = Expression::new(normalize_expr_kind(left.kind));
            let right = Expression::new(normalize_expr_kind(right.kind));
            if op == BinOp::Concat {
                return ExprKind::Binary {
                    op,
                    left: Box::new(lua_to_string_expr(left)),
                    right: Box::new(lua_to_string_expr(right)),
                };
            }
            ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            }
        }
        ExprKind::Assign { target, value } => ExprKind::Assign {
            target: Box::new(Expression::new(normalize_expr_kind(target.kind))),
            value: Box::new(Expression::new(normalize_expr_kind(value.kind))),
        },
        ExprKind::Lambda { params, body, is_async, captures } => {
            let body = match body {
                LambdaBody::Expr(expr) => {
                    LambdaBody::Expr(Box::new(Expression::new(normalize_expr_kind(expr.kind))))
                }
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
                        normalize_expr(key);
                    }
                    normalize_expr(&mut e.value);
                    e
                })
                .collect();
            ExprKind::Array(elements)
        }
        ExprKind::Call { callee, args, optional } => {
            let callee = Box::new(Expression::new(normalize_expr_kind(callee.kind)));
            let args: Vec<Argument> = args
                .into_iter()
                .map(|mut a| {
                    normalize_expr(&mut a.value);
                    a
                })
                .collect();
            if let ExprKind::Ident(name) = &callee.kind {
                if name == "type" && args.len() == 1 {
                    return ExprKind::Unary {
                        op: UnaryOp::Typeof,
                        expr: Box::new(args.into_iter().next().unwrap().value),
                    };
                }
            }
            ExprKind::Call {
                callee,
                args,
                optional,
            }
        }
        ExprKind::Member { object, field, null_safe } => {
            // Lua `t.name` → `t["name"]` so reads/writes use `common::collections`
            // (`emit_get` / `emit_set`) on ecma:array / ecma:map — same as JS `obj[k]`.
            let object = Expression::new(normalize_expr_kind(object.kind));
            ExprKind::Index {
                object: Box::new(object),
                index: Box::new(Expression::new(ExprKind::Lit(Literal::Str(field)))),
                null_safe,
            }
        }
        ExprKind::Index { object, index, null_safe } => {
            let object = Expression::new(normalize_expr_kind(object.kind));
            let index = Expression::new(normalize_expr_kind(index.kind));
            ExprKind::Index {
                object: Box::new(object),
                index: Box::new(lua_one_based_index(index)),
                null_safe,
            }
        }
        ExprKind::Sequence(values) => ExprKind::Sequence(
            values
                .into_iter()
                .map(|e| Expression::new(normalize_expr_kind(e.kind)))
                .collect(),
        ),
        other => normalize_literal(other),
    };
    kind
}

/// ECMA runtime uses f64; emit JS-shaped numeric literals like the JS walker.
fn normalize_literal(kind: ExprKind) -> ExprKind {
    match kind {
        ExprKind::Lit(Literal::Int(n)) => ExprKind::Lit(Literal::Float(n as f64)),
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

fn lua_to_string_expr(expr: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("tostring".to_string()))),
        args: vec![Argument::positional(expr)],
        optional: false,
    })
}
