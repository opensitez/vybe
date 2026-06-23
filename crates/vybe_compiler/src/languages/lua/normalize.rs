//! Lua → JS-shaped AST so the **shared compiler** lowers control flow and
//! expressions through `crates/vybe_compiler/src/emitter/` (`loops`, `expressions`,
//! `strings`, `collections`, `io`, …) — same path as JavaScript.

use crate::ast::*;

pub fn normalize_module(module: &mut Module) {
    for stmt in &mut module.body {
        normalize_stmt(&mut stmt.kind);
    }
}

fn normalize_stmt(kind: &mut StmtKind) {
    match kind {
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
        StmtKind::FunctionDecl { params: _, body, .. } => {
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
        ExprKind::Binary { op, left, right } => ExprKind::Binary {
            op,
            left: Box::new(Expression::new(normalize_expr_kind(left.kind))),
            right: Box::new(Expression::new(normalize_expr_kind(right.kind))),
        },
        ExprKind::Call { callee, args, optional } => {
            let callee = Box::new(Expression::new(normalize_expr_kind(callee.kind)));
            let args: Vec<Argument> = args
                .into_iter()
                .map(|mut a| {
                    normalize_expr(&mut a.value);
                    a
                })
                .collect();
            // Lua `type(x)` → JS `typeof x`
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
        ExprKind::Member { object, field, null_safe } => ExprKind::Member {
            object: Box::new(Expression::new(normalize_expr_kind(object.kind))),
            field,
            null_safe,
        },
        ExprKind::Index { object, index, null_safe } => ExprKind::Index {
            object: Box::new(Expression::new(normalize_expr_kind(object.kind))),
            index: Box::new(Expression::new(normalize_expr_kind(index.kind))),
            null_safe,
        },
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
