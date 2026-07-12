//! Lua → JS-shaped AST normalization.
//! Normalizes Lua-specific operations to adapter calls that handle metamethods at runtime.

use vybe_ast::*;

/// Normalize module: transform statements and expressions
pub fn normalize_module(module: &mut Module) {
    for stmt in &mut module.body {
        normalize_stmt(&mut stmt.kind);
    }
}

fn normalize_expr(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Binary { op, left, right } => {
            // Recursively normalize operands first
            normalize_expr(left);
            normalize_expr(right);

            // Wrap operations in adapters that check for metamethods at runtime
            let adapter_name = match op {
                BinOp::Add => "__lua_add",
                BinOp::Sub => "__lua_sub",
                BinOp::Mul => "__lua_mul",
                BinOp::Div => "__lua_div",
                BinOp::Mod => "__lua_mod",
                BinOp::Pow => "__lua_pow",
                BinOp::Lt => "__lua_lt",
                BinOp::LtEq => "__lua_le",
                BinOp::Eq => "__lua_eq",
                BinOp::Concat => "__lua_concat",
                _ => return, // Other ops pass through
            };

            // Replace with call to adapter: __lua_add(left, right)
            let call = ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident(adapter_name.to_string()))),
                args: vec![
                    Argument::positional(left.as_ref().clone()),
                    Argument::positional(right.as_ref().clone()),
                ],
                optional: false,
            };
            expr.kind = call;
        }
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr: inner,
        } => {
            normalize_expr(inner);
            let call = ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident("__lua_unm".to_string()))),
                args: vec![Argument::positional(inner.as_ref().clone())],
                optional: false,
            };
            expr.kind = call;
        }
        ExprKind::Index { object, index, .. } => {
            normalize_expr(object);
            normalize_expr(index);
            let call = ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident("__lua_index".to_string()))),
                args: vec![
                    Argument::positional(object.as_ref().clone()),
                    Argument::positional(index.as_ref().clone()),
                ],
                optional: false,
            };
            expr.kind = call;
        }
        ExprKind::Call { callee, args, .. } => {
            normalize_expr(callee);
            for arg in args.iter_mut() {
                normalize_expr(&mut arg.value);
            }
        }
        ExprKind::Array(elems) => {
            for elem in elems.iter_mut() {
                normalize_expr(&mut elem.value);
                if let Some(key) = &mut elem.key {
                    normalize_expr(key);
                }
            }
        }
        ExprKind::Member { object, .. } => {
            normalize_expr(object);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_expr(cond);
            normalize_expr(then);
            normalize_expr(else_);
        }
        _ => {}
    }
}

fn normalize_stmt(kind: &mut StmtKind) {
    match kind {
        StmtKind::Expr(expr) => normalize_expr(expr),
        StmtKind::Assign { targets, value } => {
            for t in targets {
                normalize_expr(t);
            }
            normalize_expr(value);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_expr(init);
                }
            }
        }
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
        StmtKind::FunctionDecl { body, .. } => {
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
        }
        StmtKind::Return(Some(expr)) => normalize_expr(expr),
        _ => {}
    }
}

/// Simple numeric for → while loop with proper condition
pub(crate) fn build_numeric_for(
    index_var: String,
    start: Expression,
    limit: Expression,
    step: Expression,
    body: Vec<Statement>,
) -> StmtKind {
    let ctrl_var = format!("__lua_for_{}", index_var);
    let limit_var = format!("__lua_limit_{}", index_var);
    let step_var = format!("__lua_step_{}", index_var);

    let ctrl_expr = Expression::new(ExprKind::Ident(ctrl_var.clone()));
    let limit_expr = Expression::new(ExprKind::Ident(limit_var.clone()));
    let step_expr_cond = Expression::new(ExprKind::Ident(step_var.clone()));

    // Lua numeric for: while (step > 0 && ctrl <= limit) || (step < 0 && ctrl >= limit)
    let step_pos = Expression::new(ExprKind::Binary {
        op: BinOp::Gt,
        left: Box::new(step_expr_cond.clone()),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(0)))),
    });

    let step_neg = Expression::new(ExprKind::Binary {
        op: BinOp::Lt,
        left: Box::new(step_expr_cond),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(0)))),
    });

    let ctrl_lte_limit = Expression::new(ExprKind::Binary {
        op: BinOp::LtEq,
        left: Box::new(ctrl_expr.clone()),
        right: Box::new(limit_expr.clone()),
    });

    let ctrl_gte_limit = Expression::new(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(ctrl_expr.clone()),
        right: Box::new(limit_expr.clone()),
    });

    // (step > 0 && ctrl <= limit) || (step < 0 && ctrl >= limit)
    let cond = Expression::new(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(step_pos),
            right: Box::new(ctrl_lte_limit),
        })),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(step_neg),
            right: Box::new(ctrl_gte_limit),
        })),
    });

    // Bind user-visible loop variable each iteration
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

    let step_var_expr = Expression::new(ExprKind::Ident(step_var.clone()));
    let increment = Statement::new(StmtKind::Assign {
        targets: vec![ctrl_expr.clone()],
        value: Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(ctrl_expr),
            right: Box::new(step_var_expr),
        }),
    });

    let while_body = vec![Statement::new(StmtKind::Block(scoped_body)), increment];

    StmtKind::Block(vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![
                VarDeclarator {
                    pattern: BindingPattern::Ident(ctrl_var),
                    type_hint: None,
                    init: Some(start),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(limit_var),
                    type_hint: None,
                    init: Some(limit),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(step_var),
                    type_hint: None,
                    init: Some(step),
                    array_bounds: None,
                    with_events: false,
                },
            ],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::While {
            cond,
            body: while_body,
            else_body: None,
        }),
    ])
}
