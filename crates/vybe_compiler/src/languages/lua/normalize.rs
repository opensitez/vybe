//! Lua → JS-shaped AST — minimal normalization.
//! The compiler + emitters (in vybe_emitter) handle semantics, metamethods, and type coercion.

use crate::ast::*;

/// Minimal module normalization - just walk statements
pub fn normalize_module(module: &mut Module) {
    for stmt in &mut module.body {
        normalize_stmt(&mut stmt.kind);
    }
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
            cond: _,
            update: _,
            body,
        } => {
            if let Some(init_stmt) = init {
                normalize_stmt(&mut init_stmt.kind);
            }
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
        }
        StmtKind::ForIn { body, else_body, .. } => {
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
            if let Some(eb) = else_body {
                for s in eb.iter_mut() {
                    normalize_stmt(&mut s.kind);
                }
            }
        }
        StmtKind::While { body, .. } => {
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
        }
        StmtKind::DoWhile { body, .. } => {
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            for s in then_body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
            for (_, body) in elifs.iter_mut() {
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
    let step_expr = Expression::new(ExprKind::Ident(step_var.clone()));
    
    // Simplified: assume positive step, use <=
    let cond = Expression::new(ExprKind::Binary {
        op: BinOp::LtEq,
        left: Box::new(ctrl_expr.clone()),
        right: Box::new(limit_expr),
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
    
    let increment = Statement::new(StmtKind::Assign {
        targets: vec![ctrl_expr.clone()],
        value: Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(ctrl_expr),
            right: Box::new(step_expr),
        }),
    });
    
    let while_body = vec![
        Statement::new(StmtKind::Block(scoped_body)),
        increment,
    ];

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
