//! Reversible Lua numeric-loop policies. Journals retain scalar changes only;
//! user expressions and closure bodies move into their lowered parents.
use vybe_ast::*;
mod predicates {
    include!(concat!(env!("OUT_DIR"), "/lua_numeric_predicates.rs"));
}

pub struct CaptureJournal {
    var: String,
    edits: Vec<(usize, Vec<(usize, String)>)>,
}
#[derive(Clone, Copy)]
enum Walk {
    Capture,
    Rename,
}
// Match the existing Lua lowering's traversal and shadowing policies exactly.
fn visit_expr(
    expr: &mut Expression,
    mode: Walk,
    var: &str,
    action: &mut impl FnMut(&mut Expression) -> bool,
) {
    if !action(expr) {
        return;
    }
    match &mut expr.kind {
        ExprKind::Lambda { params, body, .. } => {
            if matches!(mode, Walk::Rename) && !params.iter().any(|p| p.name == var) {
                match body {
                    LambdaBody::Expr(expr) => visit_expr(expr, mode, var, action),
                    LambdaBody::Block(body) => visit_stmts(body, mode, var, action),
                }
            }
        }
        ExprKind::Binary { left, right, .. } => {
            visit_expr(left, mode, var, action);
            visit_expr(right, mode, var, action);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr) => visit_expr(expr, mode, var, action),
        ExprKind::Call { callee, args, .. } => {
            visit_expr(callee, mode, var, action);
            for arg in args {
                visit_expr(&mut arg.value, mode, var, action);
            }
        }
        ExprKind::Member { object, .. } => visit_expr(object, mode, var, action),
        ExprKind::Index { object, index, .. } => {
            visit_expr(object, mode, var, action);
            visit_expr(index, mode, var, action);
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &mut element.key {
                    visit_expr(key, mode, var, action);
                }
                visit_expr(&mut element.value, mode, var, action);
            }
        }
        ExprKind::Tuple(values) | ExprKind::Sequence(values) => {
            for value in values {
                visit_expr(value, mode, var, action);
            }
        }
        ExprKind::Ternary { cond, then, else_ } => {
            visit_expr(cond, mode, var, action);
            visit_expr(then, mode, var, action);
            visit_expr(else_, mode, var, action);
        }
        _ => {}
    }
}
fn visit_stmts(
    stmts: &mut [Statement],
    mode: Walk,
    var: &str,
    action: &mut impl FnMut(&mut Expression) -> bool,
) {
    for stmt in stmts {
        match &mut stmt.kind {
            StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
                visit_expr(expr, mode, var, action);
            }
            StmtKind::Assign { targets, value, .. } => {
                if matches!(mode, Walk::Rename) {
                    for target in targets {
                        visit_expr(target, mode, var, action);
                    }
                }
                visit_expr(value, mode, var, action);
            }
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if matches!(mode, Walk::Rename)
                        && matches!(&decl.pattern, BindingPattern::Ident(name) if name == var)
                    {
                        continue;
                    }
                    if let Some(init) = &mut decl.init {
                        visit_expr(init, mode, var, action);
                    }
                }
            }
            StmtKind::Block(body)
            | StmtKind::While { body, .. }
            | StmtKind::DoWhile { body, .. }
            | StmtKind::For { body, .. } => visit_stmts(body, mode, var, action),
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                if matches!(mode, Walk::Rename) {
                    visit_expr(cond, mode, var, action);
                }
                visit_stmts(then_body, mode, var, action);
                for (cond, body) in elifs {
                    if matches!(mode, Walk::Rename) {
                        visit_expr(cond, mode, var, action);
                    }
                    visit_stmts(body, mode, var, action);
                }
                if let Some(body) = else_body {
                    visit_stmts(body, mode, var, action);
                }
            }
            _ => {}
        }
    }
}

pub fn capture(body: &mut [Statement], var: &str) -> CaptureJournal {
    let mut edits = Vec::new();
    let mut index = 0;
    visit_stmts(body, Walk::Capture, var, &mut |expr| {
        let position = index;
        index += 1;
        if !matches!(expr.kind, ExprKind::Lambda { .. }) {
            return true;
        }
        if !predicates::lua_expr_contains_unshadowed_ident(expr, var) {
            return false;
        }
        let name = format!("__lua_capture_{var}");
        let span = expr.span;
        let mut inner = std::mem::replace(expr, Expression::new(ExprKind::Lit(Literal::Null)));
        let mut renames = Vec::new();
        let mut ordinal = 0;
        visit_expr(&mut inner, Walk::Rename, var, &mut |node| {
            let position = ordinal;
            ordinal += 1;
            if let ExprKind::Ident(ident) = &mut node.kind
                && ident == var
            {
                renames.push((position, std::mem::replace(ident, name.clone())));
            }
            true
        });
        expr.kind = ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Lambda {
                params: vec![Param {
                    name,
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                }],
                body: LambdaBody::Expr(Box::new(inner)),
                is_async: false,
                captures: Vec::new(),
            })),
            args: vec![Argument::positional(Expression::ident(var))],
            optional: false,
        };
        expr.span = span;
        edits.push((position, renames));
        false
    });
    CaptureJournal {
        var: var.into(),
        edits,
    }
}
pub fn restore_capture(body: &mut [Statement], journal: CaptureJournal) {
    let mut edits = journal.edits.into_iter().peekable();
    let mut index = 0;
    visit_stmts(body, Walk::Capture, &journal.var, &mut |expr| {
        let position = index;
        index += 1;
        if edits.peek().is_none_or(|edit| edit.0 != position) {
            return true;
        }
        let (_, renames) = edits.next().unwrap();
        let ExprKind::Call { callee, .. } =
            std::mem::replace(&mut expr.kind, ExprKind::Lit(Literal::Null))
        else {
            unreachable!()
        };
        let ExprKind::Lambda {
            body: LambdaBody::Expr(inner),
            ..
        } = callee.kind
        else {
            unreachable!()
        };
        *expr = *inner;
        let mut renames = renames.into_iter().peekable();
        let mut ordinal = 0;
        visit_expr(expr, Walk::Rename, &journal.var, &mut |node| {
            let position = ordinal;
            ordinal += 1;
            if renames.peek().is_some_and(|rename| rename.0 == position) {
                let ExprKind::Ident(name) = &mut node.kind else {
                    unreachable!()
                };
                *name = renames.next().unwrap().1;
            }
            true
        });
        assert!(renames.next().is_none());
        false
    });
    assert!(edits.next().is_none());
}

// Numeric promotion follows the walker's limited traversal; a journal keeps
// exact i64 values, including those that lose precision after f64 conversion.
fn numeric_literals(expr: &mut Expression, action: &mut impl FnMut(&mut Literal)) {
    match &mut expr.kind {
        ExprKind::Lit(literal) => action(literal),
        ExprKind::Unary { expr, .. } => numeric_literals(expr, action),
        ExprKind::Binary { left, right, .. } => {
            numeric_literals(left, action);
            numeric_literals(right, action);
        }
        ExprKind::Call { args, .. } => {
            for arg in args {
                numeric_literals(&mut arg.value, action);
            }
        }
        _ => {}
    }
}
pub fn contains_float(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Float(_)) => true,
        ExprKind::Unary { expr, .. } => contains_float(expr),
        ExprKind::Binary { left, right, .. } => contains_float(left) || contains_float(right),
        ExprKind::Call { args, .. } => args.iter().any(|arg| contains_float(&arg.value)),
        _ => false,
    }
}
pub fn promote(expr: &mut Expression) -> Vec<(usize, i64)> {
    let mut journal = Vec::new();
    let mut index = 0;
    numeric_literals(expr, &mut |literal| {
        if let Literal::Int(value) = literal {
            journal.push((index, *value));
            *literal = Literal::Float(*value as f64);
        }
        index += 1;
    });
    journal
}
pub fn restore_promotion(expr: &mut Expression, journal: Vec<(usize, i64)>) {
    let mut journal = journal.into_iter().peekable();
    let mut index = 0;
    numeric_literals(expr, &mut |literal| {
        if journal.peek().is_some_and(|edit| edit.0 == index) {
            *literal = Literal::Int(journal.next().unwrap().1);
        }
        index += 1;
    });
    assert!(journal.next().is_none());
}

pub fn build(
    var: &str,
    start: Expression,
    limit: Expression,
    step: Expression,
    body: Vec<Statement>,
) -> StmtKind {
    let ctrl = format!("__lua_for_{var}");
    let limit_name = format!("__lua_limit_{var}");
    let step_name = format!("__lua_step_{var}");
    let binary = |op, left, right| {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        })
    };
    let zero = || binary(BinOp::Eq, Expression::ident(&step_name), Expression::int(0));
    let ternary = |cond, then, else_| {
        Expression::new(ExprKind::Ternary {
            cond: Box::new(cond),
            then: Box::new(then),
            else_: Box::new(else_),
        })
    };
    let declaration = |name: String, init| VarDeclarator {
        pattern: BindingPattern::Ident(name),
        type_hint: None,
        init: Some(init),
        array_bounds: None,
        with_events: false,
    };
    let cond = ternary(
        zero(),
        binary(
            BinOp::Eq,
            Expression::ident(&ctrl),
            Expression::ident(&limit_name),
        ),
        ternary(
            binary(BinOp::Gt, Expression::ident(&step_name), Expression::int(0)),
            binary(
                BinOp::LtEq,
                Expression::ident(&ctrl),
                Expression::ident(&limit_name),
            ),
            binary(
                BinOp::GtEq,
                Expression::ident(&ctrl),
                Expression::ident(&limit_name),
            ),
        ),
    );
    let update = Expression::new(ExprKind::Assign {
        target: Box::new(Expression::ident(&ctrl)),
        value: Box::new(ternary(
            zero(),
            binary(
                BinOp::Add,
                Expression::ident(&limit_name),
                Expression::int(1),
            ),
            binary(
                BinOp::Add,
                Expression::ident(&ctrl),
                Expression::ident(&step_name),
            ),
        )),
    });
    let mut scoped_body = vec![Statement::new(StmtKind::VarDecl {
        declarations: vec![declaration(var.into(), Expression::ident(&ctrl))],
        kind: VarDeclKind::Let,
    })];
    scoped_body.extend(body);
    StmtKind::For {
        init: Some(Box::new(Statement::new(StmtKind::VarDecl {
            declarations: vec![
                declaration(ctrl, start),
                declaration(limit_name, limit),
                declaration(step_name, step),
            ],
            kind: VarDeclKind::Let,
        }))),
        cond: Some(cond),
        update: Some(update),
        body: vec![Statement::new(StmtKind::Block(scoped_body))],
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn promotion_restores_exact_large_integers_and_existing_floats() {
        let mut expr = Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::int(9_007_199_254_740_993)),
            right: Box::new(Expression::float(0.5)),
        });
        let expected = format!("{expr:?}");
        let journal = promote(&mut expr);
        assert_eq!(journal, vec![(0, 9_007_199_254_740_993)]);
        restore_promotion(&mut expr, journal);
        assert_eq!(format!("{expr:?}"), expected);
    }
    #[test]
    fn capture_restores_names_without_confusing_existing_generated_names() {
        let lambda = Expression::with_span(
            ExprKind::Lambda {
                params: Vec::new(),
                body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Sequence(vec![
                    Expression::ident("i"),
                    Expression::ident("__lua_capture_i"),
                    Expression::ident("i"),
                ])))),
                is_async: false,
                captures: Vec::new(),
            },
            Span {
                start_line: 3,
                start_col: 7,
                end_line: 3,
                end_col: 29,
            },
        );
        let mut body = vec![Statement::new(StmtKind::Expr(lambda))];
        let expected = format!("{body:?}");
        let journal = capture(&mut body, "i");
        assert_eq!(journal.edits.len(), 1);
        let StmtKind::Expr(Expression {
            kind: ExprKind::Call { callee, .. },
            ..
        }) = &body[0].kind
        else {
            panic!("closure was not captured")
        };
        let ExprKind::Lambda { params, .. } = &callee.kind else {
            unreachable!()
        };
        assert_eq!(params[0].name, "__lua_capture_i");
        restore_capture(&mut body, journal);
        assert_eq!(format!("{body:?}"), expected);
    }
}
