//! Conservative liveness for state synthesized by Go's defer normalization.
//! Unknown forms retain the state; this is not general dead-code elimination.

use vybe_ast::{BindingPattern, ExprKind, Expression, LambdaBody, ObjectProperty, Statement, StmtKind};

pub(super) fn is_unused(
    body: &[Statement], final_return: Option<&Expression>, names: &[&str],
) -> bool {
    !body.iter().any(|stmt| stmt_may_use(stmt, names))
        && !final_return.is_some_and(|expr| expr_may_use(expr, names))
}

fn expr_may_use(expr: &Expression, names: &[&str]) -> bool {
    match &expr.kind {
        ExprKind::Lit(_) => false,
        ExprKind::Ident(name) => names.contains(&name.as_str()),
        ExprKind::Unary { expr, .. } | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr) => expr_may_use(expr, names),
        ExprKind::FuncRef(name) => names.contains(&name.as_str()),
        ExprKind::Tuple(values) | ExprKind::Sequence(values) =>
            values.iter().any(|value| expr_may_use(value, names)),
        ExprKind::Array(values) => values.iter().any(|element|
            element.key.as_ref().is_some_and(|key| expr_may_use(key, names))
                || expr_may_use(&element.value, names)),
        // Inspect every capture/default/body; do not infer whether a closure
        // executes. Shadowing may conservatively keep state, never remove it.
        // Any unrecognized child form still takes the retaining fallback.
        ExprKind::Lambda { params, body, captures, .. } => {
            captures.iter().any(|name| names.contains(&name.as_str()))
                || params.iter().any(|param|
                    param.default.as_ref().is_some_and(|value| expr_may_use(value, names)))
                || match body {
                    LambdaBody::Expr(value) => expr_may_use(value, names),
                    LambdaBody::Block(body) => body.iter().any(|stmt| stmt_may_use(stmt, names)),
                }
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right } => {
            expr_may_use(left, names) || expr_may_use(right, names)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            expr_may_use(cond, names) || expr_may_use(then, names)
                || expr_may_use(else_, names)
        }
        ExprKind::Call { callee, args, .. } => {
            expr_may_use(callee, names)
                || args.iter().any(|arg| expr_may_use(&arg.value, names))
        }
        ExprKind::Member { object, .. } => expr_may_use(object, names),
        ExprKind::Index { object, index, .. } => {
            expr_may_use(object, names) || expr_may_use(index, names)
        }
        ExprKind::Assign { target, value } => {
            expr_may_use(target, names) || expr_may_use(value, names)
        }
        ExprKind::Object(properties) => properties.iter().any(|property| match property {
            ObjectProperty::KeyValue { key, value } => {
                expr_may_use(key, names) || expr_may_use(value, names)
            }
            _ => true,
        }),
        // Unknown shapes retain state, including unhandled nested functions.
        // This pass proves absence of references, never invocation behavior.
        _ => true,
    }
}

fn stmt_may_use(stmt: &Statement, names: &[&str]) -> bool {
    let body_may_use = |body: &[Statement]| body.iter().any(|stmt| stmt_may_use(stmt, names));
    match &stmt.kind {
        StmtKind::Empty => false,
        StmtKind::Expr(expr) => expr_may_use(expr, names),
        StmtKind::Return(expr) => expr.as_ref().is_some_and(|expr| expr_may_use(expr, names)),
        StmtKind::Throw { expr, cause } => {
            expr.as_ref().is_some_and(|expr| expr_may_use(expr, names))
                || cause.as_ref().is_some_and(|expr| expr_may_use(expr, names))
        }
        StmtKind::Assign { targets, value, .. } => {
            targets.iter().any(|expr| expr_may_use(expr, names)) || expr_may_use(value, names)
        }
        StmtKind::CompoundAssign { target, value, .. } =>
            expr_may_use(target, names) || expr_may_use(value, names),
        StmtKind::VarDecl { declarations, .. } => declarations.iter().any(|decl| {
            let binding_may_use = match &decl.pattern {
                BindingPattern::Ident(name) => names.contains(&name.as_str()),
                _ => true,
            };
            binding_may_use
                || decl.init.as_ref().is_some_and(|expr| expr_may_use(expr, names))
                || decl.array_bounds.as_ref().is_some_and(|bounds| {
                    bounds.iter().any(|expr| expr_may_use(expr, names))
                })
        }),
        StmtKind::Block(body) => body_may_use(body),
        StmtKind::If { cond, then_body, elifs, else_body } => {
            expr_may_use(cond, names) || body_may_use(then_body)
                || elifs.iter().any(|(cond, body)| {
                    expr_may_use(cond, names) || body_may_use(body)
                })
                || else_body.as_ref().is_some_and(|body| body_may_use(body))
        }
        StmtKind::For { init, cond, update, body } => {
            init.as_ref().is_some_and(|stmt| stmt_may_use(stmt, names))
                || cond.as_ref().is_some_and(|expr| expr_may_use(expr, names))
                || update.as_ref().is_some_and(|expr| expr_may_use(expr, names))
                || body_may_use(body)
        }
        StmtKind::While { cond, body, else_body } => {
            expr_may_use(cond, names) || body_may_use(body)
                || else_body.as_ref().is_some_and(|body| body_may_use(body))
        }
        StmtKind::DoWhile { body, cond, .. } => {
            body_may_use(body) || expr_may_use(cond, names)
        }
        StmtKind::ForIn { iter, body, else_body, .. } => {
            expr_may_use(iter, names) || body_may_use(body)
                || else_body.as_ref().is_some_and(|body| body_may_use(body))
        }
        _ => true,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::{ArrayElement, Param, PassBy, PlaceExpr};

    fn unused(expression: Expression) -> bool {
        is_unused(&[Statement::new(StmtKind::Return(Some(expression)))], None,
            &["panic_value", "panic_active", "in_defer", "recover_entry"])
    }

    fn closure(body: Expression, captures: Vec<String>) -> Expression {
        Expression::new(ExprKind::Lambda { params: Vec::new(),
            body: LambdaBody::Expr(Box::new(body)), captures, is_async: false })
    }

    #[test]
    fn primitive_cast_and_literal_aggregate_do_not_require_state() {
        assert!(unused(Expression::new(ExprKind::Cast {
            type_name: "int32".into(), expr: Box::new(Expression::int(3)),
        })));
        assert!(unused(Expression::new(ExprKind::Array(vec![ArrayElement {
            key: None, value: Expression::int(3), spread: false, by_ref: false,
        }]))));
    }

    #[test]
    fn pure_nested_closure_does_not_require_outer_state() {
        assert!(unused(closure(closure(Expression::int(7), Vec::new()), Vec::new())));
    }

    #[test]
    fn explicit_state_capture_retains_state_even_without_body_read() {
        assert!(!unused(closure(Expression::int(7), vec!["recover_entry".into()])));
    }

    #[test]
    fn nested_recovery_read_retains_state() {
        assert!(!unused(closure(closure(Expression::ident("recover_entry"), Vec::new()), Vec::new())));
    }

    #[test]
    fn closure_default_expression_retains_state() {
        let parameter = Param { name: "arg".into(), type_hint: None,
            default: Some(Expression::ident("panic_value")), pass_by: PassBy::Value,
            is_rest: false, is_kwargs: false, is_optional: true, is_nullable: false };
        assert!(!unused(Expression::new(ExprKind::Lambda {
            params: vec![parameter], body: LambdaBody::Expr(Box::new(Expression::int(7))),
            captures: Vec::new(), is_async: false,
        })));
    }

    #[test]
    fn unknown_address_form_remains_conservative() {
        assert!(!unused(Expression::new(ExprKind::RefOf(Box::new(PlaceExpr::Ident("other".into()))))));
    }

    #[test]
    fn aggregate_key_and_final_return_reads_retain_state() {
        assert!(!unused(Expression::new(ExprKind::Array(vec![ArrayElement {
            key: Some(Expression::ident("panic_value")), value: Expression::int(3),
            spread: false, by_ref: false,
        }]))));
        assert!(!is_unused(&[], Some(&Expression::ident("recover_entry")), &["recover_entry"]));
    }
}
