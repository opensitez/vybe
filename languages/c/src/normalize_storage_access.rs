//! Access facts survive C normalization without changing emitted bindings.
use vybe_ast::{ExprKind, LambdaBody, Param, Statement, StmtKind, TypeBinding, TypeHint};

fn clear_hint(hint: &mut Option<TypeHint>) {
    if hint.as_ref().is_some_and(|hint| {
        hint.access.is_some() && hint.binding == TypeBinding::Descriptive
    }) {
        *hint = None;
    }
}

fn clear_params(params: &mut [Param]) {
    for param in params {
        clear_hint(&mut param.type_hint);
    }
}

// Walk statement-owned declarations only. Expression-owned declarations are
// handled by the common expression visitor below, not rescanned here.
fn finish_bindings(body: &mut [Statement]) {
    for statement in body {
        match &mut statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    clear_hint(&mut declaration.type_hint);
                }
            }
            StmtKind::FunctionDecl { params, body, .. } => {
                clear_params(params);
                finish_bindings(body);
            }
            StmtKind::For { init: Some(init), .. } => {
                finish_bindings(std::slice::from_mut(init.as_mut()));
            }
            _ => {}
        }
        crate::normalize_affine_bytes::children(statement, &mut |body| finish_bindings(body));
    }
}

pub fn finish(body: &mut [Statement]) {
    finish_bindings(body);
    for statement in body {
        statement.walk_exprs_mut(&mut |value| match &mut value.kind {
            ExprKind::FunctionExpr(declaration) => {
                finish_bindings(std::slice::from_mut(declaration.as_mut()));
            }
            ExprKind::Lambda { params, body, .. } => {
                clear_params(params);
                if let LambdaBody::Block(body) = body {
                    finish_bindings(body);
                }
            }
            _ => {}
        });
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::{BindingPattern, Expression, PassBy, StorageAccess, VarDeclKind, VarDeclarator};

    fn annotated(converting: bool, access: Option<StorageAccess>) -> TypeHint {
        let mut hint = if converting {
            TypeHint::converting("int")
        } else {
            TypeHint::descriptive("int")
        };
        hint.access = access;
        hint
    }

    fn param(hint: TypeHint) -> Param {
        Param { name: "p".into(), type_hint: Some(hint), default: None,
            pass_by: PassBy::Value, is_rest: false, is_kwargs: false,
            is_optional: false, is_nullable: false }
    }

    fn variable(hint: TypeHint, init: Option<Expression>) -> Statement {
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator { pattern: BindingPattern::Ident("v".into()),
                type_hint: Some(hint), init, array_bounds: None, with_events: false }],
            kind: VarDeclKind::Var,
        })
    }

    fn function(params: Vec<Param>, body: Vec<Statement>) -> Statement {
        Statement::new(StmtKind::FunctionDecl { name: "f".into(), params,
            return_type: None, body, modifiers: Default::default(), handles: vec![],
            is_async: false, is_generator: false, is_sub: false })
    }

    #[test]
    fn erases_only_descriptive_access_hints_in_bindings_and_parameters() {
        let hints = [annotated(false, Some(StorageAccess::Ordinary)),
            annotated(false, Some(StorageAccess::Volatile)),
            annotated(true, Some(StorageAccess::Ordinary)), annotated(false, None)];
        let mut body = vec![function(hints.iter().cloned().map(param).collect(),
            hints.iter().cloned().map(|hint| variable(hint, Some(Expression::int(1)))).collect())];
        finish(&mut body);
        let StmtKind::FunctionDecl { params, body, .. } = &body[0].kind else { panic!() };
        for index in 0..hints.len() {
            let keep = index >= 2;
            assert_eq!(params[index].type_hint.is_some(), keep);
            let StmtKind::VarDecl { declarations, .. } = &body[index].kind else { panic!() };
            assert_eq!(declarations[0].type_hint.is_some(), keep);
            if keep {
                assert_eq!(format!("{:?}", params[index].type_hint.as_ref().unwrap()), format!("{:?}", hints[index]));
                assert_eq!(format!("{:?}", declarations[0].type_hint.as_ref().unwrap()), format!("{:?}", hints[index]));
            }
        }
    }

    #[test]
    fn cleans_nested_expression_scopes_and_remains_idempotent() {
        let ordinary = || annotated(false, Some(StorageAccess::Ordinary));
        let lambda = Expression::new(ExprKind::Lambda {
            params: vec![param(ordinary())],
            body: LambdaBody::Block(vec![variable(ordinary(), Some(Expression::int(1)))]),
            is_async: false, captures: vec![],
        });
        let expression = Expression::new(ExprKind::FunctionExpr(Box::new(function(
            vec![param(ordinary())], vec![Statement::new(StmtKind::Expr(lambda))],
        ))));
        let mut body = vec![variable(ordinary(), Some(expression))];
        finish(&mut body);
        let StmtKind::VarDecl { declarations, .. } = &body[0].kind else { panic!() };
        assert!(declarations[0].type_hint.is_none());
        let ExprKind::FunctionExpr(declaration) = &declarations[0].init.as_ref().unwrap().kind else { panic!() };
        let StmtKind::FunctionDecl { params, body: inner, .. } = &declaration.kind else { panic!() };
        assert!(params[0].type_hint.is_none());
        let StmtKind::Expr(lambda) = &inner[0].kind else { panic!() };
        let ExprKind::Lambda { params, body: LambdaBody::Block(inner), .. } = &lambda.kind else { panic!() };
        assert!(params[0].type_hint.is_none());
        let StmtKind::VarDecl { declarations, .. } = &inner[0].kind else { panic!() };
        assert!(declarations[0].type_hint.is_none());
        let first = format!("{body:?}");
        finish(&mut body);
        assert_eq!(first, format!("{body:?}"));
    }
}
