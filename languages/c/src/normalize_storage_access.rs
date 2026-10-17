//! Access facts survive C normalization without changing emitted bindings.

use vybe_ast::{ExprKind, LambdaBody, Statement, StmtKind, TypeBinding};

pub fn finish(body: &mut [Statement]) {
    for statement in body {
        match &mut statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    if declaration.type_hint.as_ref().is_some_and(|hint| {
                        hint.access.is_some() && hint.binding == TypeBinding::Descriptive
                    }) {
                        declaration.type_hint = None;
                    }
                }
            }
            StmtKind::FunctionDecl { params, body, .. } => {
                for param in params {
                    if param.type_hint.as_ref().is_some_and(|hint| {
                        hint.access.is_some() && hint.binding == TypeBinding::Descriptive
                    }) {
                        param.type_hint = None;
                    }
                }
                finish(body);
            }
            StmtKind::For {
                init: Some(init), ..
            } => {
                finish(std::slice::from_mut(init.as_mut()));
            }
            _ => {}
        }
        crate::normalize_affine_bytes::children(statement, &mut |body| finish(body));
        statement.walk_exprs_mut(&mut |value| match &mut value.kind {
            ExprKind::FunctionExpr(declaration) => {
                finish(std::slice::from_mut(declaration.as_mut()))
            }
            ExprKind::Lambda { params, body, .. } => {
                for param in params {
                    if param.type_hint.as_ref().is_some_and(|hint| {
                        hint.access.is_some() && hint.binding == TypeBinding::Descriptive
                    }) {
                        param.type_hint = None;
                    }
                }
                if let LambdaBody::Block(body) = body {
                    finish(body);
                }
            }
            _ => {}
        });
    }
}
