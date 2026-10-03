//! Lower generated storage-type guards after the walker has resolved C types.

use vybe_ast::{Argument, BinOp, ExprKind, Expression, Literal, Statement, UnaryOp};

pub fn normalize(body: &mut [Statement]) {
    for statement in body {
        statement.walk_exprs_mut(&mut |value| {
            let ExprKind::Binary { op, left, right } = &value.kind else {
                return;
            };
            if !matches!(
                op,
                BinOp::Eq | BinOp::StrictEq | BinOp::NotEq | BinOp::StrictNotEq
            ) {
                return;
            }
            let tag_comparison = match (&left.kind, &right.kind) {
                (ExprKind::Member { field, .. }, ExprKind::Lit(Literal::Str(tag)))
                    if field == "__ref_kind"
                        && matches!(tag.as_str(), "cell" | "carray" | "cstruct") =>
                {
                    Some((left.as_ref().clone(), right.as_ref().clone()))
                }
                (ExprKind::Lit(Literal::Str(tag)), ExprKind::Member { field, .. })
                    if field == "__ref_kind"
                        && matches!(tag.as_str(), "cell" | "carray" | "cstruct") =>
                {
                    Some((right.as_ref().clone(), left.as_ref().clone()))
                }
                _ => None,
            };
            if let Some((actual, expected)) = tag_comparison {
                let helper = if matches!(op, BinOp::NotEq | BinOp::StrictNotEq) {
                    "__c_tag_ne"
                } else {
                    "__c_tag_eq"
                };
                value.kind = ExprKind::Call {
                    callee: Box::new(Expression::ident(helper)),
                    args: vec![Argument::positional(actual), Argument::positional(expected)],
                    optional: false,
                };
                return;
            }
            let (operand, tag) = match (&left.kind, &right.kind) {
                (
                    ExprKind::Unary {
                        op: UnaryOp::Typeof,
                        expr,
                    },
                    ExprKind::Lit(Literal::Str(tag)),
                )
                | (
                    ExprKind::Lit(Literal::Str(tag)),
                    ExprKind::Unary {
                        op: UnaryOp::Typeof,
                        expr,
                    },
                ) => (expr, tag),
                _ => return,
            };
            // `typeof null` is "object", unlike IsType's object predicate.
            // Only normalize tags whose common-AST predicates are equivalent.
            if !matches!(tag.as_str(), "number" | "string" | "boolean" | "undefined") {
                return;
            }
            let test = Expression {
                kind: ExprKind::IsType {
                    expr: operand.clone(),
                    type_name: tag.clone(),
                },
                span: value.span,
            };
            value.kind = if matches!(op, BinOp::NotEq | BinOp::StrictNotEq) {
                ExprKind::Unary {
                    op: UnaryOp::Not,
                    expr: Box::new(test),
                }
            } else {
                test.kind
            };
        });
    }
}
