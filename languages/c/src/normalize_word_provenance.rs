//! Retain signed word bits for locals observed only by explicit W32 operations.

use std::collections::{HashMap, HashSet};
use vybe_ast::{
    BinOp, BindingPattern, BitLane, ExprKind, Expression, Literal, Statement, StmtKind,
    StorageAccess, TypeBinding, VarDeclKind,
};

fn raw_word(value: &Expression) -> Option<&Expression> {
    match &value.kind {
        ExprKind::Binary {
            op: BinOp::Integer(op, BitLane::W32),
            ..
        } if !op.is_comparison() => Some(value),
        ExprKind::Call {
            callee,
            args,
            optional: false,
        } if matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_unsigned32_from_bits")
            && args.len() == 1
            && !args[0].by_ref
            && !args[0].spread
            && args[0].name.is_none() =>
        {
            raw_word(&args[0].value)
        }
        _ => None,
    }
}

fn discarded(value: &Expression, permitted: &mut HashSet<*const Expression>) {
    permitted.insert(value as *const Expression);
    if let ExprKind::Sequence(parts) = &value.kind {
        for part in parts {
            discarded(part, permitted);
        }
    }
}

fn declarations(
    body: &mut [Statement],
    counts: &mut HashMap<String, usize>,
    candidates: &mut HashSet<String>,
    permitted: &mut HashSet<*const Expression>,
    nested: &mut bool,
) {
    for statement in body {
        *nested |= matches!(statement.kind, StmtKind::FunctionDecl { .. });
        if let StmtKind::Expr(value) = &statement.kind {
            discarded(value, permitted);
        }
        if let StmtKind::VarDecl { kind, declarations } = &statement.kind {
            for declaration in declarations {
                if let BindingPattern::Ident(name) = &declaration.pattern {
                    *counts.entry(name.clone()).or_default() += 1;
                    if *kind != VarDeclKind::Static
                        && !declaration.with_events
                        && declaration.array_bounds.is_none()
                        && declaration.type_hint.as_ref().is_some_and(|hint| {
                            hint.spelling() == "uint32"
                                && hint.binding == TypeBinding::Checked
                                && hint.access == Some(StorageAccess::Ordinary)
                        })
                        && declaration.init.as_ref().is_none_or(|value| {
                            matches!(value.kind, ExprKind::Lit(Literal::Int(0)))
                                || raw_word(value).is_some()
                        })
                    {
                        candidates.insert(name.clone());
                    }
                }
            }
        }
        if let StmtKind::For {
            init: Some(init), ..
        } = &mut statement.kind
        {
            declarations(
                std::slice::from_mut(init.as_mut()),
                counts,
                candidates,
                permitted,
                nested,
            );
        }
        crate::normalize_affine_bytes::children(statement, &mut |body| {
            declarations(body, counts, candidates, permitted, nested)
        });
    }
}

/// Returns the number of bindings whose value-ABI unsigned widening is removed.
pub fn normalize(body: &mut [Statement]) -> usize {
    let mut changed = 0;
    for statement in body {
        let StmtKind::FunctionDecl { params, body, .. } = &mut statement.kind else {
            continue;
        };
        let mut counts = HashMap::new();
        for param in params {
            *counts.entry(param.name.clone()).or_default() += 1;
        }
        let mut candidates = HashSet::new();
        let mut permitted = HashSet::new();
        let mut nested = false;
        declarations(
            body,
            &mut counts,
            &mut candidates,
            &mut permitted,
            &mut nested,
        );
        if nested {
            continue;
        }
        candidates.retain(|name| counts.get(name) == Some(&1));
        if candidates.is_empty() {
            continue;
        }
        let mut allowed = HashSet::new();
        let mut reads = HashSet::new();
        for statement in body.iter() {
            statement.walk_exprs(&mut |value| match &value.kind {
                ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_) => candidates.clear(),
                ExprKind::Binary {
                    op: BinOp::Integer(_, BitLane::W32),
                    left,
                    right,
                } => {
                    for operand in [left.as_ref(), right.as_ref()] {
                        permitted.insert(operand as *const Expression);
                        if let ExprKind::Ident(name) = &operand.kind {
                            allowed.insert(operand as *const Expression);
                            reads.insert(name.clone());
                        }
                    }
                }
                ExprKind::Assign { target, value } => {
                    if let ExprKind::Ident(name) = &target.kind {
                        allowed.insert(target.as_ref() as *const Expression);
                        if raw_word(value).is_none() {
                            candidates.remove(name);
                        }
                    }
                }
                _ => {}
            });
        }
        for statement in body.iter() {
            statement.walk_exprs(&mut |value| {
                if let ExprKind::Assign { target, .. } = &value.kind
                    && let ExprKind::Ident(name) = &target.kind
                    && !permitted.contains(&(value as *const Expression))
                {
                    candidates.remove(name);
                }
                if let ExprKind::Ident(name) = &value.kind
                    && !allowed.contains(&(value as *const Expression))
                {
                    candidates.remove(name);
                }
            });
        }
        candidates.retain(|name| reads.contains(name));
        if candidates.is_empty() {
            continue;
        }
        changed += candidates.len();
        fn rewrite_declarations(body: &mut [Statement], candidates: &HashSet<String>) {
            for statement in body {
                if let StmtKind::VarDecl { declarations, .. } = &mut statement.kind {
                    for declaration in declarations {
                        if matches!(&declaration.pattern, BindingPattern::Ident(name) if candidates.contains(name))
                        {
                            let hint = declaration.type_hint.as_mut().unwrap();
                            hint.set_spelling("int32");
                            hint.word_value = Some(
                                vybe_ast::word_local_transport::WordInterpretation::Signed,
                            );
                            if let Some(value) = &mut declaration.init
                                && let Some(raw) = raw_word(value)
                            {
                                *value = raw.clone();
                            }
                        }
                    }
                }
                if let StmtKind::For {
                    init: Some(init), ..
                } = &mut statement.kind
                {
                    rewrite_declarations(std::slice::from_mut(init.as_mut()), candidates);
                }
                crate::normalize_affine_bytes::children(statement, &mut |body| {
                    rewrite_declarations(body, candidates)
                });
            }
        }
        rewrite_declarations(body, &candidates);
        for statement in body {
            statement.walk_exprs_mut(&mut |expression| {
                if let ExprKind::Assign { target, value } = &mut expression.kind
                    && matches!(&target.kind, ExprKind::Ident(name) if candidates.contains(name))
                    && let Some(raw) = raw_word(value)
                {
                    **value = raw.clone();
                }
            });
        }
    }
    changed
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::{Argument, IntOp, Modifiers, TypeHint, VarDeclarator};

    fn word(name: &str) -> Expression {
        Expression::new(ExprKind::Binary {
            op: BinOp::Integer(IntOp::ShrU, BitLane::W32),
            left: Box::new(Expression::ident(name)),
            right: Box::new(Expression::int(1)),
        })
    }
    fn exposure(value: Expression) -> Expression {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__c_unsigned32_from_bits")),
            args: vec![Argument::positional(value)],
            optional: false,
        })
    }
    fn write() -> Expression {
        Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident("bits")),
            value: Box::new(exposure(word("input"))),
        })
    }
    fn module(mut operations: Vec<Statement>, access: StorageAccess) -> Vec<Statement> {
        let mut hint = TypeHint::checked("uint32");
        hint.access = Some(access);
        operations.insert(
            0,
            Statement::new(StmtKind::VarDecl {
                kind: VarDeclKind::Let,
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident("bits".into()),
                    type_hint: Some(hint),
                    init: Some(exposure(word("input"))),
                    array_bounds: None,
                    with_events: false,
                }],
            }),
        );
        vec![Statement::new(StmtKind::FunctionDecl {
            name: "unrelated".into(),
            params: vec![],
            return_type: Some("int".into()),
            body: operations,
            modifiers: Modifiers::default(),
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        })]
    }
    #[test]
    fn retains_words_at_checked_local_boundaries_and_is_idempotent() {
        let mut body = module(
            vec![
                Statement::new(StmtKind::Expr(write())),
                Statement::new(StmtKind::Return(Some(word("bits")))),
            ],
            StorageAccess::Ordinary,
        );
        assert_eq!(normalize(&mut body), 1);
        let dump = format!("{body:?}");
        assert!(dump.contains("int32"));
        assert!(!dump.contains("__c_unsigned32_from_bits"));
        assert_eq!(normalize(&mut body), 0);
        assert_eq!(dump, format!("{body:?}"));
    }
    #[test]
    fn rejects_assignment_result_escapes_even_when_identifier_reads_are_word_only() {
        for escape in [
            Statement::new(StmtKind::Return(Some(write()))),
            Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("sink")),
                args: vec![Argument::positional(write())],
                optional: false,
            }))),
        ] {
            let mut body = module(
                vec![escape, Statement::new(StmtKind::Return(Some(word("bits"))))],
                StorageAccess::Ordinary,
            );
            let before = format!("{body:?}");
            assert_eq!(normalize(&mut body), 0);
            assert_eq!(before, format!("{body:?}"));
        }
    }
    #[test]
    fn rejects_exposed_value_volatile_storage_and_wider_lane_reads() {
        for (read, access) in [
            (Expression::ident("bits"), StorageAccess::Ordinary),
            (word("bits"), StorageAccess::Volatile),
            (
                Expression::new(ExprKind::Binary {
                    op: BinOp::Integer(IntOp::Add, BitLane::W64),
                    left: Box::new(Expression::ident("bits")),
                    right: Box::new(Expression::int(1)),
                }),
                StorageAccess::Ordinary,
            ),
        ] {
            let mut body = module(vec![Statement::new(StmtKind::Return(Some(read)))], access);
            let before = format!("{body:?}");
            assert_eq!(normalize(&mut body), 0);
            assert_eq!(before, format!("{body:?}"));
        }
    }
}
