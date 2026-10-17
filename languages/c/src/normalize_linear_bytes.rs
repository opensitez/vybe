//! Preserve the byte range already guaranteed by unsigned linear loads.

use vybe_ast::{BinOp, BitLane, ExprKind, Expression, IntOp, Literal, Statement};

fn byte_value(value: &Expression) -> bool {
    match &value.kind {
        ExprKind::Lit(Literal::Int(value)) => (0..=255).contains(value),
        ExprKind::Binary {
            op: BinOp::Integer(IntOp::And, BitLane::W32),
            left,
            right,
        } => mask(left) || mask(right),
        ExprKind::Call {
            callee,
            args,
            optional: false,
        } if matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_ptr_i32_load8_u") => {
            args.len() == 1
                && args
                    .iter()
                    .all(|arg| !arg.by_ref && !arg.spread && arg.name.is_none())
        }
        ExprKind::Cast { type_name, .. } => type_name == "uint8",
        ExprKind::Sequence(values) => values.last().is_some_and(byte_value),
        ExprKind::Ternary { then, else_, .. } => byte_value(then) && byte_value(else_),
        _ => false,
    }
}

fn mask(value: &Expression) -> bool {
    match &value.kind {
        ExprKind::Lit(Literal::Int(255)) => true,
        ExprKind::Cast { expr, type_name } if matches!(type_name.as_str(), "int32" | "uint32") => {
            mask(expr)
        }
        _ => false,
    }
}

fn narrowed_byte(value: &Expression) -> Expression {
    if byte_value(value) {
        value.clone()
    } else {
        Expression::new(ExprKind::Cast {
            expr: Box::new(value.clone()),
            type_name: "uint8".into(),
        })
    }
}

pub fn normalize(body: &mut [Statement]) {
    for statement in body {
        statement.walk_exprs_mut(&mut |value| {
            if let ExprKind::Cast { expr, type_name } = &value.kind
                && type_name == "uint8"
                && let ExprKind::Ternary { cond, then, else_ } = &expr.kind
                && (byte_value(then) || byte_value(else_))
            {
                value.kind = ExprKind::Ternary {
                    cond: cond.clone(),
                    then: Box::new(narrowed_byte(then)),
                    else_: Box::new(narrowed_byte(else_)),
                };
            }
            if let ExprKind::Call {
                callee,
                args,
                optional: false,
            } = &mut value.kind
                && matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_ptr_i32_store8")
                && args.len() == 2
                && args
                    .iter()
                    .all(|arg| !arg.by_ref && !arg.spread && arg.name.is_none())
                && byte_value(&args[1].value)
            {
                callee.kind = ExprKind::Ident("__c_ptr_byte_store".into());
            }
            let retained = match &value.kind {
                ExprKind::Cast { expr, type_name } if type_name == "uint8" && byte_value(expr) => {
                    Some(expr.as_ref())
                }
                ExprKind::Binary {
                    op: BinOp::BitAnd | BinOp::Integer(IntOp::And, BitLane::W32),
                    left,
                    right,
                } if mask(right) && byte_value(left) => Some(left.as_ref()),
                ExprKind::Binary {
                    op: BinOp::BitAnd | BinOp::Integer(IntOp::And, BitLane::W32),
                    left,
                    right,
                } if mask(left) && byte_value(right) => Some(right.as_ref()),
                _ => None,
            };
            if let Some(retained) = retained {
                value.kind = retained.kind.clone();
            }
        });
    }
}

// These C-owned load adapters emit integer loads, not arbitrary managed-array
// reads. Their signed result already survives int32 conversion unchanged.
fn signed_word_value(value: &Expression) -> bool {
    if byte_value(value) {
        return true;
    }
    match &value.kind {
        ExprKind::Lit(Literal::Int(value)) => i32::try_from(*value).is_ok(),
        ExprKind::Binary {
            op: BinOp::Integer(operation, BitLane::W32),
            ..
        } => !operation.is_comparison(),
        ExprKind::Call { callee, args, optional: false }
            if matches!(&callee.kind, ExprKind::Ident(name) if matches!(name.as_str(),
                "__c_ptr_i32_load8_s" | "__c_ptr_i32_load16_s"
                    | "__c_ptr_i32_load16_u" | "__c_ptr_i32_load")) =>
        {
            args.len() == 1
                && args.iter().all(|arg| !arg.by_ref && !arg.spread && arg.name.is_none())
        }
        ExprKind::Sequence(values) => values.last().is_some_and(signed_word_value),
        ExprKind::Ternary { then, else_, .. } =>
            signed_word_value(then) && signed_word_value(else_),
        _ => false,
    }
}

/// Reuse proven word values under the module's explicit AST policy. Unknown
/// managed values, unsigned widening and wide integer values keep their casts.
pub fn normalize_words(body: &mut [Statement], directives: &vybe_ast::Directives) {
    if directives.word_local_transport() != vybe_ast::WordLocalTransport::ProvenW32 {
        return;
    }
    for statement in body {
        statement.walk_exprs_mut(&mut |value| {
            let ExprKind::Cast { expr, type_name } = &mut value.kind else {
                return;
            };
            if type_name != "int32" {
                return;
            }
            if signed_word_value(expr) {
                value.kind = std::mem::replace(&mut expr.kind, ExprKind::Lit(Literal::Int(0)));
                return;
            }
            let ExprKind::Ternary { then, else_, .. } = &mut expr.kind else {
                return;
            };
            let known_then = signed_word_value(then);
            let known_else = signed_word_value(else_);
            if !known_then && !known_else {
                return;
            }
            // Only the unproven branch needs the original cast. Keep the
            // condition and each operand evaluated once, without cloning them.
            let unknown = if !known_then { then } else { else_ };
            let original = std::mem::replace(&mut unknown.kind, ExprKind::Lit(Literal::Int(0)));
            unknown.kind = ExprKind::Cast {
                expr: Box::new(Expression::with_span(original, unknown.span)),
                type_name: std::mem::take(type_name),
            };
            value.kind = std::mem::replace(&mut expr.kind, ExprKind::Lit(Literal::Int(0)));
        });
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::{Argument, StmtKind};

    fn load(name: &str) -> Expression {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(name)),
            args: vec![Argument::positional(Expression::ident("address"))],
            optional: false,
        })
    }

    fn narrow(value: Expression) -> Expression {
        Expression::new(ExprKind::Cast {
            expr: Box::new(value),
            type_name: "uint8".into(),
        })
    }

    fn normalized(value: Expression) -> Expression {
        let mut body = vec![Statement::new(StmtKind::Expr(value))];
        normalize(&mut body);
        let StmtKind::Expr(value) = body.remove(0).kind else {
            unreachable!()
        };
        value
    }

    #[test]
    fn removes_unsigned_byte_narrowing_but_not_signed_or_wider_loads() {
        for name in ["__c_ptr_i32_load8_s", "__c_ptr_i32_load16_u", "arbitrary"] {
            assert!(matches!(
                normalized(narrow(load(name))).kind,
                ExprKind::Cast { .. }
            ));
        }
        for op in [BinOp::BitAnd, BinOp::Integer(IntOp::And, BitLane::W32)] {
            for reverse in [false, true] {
                let byte = load("__c_ptr_i32_load8_u");
                let (left, right) = if reverse {
                    (Expression::int(255), byte)
                } else {
                    (byte, Expression::int(255))
                };
                let value = Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(left),
                    right: Box::new(right),
                });
                assert!(matches!(normalized(value).kind, ExprKind::Call { .. }));
            }
        }
        assert!(matches!(
            normalized(narrow(load("__c_ptr_i32_load8_u"))).kind,
            ExprKind::Call { .. }
        ));
    }

    #[test]
    fn retains_sequence_effects_and_requires_both_branches() {
        let sequence = Expression::new(ExprKind::Sequence(vec![
            load("effect"),
            load("__c_ptr_i32_load8_u"),
        ]));
        let result = normalized(narrow(sequence));
        let ExprKind::Sequence(values) = result.kind else {
            panic!("sequence lost")
        };
        assert_eq!(values.len(), 2);
        for other in [load("__c_ptr_i32_load8_s"), Expression::int(256)] {
            let branches = Expression::new(ExprKind::Ternary {
                cond: Box::new(load("condition")),
                then: Box::new(load("__c_ptr_i32_load8_u")),
                else_: Box::new(other),
            });
            let ExprKind::Ternary { cond, then, else_ } = normalized(narrow(branches)).kind else {
                panic!("branch selection lost")
            };
            assert!(matches!(&cond.kind, ExprKind::Call { callee, .. }
                if matches!(&callee.kind, ExprKind::Ident(name) if name == "condition")));
            assert!(matches!(then.kind, ExprKind::Call { .. }));
            assert!(matches!(else_.kind, ExprKind::Cast { .. }));
        }
        let unknown = Expression::new(ExprKind::Ternary {
            cond: Box::new(load("condition")),
            then: Box::new(load("__c_ptr_i32_load8_s")),
            else_: Box::new(load("__c_ptr_i32_load16_u")),
        });
        assert!(matches!(
            normalized(narrow(unknown)).kind,
            ExprKind::Cast { .. }
        ));
    }

    #[test]
    fn selects_byte_store_only_when_the_value_range_is_proven() {
        for (value, proven) in [
            (load("__c_ptr_i32_load8_u"), true),
            (narrow(load("effect")), true),
            (load("__c_ptr_i32_load8_s"), false),
            (load("__c_ptr_i32_load16_u"), false),
            (Expression::ident("unknown"), false),
            (Expression::int(0), true),
            (Expression::int(255), true),
            (Expression::int(-1), false),
            (Expression::int(256), false),
        ] {
            let store = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__c_ptr_i32_store8")),
                args: vec![
                    Argument::positional(load("address_effect")),
                    Argument::positional(value),
                ],
                optional: false,
            });
            let ExprKind::Call { callee, args, .. } = normalized(store).kind else {
                panic!("store lost")
            };
            let expected = if proven {
                "__c_ptr_byte_store"
            } else {
                "__c_ptr_i32_store8"
            };
            assert!(matches!(&callee.kind, ExprKind::Ident(name) if name == expected));
            assert_eq!(args.len(), 2);
            assert!(matches!(&args[0].value.kind, ExprKind::Call { callee, .. }
                if matches!(&callee.kind, ExprKind::Ident(name) if name == "address_effect")));
        }
    }

    #[test]
    fn masked_word_is_a_byte_without_discarding_operand_effects() {
        for reverse in [false, true] {
            let effect = load("effect");
            let (left, right) = if reverse {
                (Expression::int(255), effect)
            } else {
                (effect, Expression::int(255))
            };
            let masked = Expression::new(ExprKind::Binary {
                op: BinOp::Integer(IntOp::And, BitLane::W32),
                left: Box::new(left),
                right: Box::new(right),
            });
            let ExprKind::Binary { op, left, right } = normalized(narrow(masked)).kind else {
                panic!("mask or operand effect lost")
            };
            assert!(matches!(op, BinOp::Integer(IntOp::And, BitLane::W32)));
            let effect = if reverse { right } else { left };
            assert!(matches!(&effect.kind, ExprKind::Call { callee, .. }
                if matches!(&callee.kind, ExprKind::Ident(name) if name == "effect")));
        }
        let wide = Expression::new(ExprKind::Binary {
            op: BinOp::Integer(IntOp::And, BitLane::W64),
            left: Box::new(load("effect")),
            right: Box::new(Expression::int(255)),
        });
        assert!(matches!(
            normalized(narrow(wide)).kind,
            ExprKind::Cast { .. }
        ));
        for (type_name, proven) in [("uint32", true), ("int32", true), ("int8", false)] {
            let masked = Expression::new(ExprKind::Binary {
                op: BinOp::Integer(IntOp::And, BitLane::W32),
                left: Box::new(load("effect")),
                right: Box::new(Expression::new(ExprKind::Cast {
                    expr: Box::new(Expression::int(255)),
                    type_name: type_name.into(),
                })),
            });
            assert_eq!(byte_value(&masked), proven);
            assert_eq!(
                matches!(normalized(narrow(masked)).kind, ExprKind::Binary { .. }),
                proven
            );
        }
    }

    fn normalized_word(value: Expression, enabled: bool) -> Expression {
        let mut body = vec![Statement::new(StmtKind::Expr(Expression::new(ExprKind::Cast {
            expr: Box::new(value), type_name: "int32".into(),
        })))];
        normalize_words(&mut body, &vybe_ast::Directives {
            word_local_transport: Some(if enabled {
                vybe_ast::WordLocalTransport::ProvenW32
            } else {
                vybe_ast::WordLocalTransport::Boxed
            }),
            ..Default::default()
        });
        let StmtKind::Expr(value) = body.remove(0).kind else { unreachable!() };
        value
    }

    #[test]
    fn word_load_cast_reuse_requires_policy_and_exact_adapter_contract() {
        for name in ["__c_ptr_i32_load8_u", "__c_ptr_i32_load8_s",
            "__c_ptr_i32_load16_u", "__c_ptr_i32_load16_s", "__c_ptr_i32_load"]
        {
            assert!(matches!(normalized_word(load(name), true).kind, ExprKind::Call { .. }));
            assert!(matches!(normalized_word(load(name), false).kind, ExprKind::Cast { .. }));
        }
        for name in ["arbitrary", "__c_array_get", "__c_ptr_i64_load", "__c_unsigned32_from_bits"] {
            assert!(matches!(normalized_word(load(name), true).kind, ExprKind::Cast { .. }));
        }
        for spread in [false, true] {
            let mut value = load("__c_ptr_i32_load");
            let ExprKind::Call { args, .. } = &mut value.kind else { unreachable!() };
            args[0].by_ref = !spread;
            args[0].spread = spread;
            assert!(matches!(normalized_word(value, true).kind, ExprKind::Cast { .. }));
        }
    }

    #[test]
    fn word_cast_reuse_keeps_unsigned_wide_and_boolean_boundaries() {
        for value in [i64::from(i32::MIN), -1, 0, i64::from(i32::MAX)] {
            assert!(matches!(normalized_word(Expression::int(value), true).kind, ExprKind::Lit(_)));
        }
        for value in [i64::from(i32::MIN) - 1, i64::from(i32::MAX) + 1] {
            assert!(matches!(normalized_word(Expression::int(value), true).kind, ExprKind::Cast { .. }));
        }
        for (operation, lane, retained) in [
            (IntOp::Add, BitLane::W32, false),
            (IntOp::Eq, BitLane::W32, true),
            (IntOp::Add, BitLane::W64, true),
        ] {
            let value = Expression::new(ExprKind::Binary {
                op: BinOp::Integer(operation, lane),
                left: Box::new(load("effect")), right: Box::new(Expression::int(1)),
            });
            assert_eq!(matches!(normalized_word(value, true).kind, ExprKind::Cast { .. }), retained);
        }
    }

    #[test]
    fn mixed_word_load_keeps_managed_cast_and_once_only_effects() {
        let value = Expression::new(ExprKind::Ternary {
            cond: Box::new(load("condition")),
            then: Box::new(load("__c_ptr_i32_load")),
            else_: Box::new(load("managed_read")),
        });
        let ExprKind::Ternary { cond, then, else_ } = normalized_word(value, true).kind else {
            panic!("branch selection lost")
        };
        assert!(matches!(&cond.kind, ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "condition")));
        assert!(matches!(then.kind, ExprKind::Call { .. }));
        assert!(matches!(&else_.kind, ExprKind::Cast { expr, .. }
            if matches!(&expr.kind, ExprKind::Call { callee, .. }
                if matches!(&callee.kind, ExprKind::Ident(name) if name == "managed_read"))));
        let value = Expression::new(ExprKind::Sequence(vec![load("effect"), load("__c_ptr_i32_load")]));
        let ExprKind::Sequence(values) = normalized_word(value, true).kind else {
            panic!("sequence effects lost")
        };
        assert_eq!(values.len(), 2);
    }
}
