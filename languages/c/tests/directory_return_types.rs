use vybe_ast::{BinOp, BitLane, ExprKind, IntOp};

#[test]
fn directory_close_status_has_static_integer_equality_and_single_call() {
    let module = vybe_language_c::parse(include_str!(
        "../../../tests/c/doom_startup/directory_close_static_comparison.c"
    ))
    .unwrap();
    let mut comparisons = Vec::new();
    let mut closes = 0;
    let mut arguments = 0;
    for statement in &module.body {
        statement.walk_exprs(&mut |value| {
            if let ExprKind::Call { callee, .. } = &value.kind {
                if let ExprKind::Ident(name) = &callee.kind {
                    closes += usize::from(name == "closedir");
                    arguments += usize::from(name == "next_directory");
                }
            }
            if let ExprKind::Binary {
                op: BinOp::Integer(operation, lane),
                left,
                ..
            } = &value.kind
                && matches!(operation, IntOp::Eq | IntOp::Ne)
            {
                let (operand, cast_type) = if let ExprKind::Cast { expr, type_name } = &left.kind {
                    (expr.as_ref(), Some(type_name.as_str()))
                } else {
                    (left.as_ref(), None)
                };
                if let ExprKind::Call { callee, args, .. } = &operand.kind
                    && matches!(&callee.kind, ExprKind::Ident(name) if name == "closedir")
                {
                    if let Some(type_name) = cast_type {
                        assert!(
                            matches!(
                                type_name.trim(),
                                "int" | "signed int" | "integer" | "int32" | "int32_t"
                            ),
                            "closedir must retain its signed int result, not {type_name}"
                        );
                    }
                    assert_eq!(*lane, BitLane::W32);
                    assert_eq!(args.len(), 1);
                    assert!(
                        matches!(&args[0].value.kind, ExprKind::Call { callee, args, .. }
                        if matches!(&callee.kind, ExprKind::Ident(name) if name == "next_directory")
                            && args.len() == 1)
                    );
                    comparisons.push(*operation);
                }
            }
        });
    }
    assert_eq!(comparisons, vec![IntOp::Eq, IntOp::Ne]);
    assert_eq!((closes, arguments), (2, 2));
}
