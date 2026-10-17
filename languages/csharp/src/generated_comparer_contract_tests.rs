//! Preconditions for a future generated-body carrier. No reuse is enabled here.
use super::*;
use std::collections::BTreeSet;

fn parts(value: &Expression) -> (&[Param], &Expression) {
    let ExprKind::Lambda {
        params,
        body: LambdaBody::Expr(body),
        captures,
        is_async,
    } = &value.kind
    else {
        panic!("generated comparer must remain a plain expression lambda")
    };
    assert!(captures.is_empty());
    assert!(!is_async);
    (params, body)
}

#[test]
fn ordinal_comparer_body_is_closed_over_two_explicit_value_parameters() {
    let value = build_string_comparer_lambda(false);
    let (params, body) = parts(&value);
    assert_eq!(
        params.iter().map(|p| p.name.as_str()).collect::<Vec<_>>(),
        ["left", "right"]
    );
    for param in params {
        assert!(matches!(param.pass_by, PassBy::Value));
        assert!(param.default.is_none() && param.type_hint.is_none());
        assert!(!param.is_rest && !param.is_kwargs && !param.is_optional);
    }
    let mut names = BTreeSet::new();
    body.walk_exprs(&mut |node| match &node.kind {
        ExprKind::Ident(name) => {
            names.insert(name.clone());
        }
        ExprKind::Binary {
            op: BinOp::Lt | BinOp::Gt,
            ..
        }
        | ExprKind::Ternary { .. }
        | ExprKind::Lit(Literal::Int(-1 | 0 | 1)) => {}
        other => panic!("new context/effect requires revisiting provider contract: {other:?}"),
    });
    assert_eq!(
        names,
        BTreeSet::from(["left".to_string(), "right".to_string()])
    );
}

#[test]
fn ordinal_comparer_implementation_is_stable_but_casefold_provider_is_distinct() {
    let first = build_string_comparer_lambda(false);
    let second = build_string_comparer_lambda(false);
    assert_eq!(format!("{first:?}"), format!("{second:?}"));
    assert_ne!(
        format!("{first:?}"),
        format!("{:?}", build_string_comparer_lambda(true))
    );
    assert_eq!(first.span.start_line, 0);
    assert_eq!(first.span.end_line, 0);
}

#[test]
fn ignore_case_comparer_requires_resolved_helper_authority_and_original_call_count() {
    let value = build_string_comparer_lambda(true);
    let (_, body) = parts(&value);
    let mut calls = 0;
    let mut dependencies = BTreeSet::new();
    body.walk_exprs(&mut |node| {
        if let ExprKind::Call {
            callee,
            args,
            optional,
        } = &node.kind
        {
            let ExprKind::Ident(name) = &callee.kind else {
                panic!("new receiver dependency")
            };
            dependencies.insert(name.clone());
            assert_eq!(args.len(), 1);
            assert!(!optional);
            calls += 1;
        }
    });
    // Lt and Gt each evaluate both conversions. Precomputing the conversions
    // once would change a shadowed/helper callback's evaluation count.
    assert_eq!(calls, 4);
    assert_eq!(
        dependencies,
        BTreeSet::from(["__csharp_str_to_lower".to_string()])
    );
}

#[test]
fn capture_free_ref_func_is_stable_per_site_and_no_intern_is_not_equivalent() {
    use vybe_compiler::primitives::functions::{create_function_chunk, emit_ref_func};
    use vybe_runtime::{Chunk, ObjectKind, Op, VM, Value};
    let mut caller = Chunk::new("<script>");
    // The original compiler assigns one body index to each generated occurrence.
    emit_ref_func(&mut caller, 1, 0, 0);
    emit_ref_func(&mut caller, 1, 0, 0);
    emit_ref_func(&mut caller, 2, 0, 0);
    emit_ref_func(&mut caller, 1, 0x80, 0);
    emit_ref_func(&mut caller, 1, 0x80, 0);
    caller.emit_array_new_fixed(0, 5, 0);
    caller.emit_op(Op::RETURN, 0);
    let mut body = create_function_chunk("<lambda>", 0);
    body.emit_f64_const(0.0, 0);
    body.emit_op(Op::RETURN, 0);
    let result = VM::new().run(vec![caller, body.clone(), body]).unwrap();
    let Value::Object(result) = result else {
        panic!("array result")
    };
    let result = result.lock().unwrap();
    let ObjectKind::Array(values) = &result.kind else {
        panic!("array result")
    };
    let function = |value: &Value| match value {
        Value::Object(value) => value.clone(),
        _ => panic!("function carrier"),
    };
    assert!(std::sync::Arc::ptr_eq(
        &function(&values[0]),
        &function(&values[1])
    ));
    assert!(!std::sync::Arc::ptr_eq(
        &function(&values[0]),
        &function(&values[2])
    ));
    assert!(!std::sync::Arc::ptr_eq(
        &function(&values[3]),
        &function(&values[4])
    ));
}
