use super::*;

fn run_numeric_equality(lhs: f64, rhs: f64, literal: Option<i32>) -> bool {
    let mut code = Chunk::new("numeric_equality_transport");
    code.emit_f64_const(lhs, 1);
    if let Some(literal) = literal {
        emit_dyn_eq_i32_literal(&mut code, literal, 1);
    } else {
        code.emit_f64_const(rhs, 1);
        let mut imports = Chunk::new("unused_import_owner");
        emit_dyn_eq_into(&mut imports, &mut code, 1);
    }
    code.emit_op(Op::RETURN, 1);
    let mut vm = vybe_runtime::VM::new();
    crate::primitives::platforms::register_platforms_all(&mut vm);
    crate::dynamic::register_dynamic_runtime_imports(&mut vm);
    match vm.run(vec![code]) {
        Ok(Value::I32(value)) => value != 0,
        other => panic!("expected raw i32 equality result, got {other:?}"),
    }
}

#[test]
fn guarded_numeric_equality_preserves_ieee_results() {
    for (lhs, rhs) in [
        (0.0, -0.0),
        (1.0, 1.0),
        (1.0, 2.0),
        (f64::INFINITY, f64::INFINITY),
        (f64::NEG_INFINITY, f64::INFINITY),
        (f64::NAN, f64::NAN),
        (f64::NAN, 0.0),
        (9_007_199_254_740_992.0, 9_007_199_254_740_992.0),
    ] {
        assert_eq!(run_numeric_equality(lhs, rhs, None), lhs == rhs);
    }
}

#[test]
fn guarded_numeric_literal_equality_preserves_boundaries() {
    for literal in [i32::MIN, -1, 0, 1, i32::MAX] {
        for lhs in [f64::from(literal), f64::from(literal) + 0.5, f64::NAN] {
            assert_eq!(run_numeric_equality(lhs, 0.0, Some(literal)), lhs == f64::from(literal));
        }
    }
    assert!(run_numeric_equality(-0.0, 0.0, Some(0)));
}
