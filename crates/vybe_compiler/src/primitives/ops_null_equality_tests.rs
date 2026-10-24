use super::*;

#[test]
fn literal_null_comparisons_retain_one_local_on_either_side() {
    for null_on_left in [false, true] {
        for strict in [false, true] {
            for other_is_null in [false, true] {
                let mut code = Chunk::new("null_equality_transport");
                if null_on_left {
                    code.emit_ref_null(0x6f, 1); // externref
                }
                if other_is_null {
                    code.emit_ref_null(0x6f, 1);
                } else {
                    code.emit_f64_const(f64::NAN, 1);
                }
                if !null_on_left {
                    code.emit_ref_null(0x6f, 1);
                }
                if strict {
                    emit_js_strict_eq_null_operand(&mut code, null_on_left, 1);
                } else {
                    emit_dyn_eq_null_operand(&mut code, null_on_left, 1);
                }
                assert_eq!(code.local_count, 1);
                code.emit_op(Op::RETURN, 1);
                let mut vm = vybe_runtime::VM::new();
                crate::primitives::platforms::register_platforms_all(&mut vm);
                crate::dynamic::register_dynamic_runtime_imports(&mut vm);
                match vm.run(vec![code]) {
                    Ok(Value::I32(value)) => assert_eq!(value != 0, other_is_null),
                    other => panic!("expected raw null equality result, got {other:?}"),
                }
            }
        }
    }
}
