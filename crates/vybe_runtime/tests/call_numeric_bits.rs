use vybe_runtime::{Chunk, Value, VM, opcode::Op};

#[test]
fn call_ref_preserves_high_unsigned_numeric_arguments_and_wasm_wrap_bits() {
    for argc in 1usize..=5 {
        for number in [2147483648.0, 3221225472.0, 4294967295.0] {
            for wrap in [false, true] {
                let mut callee = Chunk::new("high_unsigned_argument");
                callee.arity = argc as u8;
                callee.local_count = argc as u16 + 8;
                callee.emit_op_u16(Op::LOCAL_GET, argc as u16 - 1, 0);
                if wrap {
                    callee.emit_op(Op::I64_TRUNC_F64_S, 0);
                    callee.emit_op(Op::I32_WRAP_I64, 0);
                }
                callee.emit_op(Op::RETURN, 0);
                let mut caller = Chunk::new("call_numeric_bits");
                caller.emit_op_u16(Op::REF_FUNC, 1, 0);
                caller.emit(0, 0);
                for argument in 0..argc - 1 { caller.emit_f64_const(100.0 + argument as f64, 0); }
                caller.emit_f64_const(number, 0);
                caller.emit_op_u8_u8(Op::CALL_REF, argc as u8, 1, 0);
                caller.emit_op(Op::RETURN, 0);
                let actual = VM::new().run(vec![caller, callee]).unwrap();
                if wrap {
                    assert!(matches!(actual, Value::I32(value) if value == (number as i64) as i32),
                        "argc={argc} number={number} actual={actual:?}");
                } else {
                    assert!(matches!(actual, Value::F64(value) if value.to_bits() == number.to_bits()),
                        "argc={argc} number={number} actual={actual:?}");
                }
            }
        }
    }
}
