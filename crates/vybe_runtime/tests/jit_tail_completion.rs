use vybe_runtime::native_gc::{CompositeType, RecGroup, SubType, ValueType};
use vybe_runtime::{Chunk, Op, VM, Value};

fn typed(name: &str) -> Chunk {
    let mut chunk = Chunk::new(name);
    chunk.arity = 1;
    chunk.param_count = 1;
    chunk.local_count = 1;
    chunk.result_arity = 1;
    chunk.func_sig = Some(("i32".into(), "i32".into()));
    chunk.native_func_type = RecGroup::new(
        vec![SubType {
            final_: true,
            shared: false,
            supertypes: vec![],
            descriptor: None,
            describes: None,
            composite: CompositeType::Func {
                params: vec![ValueType::Scalar(0x7f)],
                results: vec![ValueType::Scalar(0x7f)],
            },
        }],
        vec![],
    )
    .identity(0);
    chunk
}

fn fixture(op: Op, nested: bool, enabled: bool) -> (VM, Value) {
    let mut root = Chunk::new("root");
    root.emit_i32_const(0, 0);
    root.emit_op_u16(Op::REF_FUNC, 2, 0);
    root.emit(0, 0);
    root.emit_op_u8(Op::TABLE_SET, 0, 0);
    root.emit_op_u16(Op::REF_FUNC, if nested { 3 } else { 1 }, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);

    let mut tail = typed("tail");
    if op != Op::RETURN_CALL_INDIRECT {
        tail.emit_op_u16(Op::REF_FUNC, 2, 0);
        tail.emit(0, 0);
    }
    tail.emit_op_u16(Op::LOCAL_GET, 0, 0);
    if op == Op::RETURN_CALL_INDIRECT {
        tail.emit_i32_const(0, 0);
        tail.emit_op_u8(op, 1, 0);
        tail.emit(0, 0);
        tail.emit(1, 0);
    } else {
        tail.emit_op_u8_u8(op, 1, 1, 0);
    }

    let mut identity = typed("identity");
    identity.emit_op_u16(Op::LOCAL_GET, 0, 0);
    identity.emit_op(Op::RETURN, 0);

    let mut caller = typed("caller");
    caller.emit_op_u16(Op::REF_FUNC, 1, 0);
    caller.emit(0, 0);
    caller.emit_op_u16(Op::LOCAL_GET, 0, 0);
    caller.emit_op_u8_u8(Op::CALL_REF, 1, 1, 0);
    caller.emit_i32_const(1, 0);
    caller.emit_op(Op::I32_ADD, 0);
    caller.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.set_jit_enabled(enabled);
    vm.wasm_tables = vec![vec![Value::Null]];
    let function = vm.run(vec![root, tail, identity, caller]).unwrap();
    (vm, function)
}

fn verify_native_entry(vm: &VM, enabled: bool) {
    if enabled && cfg!(all(target_os = "macos", target_arch = "aarch64")) {
        assert!(
            vm.jit_counters().native_entries > 0,
            "must exercise inline JIT completion"
        );
    } else {
        assert_eq!(vm.jit_counters().native_entries, 0);
    }
}

#[test]
fn inline_tail_result_completes_host_and_callback_entries() {
    for op in [
        Op::RETURN_CALL,
        Op::RETURN_CALL_REF,
        Op::RETURN_CALL_INDIRECT,
    ] {
        for enabled in [false, true] {
            let (mut vm, function) = fixture(op, false, enabled);
            for value in [i32::MIN, -1, 0, 41, i32::MAX] {
                assert_eq!(
                    vm.invoke(&function, &[Value::I32(value)]).unwrap(),
                    Value::I32(value),
                    "{op:?}"
                );
                assert_eq!(
                    vm.invoke_callback(&function, &[Value::I32(value)]),
                    Value::I32(value),
                    "callback {op:?}"
                );
            }
            verify_native_entry(&vm, enabled);
        }
    }
}

#[test]
fn inline_tail_result_resumes_the_interpreted_caller() {
    for op in [
        Op::RETURN_CALL,
        Op::RETURN_CALL_REF,
        Op::RETURN_CALL_INDIRECT,
    ] {
        for enabled in [false, true] {
            let (mut vm, function) = fixture(op, true, enabled);
            assert_eq!(
                vm.invoke(&function, &[Value::I32(41)]).unwrap(),
                Value::I32(42),
                "{op:?}"
            );
            verify_native_entry(&vm, enabled);
        }
    }
}
