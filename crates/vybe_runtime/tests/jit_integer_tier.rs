use vybe_runtime::native_gc::{CompositeType, RecGroup, SubType, ValueType};
use vybe_runtime::{Chunk, Op, VM, Value};

fn typed(arity: u8) -> Chunk {
    let mut c = Chunk::new("integer");
    c.arity = arity;
    c.param_count = arity;
    c.local_count = arity as u16;
    c.result_arity = 1;
    c.func_sig = Some((vec!["i32"; arity as usize].join(","), "i32".into()));
    c.native_func_type = RecGroup::new(
        vec![SubType {
            final_: true,
            shared: false,
            supertypes: vec![],
            descriptor: None,
            describes: None,
            composite: CompositeType::Func {
                params: vec![ValueType::Scalar(0x7f); arity as usize],
                results: vec![ValueType::Scalar(0x7f)],
            },
        }],
        vec![],
    )
    .identity(0);
    c
}
fn function(vm: &mut VM, c: Chunk) -> Value {
    let mut root = Chunk::new("root");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    vm.run(vec![root, c]).unwrap()
}
fn result(vm: &mut VM, f: &Value, args: &[i32]) -> i32 {
    let args: Vec<_> = args.iter().copied().map(Value::I32).collect();
    let v = vm.invoke(f, &args).unwrap();
    let Value::I32(n) = v else { panic!("{v:?}") };
    n
}
fn supported() -> bool {
    cfg!(all(target_os = "macos", target_arch = "aarch64"))
}

#[test]
fn wrapping_arithmetic_comparisons_and_shift_boundaries() {
    let ops = [
        Op::I32_ADD,
        Op::I32_SUB,
        Op::I32_MUL,
        Op::I32_AND,
        Op::I32_OR,
        Op::I32_XOR,
        Op::I32_SHL,
        Op::I32_SHR_U,
        Op::I32_SHR_S,
        Op::I32_EQ,
        Op::I32_NE,
        Op::I32_LT_S,
        Op::I32_LT_U,
        Op::I32_GT_S,
        Op::I32_GT_U,
        Op::I32_LE_S,
        Op::I32_LE_U,
        Op::I32_GE_S,
        Op::I32_GE_U,
    ];
    let values = [i32::MIN, -65, -1, 0, 1, 31, 32, 65, i32::MAX];
    for op in ops {
        let mut c = typed(2);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_op_u16(Op::LOCAL_GET, 1, 0);
        c.emit_op(op, 0);
        c.emit_op(Op::RETURN, 0);
        let mut vm = VM::new();
        vm.set_jit_enabled(true);
        let f = function(&mut vm, c);
        for a in values {
            for b in values {
                let expected = match op {
                    Op::I32_ADD => a.wrapping_add(b),
                    Op::I32_SUB => a.wrapping_sub(b),
                    Op::I32_MUL => a.wrapping_mul(b),
                    Op::I32_AND => a & b,
                    Op::I32_OR => a | b,
                    Op::I32_XOR => a ^ b,
                    Op::I32_SHL => a.wrapping_shl(b as u32),
                    Op::I32_SHR_U => ((a as u32).wrapping_shr(b as u32)) as i32,
                    Op::I32_SHR_S => a.wrapping_shr(b as u32),
                    Op::I32_EQ => (a == b) as i32,
                    Op::I32_NE => (a != b) as i32,
                    Op::I32_LT_S => (a < b) as i32,
                    Op::I32_LT_U => ((a as u32) < (b as u32)) as i32,
                    Op::I32_GT_S => (a > b) as i32,
                    Op::I32_GT_U => ((a as u32) > (b as u32)) as i32,
                    Op::I32_LE_S => (a <= b) as i32,
                    Op::I32_LE_U => ((a as u32) <= (b as u32)) as i32,
                    Op::I32_GE_S => (a >= b) as i32,
                    Op::I32_GE_U => ((a as u32) >= (b as u32)) as i32,
                    _ => unreachable!(),
                };
                assert_eq!(result(&mut vm, &f, &[a, b]), expected, "{op:?} a={a} b={b}");
            }
        }
        assert_eq!(
            vm.jit_counters().native_entries,
            if supported() { 81 } else { 0 }
        );
    }
}

fn loop_body() -> Chunk {
    let mut c = typed(1);
    c.local_count = 2;
    c.emit_i32_const(0, 0);
    c.emit_op_u16(Op::LOCAL_SET, 1, 0);
    c.emit_block(0);
    c.emit_loop_s(0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_i32_const(42, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_i32_const(1, 0);
    c.emit_op(Op::I32_SUB, 0);
    c.emit_op_u16(Op::LOCAL_SET, 0, 0);
    c.emit_br(0, 0);
    c.emit_op(Op::END, 0);
    c.emit_op(Op::END, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

#[test]
fn ordinary_vm_loop_entry_and_interpreter_match() {
    for enabled in [false, true] {
        let mut vm = VM::new();
        vm.set_jit_enabled(enabled);
        let f = function(&mut vm, loop_body());
        for n in [0, 1, 17, 10000] {
            assert_eq!(result(&mut vm, &f, &[n]), n * 42);
        }
        assert_eq!(
            vm.jit_counters().native_entries,
            if enabled && supported() { 4 } else { 0 }
        );
    }
}

#[test]
fn unsupported_trapping_instruction_stays_on_interpreter() {
    let mut c = typed(2);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::I32_DIV_S, 0);
    c.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    vm.set_jit_enabled(true);
    let f = function(&mut vm, c);
    assert_eq!(result(&mut vm, &f, &[42, 2]), 21);
    assert!(vm.invoke(&f, &[Value::I32(1), Value::I32(0)]).is_err());
    assert_eq!(vm.jit_counters().native_entries, 0);
    assert_eq!(
        vm.jit_counters().compile_attempts,
        if supported() { 1 } else { 0 }
    );
}

#[test]
fn instrumentation_and_snapshot_reset_bypass_and_release_code() {
    let mut vm = VM::new();
    vm.set_jit_enabled(true);
    let snapshot = vm.snapshot();
    let f = function(&mut vm, loop_body());
    assert_eq!(result(&mut vm, &f, &[17]), 714);
    let entries = vm.jit_counters().native_entries;
    vm.record_runtime_perf(true);
    assert_eq!(result(&mut vm, &f, &[17]), 714);
    assert!(vm.take_runtime_perf().unwrap().dispatches > 0);
    assert_eq!(vm.jit_counters().native_entries, entries);
    vm.record_error_context(true);
    assert_eq!(result(&mut vm, &f, &[17]), 714);
    assert_eq!(vm.jit_counters().native_entries, entries);
    vm.record_error_context(false);
    vm.record_types(true);
    assert_eq!(result(&mut vm, &f, &[17]), 714);
    assert_eq!(vm.jit_counters().native_entries, entries);
    vm.record_types(false);
    let (_commands, rx) = std::sync::mpsc::channel();
    let (tx, _events) = std::sync::mpsc::channel();
    vm.attach_debugger(rx, tx, false);
    assert_eq!(result(&mut vm, &f, &[1]), 42);
    assert_eq!(vm.jit_counters().native_entries, entries);
    vm.detach_debugger();
    vm.reset_to(&snapshot);
    assert_eq!(vm.jit_counters().native_entries, 0);
    assert_eq!(vm.jit_counters().code_bytes, 0);
}

#[test]
fn exact_byte_mutation_and_signature_changes_invalidate() {
    let mut state = vybe_runtime::jit::State::default();
    let mut c = typed(0);
    c.emit_i32_const(42, 0);
    c.emit_op(Op::RETURN, 0);
    assert_eq!(
        state.invoke(1, &c, &[]),
        if supported() { Some(42) } else { None }
    );
    c.code[4] = 43;
    assert_eq!(
        state.invoke(1, &c, &[]),
        if supported() { Some(43) } else { None }
    );
    c.native_func_type = None;
    assert_eq!(state.invoke(1, &c, &[]), None);
    if supported() {
        assert_eq!(state.counters.admitted, 2);
    }
}

#[test]
fn bounded_native_backedges_restart_only_pure_functions() {
    let mut state = vybe_runtime::jit::State::default();
    let mut infinite = typed(0);
    infinite.emit_loop_s(0);
    infinite.emit_br(0, 0);
    infinite.emit_op(Op::END, 0);
    infinite.emit_i32_const(42, 0);
    infinite.emit_op(Op::RETURN, 0);
    assert_eq!(state.invoke(1, &infinite, &[]), None);
    assert_eq!(
        state.counters.fuel_fallbacks,
        if supported() { 1 } else { 0 }
    );
    let mut vm = VM::new();
    vm.set_jit_enabled(true);
    let f = function(&mut vm, loop_body());
    assert_eq!(result(&mut vm, &f, &[200000]), 8400000);
    assert_eq!(
        vm.jit_counters().fuel_fallbacks,
        if supported() { 1 } else { 0 }
    );
    let entries = vm.jit_counters().native_entries;
    assert_eq!(result(&mut vm, &f, &[17]), 714);
    assert_eq!(vm.jit_counters().native_entries, entries);
    assert_eq!(vm.jit_counters().code_bytes, 0);
}
