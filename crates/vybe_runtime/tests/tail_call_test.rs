//! Tests for tail-call opcodes: RETURN_CALL (0x12), RETURN_CALL_REF (0x15),
//! RETURN_CALL_INDIRECT (0x13).
//!
//! Function references are installed via GlobalInit { ConstExpr::RefFunc } so
//! that the callee is in the VM's func_table before execution starts — the
//! same pattern used by vm_reffunc_callref_test.rs.

use std::sync::Arc;
use vybe_runtime::chunk::*;
use vybe_runtime::*;

// ── RETURN_CALL_REF ───────────────────────────────────────────────────────

#[test]
fn return_call_ref_delivers_callee_result() {
    // Chunk 1: double(x) = x * 2
    let mut double_fn = Chunk::new("double");
    double_fn.arity = 1;
    double_fn.local_count = 1;
    {
        let two = double_fn.add_constant(Value::I32(2));
        double_fn.emit_op_u16(opcode::Op::LOCAL_GET, 0, 0);
        double_fn.emit_op_u16(opcode::Op::CONST, two, 0);
        double_fn.emit_op(opcode::Op::I32_MUL, 0);
        double_fn.emit_op(opcode::Op::RETURN, 0);
    }

    // Chunk 0: main — tail-calls double(21) via RETURN_CALL_REF
    let mut main = Chunk::new("<main>");
    main.local_count = 1;
    main.global_inits.push(GlobalInit {
        name: "__double".to_string(),
        init: ConstExpr::RefFunc(1),
    });
    {
        let fn_name = main.add_constant(Value::String(Arc::from("__double")));
        let arg = main.add_constant(Value::I32(21));
        main.emit_op_u16(opcode::Op::GLOBAL_GET, fn_name, 0); // push func ref
        main.emit_op_u16(opcode::Op::CONST, arg, 0); // push arg 21
        main.emit_op_u8(opcode::Op::RETURN_CALL_REF, 1, 0); // tail-call, argc=1
    }

    let r = VM::new().run(vec![main, double_fn]).expect("run failed");
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn return_call_ref_non_function_traps() {
    let mut main = Chunk::new("<main>");
    let not_func = main.add_constant(Value::I32(123));
    main.emit_op_u16(opcode::Op::CONST, not_func, 0);
    main.emit_op_u8(opcode::Op::RETURN_CALL_REF, 0, 0);

    let err = VM::new().run(vec![main]).unwrap_err().to_string();
    assert!(err.contains("call") || err.contains("function"));
}

// ── RETURN_CALL ───────────────────────────────────────────────────────────

#[test]
fn return_call_delivers_callee_result() {
    // Chunk 1: add_one(x) = x + 1
    let mut add_one = Chunk::new("add_one");
    add_one.arity = 1;
    add_one.local_count = 1;
    {
        let one = add_one.add_constant(Value::I32(1));
        add_one.emit_op_u16(opcode::Op::LOCAL_GET, 0, 0);
        add_one.emit_op_u16(opcode::Op::CONST, one, 0);
        add_one.emit_op(opcode::Op::I32_ADD, 0);
        add_one.emit_op(opcode::Op::RETURN, 0);
    }

    // Chunk 0: push func_ref for add_one, push arg 41, RETURN_CALL
    let mut main = Chunk::new("<main>");
    main.local_count = 1;
    main.global_inits.push(GlobalInit {
        name: "__add_one".to_string(),
        init: ConstExpr::RefFunc(1),
    });
    {
        let fn_name = main.add_constant(Value::String(Arc::from("__add_one")));
        let arg = main.add_constant(Value::I32(41));
        main.emit_op_u16(opcode::Op::GLOBAL_GET, fn_name, 0);
        main.emit_op_u16(opcode::Op::CONST, arg, 0);
        main.emit_op_u8(opcode::Op::RETURN_CALL, 1, 0);
    }

    let r = VM::new().run(vec![main, add_one]).expect("run failed");
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn return_call_non_function_traps() {
    let mut main = Chunk::new("<main>");
    main.local_count = 1;
    let not_func = main.add_constant(Value::I32(7));
    main.emit_op_u16(opcode::Op::CONST, not_func, 0);
    main.emit_op_u8(opcode::Op::RETURN_CALL, 0, 0);

    let err = VM::new().run(vec![main]).unwrap_err().to_string();
    assert!(err.contains("call") || err.contains("function"));
}

// ── RETURN_CALL_INDIRECT ──────────────────────────────────────────────────

#[test]
fn return_call_indirect_via_function_table() {
    // Chunk 1: triple(x) = x * 3
    let mut triple_fn = Chunk::new("triple");
    triple_fn.arity = 1;
    // `param_count`/`result_arity` are the WASM type *shape* the
    // `(return_)call_indirect` runtime check compares against the call site.
    // `arity` alone is not enough — it may include an implicit receiver.
    triple_fn.param_count = 1;
    triple_fn.result_arity = 1;
    triple_fn.local_count = 1;
    {
        let three = triple_fn.add_constant(Value::I32(3));
        triple_fn.emit_op_u16(opcode::Op::LOCAL_GET, 0, 0);
        triple_fn.emit_op_u16(opcode::Op::CONST, three, 0);
        triple_fn.emit_op(opcode::Op::I32_MUL, 0);
        triple_fn.emit_op(opcode::Op::RETURN, 0);
    }

    // Chunk 0: use REF_FUNC opcode (which registers in func_table) to get the
    // table index, then RETURN_CALL_INDIRECT.
    //
    // REF_FUNC operand format: [U16 chunk_idx][U8 upvalue_count][...upvalues]
    // With 0 upvalues: emit_op_u16 + emit(0).
    let mut main = Chunk::new("<main>");
    main.local_count = 0;
    {
        let arg = main.add_constant(Value::I32(14));

        // Populate WASM table 0 slot 0 with the funcref, exactly as the spec
        // value model prescribes: `ref.func` yields a value, `table.set` stores
        // it, `return_call_indirect` dispatches through the table. The
        // `__table_idx` on the func object indexes `func_table`, a *different*
        // space from `wasm_tables` — `(return_)call_indirect` resolves only
        // against `wasm_tables`.
        let zero = main.add_constant(Value::I32(0));
        main.emit_op_u16(opcode::Op::CONST, zero, 0); // table slot
        main.emit_op_u16(opcode::Op::REF_FUNC, 1, 0);
        main.emit(0u8, 0); // upvalue_count = 0
        main.emit_op_u8(opcode::Op::TABLE_SET, 0, 0);

        // Spec `return_call_indirect`: `[args… i32]` — the table index is on
        // TOP of the stack, above the args. Push the argument first, then the
        // table index.
        main.emit_op_u16(opcode::Op::CONST, arg, 0);
        main.emit_op_u16(opcode::Op::CONST, zero, 0);

        // Stack: [arg_14, table_idx_f64] — table index on top (spec).
        // `return_call_indirect` is U8_U8_U8: argc, tableidx, expected_results.
        // All three operand bytes are required — the VM reads them in order and
        // compares `expected_results` against the callee's `result_arity`.
        main.emit_op_u8(opcode::Op::RETURN_CALL_INDIRECT, 1, 0);
        main.emit(0u8, 0); // tableidx 0
        main.emit(1u8, 0); // expected_results: `triple` returns one value
    }

    let mut vm = VM::new();
    vm.wasm_tables = vec![vec![Value::Null]];
    let r = vm.run(vec![main, triple_fn]).expect("run failed");
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn return_call_indirect_oob_table_index_traps() {
    let mut main = Chunk::new("<main>");
    main.local_count = 1;
    let table_idx = main.add_constant(Value::I32(99));
    main.emit_op_u16(opcode::Op::CONST, table_idx, 0);
    main.emit_op_u8(opcode::Op::RETURN_CALL_INDIRECT, 0, 0);
    main.emit(0u8, 0); // tableidx 0
    main.emit(1u8, 0); // expected_results

    let err = VM::new().run(vec![main]).unwrap_err().to_string();
    assert!(err.contains("call_indirect") || err.contains("table"));
}

// ── Tail-call chain (the core spec purpose: no stack growth) ──────────────

#[test]
fn return_call_chain_does_not_overflow() {
    // count_down(n): tail-recursively calls itself until n == 0, returns 0.
    // Without tail-call elimination this would overflow at depth 10_000.
    let mut countdown = Chunk::new("countdown");
    countdown.arity = 1;
    countdown.local_count = 1;
    {
        let fn_name = countdown.add_constant(Value::String(Arc::from("__countdown")));
        let zero = countdown.add_constant(Value::I32(0));
        let one = countdown.add_constant(Value::I32(1));

        countdown.global_inits.push(GlobalInit {
            name: "__countdown".to_string(),
            init: ConstExpr::RefFunc(0),
        });

        // if n == 0, return 0
        countdown.emit_op_u16(opcode::Op::LOCAL_GET, 0, 0);
        countdown.emit_op_u16(opcode::Op::CONST, zero, 0);
        countdown.emit_op(opcode::Op::I32_EQ, 0);
        countdown.emit_op(opcode::Op::IF, 0);
        countdown.emit(0x40, 0); // block type void
        countdown.emit_op_u16(opcode::Op::CONST, zero, 0);
        countdown.emit_op(opcode::Op::RETURN, 0);
        countdown.emit_op(opcode::Op::END, 0);

        // else: tail-call countdown(n-1)
        countdown.emit_op_u16(opcode::Op::GLOBAL_GET, fn_name, 0);
        countdown.emit_op_u16(opcode::Op::LOCAL_GET, 0, 0);
        countdown.emit_op_u16(opcode::Op::CONST, one, 0);
        countdown.emit_op(opcode::Op::I32_SUB, 0);
        countdown.emit_op_u8(opcode::Op::RETURN_CALL_REF, 1, 0);
    }

    let mut main = Chunk::new("<main>");
    main.local_count = 0;
    main.global_inits.push(GlobalInit {
        name: "__countdown".to_string(),
        init: ConstExpr::RefFunc(1),
    });
    {
        let fn_name = main.add_constant(Value::String(Arc::from("__countdown")));
        let n = main.add_constant(Value::I32(10_000));
        main.emit_op_u16(opcode::Op::GLOBAL_GET, fn_name, 0);
        main.emit_op_u16(opcode::Op::CONST, n, 0);
        main.emit_op_u8(opcode::Op::RETURN_CALL_REF, 1, 0);
    }

    let r = VM::new()
        .run(vec![main, countdown])
        .expect("should not stack overflow");
    assert_eq!(r.as_i32(), 0);
}
