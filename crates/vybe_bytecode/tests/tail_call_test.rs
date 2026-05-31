//! Tests for tail-call opcodes: RETURN_CALL (0x12), RETURN_CALL_REF (0x15),
//! RETURN_CALL_INDIRECT (0x13).
//!
//! Function references are installed via GlobalInit { ConstExpr::RefFunc } so
//! that the callee is in the VM's func_table before execution starts — the
//! same pattern used by vm_reffunc_callref_test.rs.

use vybe_bytecode::*;
use vybe_bytecode::chunk::*;
use std::sync::Arc;

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
        let arg     = main.add_constant(Value::I32(21));
        main.emit_op_u16(opcode::Op::GLOBAL_GET, fn_name, 0); // push func ref
        main.emit_op_u16(opcode::Op::CONST, arg, 0);           // push arg 21
        main.emit_op_u8(opcode::Op::RETURN_CALL_REF, 1, 0);   // tail-call, argc=1
    }

    let r = VM::new().run(vec![main, double_fn]).expect("run failed");
    assert_eq!(r.as_i32(), 42);
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
        let arg     = main.add_constant(Value::I32(41));
        main.emit_op_u16(opcode::Op::GLOBAL_GET, fn_name, 0);
        main.emit_op_u16(opcode::Op::CONST, arg, 0);
        main.emit_op_u8(opcode::Op::RETURN_CALL, 1, 0);
    }

    let r = VM::new().run(vec![main, add_one]).expect("run failed");
    assert_eq!(r.as_i32(), 42);
}

// ── RETURN_CALL_INDIRECT ──────────────────────────────────────────────────

#[test]
fn return_call_indirect_via_function_table() {
    // Chunk 1: triple(x) = x * 3
    let mut triple_fn = Chunk::new("triple");
    triple_fn.arity = 1;
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
        let tidx_key = main.add_constant(Value::String(Arc::from("__table_idx")));
        let arg      = main.add_constant(Value::I32(14));

        // REF_FUNC 1 (triple_fn) with 0 upvalues → pushes func object, registers in func_table
        main.emit_op_u16(opcode::Op::REF_FUNC, 1, 0);
        main.emit(0u8, 0); // upvalue_count = 0

        // STRUCT_GET __table_idx → pops func object, pushes table index (F64)
        main.emit_op_u16(opcode::Op::STRUCT_GET, tidx_key, 0);

        // Push argument
        main.emit_op_u16(opcode::Op::CONST, arg, 0);

        // Stack: [table_idx_f64, arg_14]
        main.emit_op_u8(opcode::Op::RETURN_CALL_INDIRECT, 1, 0);
    }

    let r = VM::new().run(vec![main, triple_fn]).expect("run failed");
    assert_eq!(r.as_i32(), 42);
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
        let zero    = countdown.add_constant(Value::I32(0));
        let one     = countdown.add_constant(Value::I32(1));

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
        let n       = main.add_constant(Value::I32(10_000));
        main.emit_op_u16(opcode::Op::GLOBAL_GET, fn_name, 0);
        main.emit_op_u16(opcode::Op::CONST, n, 0);
        main.emit_op_u8(opcode::Op::RETURN_CALL_REF, 1, 0);
    }

    let r = VM::new().run(vec![main, countdown]).expect("should not stack overflow");
    assert_eq!(r.as_i32(), 0);
}
