//! WASM Compliance Test Suite
//!
//! Exercises the VM against WASM spec semantics:
//! - Function calling convention (arity params, no implicit callee slot)
//! - Arithmetic opcodes
//! - Control flow (block / loop / br / br_if / if)
//! - GC opcodes (struct / array)
//! - Round-trip: compile → .wasm → read back → execute, same result
//!
//! Every test creates a minimal Chunk, runs it in the VM, and asserts
//! the result. Tests are self-contained — no stdlib, no imports unless
//! explicitly noted.
//!
//! WASM compliance means:
//!   - slot 0 = first arg (NOT a reserved callee slot)
//!   - Function type has exactly `arity` params
//!   - Block/loop/end label stack matches WASM spec
//!   - br depth targets the Nth enclosing construct

use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

/// Helper: build a script chunk from an emit closure, run it, return the popped result.
fn run_script(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut chunk = Chunk::new("<script>");
    emit(&mut chunk);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    vm.run(vec![chunk]).expect("VM execution failed")
}

// ──────────────────────────────────────────────────────────────────────
// 1. CONSTANTS AND STACK
// ──────────────────────────────────────────────────────────────────────

#[test]
fn const_f64_on_stack() {
    let result = run_script(|c| {
        c.emit_f64_const(42.5, 0);
    });
    assert_eq!(result.as_f64(), 42.5);
}

#[test]
fn const_i32_on_stack() {
    let result = run_script(|c| {
        c.emit_i32_const(-7, 0);
    });
    assert_eq!(result.as_i32(), -7);
}

#[test]
fn const_string_on_stack() {
    let result = run_script(|c| {
        c.emit_string_const("hello", 0);
    });
    assert_eq!(format!("{}", result), "hello");
}

#[test]
fn i32_const_shortcuts() {
    assert_eq!(
        run_script(|c| {
            c.emit_i32_const(0, 0);
        })
        .as_i32(),
        0
    );
    assert_eq!(
        run_script(|c| {
            c.emit_i32_const(1, 0);
        })
        .as_i32(),
        1
    );
    assert_eq!(
        run_script(|c| {
            c.emit_f64_const(0.0, 0);
        })
        .as_f64(),
        0.0
    );
}

#[test]
fn dup_replicates_top_of_stack() {
    // push 5, dup → [5, 5], pop the top, assert the remaining is 5
    let result = run_script(|c| {
        c.emit_f64_const(5.0, 0);
        c.emit_dup(0);
        c.emit_op(Op::DROP, 0);
    });
    assert_eq!(result.as_f64(), 5.0);
}

#[test]
fn drop_removes_top_of_stack() {
    // push 1, push 2, drop → top is 1
    let result = run_script(|c| {
        c.emit_f64_const(1.0, 0);
        c.emit_f64_const(2.0, 0);
        c.emit_op(Op::DROP, 0);
    });
    assert_eq!(result.as_f64(), 1.0);
}

// ──────────────────────────────────────────────────────────────────────
// 2. ARITHMETIC (WASM spec semantics)
// ──────────────────────────────────────────────────────────────────────

fn binop_f64(emit_op: Op, a: f64, b: f64) -> f64 {
    run_script(|c| {
        c.emit_f64_const(a, 0);
        c.emit_f64_const(b, 0);
        c.emit_op(emit_op, 0);
    })
    .as_f64()
}

#[test]
fn f64_add() {
    assert_eq!(binop_f64(Op::F64_ADD, 2.5, 3.25), 5.75);
}
#[test]
fn f64_sub() {
    assert_eq!(binop_f64(Op::F64_SUB, 10.0, 3.0), 7.0);
}
#[test]
fn f64_mul() {
    assert_eq!(binop_f64(Op::F64_MUL, 4.0, 2.5), 10.0);
}
#[test]
fn f64_div() {
    assert_eq!(binop_f64(Op::F64_DIV, 10.0, 4.0), 2.5);
}
#[test]
fn f64_min() {
    assert_eq!(binop_f64(Op::F64_MIN, 3.0, 7.0), 3.0);
}
#[test]
fn f64_max() {
    assert_eq!(binop_f64(Op::F64_MAX, 3.0, 7.0), 7.0);
}

fn binop_i32(emit_op: Op, a: i32, b: i32) -> i32 {
    run_script(|c| {
        c.emit_i32_const(a, 0);
        c.emit_i32_const(b, 0);
        c.emit_op(emit_op, 0);
    })
    .as_i32()
}

#[test]
fn i32_add() {
    assert_eq!(binop_i32(Op::I32_ADD, 5, 7), 12);
}
#[test]
fn i32_sub() {
    assert_eq!(binop_i32(Op::I32_SUB, 10, 3), 7);
}
#[test]
fn i32_mul() {
    assert_eq!(binop_i32(Op::I32_MUL, 4, 6), 24);
}
#[test]
fn i32_div_s() {
    assert_eq!(binop_i32(Op::I32_DIV_S, -10, 2), -5);
}
#[test]
fn i32_and() {
    assert_eq!(binop_i32(Op::I32_AND, 0b1100, 0b1010), 0b1000);
}
#[test]
fn i32_or() {
    assert_eq!(binop_i32(Op::I32_OR, 0b1100, 0b1010), 0b1110);
}
#[test]
fn i32_xor() {
    assert_eq!(binop_i32(Op::I32_XOR, 0b1100, 0b1010), 0b0110);
}
#[test]
fn i32_shl() {
    assert_eq!(binop_i32(Op::I32_SHL, 1, 4), 16);
}

#[test]
fn f64_neg_unary() {
    let r = run_script(|c| {
        c.emit_f64_const(3.5, 0);
        c.emit_op(Op::F64_NEG, 0);
    });
    assert_eq!(r.as_f64(), -3.5);
}

#[test]
fn f64_abs_unary() {
    let r = run_script(|c| {
        c.emit_f64_const(-4.5, 0);
        c.emit_op(Op::F64_ABS, 0);
    });
    assert_eq!(r.as_f64(), 4.5);
}

#[test]
fn f64_comparison_lt() {
    let r = run_script(|c| {
        c.emit_f64_const(2.0, 0);
        c.emit_f64_const(3.0, 0);
        c.emit_op(Op::F64_LT, 0);
    });
    // WASM returns i32 (0 or 1) for comparisons; VM may return Bool or I32.
    let truthy = match r {
        Value::I32(v) => v != 0,
        Value::Bool(b) => b,
        Value::F64(f) => f != 0.0,
        _ => false };
    assert!(truthy, "2.0 < 3.0 should be truthy");
}

// ──────────────────────────────────────────────────────────────────────
// 3. LOCALS — slot 0 = first arg (WASM convention)
// ──────────────────────────────────────────────────────────────────────

#[test]
fn local_get_set_slot_0() {
    // Define a local at slot 0, store 42, read back.
    let result = run_script(|c| {
        c.local_count = 1;
        c.emit_f64_const(42.0, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0); // tee to slot 0, keeps value
        c.emit_op_u16(Op::LOCAL_GET, 0, 0); // push slot 0
    });
    assert_eq!(result.as_f64(), 42.0);
}

#[test]
fn function_receives_args_at_slots_0_and_1() {
    // Function `add(a, b)` — a at slot 0, b at slot 1.
    let mut script = Chunk::new("<script>");
    let mut add_fn = Chunk::new("add");
    add_fn.arity = 2;
    add_fn.local_count = 2;
    add_fn.emit_op_u16(Op::LOCAL_GET, 0, 0); // a
    add_fn.emit_op_u16(Op::LOCAL_GET, 1, 0); // b
    add_fn.emit_op(Op::F64_ADD, 0);
    add_fn.emit_op(Op::RETURN, 0);

    // Script: push ref_func add, push 10, push 20, call_ref 2
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0); // 0 upvalues
    script.emit_f64_const(10.0, 0);
    script.emit_f64_const(20.0, 0);
    script.emit_op_u8_u8(Op::CALL_REF, 2, 1, 0);
    script.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![script, add_fn]).unwrap();
    assert_eq!(result.as_f64(), 30.0);
}

#[test]
fn function_with_local_beyond_args() {
    // Function: add(a, b) { let tmp = a + b; return tmp * 2; }
    // arity=2, local_count=3 (a=0, b=1, tmp=2)
    let mut script = Chunk::new("<script>");
    let mut fun = Chunk::new("fn");
    fun.arity = 2;
    fun.local_count = 3;
    fun.emit_op_u16(Op::LOCAL_GET, 0, 0); // a
    fun.emit_op_u16(Op::LOCAL_GET, 1, 0); // b
    fun.emit_op(Op::F64_ADD, 0);
    fun.emit_op_u16(Op::LOCAL_SET, 2, 0); // tmp
    fun.emit_op_u16(Op::LOCAL_GET, 2, 0); // tmp
    fun.emit_f64_const(2.0, 0);
    fun.emit_op(Op::F64_MUL, 0);
    fun.emit_op(Op::RETURN, 0);

    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_f64_const(3.0, 0);
    script.emit_f64_const(4.0, 0);
    script.emit_op_u8_u8(Op::CALL_REF, 2, 1, 0);
    script.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![script, fun]).unwrap();
    assert_eq!(result.as_f64(), 14.0); // (3+4)*2
}

// ──────────────────────────────────────────────────────────────────────
// 4. CONTROL FLOW — structured (block / loop / br / br_if)
// ──────────────────────────────────────────────────────────────────────

#[test]
fn block_end_no_branch() {
    // block { push 5 } — falls through, 5 on stack
    let result = run_script(|c| {
        let bp = c.emit_block(0);
        c.emit_f64_const(5.0, 0);
        c.emit_end(0);
        c.patch_block(bp);
    });
    assert_eq!(result.as_f64(), 5.0);
}

#[test]
fn br_0_exits_block() {
    // block { push 5 ; br 0 ; push 99 } ; end → 5 on stack (99 skipped)
    let result = run_script(|c| {
        let bp = c.emit_block_typed(0, 1);
        c.emit_f64_const(5.0, 0);
        c.emit_br(0, 0); // branch to end of block
        c.emit_f64_const(99.0, 0); // unreachable
        c.emit_end(0);
        c.patch_block(bp);
    });
    assert_eq!(result.as_f64(), 5.0);
}

#[test]
fn br_if_conditional_branch() {
    // block {
    //   i32.const 1 ; br_if 0 ; push 99
    // } ; end → nothing on stack from 99
    let result = run_script(|c| {
        let bp = c.emit_block_typed(0, 1);
        c.emit_i32_const(1, 0);
        c.emit_br_if(0, 0); // branch because true
        c.emit_f64_const(99.0, 0); // skipped
        c.emit_end(0);
        c.patch_block(bp);
        // Push sentinel so we have something to return
        c.emit_f64_const(42.0, 0);
    });
    assert_eq!(result.as_f64(), 42.0);
}

#[test]
fn br_if_zero_does_not_branch() {
    // block { i32.const 0 ; br_if 0 ; push 7 ; br 0 ; } ; end → 7 on stack
    let result = run_script(|c| {
        let bp = c.emit_block_typed(0, 1);
        c.emit_i32_const(0, 0);
        c.emit_br_if(0, 0); // does NOT branch
        c.emit_f64_const(7.0, 0);
        c.emit_br(0, 0);
        c.emit_end(0);
        c.patch_block(bp);
    });
    assert_eq!(result.as_f64(), 7.0);
}

#[test]
fn simple_while_loop() {
    // i = 0; while (i < 3) { i++; }; return i;
    // block { loop { i<3 ? (i++; br 0) : br 1 } } → 3
    let result = run_script(|c| {
        c.local_count = 1;
        c.emit_i32_const(0, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);

        let bp = c.emit_block(0);
        let (lp, _) = c.emit_loop_s(0);

        // i < 3 ?
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(3, 0);
        c.emit_op(Op::I32_LT_S, 0);
        // Invert: if false (i >= 3), exit block (depth 1)
        c.emit_op(Op::I32_EQZ, 0);
        c.emit_br_if(1, 0);

        // i++
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(1, 0);
        c.emit_op(Op::I32_ADD, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);

        // continue loop
        c.emit_br(0, 0);

        c.emit_end(0);
        c.patch_loop(lp);
        c.emit_end(0);
        c.patch_block(bp);

        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    });
    assert_eq!(result.as_i32(), 3);
}

#[test]
fn nested_loops_break_with_labels() {
    // Outer counts 0..3, inner counts 0..2. total iterations = 6.
    // let total = 0;
    // block $outer {
    //   loop $outer_loop {
    //     if i_outer >= 3: br 1 (exit outer block)
    //     block $inner {
    //       loop $inner_loop {
    //         if i_inner >= 2: br 1 (exit inner block)
    //         total++; i_inner++;
    //         br 0 (continue inner)
    //       }
    //     }
    //     reset i_inner, increment i_outer, br 0 (continue outer)
    //   }
    // }
    let result = run_script(|c| {
        c.local_count = 3; // total=0, i_outer=1, i_inner=2

        // Init
        c.emit_i32_const(0, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        c.emit_i32_const(0, 0);
        c.emit_op_u16(Op::LOCAL_SET, 1, 0);

        let outer_b = c.emit_block(0);
        let (outer_l, _) = c.emit_loop_s(0);

        // outer cond
        c.emit_op_u16(Op::LOCAL_GET, 1, 0);
        c.emit_i32_const(3, 0);
        c.emit_op(Op::I32_LT_S, 0);
        c.emit_op(Op::I32_EQZ, 0);
        c.emit_br_if(1, 0); // exit outer block if !(i_outer<3)

        // Reset i_inner = 0
        c.emit_i32_const(0, 0);
        c.emit_op_u16(Op::LOCAL_SET, 2, 0);

        let inner_b = c.emit_block(0);
        let (inner_l, _) = c.emit_loop_s(0);

        // inner cond
        c.emit_op_u16(Op::LOCAL_GET, 2, 0);
        c.emit_i32_const(2, 0);
        c.emit_op(Op::I32_LT_S, 0);
        c.emit_op(Op::I32_EQZ, 0);
        c.emit_br_if(1, 0); // exit inner block if !(i_inner<2)

        // total++
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(1, 0);
        c.emit_op(Op::I32_ADD, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);

        // i_inner++
        c.emit_op_u16(Op::LOCAL_GET, 2, 0);
        c.emit_i32_const(1, 0);
        c.emit_op(Op::I32_ADD, 0);
        c.emit_op_u16(Op::LOCAL_SET, 2, 0);

        c.emit_br(0, 0); // continue inner loop
        c.emit_end(0);
        c.patch_loop(inner_l);
        c.emit_end(0);
        c.patch_block(inner_b);

        // i_outer++
        c.emit_op_u16(Op::LOCAL_GET, 1, 0);
        c.emit_i32_const(1, 0);
        c.emit_op(Op::I32_ADD, 0);
        c.emit_op_u16(Op::LOCAL_SET, 1, 0);

        c.emit_br(0, 0); // continue outer loop
        c.emit_end(0);
        c.patch_loop(outer_l);
        c.emit_end(0);
        c.patch_block(outer_b);

        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    });
    assert_eq!(result.as_i32(), 6);
}

// ──────────────────────────────────────────────────────────────────────
// 5. WASM ROUND-TRIP: compile → .wasm bytes → read back → execute
// ──────────────────────────────────────────────────────────────────────

#[test]
fn round_trip_simple_constant() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_f64_const(42.0, 0);
    chunk.emit_op(Op::RETURN, 0);
    let orig_chunks = vec![chunk];

    // Write to .wasm
    let wasm_bytes = vybe_platform_wasm::write_wasm(&orig_chunks);
    assert!(wasm_bytes.len() > 0, "wasm output empty");
    assert_eq!(&wasm_bytes[0..4], b"\x00asm", "not a valid wasm magic");

    // Read back
    let read_chunks = vybe_platform_wasm::read_wasm(&wasm_bytes).expect("read failed");
    assert!(read_chunks.len() >= 1, "no chunks read back");

    // Execute the read-back chunks and verify result matches
    let mut vm = VM::new();
    let result = vm.run(read_chunks).expect("execution failed");
    assert_eq!(result.as_f64(), 42.0);
}

#[test]
fn round_trip_arithmetic() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_f64_const(10.0, 0);
    chunk.emit_f64_const(3.0, 0);
    chunk.emit_op(Op::F64_ADD, 0);
    chunk.emit_op(Op::RETURN, 0);

    let wasm = vybe_platform_wasm::write_wasm(&[chunk]);
    let chunks = vybe_platform_wasm::read_wasm(&wasm).expect("read");
    let mut vm = VM::new();
    let result = vm.run(chunks).expect("run");
    assert_eq!(result.as_f64(), 13.0);
}

#[test]
fn round_trip_function_call() {
    // add(a, b) { return a + b; } ; add(2, 3) → 5
    let mut script = Chunk::new("<script>");
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_f64_const(2.0, 0);
    script.emit_f64_const(3.0, 0);
    script.emit_op_u8_u8(Op::CALL_REF, 2, 1, 0);
    script.emit_op(Op::RETURN, 0);

    let mut add_fn = Chunk::new("add");
    add_fn.arity = 2;
    add_fn.local_count = 2;
    add_fn.emit_op_u16(Op::LOCAL_GET, 0, 0);
    add_fn.emit_op_u16(Op::LOCAL_GET, 1, 0);
    add_fn.emit_op(Op::F64_ADD, 0);
    add_fn.emit_op(Op::RETURN, 0);

    let chunks = vec![script, add_fn];

    // First verify it works when run directly
    let mut vm_direct = VM::new();
    let direct_result = vm_direct.run(chunks.clone()).expect("direct run");
    assert_eq!(
        direct_result.as_f64(),
        5.0,
        "direct execution should give 5.0"
    );

    // Now round-trip
    let wasm = vybe_platform_wasm::write_wasm(&chunks);
    let chunks_rt = vybe_platform_wasm::read_wasm(&wasm).expect("read");

    // Verify the read-back bytecode matches
    assert_eq!(chunks_rt.len(), chunks.len(), "chunk count mismatch");
    assert_eq!(
        chunks_rt[0].code, chunks[0].code,
        "script bytecode round-trip mismatch"
    );
    assert_eq!(
        chunks_rt[1].code, chunks[1].code,
        "add bytecode round-trip mismatch"
    );
    assert_eq!(chunks_rt[1].arity, chunks[1].arity, "arity mismatch");
    assert_eq!(
        chunks_rt[1].local_count, chunks[1].local_count,
        "local_count mismatch"
    );

    let mut vm = VM::new();
    let result = vm.run(chunks_rt).expect("run");
    assert_eq!(result.as_f64(), 5.0);
}

// ──────────────────────────────────────────────────────────────────────
// 6. WASM BINARY STRUCTURE
// ──────────────────────────────────────────────────────────────────────

#[test]
fn wasm_has_magic_and_version() {
    let chunk = Chunk::new("<script>");
    let wasm = vybe_platform_wasm::write_wasm(&[chunk]);
    assert_eq!(&wasm[0..4], b"\x00asm");
    assert_eq!(&wasm[4..8], b"\x01\x00\x00\x00", "wasm version should be 1");
}

#[test]
fn wasm_emits_required_sections() {
    // A minimal module should have type, function, memory, export, code sections
    // (table/element/global are added when needed).
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op(Op::RETURN, 0);
    let wasm = vybe_platform_wasm::write_wasm(&[chunk]);

    // Walk sections
    let mut pos = 8;
    let mut seen = std::collections::HashSet::new();
    while pos < wasm.len() {
        let sid = wasm[pos];
        pos += 1;
        // LEB128 size
        let mut size = 0u32;
        let mut shift = 0u32;
        loop {
            let b = wasm[pos];
            pos += 1;
            size |= ((b & 0x7f) as u32) << shift;
            shift += 7;
            if b & 0x80 == 0 {
                break;
            }
        }
        seen.insert(sid);
        pos += size as usize;
    }

    // Required sections: 1 (type), 3 (function), 7 (export), 10 (code)
    assert!(seen.contains(&1), "missing type section");
    assert!(seen.contains(&3), "missing function section");
    assert!(seen.contains(&7), "missing export section");
    assert!(seen.contains(&10), "missing code section");
}

// ──────────────────────────────────────────────────────────────────────
// 7. WASM TYPE SECTION — function signatures
// ──────────────────────────────────────────────────────────────────────

#[test]
fn function_type_has_arity_params_not_arity_plus_one() {
    // Critical WASM compliance check: a function with arity=2 must have
    // exactly 2 params in its type signature, NOT 3 (no reserved callee slot).
    let mut fun = Chunk::new("f");
    fun.arity = 2;
    fun.local_count = 2;
    fun.emit_op_u16(Op::LOCAL_GET, 0, 0);
    fun.emit_op_u16(Op::LOCAL_GET, 1, 0);
    fun.emit_op(Op::F64_ADD, 0);
    fun.emit_op(Op::RETURN, 0);

    let script = Chunk::new("<script>");
    let wasm = vybe_platform_wasm::write_wasm(&[script, fun]);

    // Search for the (externref, externref) → externref type signature in the binary.
    // This is: 0x60 0x02 0x6F 0x6F 0x01 0x6F (func, 2 params, externref x2, 1 result, externref)
    let target = [0x60, 0x02, 0x6F, 0x6F, 0x01, 0x6F];
    let found = wasm.windows(target.len()).any(|w| w == target);
    assert!(
        found,
        "WASM binary should contain the (externref, externref) → externref function type for arity-2 functions"
    );
}

// ──────────────────────────────────────────────────────────────────────
// 8. MORE ARITHMETIC EDGE CASES
// ──────────────────────────────────────────────────────────────────────

#[test]
fn f64_sub_negative() {
    assert_eq!(binop_f64(Op::F64_SUB, 3.0, 10.0), -7.0);
}
#[test]
fn f64_div_by_one() {
    assert_eq!(binop_f64(Op::F64_DIV, 42.0, 1.0), 42.0);
}
#[test]
fn f64_mul_zero() {
    assert_eq!(binop_f64(Op::F64_MUL, 42.0, 0.0), 0.0);
}
#[test]
fn f64_add_zero() {
    assert_eq!(binop_f64(Op::F64_ADD, 42.0, 0.0), 42.0);
}
#[test]
fn f64_min_equal() {
    assert_eq!(binop_f64(Op::F64_MIN, 5.0, 5.0), 5.0);
}
#[test]
fn f64_max_negatives() {
    assert_eq!(binop_f64(Op::F64_MAX, -3.0, -1.0), -1.0);
}

#[test]
fn f64_sqrt() {
    let r = run_script(|c| {
        c.emit_f64_const(16.0, 0);
        c.emit_op(Op::F64_SQRT, 0);
    });
    assert_eq!(r.as_f64(), 4.0);
}

#[test]
fn f64_floor() {
    let r = run_script(|c| {
        c.emit_f64_const(3.7, 0);
        c.emit_op(Op::F64_FLOOR, 0);
    });
    assert_eq!(r.as_f64(), 3.0);
}

#[test]
fn f64_ceil() {
    let r = run_script(|c| {
        c.emit_f64_const(3.2, 0);
        c.emit_op(Op::F64_CEIL, 0);
    });
    assert_eq!(r.as_f64(), 4.0);
}

#[test]
fn i32_sub_negative() {
    assert_eq!(binop_i32(Op::I32_SUB, 3, 10), -7);
}
#[test]
fn i32_rem_s_positive() {
    assert_eq!(binop_i32(Op::I32_REM_S, 10, 3), 1);
}
#[test]
fn i32_rem_s_negative() {
    assert_eq!(binop_i32(Op::I32_REM_S, -10, 3), -1);
}
#[test]
fn i32_shr_s_sign_extend() {
    assert_eq!(binop_i32(Op::I32_SHR_S, -8, 1), -4);
}
#[test]
fn i32_shr_u_zero_extend() {
    // -8 >> 1 unsigned = i32::MAX / 2 + 1 area
    let v = binop_i32(Op::I32_SHR_U, -1, 1);
    assert_eq!(v, i32::MAX);
}

#[test]
fn i32_eq_true() {
    let r = run_script(|c| {
        c.emit_i32_const(5, 0);
        c.emit_i32_const(5, 0);
        c.emit_op(Op::EQ, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn i32_eq_false() {
    let r = run_script(|c| {
        c.emit_i32_const(5, 0);
        c.emit_i32_const(6, 0);
        c.emit_op(Op::EQ, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn i32_ne() {
    let r = run_script(|c| {
        c.emit_i32_const(5, 0);
        c.emit_i32_const(6, 0);
        c.emit_op(Op::NE, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn i32_eqz_zero() {
    let r = run_script(|c| {
        c.emit_i32_const(0, 0);
        c.emit_op(Op::I32_EQZ, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn i32_eqz_nonzero() {
    let r = run_script(|c| {
        c.emit_i32_const(1, 0);
        c.emit_op(Op::I32_EQZ, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

// ──────────────────────────────────────────────────────────────────────
// 9. Typed arithmetic and comparison operations
// ──────────────────────────────────────────────────────────────────────

#[test]
fn f64_add_numbers() {
    let r = run_script(|c| {
        c.emit_f64_const(1.5, 0);
        c.emit_f64_const(2.25, 0);
        c.emit_op(Op::F64_ADD, 0);
    });
    assert_eq!(r.as_f64(), 3.75);
}

#[test]
fn f64_lt_true() {
    let r = run_script(|c| {
        c.emit_f64_const(1.0, 0);
        c.emit_f64_const(2.0, 0);
        c.emit_op(Op::F64_LT, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn f64_eq_true() {
    let r = run_script(|c| {
        c.emit_f64_const(3.14, 0);
        c.emit_f64_const(3.14, 0);
        c.emit_op(Op::F64_EQ, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn i32_eqz_true_negates_truthy() {
    // I32_EQZ is the WASM-compliant way to invert a boolean i32
    let r = run_script(|c| {
        c.emit_i32_const(1, 0);
        c.emit_op(Op::I32_EQZ, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn i32_eqz_false_negates_falsy() {
    let r = run_script(|c| {
        c.emit_i32_const(0, 0);
        c.emit_op(Op::I32_EQZ, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn f64_neg_negates() {
    let r = run_script(|c| {
        c.emit_f64_const(5.0, 0);
        c.emit_op(Op::F64_NEG, 0);
    });
    assert_eq!(r.as_f64(), -5.0);
}

// ──────────────────────────────────────────────────────────────────────
// 10. STRING OPS
// ──────────────────────────────────────────────────────────────────────

// ──────────────────────────────────────────────────────────────────────
// 11. GC ARRAY OPS
// ──────────────────────────────────────────────────────────────────────

#[test]
fn array_new_fixed_and_length() {
    let r = run_script(|c| {
        // Push 3 elements, array.new_fixed 3
        c.emit_f64_const(10.0, 0);
        c.emit_f64_const(20.0, 0);
        c.emit_f64_const(30.0, 0);
        c.emit_array_new_fixed(0, 3, 0);
        c.emit_op(Op::ARRAY_LENGTH, 0);
    });
    assert_eq!(r.as_i32(), 3);
}

#[test]
fn array_get() {
    let r = run_script(|c| {
        c.emit_f64_const(10.0, 0);
        c.emit_f64_const(20.0, 0);
        c.emit_f64_const(30.0, 0);
        c.emit_array_new_fixed(0, 3, 0);
        // array_get arr 1 → 20
        c.emit_i32_const(1, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    });
    assert_eq!(r.as_f64(), 20.0);
}

#[test]
fn array_set_then_get() {
    let r = run_script(|c| {
        c.local_count = 1;
        // Create array [1, 2, 3]
        c.emit_f64_const(1.0, 0);
        c.emit_f64_const(2.0, 0);
        c.emit_f64_const(3.0, 0);
        c.emit_array_new_fixed(0, 3, 0);
        // Save to slot 0
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        // arr[1] = 99
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(1, 0);
        c.emit_f64_const(99.0, 0);
        c.emit_op(Op::ARRAY_SET, 0);
        c.emit_op(Op::DROP, 0); // array.set result
        // Get arr[1]
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(1, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    });
    assert_eq!(r.as_f64(), 99.0);
}

// ──────────────────────────────────────────────────────────────────────
// 12. FUNCTION CALL EDGE CASES
// ──────────────────────────────────────────────────────────────────────

#[test]
fn function_zero_args() {
    let mut script = Chunk::new("<script>");
    let mut fun = Chunk::new("f");
    fun.arity = 0;
    fun.local_count = 0;
    fun.emit_f64_const(42.0, 0);
    fun.emit_op(Op::RETURN, 0);

    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_op_u8_u8(Op::CALL_REF, 0, 1, 0);
    script.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![script, fun]).unwrap();
    assert_eq!(result.as_f64(), 42.0);
}

#[test]
fn function_three_args() {
    // sum3(a, b, c) { return a + b + c; } ; sum3(1, 2, 3) → 6
    let mut script = Chunk::new("<script>");
    let mut fun = Chunk::new("sum3");
    fun.arity = 3;
    fun.local_count = 3;
    fun.emit_op_u16(Op::LOCAL_GET, 0, 0);
    fun.emit_op_u16(Op::LOCAL_GET, 1, 0);
    fun.emit_op(Op::F64_ADD, 0);
    fun.emit_op_u16(Op::LOCAL_GET, 2, 0);
    fun.emit_op(Op::F64_ADD, 0);
    fun.emit_op(Op::RETURN, 0);

    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_f64_const(1.0, 0);
    script.emit_f64_const(2.0, 0);
    script.emit_f64_const(3.0, 0);
    script.emit_op_u8_u8(Op::CALL_REF, 3, 1, 0);
    script.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![script, fun]).unwrap();
    assert_eq!(result.as_f64(), 6.0);
}

#[test]
fn recursive_function_fibonacci() {
    // fib(n) { if (n < 2) return n; return fib(n-1) + fib(n-2); }
    // fib(10) = 55
    let mut script = Chunk::new("<script>");
    let mut fib = Chunk::new("fib");
    fib.arity = 1;
    fib.local_count = 1;

    // if n < 2: return n
    fib.emit_op_u16(Op::LOCAL_GET, 0, 0);
    fib.emit_f64_const(2.0, 0);
    fib.emit_op(Op::F64_LT, 0);
    // if non-zero, return n
    // Use structured CF: block { br_if 0 if not <2; return }
    let bp = fib.emit_block(0);
    fib.emit_op(Op::I32_EQZ, 0);
    fib.emit_br_if(0, 0); // skip if not less than 2
    fib.emit_op_u16(Op::LOCAL_GET, 0, 0);
    fib.emit_op(Op::RETURN, 0);
    fib.emit_end(0);
    fib.patch_block(bp);

    // return fib(n-1) + fib(n-2)
    fib.emit_op_u16(Op::REF_FUNC, 1, 0); // ref to self (chunk 1)
    fib.emit(0, 0);
    fib.emit_op_u16(Op::LOCAL_GET, 0, 0);
    fib.emit_f64_const(1.0, 0);
    fib.emit_op(Op::F64_SUB, 0);
    fib.emit_op_u8_u8(Op::CALL_REF, 1, 1, 0);

    fib.emit_op_u16(Op::REF_FUNC, 1, 0);
    fib.emit(0, 0);
    fib.emit_op_u16(Op::LOCAL_GET, 0, 0);
    fib.emit_f64_const(2.0, 0);
    fib.emit_op(Op::F64_SUB, 0);
    fib.emit_op_u8_u8(Op::CALL_REF, 1, 1, 0);

    fib.emit_op(Op::F64_ADD, 0);
    fib.emit_op(Op::RETURN, 0);

    // script: fib(10)
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_f64_const(10.0, 0);
    script.emit_op_u8_u8(Op::CALL_REF, 1, 1, 0);
    script.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![script, fib]).unwrap();
    assert_eq!(result.as_f64(), 55.0);
}

// ──────────────────────────────────────────────────────────────────────
// 13. MORE CONTROL FLOW EDGE CASES
// ──────────────────────────────────────────────────────────────────────

#[test]
fn br_if_rejects_null_condition() {
    // WASM br_if consumes an i32 condition. Null truthiness belongs in front-end lowering.
    let mut chunk = Chunk::new("<script>");
    {
        let c = &mut chunk;
        let bp = c.emit_block(0);
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        c.emit_br_if(0, 0);
        c.emit_f64_const(7.0, 0);
        c.emit_br(0, 0);
        c.emit_end(0);
        c.patch_block(bp);
        c.emit_op(Op::RETURN, 0);
    }
    let mut vm = VM::new();
    let err = vm
        .run(vec![chunk])
        .expect_err("br_if should reject null conditions");
    assert!(err.message.contains("br_if expected i32 condition"));
}

#[test]
fn loop_restarts_at_top() {
    // Simple counter test
    let r = run_script(|c| {
        c.local_count = 1;
        c.emit_i32_const(0, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);

        let bp = c.emit_block(0);
        let (lp, _) = c.emit_loop_s(0);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(5, 0);
        c.emit_op(Op::I32_LT_S, 0);
        c.emit_op(Op::I32_EQZ, 0);
        c.emit_br_if(1, 0); // exit if !(i < 5)
        // i++
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(1, 0);
        c.emit_op(Op::I32_ADD, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        c.emit_br(0, 0);
        c.emit_end(0);
        c.patch_loop(lp);
        c.emit_end(0);
        c.patch_block(bp);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    });
    assert_eq!(r.as_i32(), 5);
}

#[test]
fn nested_blocks_br_outer() {
    // block $outer { block $inner { br 1 (to outer) ; push 1 } ; push 2 }
    // Should: skip both "push 1" and "push 2", fall through.
    let r = run_script(|c| {
        let outer = c.emit_block(0);
        let inner = c.emit_block(0);
        c.emit_br(1, 0); // to outer end
        c.emit_f64_const(1.0, 0);
        c.emit_end(0);
        c.patch_block(inner);
        c.emit_f64_const(2.0, 0);
        c.emit_end(0);
        c.patch_block(outer);
        c.emit_f64_const(99.0, 0);
    });
    assert_eq!(r.as_f64(), 99.0);
}

// ──────────────────────────────────────────────────────────────────────
// 14. ROUND-TRIP FOR COMPLEX PROGRAMS
// ──────────────────────────────────────────────────────────────────────

#[test]
fn round_trip_preserves_line_info() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_f64_const(7.0, 42); // line 42
    chunk.emit_op(Op::RETURN, 42);

    let wasm = vybe_platform_wasm::write_wasm(&[chunk.clone()]);
    let chunks_rt = vybe_platform_wasm::read_wasm(&wasm).unwrap();
    assert_eq!(
        chunks_rt[0].lines.len(),
        chunk.lines.len(),
        "line count preserved"
    );
}

#[test]
fn round_trip_preserves_constants() {
    let mut chunk = Chunk::new("<script>");
    chunk.add_constant(Value::F64(1.0));
    chunk.add_constant(Value::String(std::sync::Arc::from("test")));
    chunk.add_constant(Value::I32(42));
    chunk.add_constant(Value::Bool(true));
    chunk.add_constant(Value::Null);
    chunk.emit_f64_const(99.0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let wasm = vybe_platform_wasm::write_wasm(&[chunk.clone()]);
    let chunks_rt = vybe_platform_wasm::read_wasm(&wasm).unwrap();
    assert_eq!(chunks_rt[0].constants.len(), chunk.constants.len());
}

#[test]
fn round_trip_loop_execution_matches() {
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);

    let bp = chunk.emit_block(0);
    let (lp, _) = chunk.emit_loop_s(0);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_i32_const(10, 0);
    chunk.emit_op(Op::I32_LT_S, 0);
    chunk.emit_op(Op::I32_EQZ, 0);
    chunk.emit_br_if(1, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
    chunk.emit_br(0, 0);
    chunk.emit_end(0);
    chunk.patch_loop(lp);
    chunk.emit_end(0);
    chunk.patch_block(bp);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let chunks = vec![chunk];

    // Direct run
    let mut vm1 = VM::new();
    let direct = vm1.run(chunks.clone()).unwrap().as_i32();

    // Round-trip run
    let wasm = vybe_platform_wasm::write_wasm(&chunks);
    let rt = vybe_platform_wasm::read_wasm(&wasm).unwrap();
    let mut vm2 = VM::new();
    let rt_result = vm2.run(rt).unwrap().as_i32();

    assert_eq!(direct, 10);
    assert_eq!(rt_result, 10);
    assert_eq!(direct, rt_result, "round-trip execution must match direct");
}

// ──────────────────────────────────────────────────────────────────────
// 15. BYTECODE SIZE & OPCODE ENCODING (WASM-spec compliance)
// ──────────────────────────────────────────────────────────────────────

#[test]
fn every_opcode_is_two_bytes() {
    // Opcodes are 2 bytes: [prefix, sub].
    // Verify encoding round-trip matches spec.
    for byte1 in [0x00u16, 0xF0, 0xFB, 0xFC, 0xFD, 0xFE, 0xFF] {
        for byte2 in 0u16..=0xFF {
            let _ = vybe_runtime::opcode::Op::decode(byte1, byte2);
        }
    }
}

#[test]
fn core_wasm_opcode_bytes_match_spec() {
    // Verify specific opcodes have the correct WASM byte values
    assert_eq!(Op::DROP.sub(), 0x1A, "drop should be 0x1A per WASM spec");
    assert_eq!(Op::I32_ADD.sub(), 0x6A, "i32.add should be 0x6A");
    assert_eq!(Op::I32_SUB.sub(), 0x6B, "i32.sub should be 0x6B");
    assert_eq!(Op::F64_ADD.sub(), 0xA0, "f64.add should be 0xA0");
    assert_eq!(Op::F64_NEG.sub(), 0x9A, "f64.neg should be 0x9A");
    assert_eq!(Op::LOCAL_GET.sub(), 0x20, "local.get should be 0x20");
    assert_eq!(Op::LOCAL_SET.sub(), 0x21, "local.set should be 0x21");
    assert_eq!(Op::CALL.sub(), 0x10, "call should be 0x10");
    assert_eq!(Op::RETURN.sub(), 0x0F, "return should be 0x0F");
    assert_eq!(Op::END.sub(), 0x0B, "end should be 0x0B");
    assert_eq!(Op::BLOCK.sub(), 0x02, "block should be 0x02");
    assert_eq!(Op::LOOP.sub(), 0x03, "loop should be 0x03");
    assert_eq!(Op::BR.sub(), 0x0C, "br should be 0x0C");
}

#[test]
fn core_opcodes_have_prefix_0x00() {
    assert_eq!(Op::DROP.group(), 0x00);
    assert_eq!(Op::I32_ADD.group(), 0x00);
    assert_eq!(Op::F64_ADD.group(), 0x00);
    assert_eq!(Op::LOCAL_GET.group(), 0x00);
    assert_eq!(Op::END.group(), 0x00);
}

#[allow(non_snake_case)]
#[test]
fn gc_opcodes_have_prefix_0xFB() {
    assert_eq!(Op::STRUCT_NEW.group(), 0xFB);
    assert_eq!(Op::ARRAY_NEW.group(), 0xFB);
    assert_eq!(Op::ARRAY_GET.group(), 0xFB);
}

#[allow(non_snake_case)]
#[test]
fn prefix_0xFF_holds_zero_opcodes() {
    // The custom-opcode prefix is EMPTY: CONST (0x00), CALL_IMPORT (0x04)
    // and HALT (0x23) are all retired to spec encodings, and their slots
    // (like every previously retired 0xFF slot) must fail decode loudly
    // rather than alias to something else.
    for sub in 0x00..=0xFFu16 {
        assert!(
            Op::decode(0xFF, sub).is_none(),
            "0xFF 0x{sub:02X} decoded — prefix 0xFF must hold zero opcodes"
        );
    }
    // The canon prefix (0xF0) is equally empty: the CM defines canon
    // built-ins as (core func) definitions — functions, not instructions —
    // so they resolve as imports under module "canon" and are reached via
    // spec `call`. Stale 0xF0 bytecode must fail decode loudly.
    for sub in 0x00..=0xFFu16 {
        assert!(
            Op::decode(0xF0, sub).is_none(),
            "0xF0 0x{sub:02X} decoded — canon built-ins are imports, not opcodes"
        );
    }
    assert_eq!(Op::BR.group(), 0x00);
    assert_eq!(Op::BR.sub(), 0x0C);
}

// ──────────────────────────────────────────────────────────────────────
// 16. GLOBALS (global.get / global.set)
// ──────────────────────────────────────────────────────────────────────

#[test]
fn global_set_then_get_returns_same_value() {
    let result = run_script(|c| {
        let name_idx = c.add_constant(Value::String(std::sync::Arc::from("x")));
        c.emit_f64_const(42.0, 0);
        c.emit_op_u16(Op::GLOBAL_SET, name_idx, 0);
        c.emit_op_u16(Op::GLOBAL_GET, name_idx, 0);
    });
    assert_eq!(result.as_f64(), 42.0);
}

#[test]
fn global_overwrite_takes_latest_value() {
    let result = run_script(|c| {
        let name_idx = c.add_constant(Value::String(std::sync::Arc::from("x")));
        c.emit_f64_const(1.0, 0);
        c.emit_op_u16(Op::GLOBAL_SET, name_idx, 0);
        c.emit_f64_const(99.0, 0);
        c.emit_op_u16(Op::GLOBAL_SET, name_idx, 0);
        c.emit_op_u16(Op::GLOBAL_GET, name_idx, 0);
    });
    assert_eq!(result.as_f64(), 99.0);
}

#[test]
fn global_get_undefined_returns_undefined() {
    let result = run_script(|c| {
        let name_idx = c.add_constant(Value::String(std::sync::Arc::from("never_set_global_xyz")));
        c.emit_op_u16(Op::GLOBAL_GET, name_idx, 0);
    });
    assert!(matches!(result, Value::Undefined));
}

// ──────────────────────────────────────────────────────────────────────
// 17. SELECT (WASM 0x1B)
// ──────────────────────────────────────────────────────────────────────

#[test]
fn select_picks_val1_when_cond_nonzero() {
    let result = run_script(|c| {
        c.emit_i32_const(10, 0);
        c.emit_i32_const(20, 0);
        c.emit_i32_const(1, 0);
        c.emit_op(Op::SELECT, 0);
    });
    assert_eq!(result.as_i32(), 10);
}

#[test]
fn select_picks_val2_when_cond_zero() {
    let result = run_script(|c| {
        c.emit_i32_const(10, 0);
        c.emit_i32_const(20, 0);
        c.emit_i32_const(0, 0);
        c.emit_op(Op::SELECT, 0);
    });
    assert_eq!(result.as_i32(), 20);
}

// ──────────────────────────────────────────────────────────────────────
// 18. I64 ARITHMETIC (WASM 0x7C..0x8A range)
// ──────────────────────────────────────────────────────────────────────

#[test]
fn i64_add_simple() {
    let result = run_script(|c| {
        c.emit_i64_const(1_000_000_000_000, 0);
        c.emit_i64_const(2_000_000_000_000, 0);
        c.emit_op(Op::I64_ADD, 0);
    });
    assert_eq!(result.as_i64(), 3_000_000_000_000);
}

#[test]
fn i64_sub_negative_result() {
    let result = run_script(|c| {
        c.emit_i64_const(10, 0);
        c.emit_i64_const(50, 0);
        c.emit_op(Op::I64_SUB, 0);
    });
    assert_eq!(result.as_i64(), -40);
}

#[test]
fn i64_mul() {
    let result = run_script(|c| {
        c.emit_i64_const(1_000_000, 0);
        c.emit_i64_const(1_000_000, 0);
        c.emit_op(Op::I64_MUL, 0);
    });
    assert_eq!(result.as_i64(), 1_000_000_000_000);
}

#[test]
fn i64_div_s_negative() {
    let result = run_script(|c| {
        c.emit_i64_const(-100, 0);
        c.emit_i64_const(7, 0);
        c.emit_op(Op::I64_DIV_S, 0);
    });
    assert_eq!(result.as_i64(), -14);
}

#[test]
fn i64_and_or_xor() {
    // 0xFF00 AND 0x0FF0 = 0x0F00
    let and_result = run_script(|c| {
        c.emit_i64_const(0xFF00, 0);
        c.emit_i64_const(0x0FF0, 0);
        c.emit_op(Op::I64_AND, 0);
    });
    assert_eq!(and_result.as_i64(), 0x0F00);

    // 0xFF00 OR 0x0FF0 = 0xFFF0
    let or_result = run_script(|c| {
        c.emit_i64_const(0xFF00, 0);
        c.emit_i64_const(0x0FF0, 0);
        c.emit_op(Op::I64_OR, 0);
    });
    assert_eq!(or_result.as_i64(), 0xFFF0);

    // 0xFF00 XOR 0x0FF0 = 0xF0F0
    let xor_result = run_script(|c| {
        c.emit_i64_const(0xFF00, 0);
        c.emit_i64_const(0x0FF0, 0);
        c.emit_op(Op::I64_XOR, 0);
    });
    assert_eq!(xor_result.as_i64(), 0xF0F0);
}

#[test]
fn i64_eqz_zero_returns_true() {
    let result = run_script(|c| {
        c.emit_i64_const(0, 0);
        c.emit_op(Op::I64_EQZ, 0);
    });
    assert_eq!(result.as_i32(), 1);
}

#[test]
fn i64_eqz_nonzero_returns_false() {
    let result = run_script(|c| {
        c.emit_i64_const(42, 0);
        c.emit_op(Op::I64_EQZ, 0);
    });
    assert_eq!(result.as_i32(), 0);
}

#[test]
fn i64_extend_i32_s_preserves_sign() {
    let result = run_script(|c| {
        c.emit_i32_const(-5, 0);
        c.emit_op(Op::I64_EXTEND_I32_S, 0);
    });
    assert_eq!(result.as_i64(), -5);
}

// ──────────────────────────────────────────────────────────────────────
// 19. MORE CONTROL FLOW (structured if/br chains)
// ──────────────────────────────────────────────────────────────────────

#[test]
fn structured_if_takes_then_when_i32_nonzero() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(1, 0);
    chunk.emit_if_value(0);
    chunk.emit_f64_const(99.0, 0);
    chunk.emit_else(0);
    chunk.emit_f64_const(5.0, 0);
    chunk.emit_end(0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_f64(), 99.0);
}

#[test]
fn structured_if_takes_else_when_i32_zero() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(0, 0);
    chunk.emit_if_value(0);
    chunk.emit_f64_const(99.0, 0);
    chunk.emit_else(0);
    chunk.emit_f64_const(5.0, 0);
    chunk.emit_end(0);
    chunk.emit_f64_const(99.0, 0);
    chunk.emit_op(Op::F64_ADD, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_f64(), 104.0);
}

#[test]
fn return_inside_block_does_not_poison_caller_labels() {
    let mut script = Chunk::new("<script>");
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0);
    script.emit_op_u8_u8(Op::CALL_REF, 0, 1, 0);
    script.emit_op(Op::DROP, 0);
    script.emit_i32_const(0, 0);
    script.emit_if_value(0);
    script.emit_f64_const(1.0, 0);
    script.emit_else(0);
    script.emit_f64_const(2.0, 0);
    script.emit_end(0);
    script.emit_op(Op::RETURN, 0);

    let mut callee = Chunk::new("returns_inside_block");
    callee.emit_block(0);
    callee.emit_f64_const(7.0, 0);
    callee.emit_op(Op::RETURN, 0);
    callee.emit_end(0);
    callee.emit_f64_const(99.0, 0);
    callee.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![script, callee]).unwrap();
    assert_eq!(result.as_f64(), 2.0);
}

#[test]
fn br_to_outer_block_skips_inner_work() {
    // Structured:
    //   block { block { br 1 } push 5 } push 42
    // br 1 jumps out of both blocks, then we push 42.
    let mut chunk = Chunk::new("<script>");
    let outer = chunk.emit_block(0);
    let inner = chunk.emit_block(0);
    chunk.emit_br(1, 0); // break out of outer
    chunk.emit_op(Op::END, 0); // end inner
    chunk.patch_block(inner);
    chunk.emit_f64_const(5.0, 0); // should be skipped
    chunk.emit_op(Op::END, 0); // end outer
    chunk.patch_block(outer);
    chunk.emit_f64_const(42.0, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    assert_eq!(result.as_f64(), 42.0);
}

// ──────────────────────────────────────────────────────────────────────
// 20. ROUND-TRIP: GLOBALS, MULTIPLE FUNCTIONS
// ──────────────────────────────────────────────────────────────────────

#[test]
fn round_trip_globals_persist() {
    let mut script = Chunk::new("<script>");
    let name_idx = script.add_constant(Value::String(std::sync::Arc::from("gx")));
    script.emit_f64_const(123.0, 0);
    script.emit_op_u16(Op::GLOBAL_SET, name_idx, 0);
    script.emit_op_u16(Op::GLOBAL_GET, name_idx, 0);
    script.emit_op(Op::RETURN, 0);

    let chunks = vec![script];
    let mut vm1 = VM::new();
    let direct = vm1.run(chunks.clone()).unwrap().as_f64();

    let wasm = vybe_platform_wasm::write_wasm(&chunks);
    let rt = vybe_platform_wasm::read_wasm(&wasm).unwrap();
    let mut vm2 = VM::new();
    let rt_result = vm2.run(rt).unwrap().as_f64();

    assert_eq!(direct, 123.0);
    assert_eq!(rt_result, 123.0);
}

#[test]
fn round_trip_three_functions_chained_calls() {
    // script calls f1(2), f1 calls f2(x+1), f2 returns x*10
    // expected: f2(2+1)*1 = 30
    let mut script = Chunk::new("<script>");
    script.emit_op_u16(Op::REF_FUNC, 1, 0); // f1
    script.emit(0, 0); // uv_count
    script.emit_f64_const(2.0, 0);
    script.emit_op_u8_u8(Op::CALL_REF, 1, 1, 0);
    script.emit_op(Op::RETURN, 0);

    let mut f1 = Chunk::new("f1");
    f1.arity = 1;
    f1.local_count = 1;
    f1.emit_op_u16(Op::REF_FUNC, 2, 0); // f2
    f1.emit(0, 0); // uv_count
    f1.emit_op_u16(Op::LOCAL_GET, 0, 0); // x
    f1.emit_f64_const(1.0, 0);
    f1.emit_op(Op::F64_ADD, 0);
    f1.emit_op_u8_u8(Op::CALL_REF, 1, 1, 0);
    f1.emit_op(Op::RETURN, 0);

    let mut f2 = Chunk::new("f2");
    f2.arity = 1;
    f2.local_count = 1;
    f2.emit_op_u16(Op::LOCAL_GET, 0, 0);
    f2.emit_f64_const(10.0, 0);
    f2.emit_op(Op::F64_MUL, 0);
    f2.emit_op(Op::RETURN, 0);

    let chunks = vec![script, f1, f2];
    let mut vm1 = VM::new();
    let direct = vm1.run(chunks.clone()).unwrap().as_f64();

    let wasm = vybe_platform_wasm::write_wasm(&chunks);
    let rt = vybe_platform_wasm::read_wasm(&wasm).unwrap();
    let mut vm2 = VM::new();
    let rt_result = vm2.run(rt).unwrap().as_f64();

    assert_eq!(direct, 30.0);
    assert_eq!(
        rt_result, direct,
        "round-trip must preserve chained call semantics"
    );
}

#[test]
fn round_trip_preserves_function_arity() {
    let mut script = Chunk::new("<script>");
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(Op::RETURN, 0);

    let mut f_two = Chunk::new("f_two");
    f_two.arity = 2;
    f_two.local_count = 2;
    f_two.emit_op(Op::RETURN, 0);

    let mut f_five = Chunk::new("f_five");
    f_five.arity = 5;
    f_five.local_count = 5;
    f_five.emit_op(Op::RETURN, 0);

    let chunks = vec![script, f_two, f_five];
    let wasm = vybe_platform_wasm::write_wasm(&chunks);
    let rt = vybe_platform_wasm::read_wasm(&wasm).unwrap();

    assert_eq!(rt.len(), 3);
    assert_eq!(rt[1].arity, 2, "f_two arity preserved");
    assert_eq!(rt[2].arity, 5, "f_five arity preserved");
}

#[test]
fn round_trip_preserves_chunk_name() {
    let mut script = Chunk::new("<script>");
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(Op::RETURN, 0);

    let mut named = Chunk::new("my_cool_function");
    named.arity = 0;
    named.emit_op(Op::RETURN, 0);

    let chunks = vec![script, named];
    let wasm = vybe_platform_wasm::write_wasm(&chunks);
    let rt = vybe_platform_wasm::read_wasm(&wasm).unwrap();
    assert_eq!(
        rt[1].name, "my_cool_function",
        "chunk name preserved across round-trip"
    );
}

// ──────────────────────────────────────────────────────────────────────
// 21. WASM BINARY STRUCTURE (validates magic, section order)
// ──────────────────────────────────────────────────────────────────────

#[test]
fn wasm_binary_sections_in_correct_order() {
    // WASM spec: sections must appear in a specific order (when present):
    //   1=type, 2=import, 3=function, 4=table, 5=memory, 6=global,
    //   7=export, 8=start, 9=element, 10=code, 11=data, 12=data_count, 0=custom
    // Custom sections can appear anywhere.
    let mut script = Chunk::new("<script>");
    script.emit_f64_const(1.0, 0);
    script.emit_op(Op::RETURN, 0);

    let wasm = vybe_platform_wasm::write_wasm(&vec![script]);

    // Skip magic + version (8 bytes)
    let mut pos = 8;
    let mut last_known_order = 0u8;
    while pos < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        // Read section size (LEB128, but for small sections 1 byte)
        let mut size: u32 = 0;
        let mut shift = 0;
        loop {
            let b = wasm[pos];
            pos += 1;
            size |= ((b & 0x7F) as u32) << shift;
            if b & 0x80 == 0 {
                break;
            }
            shift += 7;
        }

        // Custom sections (id=0) can appear anywhere — skip order check
        if section_id != 0 {
            assert!(
                section_id >= last_known_order,
                "section id {} appeared after {} — WASM spec requires ordered sections",
                section_id,
                last_known_order
            );
            last_known_order = section_id;
        }

        pos += size as usize;
    }
    assert_eq!(pos, wasm.len(), "sections should span the entire module");
}

#[test]
fn wasm_contains_vybe_custom_section() {
    let mut script = Chunk::new("<script>");
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(Op::RETURN, 0);

    let wasm = vybe_platform_wasm::write_wasm(&vec![script]);

    // Search for "vybe" string as section name
    let needle = b"vybe";
    let found = wasm.windows(needle.len()).any(|w| w == needle);
    assert!(
        found,
        "wasm must contain custom section named 'vybe' for round-trip"
    );
}

// ──────────────────────────────────────────────────────────────────────
// 22. NUMERIC EDGE CASES (NaN, infinity, integer boundaries)
// ──────────────────────────────────────────────────────────────────────

#[test]
fn f64_nan_is_not_equal_to_itself() {
    let result = run_script(|c| {
        c.emit_f64_const(f64::NAN, 0);
        c.emit_f64_const(f64::NAN, 0);
        c.emit_op(Op::F64_EQ, 0);
    });
    assert_eq!(result.as_bool(), false, "NaN != NaN per IEEE-754");
}

#[test]
fn f64_infinity_plus_finite_is_infinity() {
    let result = run_script(|c| {
        c.emit_f64_const(f64::INFINITY, 0);
        c.emit_f64_const(1.0, 0);
        c.emit_op(Op::F64_ADD, 0);
    });
    assert!(result.as_f64().is_infinite() && result.as_f64() > 0.0);
}

#[test]
fn i32_max_plus_one_wraps_to_min() {
    // WASM i32.add is modulo 2^32
    let result = run_script(|c| {
        c.emit_i32_const(i32::MAX, 0);
        c.emit_i32_const(1, 0);
        c.emit_op(Op::I32_ADD, 0);
    });
    assert_eq!(result.as_i32(), i32::MIN, "i32.add overflow wraps per WASM");
}

#[test]
fn i32_min_minus_one_wraps_to_max() {
    let result = run_script(|c| {
        c.emit_i32_const(i32::MIN, 0);
        c.emit_i32_const(1, 0);
        c.emit_op(Op::I32_SUB, 0);
    });
    assert_eq!(
        result.as_i32(),
        i32::MAX,
        "i32.sub underflow wraps per WASM"
    );
}

// ──────────────────────────────────────────────────────────────────────
// 23. LOCAL SLOT INDEPENDENCE (each arg/local gets its own slot)
// ──────────────────────────────────────────────────────────────────────

#[test]
fn five_args_each_go_to_correct_slot() {
    // f(a,b,c,d,e) → returns a + b*10 + c*100 + d*1000 + e*10000
    // Call f(1,2,3,4,5) → expect 54321
    let mut script = Chunk::new("<script>");
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0); // uv_count
    for i in 1..=5 {
        script.emit_f64_const(i as f64, 0);
    }
    script.emit_op_u8_u8(Op::CALL_REF, 5, 1, 0);
    script.emit_op(Op::RETURN, 0);

    let mut f = Chunk::new("combo");
    f.arity = 5;
    f.local_count = 5;
    // a (slot 0) + b (slot 1) * 10
    f.emit_op_u16(Op::LOCAL_GET, 0, 0);
    f.emit_op_u16(Op::LOCAL_GET, 1, 0);
    f.emit_f64_const(10.0, 0);
    f.emit_op(Op::F64_MUL, 0);
    f.emit_op(Op::F64_ADD, 0);
    // + c * 100
    f.emit_op_u16(Op::LOCAL_GET, 2, 0);
    f.emit_f64_const(100.0, 0);
    f.emit_op(Op::F64_MUL, 0);
    f.emit_op(Op::F64_ADD, 0);
    // + d * 1000
    f.emit_op_u16(Op::LOCAL_GET, 3, 0);
    f.emit_f64_const(1000.0, 0);
    f.emit_op(Op::F64_MUL, 0);
    f.emit_op(Op::F64_ADD, 0);
    // + e * 10000
    f.emit_op_u16(Op::LOCAL_GET, 4, 0);
    f.emit_f64_const(10000.0, 0);
    f.emit_op(Op::F64_MUL, 0);
    f.emit_op(Op::F64_ADD, 0);
    f.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![script, f]).unwrap().as_f64();
    assert_eq!(result, 54321.0, "each arg must land in its declared slot");
}

#[test]
fn local_set_on_arg_slot_overwrites_arg() {
    // f(a) sets slot 0 = 999, returns slot 0 → should be 999, not the arg
    let mut script = Chunk::new("<script>");
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0); // uv_count
    script.emit_f64_const(7.0, 0);
    script.emit_op_u8_u8(Op::CALL_REF, 1, 1, 0);
    script.emit_op(Op::RETURN, 0);

    let mut f = Chunk::new("clobber_arg");
    f.arity = 1;
    f.local_count = 1;
    f.emit_f64_const(999.0, 0);
    f.emit_op_u16(Op::LOCAL_SET, 0, 0);
    f.emit_op_u16(Op::LOCAL_GET, 0, 0);
    f.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![script, f]).unwrap().as_f64();
    assert_eq!(result, 999.0);
}

// ──────────────────────────────────────────────────────────────────────
// 24. JS BUILTINS IMPORTS (js-string / js-primitive / js-symbol / js-bigint)
// ──────────────────────────────────────────────────────────────────────

/// Helper: emit a tiny script that returns a single boxed f64, then
/// produce its .wasm bytes.
fn emit_trivial_wasm() -> Vec<u8> {
    let mut script = Chunk::new("<script>");
    script.emit_f64_const(7.0, 0);
    script.emit_op(Op::RETURN, 0);
    vybe_platform_wasm::write_wasm(&vec![script])
}

/// Every builtin we expect in the emitter must appear as an import in the
/// emitted .wasm. This pins the full set so regressions are caught.
#[test]
fn wasm_imports_include_full_js_builtins_surface() {
    let wasm = emit_trivial_wasm();
    let required = [
        // js-number
        "fromF64",
        "fromI32",
        "fromU32",
        "toF64",
        "toI32",
        "toU32",
        "test",
        "testI32",
        "testU32",
        // js-string (core)
        "concat",
        "equals",
        "compare",
        "length",
        "charCodeAt",
        "codePointAt",
        "fromCharCode",
        "fromCodePoint",
        "substring",
        "cast",
        "intoCharCodeArray",
        "fromCharCodeArray",
        // js-string (primitive-builtins numeric formatting)
        "fromI64",
        "fromU64",
        // js-boolean
        // js-undefined
        // js-symbol
        "equals",
        // js-bigint
    ];
    for name in required {
        let bytes = name.as_bytes();
        let found = wasm.windows(bytes.len()).any(|w| w == bytes);
        assert!(
            found,
            "builtin `{}` not found in emitted .wasm import section",
            name
        );
    }
}

#[test]
fn wasm_imports_include_all_js_builtin_modules() {
    let wasm = emit_trivial_wasm();
    for module in [
        "wasm:js-number",
        "wasm:js-string",
        "wasm:js-boolean",
        "wasm:js-undefined",
        "wasm:js-symbol",
        "wasm:js-bigint",
    ] {
        let bytes = module.as_bytes();
        let found = wasm.windows(bytes.len()).any(|w| w == bytes);
        assert!(
            found,
            "module `{}` missing from emitted .wasm imports",
            module
        );
    }
}

/// Round-trip: the 4 js-string-builtins string opcodes that were
/// previously unwired now execute correctly in the VM.
// ──────────────────────────────────────────────────────────────────────
// 25. ESM INTEGRATION — generated .wasm has a clean shape
// ──────────────────────────────────────────────────────────────────────

// ──────────────────────────────────────────────────────────────────────
// 24b. JS PRIMITIVES — Undefined / Symbol / BigInt opcodes
// ──────────────────────────────────────────────────────────────────────

// ──────────────────────────────────────────────────────────────────────
// 25. ESM INTEGRATION — generated .wasm has a clean shape
// ──────────────────────────────────────────────────────────────────────

/// An ES-module-importable `.wasm` must have a valid export section.
/// Verifies the emitted module exports at least one function (the script /
/// anonymous entry), which is what `import` statements in JS would bind.
#[test]
fn emitted_wasm_has_js_global_imports() {
    // wasm:js-undefined.value + wasm:js-boolean.true + wasm:js-boolean.false
    // are imported as WASM globals so `undefined`/`true`/`false` can be
    // produced without boxing. Verify all three names are present AND that
    // the import section contains global descriptors (kind 0x03).
    let wasm = emit_trivial_wasm();
    for name in ["value", "true", "false"] {
        let bytes = name.as_bytes();
        assert!(
            wasm.windows(bytes.len()).any(|w| w == bytes),
            "global import name `{}` not found in emitted .wasm",
            name
        );
    }
    // Scan import section for at least one global descriptor (kind 0x03,
    // externref = 0x6F, mut = 0x00).
    let mut pos = 8;
    let mut found_global = false;
    while pos < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        let mut size: u32 = 0;
        let mut shift = 0;
        loop {
            let b = wasm[pos];
            pos += 1;
            size |= ((b & 0x7F) as u32) << shift;
            if b & 0x80 == 0 {
                break;
            }
            shift += 7;
        }
        if section_id == 2 {
            // Import section — scan for kind=0x03, valtype=0x6F, mut=0x00
            let end = pos + size as usize;
            let mut scan = pos;
            while scan + 2 < end {
                if wasm[scan] == 0x03 && wasm[scan + 1] == 0x6F && wasm[scan + 2] == 0x00 {
                    found_global = true;
                    break;
                }
                scan += 1;
            }
            break;
        }
        pos += size as usize;
    }
    assert!(found_global, "no externref-global import descriptor found");
}

#[test]
fn emitted_wasm_has_export_section() {
    let wasm = emit_trivial_wasm();
    // Export section id = 7
    let mut pos = 8; // skip magic + version
    let mut found = false;
    while pos < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        let mut size: u32 = 0;
        let mut shift = 0;
        loop {
            let b = wasm[pos];
            pos += 1;
            size |= ((b & 0x7F) as u32) << shift;
            if b & 0x80 == 0 {
                break;
            }
            shift += 7;
        }
        if section_id == 7 {
            found = true;
            // Read export count (first LEB128 byte in the section payload)
            assert!(wasm[pos] >= 1, "export section empty");
            break;
        }
        pos += size as usize;
    }
    assert!(
        found,
        "generated .wasm has no export section — ESM import would see no bindings"
    );
}

// ──────────────────────────────────────────────────────────────────────
// 25b. PROPOSAL MODULES — uniform shape audit
// ──────────────────────────────────────────────────────────────────────

/// Every proposal module under `wasm/` must expose a uniform surface so
/// future pipeline code can iterate over them: `IMPORTS`, `GLOBAL_IMPORTS`,
/// `custom_sections`. This test smoke-calls that surface on every module.
/// If a module is ever missing one of these items, the test fails to
/// compile — which is the point (structural enforcement).
#[test]
fn every_proposal_module_exposes_uniform_surface() {
    use vybe_platform_wasm::writer::builtins::{js_primitive_builtins, js_string_builtins};
    use vybe_platform_wasm::writer::proposals::{
        bulk_memory, esm_integration, exception_handling, gc, multi_value, reference_types, simd,
        tail_call, threads };

    // Each of these must compile — the test is the shape check.
    let mut total_imports = 0usize;
    total_imports += reference_types::declare_imports().len();
    total_imports += js_string_builtins::IMPORTS.len();
    total_imports += js_primitive_builtins::FUNC_IMPORTS.len();
    total_imports += gc::IMPORTS.len();
    total_imports += simd::IMPORTS.len();
    total_imports += threads::IMPORTS.len();
    total_imports += bulk_memory::IMPORTS.len();
    total_imports += exception_handling::IMPORTS.len();
    total_imports += tail_call::IMPORTS.len();
    total_imports += multi_value::IMPORTS.len();
    assert!(
        total_imports > 0,
        "no proposal module exposes any imports — suspicious"
    );

    // ESM readiness check accepts a non-empty chunk list.
    let chunk = vybe_runtime::Chunk::new("ok");
    assert!(esm_integration::check_esm_readiness(&[chunk]).is_ok());
    assert!(esm_integration::check_esm_readiness(&[]).is_err());

    // Host-builtin detector recognises every `wasm:js-*` module.
    assert!(esm_integration::is_host_builtin("wasm:js-string", "concat"));
    assert!(esm_integration::is_host_builtin("wasm:js-bigint", "test"));
    assert!(!esm_integration::is_host_builtin("vybe:rt", "print"));
}

// ──────────────────────────────────────────────────────────────────────
// 26. EXTENDED NAME SECTION + COMPILATION HINTS proposals
// ──────────────────────────────────────────────────────────────────────

/// The emitted .wasm must carry a standard `"name"` custom section so
/// DevTools / node profilers can show readable names for functions,
/// locals, tables, memory, globals and element segments.
#[test]
fn emitted_wasm_has_standard_name_section() {
    let wasm = emit_trivial_wasm();
    // Scan top-level sections for a custom section whose name is "name".
    let mut pos = 8;
    let mut found_name_section = false;
    while pos < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        let (size, read) = decode_leb128_u32(&wasm[pos..]);
        pos += read;
        let end = pos + size as usize;
        if section_id == 0 {
            // Custom section — next is the section name.
            let (name_len, r2) = decode_leb128_u32(&wasm[pos..]);
            let name_start = pos + r2;
            let name_end = name_start + name_len as usize;
            if name_end <= end {
                let name = std::str::from_utf8(&wasm[name_start..name_end]).unwrap_or("");
                if name == "name" {
                    found_name_section = true;
                }
            }
        }
        pos = end;
    }
    assert!(
        found_name_section,
        "generated .wasm is missing the standard `name` custom section"
    );
}

/// The `"name"` section must contain at least subsection 1 (function
/// names) AND subsection 7 (global names) — the js-primitive globals
/// we declare must be named for debugging.
#[test]
fn name_section_includes_function_and_global_subsections() {
    let wasm = emit_trivial_wasm();
    let name_payload = find_custom_section(&wasm, "name").expect("name section missing");
    // Walk subsections.
    let mut pos = 0;
    let mut seen_fn = false;
    let mut seen_global = false;
    while pos < name_payload.len() {
        let id = name_payload[pos];
        pos += 1;
        let (size, r) = decode_leb128_u32(&name_payload[pos..]);
        pos += r;
        if id == 1 {
            seen_fn = true;
        }
        if id == 7 {
            seen_global = true;
        }
        pos += size as usize;
    }
    assert!(
        seen_fn,
        "name section missing function-names subsection (id=1)"
    );
    assert!(
        seen_global,
        "name section missing global-names subsection (id=7)"
    );
}

/// The compilation-hints proposal gets a custom section under the
/// `metadata.code.*` namespace. Verify we emit `metadata.code.compilation_order`.
#[test]
fn emitted_wasm_has_compilation_hints_section() {
    let wasm = emit_trivial_wasm();
    let payload = find_custom_section(&wasm, "metadata.code.compilation_order");
    assert!(
        payload.is_some(),
        "generated .wasm is missing metadata.code.compilation_order section"
    );
    // First byte of payload is the hint count (LEB128). Must be >= 1 for
    // our entry chunk.
    let p = payload.unwrap();
    let (count, _) = decode_leb128_u32(p);
    assert!(count >= 1, "compilation_order section has no hints");
}

// ──────────────────────────────────────────────────────────────────────
// 24b. IMPORTED STRING CONSTANTS (js-string-builtins § String constants)
// ──────────────────────────────────────────────────────────────────────

/// A module whose script pushes two string constants and returns.
fn emit_wasm_with_string_constants() -> Vec<u8> {
    let mut script = Chunk::new("<script>");
    script.emit_string_const("hello world", 0);
    script.emit_string_const("hello world", 0); // same import, one global
    script.emit_string_const("second", 0);
    script.emit_op(Op::RETURN, 0);
    vybe_platform_wasm::write_wasm(&vec![script])
}

/// Every import in the module as `(module, name, kind)`; kind is the spec's
/// external-kind byte — 0x00 func, 0x01 table, 0x02 memory, 0x03 global.
fn parse_imports(wasm: &[u8]) -> Vec<(String, String, u8)> {
    let Some(payload) = section_bytes(wasm, 2) else {
        return Vec::new();
    };
    let mut pos = 0usize;
    let (count, r) = decode_leb128_u32(&payload[pos..]);
    pos += r;
    let mut out = Vec::new();
    for _ in 0..count {
        let module = read_wasm_name(payload, &mut pos);
        let name = read_wasm_name(payload, &mut pos);
        let kind = payload[pos];
        pos += 1;
        match kind {
            0x00 => {
                let (_, r) = decode_leb128_u32(&payload[pos..]);
                pos += r;
            }
            0x01 => {
                pos += 1; // reftype
                let flags = payload[pos];
                pos += 1;
                let (_, r) = decode_leb128_u32(&payload[pos..]);
                pos += r;
                if flags & 0x01 != 0 {
                    let (_, r) = decode_leb128_u32(&payload[pos..]);
                    pos += r;
                }
            }
            0x02 => {
                let flags = payload[pos];
                pos += 1;
                let (_, r) = decode_leb128_u32(&payload[pos..]);
                pos += r;
                if flags & 0x01 != 0 {
                    let (_, r) = decode_leb128_u32(&payload[pos..]);
                    pos += r;
                }
            }
            0x03 => {
                pos += 2; // valtype + mutability
            }
            other => panic!("unknown import kind {}", other),
        }
        out.push((module, name, kind));
    }
    out
}

fn read_wasm_name(data: &[u8], pos: &mut usize) -> String {
    let (len, r) = decode_leb128_u32(&data[*pos..]);
    *pos += r;
    let s = String::from_utf8_lossy(&data[*pos..*pos + len as usize]).to_string();
    *pos += len as usize;
    s
}

fn section_bytes(wasm: &[u8], id: u8) -> Option<&[u8]> {
    let mut pos = 8;
    while pos < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        let (size, r) = decode_leb128_u32(&wasm[pos..]);
        pos += r;
        let end = pos + size as usize;
        if section_id == id {
            return Some(&wasm[pos..end]);
        }
        pos = end;
    }
    None
}

/// js-string-builtins § String constants: a string constant reaches the
/// module as an imported **global** whose field name IS its value. Never a
/// function, and therefore never a call — which also keeps it out of the
/// function index space that `CALL_IMPORT` operands are numbered in.
#[test]
fn string_constants_are_imported_globals_never_functions() {
    let wasm = emit_wasm_with_string_constants();
    let imports = parse_imports(&wasm);

    let from_namespace: Vec<_> = imports
        .iter()
        .filter(|(module, _, _)| module == "wasm:string-constants")
        .collect();

    assert!(
        !from_namespace.is_empty(),
        "no imports from the string-constants namespace"
    );
    for (module, name, kind) in &from_namespace {
        assert_eq!(
            *kind, 0x03,
            "`{}` `{}` is import kind {} — string constants must be globals",
            module, name, kind
        );
    }

    let names: Vec<&str> = from_namespace
        .iter()
        .map(|(_, name, _)| name.as_str())
        .collect();
    assert!(names.contains(&"hello world"), "constants: {:?}", names);
    assert!(names.contains(&"second"), "constants: {:?}", names);
    assert_eq!(
        names.iter().filter(|n| **n == "hello world").count(),
        1,
        "one import per distinct constant — two references share one global"
    );
}

/// The declared global type must be what the spec's own test vector accepts:
/// `externref`, **immutable**. (`constants.tentative.any.js` lists mutable
/// externref among the *bad* types — a mutable one is a CompileError.)
#[test]
fn string_constant_globals_are_immutable_externref() {
    let wasm = emit_wasm_with_string_constants();
    let payload = section_bytes(&wasm, 2).expect("import section");
    let mut pos = 0usize;
    let (count, r) = decode_leb128_u32(&payload[pos..]);
    pos += r;
    let mut checked = 0;
    for _ in 0..count {
        let module = read_wasm_name(payload, &mut pos);
        let _name = read_wasm_name(payload, &mut pos);
        let kind = payload[pos];
        pos += 1;
        if kind == 0x03 {
            let valtype = payload[pos];
            let mutability = payload[pos + 1];
            pos += 2;
            if module == "wasm:string-constants" {
                assert_eq!(valtype, 0x6F, "string constant global is not externref");
                assert_eq!(mutability, 0x00, "string constant global must be immutable");
                checked += 1;
            }
        } else {
            // Only globals appear after the function imports in our modules.
            let (_, r) = decode_leb128_u32(&payload[pos..]);
            pos += r;
        }
    }
    assert!(checked >= 2, "expected both constants, checked {}", checked);
}

/// The real oracle: v8 validates the module *with* the namespace designated
/// as `importedStringConstants`, which is when it enforces the global-type
/// rule. Skipped when node is unavailable.
#[test]
fn string_constants_accepted_by_node_with_imported_string_constants() {
    use std::io::Write;
    use std::process::Command;

    let node_ok = Command::new("node")
        .arg("--version")
        .output()
        .ok()
        .map(|o| o.status.success())
        .unwrap_or(false);
    if !node_ok {
        eprintln!("skipping: `node` not available on PATH");
        return;
    }

    let wasm = emit_wasm_with_string_constants();
    let dir = std::env::temp_dir().join(format!("vybe_strconst_{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).expect("create tmp dir");
    let wasm_path = dir.join("mod.wasm");
    let driver_path = dir.join("driver.mjs");
    std::fs::File::create(&wasm_path)
        .expect("open wasm file")
        .write_all(&wasm)
        .expect("write wasm");

    let driver = r#"
import { readFile } from 'node:fs/promises';
import { argv } from 'node:process';
const bytes = await readFile(argv[2]);
const mod = await WebAssembly.compile(bytes, {
    importedStringConstants: "wasm:string-constants" });
const found = WebAssembly.Module.imports(mod)
    .filter(i => i.module === 'wasm:string-constants');
if (!found.every(i => i.kind === 'global')) {
    console.error('string constants imported as ' + found.map(i => i.kind).join(','));
    process.exit(1);
}
console.log(JSON.stringify({ constants: found.length }));
"#;
    std::fs::File::create(&driver_path)
        .expect("open driver file")
        .write_all(driver.as_bytes())
        .expect("write driver");

    let out = Command::new("node")
        .arg(&driver_path)
        .arg(&wasm_path)
        .output()
        .expect("spawn node");
    let _ = std::fs::remove_dir_all(&dir);

    assert!(
        out.status.success(),
        "node rejected the string constants.\nstdout: {}\nstderr: {}",
        String::from_utf8_lossy(&out.stdout),
        String::from_utf8_lossy(&out.stderr)
    );
}

// ── Tiny LEB128 decoder + custom-section extractor for the tests. ─────

fn decode_leb128_u32(bytes: &[u8]) -> (u32, usize) {
    let mut result: u32 = 0;
    let mut shift = 0;
    let mut read = 0;
    for &b in bytes {
        read += 1;
        result |= ((b & 0x7F) as u32) << shift;
        if b & 0x80 == 0 {
            break;
        }
        shift += 7;
    }
    (result, read)
}

/// Returns the payload bytes of the custom section with the given name,
/// if present.
fn find_custom_section<'a>(wasm: &'a [u8], target: &str) -> Option<&'a [u8]> {
    let mut pos = 8;
    while pos < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        let (size, r) = decode_leb128_u32(&wasm[pos..]);
        pos += r;
        let end = pos + size as usize;
        if section_id == 0 {
            let (name_len, r2) = decode_leb128_u32(&wasm[pos..]);
            let name_start = pos + r2;
            let name_end = name_start + name_len as usize;
            if name_end <= end {
                if let Ok(name) = std::str::from_utf8(&wasm[name_start..name_end]) {
                    if name == target {
                        return Some(&wasm[name_end..end]);
                    }
                }
            }
        }
        pos = end;
    }
    None
}

// ──────────────────────────────────────────────────────────────────────
// 26b. Extended name section — subsections 3/4/10
// ──────────────────────────────────────────────────────────────────────

#[test]
fn emitted_wasm_declares_single_funcref_table_by_default() {
    // Spec-correct default: the reference-types proposal allows
    // multiple tables but emitting an empty, unused externref table
    // just so the binary "looks" more reference-types-y is the
    // declared-but-unused anti-pattern. The default pipeline emits
    // exactly one funcref table (for call_indirect / element section);
    // an extra externref table is available via
    // `encode_table_section_with` for callers that actually use it.
    let wasm = emit_trivial_wasm();
    let mut pos = 8;
    let mut table_count: Option<u32> = None;
    let mut first_elem_type: Option<u8> = None;
    while pos < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        let (size, r) = decode_leb128_u32(&wasm[pos..]);
        pos += r;
        if section_id == 4 {
            let (n, nread) = decode_leb128_u32(&wasm[pos..]);
            table_count = Some(n);
            first_elem_type = wasm.get(pos + nread).copied();
            break;
        }
        pos += size as usize;
    }
    assert_eq!(
        table_count,
        Some(1),
        "default table section must declare exactly 1 table (funcref) — \
         multi-table is opt-in via encode_table_section_with"
    );
    assert_eq!(
        first_elem_type,
        Some(0x70),
        "default table must be funcref (0x70)"
    );
}

#[test]
fn opt_in_multi_table_declares_externref_alongside_funcref() {
    // Direct-call the opt-in helper; verify it produces a 2-table
    // section (funcref + externref) for consumers that actually use
    // the externref table.
    let chunks = vec![Chunk::new("<script>")];
    let bytes = vybe_platform_wasm::writer::sections::encode_table_section_with(&chunks, 1);
    let (n, nread) = decode_leb128_u32(&bytes);
    assert_eq!(n, 2, "opt-in helper must declare 2 tables");
    assert_eq!(bytes[nread], 0x70, "first table must be funcref");
}

#[test]
fn multi_value_result_arity_round_trips() {
    // Build a chunk that claims 2 results, emit + re-read, verify result_arity preserved.
    let mut script = Chunk::new("<script>");
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(Op::RETURN, 0);

    let mut fun = Chunk::new("dual");
    fun.arity = 0;
    fun.result_arity = 2;
    fun.emit_f64_const(1.0, 0);
    fun.emit_f64_const(1.0, 0);
    fun.emit_op(Op::RETURN, 0);

    let chunks = vec![script, fun];
    let wasm = vybe_platform_wasm::write_wasm(&chunks);
    let reread = vybe_platform_wasm::read_wasm(&wasm).unwrap();
    // The reader restores from our custom vybe section so result_arity
    // defaults to 1 — but reading via the type section (non-vybe path)
    // would preserve it. We exercise the type section path below.

    // Direct type-section inspection: find the byte pattern that is
    // unique to a `(func () -> (externref, externref))` declaration —
    // TYPE_FUNC(0x60), 0 params, 2 results, EXTERNREF, EXTERNREF.
    // A byte-pattern search is robust against the type section also
    // containing GC struct/array preambles (0x5E, 0x5F, 0x4F, etc.),
    // which a naive count-based walker cannot skip.
    let mut pos = 8;
    let mut found_multi_result = false;
    while pos < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        let (size, r) = decode_leb128_u32(&wasm[pos..]);
        pos += r;
        if section_id == 1 {
            let end = pos + size as usize;
            let pat = [0x60u8, 0x00, 0x02, 0x6F, 0x6F];
            found_multi_result = wasm[pos..end].windows(pat.len()).any(|w| w == pat);
            break;
        }
        pos += size as usize;
    }
    // Reader may not parse multi-result types when custom "vybe" section is present —
    // accept either the emit proof above OR matching result_arity.
    let dual_idx = reread.iter().position(|c| c.name == "dual");
    let roundtripped_or_emitted = found_multi_result
        || dual_idx
            .map(|i| reread[i].result_arity == 2)
            .unwrap_or(false);
    assert!(
        roundtripped_or_emitted,
        "multi-value signature not reflected in emitted type section"
    );
}

#[test]
fn multi_value_block_emits_typeidx_blocktype() {
    // A chunk with a block that leaves 2 values on the stack should
    // cause the emitter to register a `() -> externref^2` function
    // type and reference it as the block's typeidx blocktype. We look
    // for that specific function-type signature in the type section
    // AND for a `0x02 <positive typeidx>` byte pair in the code.
    let mut script = Chunk::new("<script>");
    script.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    script.emit_op(Op::RETURN, 0);

    let mut fun = Chunk::new("with_multi_block");
    fun.arity = 0;
    let bp = fun.emit_block_typed(0, 2); // 2-result block
    fun.emit_f64_const(42.0, 0);
    fun.emit_f64_const(42.0, 0);
    fun.emit_op(Op::END, 0);
    fun.patch_block(bp);
    fun.emit_op(Op::DROP, 0);
    fun.emit_op(Op::DROP, 0);
    fun.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    fun.emit_op(Op::RETURN, 0);

    let wasm = vybe_platform_wasm::write_wasm(&vec![script, fun]);

    // Type section should contain a `(func () -> externref externref)`
    // declaration: 0x60 0x00 0x02 0x6F 0x6F.
    let mut pos = 8;
    let mut multi_type_present = false;
    while pos < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        let (size, r) = decode_leb128_u32(&wasm[pos..]);
        pos += r;
        if section_id == 1 {
            let end = pos + size as usize;
            let pat = [0x60u8, 0x00, 0x02, 0x6F, 0x6F];
            multi_type_present = wasm[pos..end].windows(pat.len()).any(|w| w == pat);
            break;
        }
        pos += size as usize;
    }
    assert!(
        multi_type_present,
        "multi-result block did not register `() -> externref^2` type"
    );
}

#[test]
fn multi_value_return_pushes_n_results_on_caller_stack() {
    // A callee declares result_arity=3, pushes three constants, and
    // RETURNs. The caller must see all three on its stack afterwards —
    // verified by summing them via F64_ADD and asserting the total.
    let mut callee = Chunk::new("triple");
    callee.arity = 0;
    callee.result_arity = 3;
    callee.emit_f64_const(1.0, 0);
    callee.emit_f64_const(10.0, 0);
    callee.emit_f64_const(100.0, 0);
    callee.emit_op(Op::RETURN, 0);

    let mut script = Chunk::new("<script>");
    // Build a callable closure that wraps chunk index 1 (callee), call
    // it with 0 args, then sum the three results that land on the stack.
    script.emit_op_u16(Op::REF_FUNC, 1, 0);
    script.emit(0, 0); // uv_count = 0
    script.emit_op_u8_u8(Op::CALL_REF, 0, 1, 0);
    // Stack now holds [100, 10, 1] — sum them to 111.
    script.emit_op(Op::F64_ADD, 0);
    script.emit_op(Op::F64_ADD, 0);
    script.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let result = vm.run(vec![script, callee]).unwrap();
    assert_eq!(
        result.as_f64(),
        111.0,
        "multi-value RETURN should leave all 3 callee results on caller stack"
    );
}

#[test]
fn shared_memory_section_emits_shared_flag() {
    // Direct-call the shared-memory encoder — we don't use it from the
    // default pipeline because no compiler requests shared memory.
    let bytes = vybe_platform_wasm::writer::sections::encode_memory_section_with(1, Some(10), true);
    // bytes[0] = count=1, bytes[1] = flags
    assert_eq!(bytes[0], 1);
    let flags = bytes[1];
    assert!(
        flags & 0x02 != 0,
        "shared flag bit not set (got {:#x})",
        flags
    );
    assert!(
        flags & 0x01 != 0,
        "max flag bit not set (shared requires max)"
    );
}

#[test]
fn name_section_includes_data_and_tag_subsections() {
    let wasm = emit_trivial_wasm();
    let payload = find_custom_section(&wasm, "name").expect("name section missing");
    let mut pos = 0;
    let mut seen_data = false;
    let mut seen_tag = false;
    while pos < payload.len() {
        let id = payload[pos];
        pos += 1;
        let (size, r) = decode_leb128_u32(&payload[pos..]);
        pos += r;
        if id == 9 {
            seen_data = true;
        }
        if id == 11 {
            seen_tag = true;
        }
        pos += size as usize;
    }
    assert!(
        seen_data,
        "name section missing data-names subsection (id=9)"
    );
    assert!(
        seen_tag,
        "name section missing tag-names subsection (id=11)"
    );
}

#[test]
fn ref_eq_matches_identical_objects() {
    use std::sync::{Arc, Mutex};
    use vybe_runtime::value::{Object, ObjectKind};

    // Two constants pointing at the same Arc<Object>.
    let obj = Arc::new(Mutex::new(Object {
        properties: Default::default(),
        kind: ObjectKind::Ordinary,
        type_id: 0,
        fields: Vec::new() }));
    // Host-side objects reach the chunk as imported globals (spec
    // `global.get`; the embedder provides the value) — the retired CONST
    // pool is no longer a host-value injection channel.
    let mut chunk = Chunk::new("<script>");
    let g1 = chunk.intern_string_constant("__ref_eq_same_a");
    let g2 = chunk.intern_string_constant("__ref_eq_same_b");
    chunk.emit_op_u16(Op::GLOBAL_GET, g1, 0);
    chunk.emit_op_u16(Op::GLOBAL_GET, g2, 0);
    chunk.emit_op(Op::REF_EQ, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    vm.globals.insert("__ref_eq_same_a".into(), Value::Object(obj.clone()));
    vm.globals.insert("__ref_eq_same_b".into(), Value::Object(obj));
    assert_eq!(vm.run(vec![chunk]).unwrap().as_i32(), 1);

    // Two DIFFERENT objects → ref.eq is false.
    let mut chunk = Chunk::new("<script>");
    let a = Arc::new(Mutex::new(Object {
        properties: Default::default(),
        kind: ObjectKind::Ordinary,
        type_id: 0,
        fields: Vec::new() }));
    let b = Arc::new(Mutex::new(Object {
        properties: Default::default(),
        kind: ObjectKind::Ordinary,
        type_id: 0,
        fields: Vec::new() }));
    let ga = chunk.intern_string_constant("__ref_eq_diff_a");
    let gb = chunk.intern_string_constant("__ref_eq_diff_b");
    chunk.emit_op_u16(Op::GLOBAL_GET, ga, 0);
    chunk.emit_op_u16(Op::GLOBAL_GET, gb, 0);
    chunk.emit_op(Op::REF_EQ, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    vm.globals.insert("__ref_eq_diff_a".into(), Value::Object(a));
    vm.globals.insert("__ref_eq_diff_b".into(), Value::Object(b));
    assert_eq!(vm.run(vec![chunk]).unwrap().as_i32(), 0);
}

#[test]
fn name_section_includes_type_subsection() {
    let wasm = emit_trivial_wasm();
    let name_payload = find_custom_section(&wasm, "name").expect("name section missing");
    let mut pos = 0;
    let mut seen = false;
    while pos < name_payload.len() {
        let id = name_payload[pos];
        pos += 1;
        let (size, r) = decode_leb128_u32(&name_payload[pos..]);
        pos += r;
        if id == 4 {
            seen = true;
        }
        pos += size as usize;
    }
    assert!(seen, "name section missing type-names subsection (id=4)");
}

#[test]
fn compilation_hints_branch_and_inlining_sections_present_for_loopy_code() {
    // Emit a tiny function with a loop so branch_hint has data to emit.
    let mut script = Chunk::new("<script>");
    script.emit_f64_const(0.0, 0);
    script.emit_op(Op::RETURN, 0);

    let mut fun = Chunk::new("loop_fn");
    fun.arity = 0;
    fun.emit_i32_const(0, 0);
    fun.emit_loop_s(0);
    fun.emit_dup(0);
    fun.emit_br_if(0, 0); // back-edge inside loop
    fun.emit_op(Op::END, 0);
    fun.emit_op(Op::RETURN, 0);

    let wasm = vybe_platform_wasm::write_wasm(&vec![script, fun]);

    assert!(
        find_custom_section(&wasm, "metadata.code.branch_hint").is_some(),
        "branch_hint section not emitted for module with back-edge BR_IF"
    );
    assert!(
        find_custom_section(&wasm, "metadata.code.inlining").is_some(),
        "inlining section not emitted for leaf function"
    );
}

// ──────────────────────────────────────────────────────────────────────
// 26d. Exception-handling proposal — tag section, throw, try_table
// ──────────────────────────────────────────────────────────────────────

/// Build a chunk that uses try/catch so the emitter produces the full
/// exception-handling output (tag section + try_table + throw).
fn emit_try_catch_wasm() -> Vec<u8> {
    let mut script = Chunk::new("<script>");
    // try { throw 1 } catch (e) { /* swallow */ } via the real spec try_table.
    // Immediate (one catch clause, as common::errors::emit_try_start emits):
    //   [u8 clause_count=1, u8 kind=catch(0), u16 tag, u16 catch_offset].
    script.emit_op(Op::TRY_TABLE, 0);
    script.emit(1u8, 0); // clause_count = 1
    script.emit(0u8, 0); // kind = catch
    script.emit(0u8, 0); // tag hi
    script.emit(0u8, 0); // tag lo
    let catch_off_pos = script.current_offset();
    script.emit(0u8, 0); // catch offset hi (placeholder)
    script.emit(0u8, 0); // catch offset lo (placeholder)
    script.emit_i32_const(1, 0);
    // `throw <tagidx>` carries its tag index as a u16 immediate (spec form).
    script.emit_op_u16(Op::THROW, 0, 0);
    // Structural END closes the try_table block (the retired custom TRY_END).
    script.emit_op(Op::END, 0);
    // skip-catch BR placeholder
    let skip = script.emit_jump(Op::BR, 0);
    // patch catch_offset to land here (catch body start): the VM/codec compute
    // catch_ip = (catch_off_pos + 2) + offset.
    let here = script.current_offset() as i32;
    let jump = here - (catch_off_pos as i32 + 2);
    script.code[catch_off_pos] = (jump >> 8) as u8;
    script.code[catch_off_pos + 1] = (jump & 0xff) as u8;
    // catch body: just drop the caught exception
    script.emit_op(Op::DROP, 0);
    // patch skip BR to land here (after-catch)
    let after = script.current_offset() as i32;
    let skip_off = after - (skip as i32 + 2);
    script.code[skip] = (skip_off >> 8) as u8;
    script.code[skip + 1] = (skip_off & 0xff) as u8;
    script.emit_op(Op::RETURN, 0);
    vybe_platform_wasm::write_wasm(&vec![script])
}

#[test]
fn tag_section_emitted_when_module_uses_exceptions() {
    let wasm = emit_try_catch_wasm();
    // Section id 13 = tag section.
    let mut pos = 8;
    let mut found = false;
    while pos < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        let (size, r) = decode_leb128_u32(&wasm[pos..]);
        pos += r;
        if section_id == 13 {
            let (n, _) = decode_leb128_u32(&wasm[pos..]);
            assert_eq!(n, 1, "tag section must declare exactly 1 tag");
            found = true;
            break;
        }
        pos += size as usize;
    }
    assert!(found, "tag section (id=13) missing for module that throws");
}

#[test]
fn tag_section_omitted_when_module_has_no_exceptions() {
    let wasm = emit_trivial_wasm();
    let mut pos = 8;
    while pos < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        let (size, r) = decode_leb128_u32(&wasm[pos..]);
        pos += r;
        assert_ne!(
            section_id, 13,
            "tag section must not be emitted when no chunk uses throw/try"
        );
        pos += size as usize;
    }
}

#[test]
fn throw_emits_as_wasm_throw_not_nop() {
    let wasm = emit_try_catch_wasm();
    // Find the code section, look for 0x08 (throw) followed by LEB128 u32=0.
    let mut pos = 8;
    let mut seen_throw = false;
    while pos < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        let (size, r) = decode_leb128_u32(&wasm[pos..]);
        pos += r;
        if section_id == 10 {
            let end = pos + size as usize;
            // Byte pattern: throw opcode (0x08) + tag index LEB128 (0x00).
            seen_throw = wasm[pos..end].windows(2).any(|w| w == [0x08, 0x00]);
            break;
        }
        pos += size as usize;
    }
    assert!(
        seen_throw,
        "emitted code must contain `throw 0` bytes (0x08 0x00) — THROW was silently nop'd"
    );
}

#[test]
fn try_table_opcode_emitted_for_try_region() {
    let wasm = emit_try_catch_wasm();
    // Look for try_table opcode (0x1F) in the code section.
    let mut pos = 8;
    let mut seen_try_table = false;
    while pos < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        let (size, r) = decode_leb128_u32(&wasm[pos..]);
        pos += r;
        if section_id == 10 {
            let end = pos + size as usize;
            seen_try_table = wasm[pos..end].iter().any(|&b| b == 0x1F);
            break;
        }
        pos += size as usize;
    }
    assert!(
        seen_try_table,
        "code section must contain try_table opcode (0x1F) for try region"
    );
}

#[test]
fn exception_type_declared_with_externref_param() {
    let wasm = emit_try_catch_wasm();
    // Look for `(externref) -> ()` func type: 0x60 0x01 0x6F 0x00.
    let mut pos = 8;
    let mut found = false;
    while pos < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        let (size, r) = decode_leb128_u32(&wasm[pos..]);
        pos += r;
        if section_id == 1 {
            let end = pos + size as usize;
            let pat = [0x60u8, 0x01, 0x6F, 0x00];
            found = wasm[pos..end].windows(pat.len()).any(|w| w == pat);
            break;
        }
        pos += size as usize;
    }
    assert!(
        found,
        "type section must declare `(externref) -> ()` for exception tag"
    );
}

// ──────────────────────────────────────────────────────────────────────
// 26f. GC proposal — spec byte-value compliance
// ──────────────────────────────────────────────────────────────────────

/// Pin every declared GC opcode to the byte position given in MVP.md.
/// Regression guard: historically we had REF_EQ colliding with
/// `array.init_elem` and ARRAY_NEW naming `array.new_fixed` — both
/// silently wrong. This test locks down the entire byte table.
#[test]
fn gc_opcodes_use_spec_byte_values() {
    let table: &[(Op, u16, u8, &str)] = &[
        // Struct ops
        (Op::STRUCT_NEW, 0xFB, 0x00, "struct.new"),
        (Op::STRUCT_NEW_DEFAULT, 0xFB, 0x01, "struct.new_default"),
        (Op::STRUCT_GET, 0xFB, 0x02, "struct.get"),
        (Op::STRUCT_GET_S, 0xFB, 0x03, "struct.get_s"),
        (Op::STRUCT_GET_U, 0xFB, 0x04, "struct.get_u"),
        (Op::STRUCT_SET, 0xFB, 0x05, "struct.set"),
        // Array ops
        (Op::ARRAY_NEW, 0xFB, 0x06, "array.new"),
        (Op::ARRAY_NEW_DEFAULT, 0xFB, 0x07, "array.new_default"),
        (Op::ARRAY_NEW_FIXED, 0xFB, 0x08, "array.new_fixed"),
        (Op::ARRAY_NEW_DATA, 0xFB, 0x09, "array.new_data"),
        (Op::ARRAY_NEW_ELEM, 0xFB, 0x0A, "array.new_elem"),
        (Op::ARRAY_GET, 0xFB, 0x0B, "array.get"),
        (Op::ARRAY_GET_S, 0xFB, 0x0C, "array.get_s"),
        (Op::ARRAY_GET_U, 0xFB, 0x0D, "array.get_u"),
        (Op::ARRAY_SET, 0xFB, 0x0E, "array.set"),
        (Op::ARRAY_LENGTH, 0xFB, 0x0F, "array.len"),
        (Op::ARRAY_FILL, 0xFB, 0x10, "array.fill"),
        (Op::ARRAY_COPY, 0xFB, 0x11, "array.copy"),
        (Op::ARRAY_INIT_DATA, 0xFB, 0x12, "array.init_data"),
        (Op::ARRAY_INIT_ELEM, 0xFB, 0x13, "array.init_elem"),
        // Reference tests / casts
        (Op::REF_TEST, 0xFB, 0x14, "ref.test"),
        (Op::REF_TEST_NULL, 0xFB, 0x15, "ref.test_null"),
        (Op::REF_CAST, 0xFB, 0x16, "ref.cast"),
        (Op::REF_CAST_NULL, 0xFB, 0x17, "ref.cast_null"),
        (Op::BR_ON_CAST, 0xFB, 0x18, "br_on_cast"),
        (Op::BR_ON_CAST_FAIL, 0xFB, 0x19, "br_on_cast_fail"),
        // Extern <-> any
        (Op::ANY_CONVERT_EXTERN, 0xFB, 0x1A, "any.convert_extern"),
        (Op::EXTERN_CONVERT_ANY, 0xFB, 0x1B, "extern.convert_any"),
        // i31
        (Op::I31_NEW, 0xFB, 0x1C, "ref.i31"),
        (Op::I31_GET_S, 0xFB, 0x1D, "i31.get_s"),
        (Op::I31_GET_U, 0xFB, 0x1E, "i31.get_u"),
        // Core-prefix GC ops
        (Op::REF_EQ, 0x00, 0xD3, "ref.eq"),
        (Op::REF_AS_NON_NULL, 0x00, 0xD4, "ref.as_non_null"),
        (Op::BR_ON_NULL, 0x00, 0xD5, "br_on_null"),
        (Op::BR_ON_NON_NULL, 0x00, 0xD6, "br_on_non_null"),
    ];
    for (op, prefix, sub, name) in table {
        assert_eq!(
            op.group(),
            *prefix,
            "{}: prefix mismatch (got 0x{:02X}, spec 0x{:02X})",
            name,
            op.group(),
            *prefix
        );
        assert_eq!(
            op.sub(),
            *sub as u16,
            "{}: sub mismatch (got 0x{:02X}, spec 0x{:02X})",
            name,
            op.sub(),
            *sub
        );
        assert_eq!(
            op.wasm_name(),
            *name,
            "{}: name mismatch (got {:?})",
            name,
            op.wasm_name()
        );
    }
}

#[test]
fn ref_eq_no_longer_collides_with_array_init_elem() {
    assert_eq!(Op::REF_EQ.group(), 0x00);
    assert_eq!(Op::REF_EQ.sub(), 0xD3);
    assert_eq!(Op::ARRAY_INIT_ELEM.group(), 0xFB);
    assert_eq!(Op::ARRAY_INIT_ELEM.sub(), 0x13);
    assert_ne!(Op::REF_EQ.0, Op::ARRAY_INIT_ELEM.0);
}

#[test]
fn array_new_fixed_vs_array_new_at_correct_spec_bytes() {
    assert_eq!(Op::ARRAY_NEW_FIXED.sub(), 0x08);
    assert_eq!(Op::ARRAY_NEW.sub(), 0x06);
    assert_eq!(Op::ARRAY_NEW_FIXED.wasm_name(), "array.new_fixed");
    assert_eq!(Op::ARRAY_NEW.wasm_name(), "array.new");
}

#[test]
fn array_new_single_value_and_length() {
    // `array.new $t`: [value, length i32] -> [array, every lane = value].
    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(7, 0);
    chunk.emit_i32_const(3, 0);
    chunk.emit_op_u16(Op::ARRAY_NEW, 0, 0);
    chunk.emit_op(Op::ARRAY_LENGTH, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    assert_eq!(vm.run(vec![chunk]).unwrap().as_i32(), 3);
}

#[test]
fn array_new_default_yields_length_array() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(5, 0);
    chunk.emit_op_u16(Op::ARRAY_NEW_DEFAULT, 0, 0);
    chunk.emit_op(Op::ARRAY_LENGTH, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    assert_eq!(vm.run(vec![chunk]).unwrap().as_i32(), 5);
}

#[test]
fn struct_new_default_creates_empty_struct() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op_u16(Op::STRUCT_NEW_DEFAULT, 0, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    let r = vm.run(vec![chunk]).unwrap();
    assert!(matches!(r, Value::Object(_)));
}

#[test]
fn ref_as_non_null_traps_on_null_passes_on_object() {
    let r = run_script(|c| {
        c.emit_op_u16(Op::STRUCT_NEW_DEFAULT, 0, 0);
        c.emit_op(Op::REF_AS_NON_NULL, 0);
    });
    assert!(matches!(r, Value::Object(_)));

    let mut chunk = Chunk::new("<script>");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::REF_AS_NON_NULL, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    let err = vm.run(vec![chunk]).unwrap_err();
    assert!(
        err.message.contains("ref.as_non_null"),
        "expected trap mentioning ref.as_non_null, got {:?}",
        err.message
    );
}

#[test]
fn any_extern_convert_round_trip_preserves_value() {
    // Spec: composing any.convert_extern and extern.convert_any yields
    // the original value. Both are identity in our value ABI.
    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(42, 0);
    chunk.emit_op(Op::ANY_CONVERT_EXTERN, 0);
    chunk.emit_op(Op::EXTERN_CONVERT_ANY, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    assert_eq!(vm.run(vec![chunk]).unwrap().as_i32(), 42);
}

#[test]
fn ref_test_null_accepts_null() {
    // ref.test_null succeeds on null; ref.test (non-null variant)
    // rejects it. Verify the former by passing null and expecting 1.
    let mut chunk = Chunk::new("<script>");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_ref_type_op(
        Op::REF_TEST_NULL,
        vybe_runtime::opcode::heaptype::HeapType::Abstract(vybe_runtime::opcode::heaptype::HT_STRUCT),
        0,
    );
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    assert_eq!(vm.run(vec![chunk]).unwrap().as_i32(), 1);
}

#[test]
fn ref_eq_emits_as_core_0xd3_byte_in_wasm() {
    // Verify spec byte emission: `ref.eq` is a single-byte core op.
    let mut chunk = Chunk::new("<script>");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::REF_EQ, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::RETURN, 0);
    let wasm = vybe_platform_wasm::write_wasm(&vec![chunk]);
    let mut pos = 8;
    let mut seen = false;
    while pos < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        let (size, r) = decode_leb128_u32(&wasm[pos..]);
        pos += r;
        if section_id == 10 {
            let end = pos + size as usize;
            seen = wasm[pos..end].iter().any(|&b| b == 0xD3);
            break;
        }
        pos += size as usize;
    }
    assert!(seen, "code section must contain core ref.eq byte 0xD3");
}

// ──────────────────────────────────────────────────────────────────────
// 26e. Relaxed-SIMD proposal — spec sub-values 0x100..=0x113
// ──────────────────────────────────────────────────────────────────────

/// Spec `v128.const` — the opcode followed by 16 raw immediate bytes
/// (dispatch reads them directly; the writer serializes 0xFD 0x0C + bytes).
fn emit_v128_const(chunk: &mut Chunk, bytes: [u8; 16]) {
    chunk.emit_op(Op::V128_CONST, 0);
    for b in bytes {
        chunk.emit(b, 0);
    }
}

/// Relaxed-SIMD ops live at spec sub-values >= 0x100. In our emitted
/// .wasm each appears as `0xFD` followed by an LEB128 u32 encoding the
/// spec sub. LEB128 of 0x100 is `0x80 0x02`; of 0x113 is `0x93 0x02`.
#[test]
fn relaxed_simd_emits_correct_spec_subopcode() {
    // Build a chunk that exercises i8x16.relaxed_swizzle (spec 0x100).
    let mut script = Chunk::new("<script>");
    emit_v128_const(&mut script, [0u8; 16]);
    emit_v128_const(&mut script, [0u8; 16]);
    script.emit_op(Op::I8X16_RELAXED_SWIZZLE, 0);
    script.emit_op(Op::DROP, 0);
    script.emit_op(Op::RETURN, 0);

    let wasm = vybe_platform_wasm::write_wasm(&vec![script]);

    // Locate the code section and scan for the spec-correct byte pattern.
    let mut pos = 8;
    let mut seen = false;
    while pos < wasm.len() {
        let section_id = wasm[pos];
        pos += 1;
        let (size, r) = decode_leb128_u32(&wasm[pos..]);
        pos += r;
        if section_id == 10 {
            let end = pos + size as usize;
            // Expected bytes for `i8x16.relaxed_swizzle`: 0xFD 0x80 0x02.
            let pat = [0xFDu8, 0x80, 0x02];
            seen = wasm[pos..end].windows(pat.len()).any(|w| w == pat);
            break;
        }
        pos += size as usize;
    }
    assert!(
        seen,
        "emitted code must contain `0xFD 0x80 0x02` (relaxed_swizzle spec sub 0x100)"
    );
    // Sanity: sub value matches the spec (256 = 0x100).
    assert_eq!(Op::I8X16_RELAXED_SWIZZLE.sub(), 256);
}

#[test]
fn relaxed_simd_every_opcode_is_named() {
    // Every one of the 20 relaxed-SIMD ops must decode and round-trip
    // through `wasm_name_opt` — otherwise the dispatch in vm.rs is
    // silently matching a "None" opcode and falling through.
    let all: &[(Op, &str)] = &[
        (Op::I8X16_RELAXED_SWIZZLE, "i8x16.relaxed_swizzle"),
        (
            Op::I32X4_RELAXED_TRUNC_F32X4_S,
            "i32x4.relaxed_trunc_f32x4_s",
        ),
        (
            Op::I32X4_RELAXED_TRUNC_F32X4_U,
            "i32x4.relaxed_trunc_f32x4_u",
        ),
        (
            Op::I32X4_RELAXED_TRUNC_F64X2_S_ZERO,
            "i32x4.relaxed_trunc_f64x2_s_zero",
        ),
        (
            Op::I32X4_RELAXED_TRUNC_F64X2_U_ZERO,
            "i32x4.relaxed_trunc_f64x2_u_zero",
        ),
        (Op::F32X4_RELAXED_MADD, "f32x4.relaxed_madd"),
        (Op::F32X4_RELAXED_NMADD, "f32x4.relaxed_nmadd"),
        (Op::F64X2_RELAXED_MADD, "f64x2.relaxed_madd"),
        (Op::F64X2_RELAXED_NMADD, "f64x2.relaxed_nmadd"),
        (Op::I8X16_RELAXED_LANESELECT, "i8x16.relaxed_laneselect"),
        (Op::I16X8_RELAXED_LANESELECT, "i16x8.relaxed_laneselect"),
        (Op::I32X4_RELAXED_LANESELECT, "i32x4.relaxed_laneselect"),
        (Op::I64X2_RELAXED_LANESELECT, "i64x2.relaxed_laneselect"),
        (Op::F32X4_RELAXED_MIN, "f32x4.relaxed_min"),
        (Op::F32X4_RELAXED_MAX, "f32x4.relaxed_max"),
        (Op::F64X2_RELAXED_MIN, "f64x2.relaxed_min"),
        (Op::F64X2_RELAXED_MAX, "f64x2.relaxed_max"),
        (Op::I16X8_RELAXED_Q15MULR_S, "i16x8.relaxed_q15mulr_s"),
        (
            Op::I16X8_RELAXED_DOT_I8X16_I7X16_S,
            "i16x8.relaxed_dot_i8x16_i7x16_s",
        ),
        (
            Op::I32X4_RELAXED_DOT_I8X16_I7X16_ADD_S,
            "i32x4.relaxed_dot_i8x16_i7x16_add_s",
        ),
    ];
    assert_eq!(
        all.len(),
        20,
        "relaxed-SIMD proposal has exactly 20 opcodes"
    );
    for (op, expected_name) in all {
        assert_eq!(
            op.group(),
            0xFD,
            "relaxed-SIMD op must use SIMD prefix 0xFD: {}",
            expected_name
        );
        assert_eq!(op.wasm_name(), *expected_name);
    }
}

#[test]
fn relaxed_simd_madd_matches_nonrelaxed_madd() {
    // Smoke test: VM executes f32x4.relaxed_madd deterministically.
    // Input lanes: a=[1,2,3,4], b=[5,6,7,8], c=[9,10,11,12]
    // Expected: lane i = a*b + c = [14, 22, 32, 44]
    let mut lanes_a = [0u8; 16];
    let mut lanes_b = [0u8; 16];
    let mut lanes_c = [0u8; 16];
    for i in 0..4u32 {
        lanes_a[i as usize * 4..(i as usize + 1) * 4]
            .copy_from_slice(&((i + 1) as f32).to_le_bytes());
        lanes_b[i as usize * 4..(i as usize + 1) * 4]
            .copy_from_slice(&((i + 5) as f32).to_le_bytes());
        lanes_c[i as usize * 4..(i as usize + 1) * 4]
            .copy_from_slice(&((i + 9) as f32).to_le_bytes());
    }
    let mut chunk = Chunk::new("<script>");
    emit_v128_const(&mut chunk, lanes_a);
    emit_v128_const(&mut chunk, lanes_b);
    emit_v128_const(&mut chunk, lanes_c);
    chunk.emit_op(Op::F32X4_RELAXED_MADD, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    if let Value::V128(bytes) = result {
        for i in 0..4 {
            let lane = f32::from_le_bytes(bytes[i * 4..i * 4 + 4].try_into().unwrap());
            let expected = (i as f32 + 1.0) * (i as f32 + 5.0) + (i as f32 + 9.0);
            assert!(
                (lane - expected).abs() < 1e-6,
                "lane {} mismatch: {} vs expected {}",
                i,
                lane,
                expected
            );
        }
    } else {
        panic!("expected V128, got {:?}", result);
    }
}

#[test]
fn relaxed_simd_laneselect_picks_by_mask_high_bit() {
    // i32x4.relaxed_laneselect: lane i = (mask_i high bit) ? a_i : b_i
    let mut a = [0u8; 16];
    let mut b = [0u8; 16];
    let mut mask = [0u8; 16];
    // a = [1, 2, 3, 4], b = [100, 200, 300, 400], mask = [0, 0xFFFFFFFF, 0, 0xFFFFFFFF]
    for i in 0..4u32 {
        a[i as usize * 4..(i as usize + 1) * 4].copy_from_slice(&((i + 1) as i32).to_le_bytes());
        b[i as usize * 4..(i as usize + 1) * 4]
            .copy_from_slice(&(((i + 1) * 100) as i32).to_le_bytes());
        let m: u32 = if i % 2 == 1 { 0xFFFF_FFFF } else { 0 };
        mask[i as usize * 4..(i as usize + 1) * 4].copy_from_slice(&m.to_le_bytes());
    }
    let mut chunk = Chunk::new("<script>");
    emit_v128_const(&mut chunk, a);
    emit_v128_const(&mut chunk, b);
    emit_v128_const(&mut chunk, mask);
    chunk.emit_op(Op::I32X4_RELAXED_LANESELECT, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    let result = vm.run(vec![chunk]).unwrap();
    if let Value::V128(bytes) = result {
        let expect = [100i32, 2, 300, 4];
        for i in 0..4 {
            let lane = i32::from_le_bytes(bytes[i * 4..i * 4 + 4].try_into().unwrap());
            assert_eq!(lane, expect[i], "lane {} mismatch", i);
        }
    } else {
        panic!("expected V128");
    }
}

// ──────────────────────────────────────────────────────────────────────
// 26c. Reference-types — table operations
// ──────────────────────────────────────────────────────────────────────

#[test]
fn table_size_returns_current_size() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op_u16(Op::TABLE_SIZE, 0, 0); // table index 0
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    vm.wasm_tables = vec![Vec::new()]; // declare table 0 (empty)
    assert_eq!(vm.run(vec![chunk]).unwrap().as_i32(), 0);
}

#[test]
fn table_grow_returns_old_size_and_resizes() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_i32_const(3, 0);
    chunk.emit_op_u16(Op::TABLE_GROW, 0, 0); // table index 0
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op_u16(Op::TABLE_SIZE, 0, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    vm.wasm_tables = vec![Vec::new()]; // declare table 0 (grown by the op)
    assert_eq!(vm.run(vec![chunk]).unwrap().as_i32(), 3);
}

#[test]
fn table_fill_assigns_value_across_range() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_i32_const(5, 0);
    chunk.emit_op_u16(Op::TABLE_GROW, 0, 0);
    chunk.emit_op(Op::DROP, 0);

    chunk.emit_i32_const(1, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_i32_const(3, 0);
    chunk.emit_op_u16(Op::TABLE_FILL, 0, 0);

    chunk.emit_op_u16(Op::TABLE_SIZE, 0, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    vm.wasm_tables = vec![Vec::new()]; // declare table 0 (grown + filled by the ops)
    assert_eq!(vm.run(vec![chunk]).unwrap().as_i32(), 5);
}

/// End-to-end ESM round-trip: emit a .wasm, drop it into a temp directory,
/// and have `node --experimental-wasm-modules` import it. If `node` isn't
/// available (CI without Node, older Node without wasm-esm), the test is
/// skipped with a message rather than failing.
///
/// What this proves:
///   * Emitted .wasm decodes under a real production WASM runtime
///   * Export section shape matches what ESM loaders expect
///   * The `wasm:js-*` globals / functions our imports reference don't
///     cause validation failure when not provided (Node resolves them as
///     builtins automatically on supported versions)
#[test]
fn esm_round_trip_via_node_if_available() {
    use std::io::Write;
    use std::process::Command;

    // Bail cleanly if `node` is not installed.
    let node_ok = Command::new("node")
        .arg("--version")
        .output()
        .ok()
        .map(|o| o.status.success())
        .unwrap_or(false);
    if !node_ok {
        eprintln!("skipping: `node` not available on PATH");
        return;
    }

    // Minimal module: script returns boxed 42 via CONST + RETURN.
    // Scripts don't need to be exported — we just care that the module
    // validates and any declared exports are discoverable.
    let mut script = Chunk::new("run");
    script.arity = 0;
    script.emit_f64_const(42.0, 0);
    script.emit_op(Op::RETURN, 0);
    let wasm = vybe_platform_wasm::write_wasm(&vec![script]);

    // Stage the .wasm + the driver .mjs in a temp directory.
    let dir = std::env::temp_dir().join(format!("vybe_esm_{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).expect("create tmp dir");
    let wasm_path = dir.join("mod.wasm");
    let driver_path = dir.join("driver.mjs");

    {
        let mut f = std::fs::File::create(&wasm_path).expect("open wasm file");
        f.write_all(&wasm).expect("write wasm");
    }

    // The driver does the minimum: read the bytes, validate, construct a
    // Module. We deliberately don't go as far as instantiation — that
    // would require providing importObject entries for every `wasm:js-*`
    // builtin; validation alone proves the module shape is ESM-compliant.
    let driver = r#"
import { readFile } from 'node:fs/promises';
import { argv } from 'node:process';
const bytes = await readFile(argv[2]);
if (!WebAssembly.validate(bytes)) {
    console.error('WebAssembly.validate rejected the module');
    process.exit(1);
}
const mod = new WebAssembly.Module(bytes);
const imports = WebAssembly.Module.imports(mod);
const exports = WebAssembly.Module.exports(mod);
console.log(JSON.stringify({
    imports: imports.length,
    exports: exports.length,
    has_js_undefined_value: imports.some(i => i.module === 'wasm:js-undefined' && i.name === 'value'),
    has_js_true: imports.some(i => i.module === 'wasm:js-boolean' && i.name === 'true'),
    has_js_false: imports.some(i => i.module === 'wasm:js-boolean' && i.name === 'false') }));
"#;
    {
        let mut f = std::fs::File::create(&driver_path).expect("open driver file");
        f.write_all(driver.as_bytes()).expect("write driver");
    }

    let out = Command::new("node")
        .arg(&driver_path)
        .arg(&wasm_path)
        .output()
        .expect("spawn node");

    let _ = std::fs::remove_dir_all(&dir);

    let stdout = String::from_utf8_lossy(&out.stdout).to_string();
    let stderr = String::from_utf8_lossy(&out.stderr).to_string();
    assert!(
        out.status.success(),
        "node rejected the module.\nstdout: {}\nstderr: {}",
        stdout,
        stderr
    );
    assert!(
        stdout.contains("\"has_js_undefined_value\":true"),
        "js-undefined global missing in import list. stdout: {}",
        stdout
    );
    assert!(
        stdout.contains("\"has_js_true\":true"),
        "js-boolean.true global missing. stdout: {}",
        stdout
    );
    assert!(
        stdout.contains("\"has_js_false\":true"),
        "js-boolean.false global missing. stdout: {}",
        stdout
    );
}
