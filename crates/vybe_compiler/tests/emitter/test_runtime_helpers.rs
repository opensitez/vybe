use std::sync::Arc;
/// Tests for the runtime helpers — pure bytecode implementations that run on any runtime.
use vybe_bytecode::{Chunk, Op, VM, Value};
use vybe_compiler::emitter::runtime_helpers::build_runtime_helpers;

/// Link a pre-built helper into a VM and run the script chunk.
/// Callers build the helper themselves (so they can read out chunk indices
/// before emitting the script body that references them).
fn run_with_prebuilt(script: Chunk, helper_chunks: Vec<Chunk>) -> Value {
    let mut vm = VM::new();
    vybe_host::register_all(&mut vm);
    let mut all_chunks = vec![script];
    all_chunks.extend(helper_chunks);
    vm.run(all_chunks).unwrap()
}

// ── reversed ────────────────────────────────────────────────

#[test]
fn helper_reversed() {
    let mut script = Chunk::new("<script>");
    script.local_count = 3;
    let helpers = build_runtime_helpers(&mut script);
    let rev_idx = helpers.get("__stdlib_reversed").unwrap() + 1;

    // reversed([1, 2, 3]) → [3, 2, 1]
    let v1 = script.add_constant(Value::I32(1));
    let v2 = script.add_constant(Value::I32(2));
    let v3 = script.add_constant(Value::I32(3));
    script.emit_op_u16(Op::CONST, v1, 0);
    script.emit_op_u16(Op::CONST, v2, 0);
    script.emit_op_u16(Op::CONST, v3, 0);
    script.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, 0);
    script.emit_op_u16(Op::LOCAL_SET, 1, 0);

    script.emit_op_u16(Op::REF_FUNC, rev_idx as u16, 0);
    script.emit(0, 0);
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op_u8(Op::CALL_REF, 1, 0);

    // Get first element — should be 3
    script.emit_i32_const(0, 0);
    script.emit_op(Op::ARRAY_GET, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, helpers.chunks);
    assert_eq!(
        result.as_i32(),
        3,
        "first element of reversed [1,2,3] should be 3"
    );
}

// ── sorted ──────────────────────────────────────────────────

#[test]
fn helper_sorted() {
    let mut script = Chunk::new("<script>");
    script.local_count = 3;
    let helpers = build_runtime_helpers(&mut script);
    let sort_idx = helpers.get("__stdlib_sorted").unwrap() + 1;

    // sorted([3, 1, 2]) → [1, 2, 3]
    let v3 = script.add_constant(Value::I32(3));
    let v1 = script.add_constant(Value::I32(1));
    let v2 = script.add_constant(Value::I32(2));
    script.emit_op_u16(Op::CONST, v3, 0);
    script.emit_op_u16(Op::CONST, v1, 0);
    script.emit_op_u16(Op::CONST, v2, 0);
    script.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, 0);
    script.emit_op_u16(Op::LOCAL_SET, 1, 0);

    script.emit_op_u16(Op::REF_FUNC, sort_idx as u16, 0);
    script.emit(0, 0);
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op_u8(Op::CALL_REF, 1, 0);

    // First element should be 1 (smallest)
    script.emit_i32_const(0, 0);
    script.emit_op(Op::ARRAY_GET, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, helpers.chunks);
    assert_eq!(
        result.as_i32(),
        1,
        "first element of sorted [3,1,2] should be 1"
    );
}

#[test]
fn helper_sorted_preserves_original() {
    let mut script = Chunk::new("<script>");
    script.local_count = 3;
    let helpers = build_runtime_helpers(&mut script);
    let sort_idx = helpers.get("__stdlib_sorted").unwrap() + 1;

    // sorted should not mutate the original
    let v3 = script.add_constant(Value::I32(3));
    let v1 = script.add_constant(Value::I32(1));
    script.emit_op_u16(Op::CONST, v3, 0);
    script.emit_op_u16(Op::CONST, v1, 0);
    script.emit_op_u16(Op::ARRAY_NEW_FIXED, 2, 0);
    script.emit_op_u16(Op::LOCAL_SET, 1, 0);

    // Call sorted
    script.emit_op_u16(Op::REF_FUNC, sort_idx as u16, 0);
    script.emit(0, 0);
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op_u8(Op::CALL_REF, 1, 0);
    script.emit_op(Op::DROP, 0); // discard sorted result

    // Original first element should still be 3
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_i32_const(0, 0);
    script.emit_op(Op::ARRAY_GET, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, helpers.chunks);
    assert_eq!(result.as_i32(), 3, "original array should be unchanged");
}

// ── min / max ───────────────────────────────────────────────

#[test]
fn helper_min() {
    let mut script = Chunk::new("<script>");
    script.local_count = 3;
    let helpers = build_runtime_helpers(&mut script);
    let min_idx = helpers.get("__stdlib_min").unwrap() + 1;

    let v5 = script.add_constant(Value::I32(5));
    let v2 = script.add_constant(Value::I32(2));
    let v8 = script.add_constant(Value::I32(8));
    script.emit_op_u16(Op::CONST, v5, 0);
    script.emit_op_u16(Op::CONST, v2, 0);
    script.emit_op_u16(Op::CONST, v8, 0);
    script.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, 0);
    script.emit_op_u16(Op::LOCAL_SET, 1, 0);

    script.emit_op_u16(Op::REF_FUNC, min_idx as u16, 0);
    script.emit(0, 0);
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op_u8(Op::CALL_REF, 1, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, helpers.chunks);
    assert_eq!(result.as_i32(), 2);
}

#[test]
fn helper_max() {
    let mut script = Chunk::new("<script>");
    script.local_count = 3;
    let helpers = build_runtime_helpers(&mut script);
    let max_idx = helpers.get("__stdlib_max").unwrap() + 1;

    let v5 = script.add_constant(Value::I32(5));
    let v2 = script.add_constant(Value::I32(2));
    let v8 = script.add_constant(Value::I32(8));
    script.emit_op_u16(Op::CONST, v5, 0);
    script.emit_op_u16(Op::CONST, v2, 0);
    script.emit_op_u16(Op::CONST, v8, 0);
    script.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, 0);
    script.emit_op_u16(Op::LOCAL_SET, 1, 0);

    script.emit_op_u16(Op::REF_FUNC, max_idx as u16, 0);
    script.emit(0, 0);
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op_u8(Op::CALL_REF, 1, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, helpers.chunks);
    assert_eq!(result.as_i32(), 8);
}

// ── pow ─────────────────────────────────────────────────────

#[test]
fn helper_pow() {
    let mut script = Chunk::new("<script>");
    script.local_count = 2;
    let helpers = build_runtime_helpers(&mut script);
    let pow_idx = helpers.get("__stdlib_pow").unwrap() + 1;

    // 2^10 = 1024
    script.emit_op_u16(Op::REF_FUNC, pow_idx as u16, 0);
    script.emit(0, 0);
    let base = script.add_constant(Value::F64(2.0));
    let exp = script.add_constant(Value::I32(10));
    script.emit_op_u16(Op::CONST, base, 0);
    script.emit_op_u16(Op::CONST, exp, 0);
    script.emit_op_u8(Op::CALL_REF, 2, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, helpers.chunks);
    assert_eq!(result.as_f64() as i32, 1024);
}

// ── enumerate ───────────────────────────────────────────────

#[test]
fn helper_enumerate() {
    let mut script = Chunk::new("<script>");
    script.local_count = 3;
    let helpers = build_runtime_helpers(&mut script);
    let enum_idx = helpers.get("__stdlib_enumerate").unwrap() + 1;

    // enumerate(["a", "b"]) → [[0,"a"], [1,"b"]]
    let va = script.add_constant(Value::String(Arc::from("a")));
    let vb = script.add_constant(Value::String(Arc::from("b")));
    script.emit_op_u16(Op::CONST, va, 0);
    script.emit_op_u16(Op::CONST, vb, 0);
    script.emit_op_u16(Op::ARRAY_NEW_FIXED, 2, 0);
    script.emit_op_u16(Op::LOCAL_SET, 1, 0);

    script.emit_op_u16(Op::REF_FUNC, enum_idx as u16, 0);
    script.emit(0, 0);
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op_u8(Op::CALL_REF, 1, 0);

    // Result length should be 2
    script.emit_op(Op::ARRAY_LENGTH, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, helpers.chunks);
    assert_eq!(result.as_i32(), 2);
}

// ── helper building ─────────────────────────────────────────

#[test]
fn helper_has_all_functions() {
    let mut script = Chunk::new("<script>");
    let helpers = build_runtime_helpers(&mut script);
    assert!(
        helpers.chunks.len() >= 13,
        "helper should export >= 13 chunks, got {}",
        helpers.chunks.len()
    );
    assert!(helpers.get("__stdlib_sorted").is_some());
    assert!(helpers.get("__stdlib_reversed").is_some());
    assert!(helpers.get("__stdlib_enumerate").is_some());
    assert!(helpers.get("__stdlib_sum").is_some());
    assert!(helpers.get("__stdlib_min").is_some());
    assert!(helpers.get("__stdlib_max").is_some());
    assert!(helpers.get("__stdlib_pow").is_some());
    assert!(helpers.get("__stdlib_slicestep").is_some());
    assert!(helpers.get("__stdlib_dynmul").is_some());
}
