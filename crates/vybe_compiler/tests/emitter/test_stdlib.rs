use std::sync::Arc;
/// Tests for the WASM stdlib — pure bytecode implementations that run on any runtime.
use vybe_bytecode::{Chunk, Op, VM, Value};
use vybe_compiler::emitter::stdlib::build_stdlib;

/// Link a pre-built stdlib into a VM and run the script chunk.
/// Callers build the stdlib themselves (so they can read out chunk indices
/// before emitting the script body that references them).
fn run_with_prebuilt(script: Chunk, stdlib_chunks: Vec<Chunk>) -> Value {
    let mut vm = VM::new();
    vybe_host::register_all(&mut vm);
    let mut all_chunks = vec![script];
    all_chunks.extend(stdlib_chunks);
    vm.run(all_chunks).unwrap()
}

// ── range ───────────────────────────────────────────────────

#[test]
fn stdlib_range_basic() {
    let mut script = Chunk::new("<script>");
    script.local_count = 2;
    let stdlib = build_stdlib(&mut script);
    let range_idx = stdlib.get("__stdlib_range").unwrap();
    let chunk_idx = range_idx + 1; // +1 because script is chunk 0

    // range(0, 5, 1)
    let zero = script.add_constant(Value::I32(0));
    let five = script.add_constant(Value::I32(5));
    let one = script.add_constant(Value::I32(1));
    script.emit_op_u16(Op::REF_FUNC, chunk_idx as u16, 0);
    script.emit(0, 0); // 0 upvalues
    script.emit_op_u16(Op::CONST, zero, 0);
    script.emit_op_u16(Op::CONST, five, 0);
    script.emit_op_u16(Op::CONST, one, 0);
    script.emit_op_u8(Op::CALL_REF, 3, 0);
    // Result should be array [0,1,2,3,4]
    script.emit_op(Op::ARRAY_LENGTH, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, stdlib.chunks);
    assert_eq!(result.as_i32(), 5, "range(0,5,1) should produce 5 elements");
}

#[test]
fn stdlib_range_step() {
    let mut script = Chunk::new("<script>");
    script.local_count = 2;
    let stdlib = build_stdlib(&mut script);
    let range_idx = stdlib.get("__stdlib_range").unwrap() + 1;

    // range(0, 10, 2) → [0,2,4,6,8]
    let zero = script.add_constant(Value::I32(0));
    let ten = script.add_constant(Value::I32(10));
    let two = script.add_constant(Value::I32(2));
    script.emit_op_u16(Op::REF_FUNC, range_idx as u16, 0);
    script.emit(0, 0);
    script.emit_op_u16(Op::CONST, zero, 0);
    script.emit_op_u16(Op::CONST, ten, 0);
    script.emit_op_u16(Op::CONST, two, 0);
    script.emit_op_u8(Op::CALL_REF, 3, 0);
    script.emit_op(Op::ARRAY_LENGTH, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, stdlib.chunks);
    assert_eq!(
        result.as_i32(),
        5,
        "range(0,10,2) should produce 5 elements"
    );
}

// ── reversed ────────────────────────────────────────────────

#[test]
fn stdlib_reversed() {
    let mut script = Chunk::new("<script>");
    script.local_count = 3;
    let stdlib = build_stdlib(&mut script);
    let rev_idx = stdlib.get("__stdlib_reversed").unwrap() + 1;

    // reversed([1, 2, 3]) → [3, 2, 1]
    let v1 = script.add_constant(Value::I32(1));
    let v2 = script.add_constant(Value::I32(2));
    let v3 = script.add_constant(Value::I32(3));
    script.emit_op_u16(Op::CONST, v1, 0);
    script.emit_op_u16(Op::CONST, v2, 0);
    script.emit_op_u16(Op::CONST, v3, 0);
    script.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, 0);
    script.emit_op_u16(Op::LOCAL_SET, 1, 0);
    script.emit_op(Op::DROP, 0);

    script.emit_op_u16(Op::REF_FUNC, rev_idx as u16, 0);
    script.emit(0, 0);
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op_u8(Op::CALL_REF, 1, 0);

    // Get first element — should be 3
    script.emit_op(Op::I32_CONST_0, 0);
    script.emit_op(Op::ARRAY_GET, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, stdlib.chunks);
    assert_eq!(
        result.as_i32(),
        3,
        "first element of reversed [1,2,3] should be 3"
    );
}

// ── sorted ──────────────────────────────────────────────────

#[test]
fn stdlib_sorted() {
    let mut script = Chunk::new("<script>");
    script.local_count = 3;
    let stdlib = build_stdlib(&mut script);
    let sort_idx = stdlib.get("__stdlib_sorted").unwrap() + 1;

    // sorted([3, 1, 2]) → [1, 2, 3]
    let v3 = script.add_constant(Value::I32(3));
    let v1 = script.add_constant(Value::I32(1));
    let v2 = script.add_constant(Value::I32(2));
    script.emit_op_u16(Op::CONST, v3, 0);
    script.emit_op_u16(Op::CONST, v1, 0);
    script.emit_op_u16(Op::CONST, v2, 0);
    script.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, 0);
    script.emit_op_u16(Op::LOCAL_SET, 1, 0);
    script.emit_op(Op::DROP, 0);

    script.emit_op_u16(Op::REF_FUNC, sort_idx as u16, 0);
    script.emit(0, 0);
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op_u8(Op::CALL_REF, 1, 0);

    // First element should be 1 (smallest)
    script.emit_op(Op::I32_CONST_0, 0);
    script.emit_op(Op::ARRAY_GET, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, stdlib.chunks);
    assert_eq!(
        result.as_i32(),
        1,
        "first element of sorted [3,1,2] should be 1"
    );
}

#[test]
fn stdlib_sorted_preserves_original() {
    let mut script = Chunk::new("<script>");
    script.local_count = 3;
    let stdlib = build_stdlib(&mut script);
    let sort_idx = stdlib.get("__stdlib_sorted").unwrap() + 1;

    // sorted should not mutate the original
    let v3 = script.add_constant(Value::I32(3));
    let v1 = script.add_constant(Value::I32(1));
    script.emit_op_u16(Op::CONST, v3, 0);
    script.emit_op_u16(Op::CONST, v1, 0);
    script.emit_op_u16(Op::ARRAY_NEW_FIXED, 2, 0);
    script.emit_op_u16(Op::LOCAL_SET, 1, 0);
    script.emit_op(Op::DROP, 0);

    // Call sorted
    script.emit_op_u16(Op::REF_FUNC, sort_idx as u16, 0);
    script.emit(0, 0);
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op_u8(Op::CALL_REF, 1, 0);
    script.emit_op(Op::DROP, 0); // discard sorted result

    // Original first element should still be 3
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op(Op::I32_CONST_0, 0);
    script.emit_op(Op::ARRAY_GET, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, stdlib.chunks);
    assert_eq!(result.as_i32(), 3, "original array should be unchanged");
}

// ── min / max ───────────────────────────────────────────────

#[test]
fn stdlib_min() {
    let mut script = Chunk::new("<script>");
    script.local_count = 3;
    let stdlib = build_stdlib(&mut script);
    let min_idx = stdlib.get("__stdlib_min").unwrap() + 1;

    let v5 = script.add_constant(Value::I32(5));
    let v2 = script.add_constant(Value::I32(2));
    let v8 = script.add_constant(Value::I32(8));
    script.emit_op_u16(Op::CONST, v5, 0);
    script.emit_op_u16(Op::CONST, v2, 0);
    script.emit_op_u16(Op::CONST, v8, 0);
    script.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, 0);
    script.emit_op_u16(Op::LOCAL_SET, 1, 0);
    script.emit_op(Op::DROP, 0);

    script.emit_op_u16(Op::REF_FUNC, min_idx as u16, 0);
    script.emit(0, 0);
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op_u8(Op::CALL_REF, 1, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, stdlib.chunks);
    assert_eq!(result.as_i32(), 2);
}

#[test]
fn stdlib_max() {
    let mut script = Chunk::new("<script>");
    script.local_count = 3;
    let stdlib = build_stdlib(&mut script);
    let max_idx = stdlib.get("__stdlib_max").unwrap() + 1;

    let v5 = script.add_constant(Value::I32(5));
    let v2 = script.add_constant(Value::I32(2));
    let v8 = script.add_constant(Value::I32(8));
    script.emit_op_u16(Op::CONST, v5, 0);
    script.emit_op_u16(Op::CONST, v2, 0);
    script.emit_op_u16(Op::CONST, v8, 0);
    script.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, 0);
    script.emit_op_u16(Op::LOCAL_SET, 1, 0);
    script.emit_op(Op::DROP, 0);

    script.emit_op_u16(Op::REF_FUNC, max_idx as u16, 0);
    script.emit(0, 0);
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op_u8(Op::CALL_REF, 1, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, stdlib.chunks);
    assert_eq!(result.as_i32(), 8);
}

// ── pow ─────────────────────────────────────────────────────

#[test]
fn stdlib_pow() {
    let mut script = Chunk::new("<script>");
    script.local_count = 2;
    let stdlib = build_stdlib(&mut script);
    let pow_idx = stdlib.get("__stdlib_pow").unwrap() + 1;

    // 2^10 = 1024
    script.emit_op_u16(Op::REF_FUNC, pow_idx as u16, 0);
    script.emit(0, 0);
    let base = script.add_constant(Value::F64(2.0));
    let exp = script.add_constant(Value::I32(10));
    script.emit_op_u16(Op::CONST, base, 0);
    script.emit_op_u16(Op::CONST, exp, 0);
    script.emit_op_u8(Op::CALL_REF, 2, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, stdlib.chunks);
    assert_eq!(result.as_f64() as i32, 1024);
}

// ── enumerate ───────────────────────────────────────────────

#[test]
fn stdlib_enumerate() {
    let mut script = Chunk::new("<script>");
    script.local_count = 3;
    let stdlib = build_stdlib(&mut script);
    let enum_idx = stdlib.get("__stdlib_enumerate").unwrap() + 1;

    // enumerate(["a", "b"]) → [[0,"a"], [1,"b"]]
    let va = script.add_constant(Value::String(Arc::from("a")));
    let vb = script.add_constant(Value::String(Arc::from("b")));
    script.emit_op_u16(Op::CONST, va, 0);
    script.emit_op_u16(Op::CONST, vb, 0);
    script.emit_op_u16(Op::ARRAY_NEW_FIXED, 2, 0);
    script.emit_op_u16(Op::LOCAL_SET, 1, 0);
    script.emit_op(Op::DROP, 0);

    script.emit_op_u16(Op::REF_FUNC, enum_idx as u16, 0);
    script.emit(0, 0);
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op_u8(Op::CALL_REF, 1, 0);

    // Result length should be 2
    script.emit_op(Op::ARRAY_LENGTH, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, stdlib.chunks);
    assert_eq!(result.as_i32(), 2);
}

// ── zip ─────────────────────────────────────────────────────

#[test]
fn stdlib_zip() {
    let mut script = Chunk::new("<script>");
    script.local_count = 4;
    let stdlib = build_stdlib(&mut script);
    let zip_idx = stdlib.get("__stdlib_zip").unwrap() + 1;

    // zip([1,2], [3,4]) → [[1,3], [2,4]]
    let v1 = script.add_constant(Value::I32(1));
    let v2 = script.add_constant(Value::I32(2));
    let v3 = script.add_constant(Value::I32(3));
    let v4 = script.add_constant(Value::I32(4));

    script.emit_op_u16(Op::CONST, v1, 0);
    script.emit_op_u16(Op::CONST, v2, 0);
    script.emit_op_u16(Op::ARRAY_NEW_FIXED, 2, 0);
    script.emit_op_u16(Op::LOCAL_SET, 1, 0);
    script.emit_op(Op::DROP, 0);

    script.emit_op_u16(Op::CONST, v3, 0);
    script.emit_op_u16(Op::CONST, v4, 0);
    script.emit_op_u16(Op::ARRAY_NEW_FIXED, 2, 0);
    script.emit_op_u16(Op::LOCAL_SET, 2, 0);
    script.emit_op(Op::DROP, 0);

    script.emit_op_u16(Op::REF_FUNC, zip_idx as u16, 0);
    script.emit(0, 0);
    script.emit_op_u16(Op::LOCAL_GET, 1, 0);
    script.emit_op_u16(Op::LOCAL_GET, 2, 0);
    script.emit_op_u8(Op::CALL_REF, 2, 0);

    script.emit_op(Op::ARRAY_LENGTH, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_prebuilt(script, stdlib.chunks);
    assert_eq!(result.as_i32(), 2);
}

// ── stdlib building ─────────────────────────────────────────

#[test]
fn stdlib_has_all_functions() {
    let mut script = Chunk::new("<script>");
    let stdlib = build_stdlib(&mut script);
    assert!(
        stdlib.chunks.len() >= 24,
        "stdlib should export >= 24 chunks, got {}",
        stdlib.chunks.len()
    );
    assert!(stdlib.get("__stdlib_range").is_some());
    assert!(stdlib.get("__stdlib_sorted").is_some());
    assert!(stdlib.get("__stdlib_reversed").is_some());
    assert!(stdlib.get("__stdlib_enumerate").is_some());
    assert!(stdlib.get("__stdlib_zip").is_some());
    assert!(stdlib.get("__stdlib_sum").is_some());
    assert!(stdlib.get("__stdlib_min").is_some());
    assert!(stdlib.get("__stdlib_max").is_some());
    assert!(stdlib.get("__stdlib_pow").is_some());
    assert!(stdlib.get("__stdlib_slicestep").is_some());
    assert!(stdlib.get("__stdlib_dynmul").is_some());
}
