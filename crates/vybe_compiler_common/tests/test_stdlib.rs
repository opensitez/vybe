/// Tests for the WASM stdlib — pure bytecode implementations that run on any runtime.

use vybe_bytecode::{VM, Value, Chunk, Op};
use vybe_compiler_common::stdlib::build_stdlib;
use std::rc::Rc;

/// Helper: link stdlib chunks into VM, then run a script chunk that calls them.
fn run_with_stdlib(script: Chunk) -> Value {
    let mut vm = VM::new();
    vybe_host::register_all(&mut vm);

    let stdlib = build_stdlib();
    let stdlib_base = 1; // script is chunk 0, stdlib starts at 1
    let mut all_chunks = vec![script];
    all_chunks.extend(stdlib.chunks);
    vm.run(all_chunks).unwrap()
}

// ── range ───────────────────────────────────────────────────

#[test]
fn stdlib_range_basic() {
    let stdlib = build_stdlib();
    let range_idx = stdlib.get("__stdlib_range").unwrap();
    let chunk_idx = range_idx + 1; // +1 because script is chunk 0

    let mut script = Chunk::new("<script>");
    script.local_count = 2;

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

    let result = run_with_stdlib(script);
    assert_eq!(result.as_i32(), 5, "range(0,5,1) should produce 5 elements");
}

#[test]
fn stdlib_range_step() {
    let stdlib = build_stdlib();
    let range_idx = stdlib.get("__stdlib_range").unwrap() + 1;

    let mut script = Chunk::new("<script>");
    script.local_count = 2;

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

    let result = run_with_stdlib(script);
    assert_eq!(result.as_i32(), 5, "range(0,10,2) should produce 5 elements");
}

// ── sum ─────────────────────────────────────────────────────

#[test]
fn stdlib_sum() {
    let stdlib = build_stdlib();
    let sum_idx = stdlib.get("__stdlib_sum").unwrap() + 1;

    let mut script = Chunk::new("<script>");
    script.local_count = 2;

    // sum([1, 2, 3, 4, 5]) → 15
    let v1 = script.add_constant(Value::I32(1));
    let v2 = script.add_constant(Value::I32(2));
    let v3 = script.add_constant(Value::I32(3));
    let v4 = script.add_constant(Value::I32(4));
    let v5 = script.add_constant(Value::I32(5));
    script.emit_op_u16(Op::CONST, v1, 0);
    script.emit_op_u16(Op::CONST, v2, 0);
    script.emit_op_u16(Op::CONST, v3, 0);
    script.emit_op_u16(Op::CONST, v4, 0);
    script.emit_op_u16(Op::CONST, v5, 0);
    script.emit_op_u16(Op::ARRAY_NEW_FIXED, 5, 0);

    script.emit_op_u16(Op::REF_FUNC, sum_idx as u16, 0);
    script.emit(0, 0);
    // swap: need [func, arr] but have [arr, func]
    // store arr, push func, push arr
    script.emit_op_u16(Op::LOCAL_SET, 1, 0);
    script.emit_op(Op::DROP, 0);

    // Restart: build array, store, then call
    // Actually let me restructure properly
    let mut script = Chunk::new("<script>");
    script.local_count = 3;

    let v1 = script.add_constant(Value::I32(1));
    let v2 = script.add_constant(Value::I32(2));
    let v3 = script.add_constant(Value::I32(3));
    let v4 = script.add_constant(Value::I32(4));
    let v5 = script.add_constant(Value::I32(5));
    script.emit_op_u16(Op::CONST, v1, 0);
    script.emit_op_u16(Op::CONST, v2, 0);
    script.emit_op_u16(Op::CONST, v3, 0);
    script.emit_op_u16(Op::CONST, v4, 0);
    script.emit_op_u16(Op::CONST, v5, 0);
    script.emit_op_u16(Op::ARRAY_NEW_FIXED, 5, 0);
    script.emit_op_u16(Op::LOCAL_SET, 1, 0); // store array
    script.emit_op(Op::DROP, 0);

    script.emit_op_u16(Op::REF_FUNC, sum_idx as u16, 0);
    script.emit(0, 0); // func ref
    script.emit_op_u16(Op::LOCAL_GET, 1, 0); // array arg
    script.emit_op_u8(Op::CALL_REF, 1, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_stdlib(script);
    assert_eq!(result.as_i32(), 15);
}

// ── reversed ────────────────────────────────────────────────

#[test]
fn stdlib_reversed() {
    let stdlib = build_stdlib();
    let rev_idx = stdlib.get("__stdlib_reversed").unwrap() + 1;

    let mut script = Chunk::new("<script>");
    script.local_count = 3;

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

    let result = run_with_stdlib(script);
    assert_eq!(result.as_i32(), 3, "first element of reversed [1,2,3] should be 3");
}

// ── sorted ──────────────────────────────────────────────────

#[test]
fn stdlib_sorted() {
    let stdlib = build_stdlib();
    let sort_idx = stdlib.get("__stdlib_sorted").unwrap() + 1;

    let mut script = Chunk::new("<script>");
    script.local_count = 3;

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

    let result = run_with_stdlib(script);
    assert_eq!(result.as_i32(), 1, "first element of sorted [3,1,2] should be 1");
}

#[test]
fn stdlib_sorted_preserves_original() {
    let stdlib = build_stdlib();
    let sort_idx = stdlib.get("__stdlib_sorted").unwrap() + 1;

    let mut script = Chunk::new("<script>");
    script.local_count = 3;

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

    let result = run_with_stdlib(script);
    assert_eq!(result.as_i32(), 3, "original array should be unchanged");
}

// ── min / max ───────────────────────────────────────────────

#[test]
fn stdlib_min() {
    let stdlib = build_stdlib();
    let min_idx = stdlib.get("__stdlib_min").unwrap() + 1;

    let mut script = Chunk::new("<script>");
    script.local_count = 3;

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

    let result = run_with_stdlib(script);
    assert_eq!(result.as_i32(), 2);
}

#[test]
fn stdlib_max() {
    let stdlib = build_stdlib();
    let max_idx = stdlib.get("__stdlib_max").unwrap() + 1;

    let mut script = Chunk::new("<script>");
    script.local_count = 3;

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

    let result = run_with_stdlib(script);
    assert_eq!(result.as_i32(), 8);
}

// ── pow ─────────────────────────────────────────────────────

#[test]
fn stdlib_pow() {
    let stdlib = build_stdlib();
    let pow_idx = stdlib.get("__stdlib_pow").unwrap() + 1;

    let mut script = Chunk::new("<script>");
    script.local_count = 2;

    // 2^10 = 1024
    script.emit_op_u16(Op::REF_FUNC, pow_idx as u16, 0);
    script.emit(0, 0);
    let base = script.add_constant(Value::F64(2.0));
    let exp = script.add_constant(Value::I32(10));
    script.emit_op_u16(Op::CONST, base, 0);
    script.emit_op_u16(Op::CONST, exp, 0);
    script.emit_op_u8(Op::CALL_REF, 2, 0);
    script.emit_op(Op::HALT, 0);

    let result = run_with_stdlib(script);
    assert_eq!(result.as_f64() as i32, 1024);
}

// ── enumerate ───────────────────────────────────────────────

#[test]
fn stdlib_enumerate() {
    let stdlib = build_stdlib();
    let enum_idx = stdlib.get("__stdlib_enumerate").unwrap() + 1;

    let mut script = Chunk::new("<script>");
    script.local_count = 3;

    // enumerate(["a", "b"]) → [[0,"a"], [1,"b"]]
    let va = script.add_constant(Value::String(Rc::from("a")));
    let vb = script.add_constant(Value::String(Rc::from("b")));
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

    let result = run_with_stdlib(script);
    assert_eq!(result.as_i32(), 2);
}

// ── zip ─────────────────────────────────────────────────────

#[test]
fn stdlib_zip() {
    let stdlib = build_stdlib();
    let zip_idx = stdlib.get("__stdlib_zip").unwrap() + 1;

    let mut script = Chunk::new("<script>");
    script.local_count = 4;

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

    let result = run_with_stdlib(script);
    assert_eq!(result.as_i32(), 2);
}

// ── stdlib building ─────────────────────────────────────────

#[test]
fn stdlib_has_all_functions() {
    let stdlib = build_stdlib();
    assert_eq!(stdlib.chunks.len(), 24);
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
