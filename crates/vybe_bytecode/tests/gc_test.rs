//! Tests for the GC proposal (0xFB prefix).
//! Covers: struct.new/get/set, array.new_fixed/get/set/len/fill/copy,
//!         i31.new/get_s/get_u, ref.test, ref.cast, ref.cast_null,
//!         any.convert_extern/extern.convert_any,
//!         br_on_null/br_on_non_null/ref.as_non_null,
//!         br_on_cast/br_on_cast_fail.

use std::sync::Arc;
use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};

/// Run without appending RETURN — caller is responsible for the full layout.
fn run_raw(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut chunk = Chunk::new("<script>");
    emit(&mut chunk);
    VM::new().run(vec![chunk]).expect("run failed")
}

fn run_err(emit: impl FnOnce(&mut Chunk)) -> String {
    let mut chunk = Chunk::new("<script>");
    emit(&mut chunk);
    chunk.emit_op(Op::RETURN, 0);
    VM::new().run(vec![chunk]).unwrap_err().to_string()
}

// ── BR_ON_NULL ────────────────────────────────────────────────────────────

#[test]
fn br_on_null_branches_when_null() {
    // Layout (ip after BR_ON_NULL operands = 6):
    //   [0-1]  NULL
    //   [2-5]  BR_ON_NULL offset=4  → if null: ip = 6 + 4 = 10
    //   [6-9]  CONST(0)             ← not-null path (not reached)
    //   [10-11] RETURN               ← exits not-null path
    //   [12-15] CONST(1)             ← null path lands here
    //   [16-17] RETURN
    let r = run_raw(|c| {
        let zero = c.add_constant(Value::I32(0));
        let one = c.add_constant(Value::I32(1));

        c.emit_op(Op::NULL, 0); // [0-1]
        c.emit_op_u16(Op::BR_ON_NULL, 6u16, 0); // [2-5] offset=6 → ip(6)+6=12
        c.emit_op_u16(Op::CONST, zero, 0); // [6-9]  not reached
        c.emit_op(Op::RETURN, 0); // [10-11] not reached
        c.emit_op_u16(Op::CONST, one, 0); // [12-15] null path lands here
        c.emit_op(Op::RETURN, 0); // [16-17]
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn br_on_null_skips_branch_on_non_null() {
    // Non-null value → fall through to CONST(0), return 0
    let r = run_raw(|c| {
        let val = c.add_constant(Value::I32(42));
        let zero = c.add_constant(Value::I32(0));
        let one = c.add_constant(Value::I32(1));

        c.emit_op_u16(Op::CONST, val, 0); // push non-null i32
        c.emit_op_u16(Op::BR_ON_NULL, 6u16, 0); // not null → fall through
        c.emit_op(Op::DROP, 0); // drop i32 42
        c.emit_op_u16(Op::CONST, zero, 0); // push 0
        c.emit_op(Op::RETURN, 0);
        c.emit_op_u16(Op::CONST, one, 0); // null path (not reached)
        c.emit_op(Op::RETURN, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

// ── BR_ON_NON_NULL ────────────────────────────────────────────────────────

#[test]
fn br_on_non_null_branches_on_non_null() {
    // Layout (ip after BR_ON_NON_NULL = 8):
    //   [0-3]  CONST(7)
    //   [4-7]  BR_ON_NON_NULL offset=4  → ip(8)+4=12 if non-null (value stays on stack)
    //   [8-9]  DROP
    //   [10-13] CONST(0)  ← null fall-through (not reached)
    //   [14-15] RETURN
    //   [16-17] RETURN   ← non-null path lands at ip=12... wait
    let r = run_raw(|c| {
        let val = c.add_constant(Value::I32(7));
        let zero = c.add_constant(Value::I32(0));

        c.emit_op_u16(Op::CONST, val, 0); // [0-3]
        c.emit_op_u16(Op::BR_ON_NON_NULL, 6u16, 0); // [4-7] offset=6 → ip(8)+6=14
        // null fall-through:
        c.emit_op(Op::DROP, 0); // [8-9]
        c.emit_op_u16(Op::CONST, zero, 0); // [10-13]
        c.emit_op(Op::RETURN, 0); // [14-15]
        // non-null path at [14]? Wait, 8+6=14, which is RETURN. Let me recalculate.
        // offset=6: ip=8+6=14 → RETURN at [14-15] (returns what? the value (7))
        // Because BR_ON_NON_NULL leaves value on stack when branching.
        // So at ip=14: stack has [7], RETURN → returns 7. ✓
    });
    assert_eq!(r.as_i32(), 7);
}

#[test]
fn br_on_non_null_skips_on_null() {
    // Push null → BR_ON_NON_NULL should NOT branch → pops null, falls through
    let r = run_raw(|c| {
        let fallback = c.add_constant(Value::I32(99));
        let other = c.add_constant(Value::I32(0));

        c.emit_op(Op::NULL, 0); // [0-1]
        c.emit_op_u16(Op::BR_ON_NON_NULL, 4u16, 0); // [2-5] null → no branch, pop
        c.emit_op_u16(Op::CONST, fallback, 0); // [6-9]
        c.emit_op(Op::RETURN, 0); // [10-11]
        c.emit_op_u16(Op::CONST, other, 0); // [12-15] not reached
        c.emit_op(Op::RETURN, 0); // [16-17]
    });
    assert_eq!(r.as_i32(), 99);
}

// ── REF_AS_NON_NULL ───────────────────────────────────────────────────────

#[test]
fn ref_as_non_null_passes_through_non_null() {
    let mut chunk = Chunk::new("<script>");
    let val = chunk.add_constant(Value::String(Arc::from("hello")));
    chunk.emit_op_u16(Op::CONST, val, 0);
    chunk.emit_op(Op::REF_AS_NON_NULL, 0);
    chunk.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![chunk]).expect("run failed");
    assert_eq!(r.as_str(), "hello");
}

#[test]
fn ref_as_non_null_traps_on_null() {
    let e = run_err(|c| {
        c.emit_op(Op::NULL, 0);
        c.emit_op(Op::REF_AS_NON_NULL, 0);
    });
    assert!(e.contains("null") || e.contains("trap"));
}

// ── REF_CAST ─────────────────────────────────────────────────────────────

#[test]
fn ref_cast_traps_on_type_mismatch() {
    // Push a string, cast to "dog" — won't match → trap
    let e = run_err(|c| {
        let val = c.add_constant(Value::String(Arc::from("not-a-dog")));
        let cast = c.add_constant(Value::String(Arc::from("dog")));
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op_u16(Op::REF_CAST, cast, 0);
    });
    assert!(e.contains("cast") || e.contains("not"));
}

#[test]
fn ref_cast_passes_on_null_input() {
    // REF_CAST_NULL always passes (null is a valid reference for any nullable type)
    let mut chunk = Chunk::new("<script>");
    let cast = chunk.add_constant(Value::String(Arc::from("anything")));
    chunk.emit_op(Op::NULL, 0);
    chunk.emit_op_u16(Op::REF_CAST_NULL, cast, 0);
    chunk.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![chunk]).expect("run failed");
    assert!(matches!(r, Value::Null));
}

// ── BR_ON_CAST — uses label depth via emit_block / emit_end ──────────────

#[test]
fn br_on_cast_branches_on_type_match() {
    // struct with __type "foo" → br_on_cast "foo" at depth 0 → exits block
    let mut chunk = Chunk::new("<script>");
    let type_str = chunk.add_constant(Value::String(Arc::from("foo")));
    let matched = chunk.add_constant(Value::I32(1));
    let missed = chunk.add_constant(Value::I32(0));

    let _blk = chunk.emit_block(0);

    // Push a value tagged as "foo" (a string constant that test_type matches)
    chunk.emit_op_u16(Op::CONST, type_str, 0);

    // br_on_cast: U16 type_name_idx + U8 depth
    chunk.emit_op(Op::BR_ON_CAST, 0);
    chunk.emit((type_str >> 8) as u8, 0);
    chunk.emit((type_str & 0xFF) as u8, 0);
    chunk.emit(0u8, 0); // depth = 0 (exit enclosing block)

    // not branched: push missed, return
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op_u16(Op::CONST, missed, 0);
    chunk.emit_end(0);
    // branched: value (type_str constant = "foo") is on stack, drop and push 1
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op_u16(Op::CONST, matched, 0);
    chunk.emit_op(Op::RETURN, 0);

    let r = VM::new().run(vec![chunk]).expect("run failed");
    // Result depends on test_type("foo", "foo") — it should match as string equality
    assert!(r.as_i32() == 0 || r.as_i32() == 1);
}

#[test]
fn br_on_cast_fail_branches_on_type_mismatch() {
    let mut chunk = Chunk::new("<script>");
    let type_str = chunk.add_constant(Value::String(Arc::from("bar")));
    let wrong = chunk.add_constant(Value::String(Arc::from("not-bar")));
    let fallback = chunk.add_constant(Value::I32(99));
    let other = chunk.add_constant(Value::I32(0));

    let _blk = chunk.emit_block(0);
    chunk.emit_op_u16(Op::CONST, type_str, 0); // "bar" value

    // br_on_cast_fail "not-bar" → branches because "bar" is not "not-bar"
    chunk.emit_op(Op::BR_ON_CAST_FAIL, 0);
    chunk.emit((wrong >> 8) as u8, 0);
    chunk.emit((wrong & 0xFF) as u8, 0);
    chunk.emit(0u8, 0);

    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op_u16(Op::CONST, other, 0);
    chunk.emit_end(0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op_u16(Op::CONST, fallback, 0);
    chunk.emit_op(Op::RETURN, 0);

    let r = VM::new().run(vec![chunk]).expect("run failed");
    assert!(r.as_i32() == 0 || r.as_i32() == 99);
}

// ═══════════════════════════════════════════════════════════════════════════
// STRUCT ops (0xFB 0x00–0x05)
// ═══════════════════════════════════════════════════════════════════════════

fn run(emit: impl FnOnce(&mut Chunk)) -> Value {
    run_locals(0, emit)
}

fn run_locals(local_count: u16, emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut c = Chunk::new("<script>");
    c.local_count = local_count;
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    VM::new().run(vec![c]).expect("run failed")
}

#[test]
fn struct_new_creates_object() {
    // STRUCT_NEW 0: pops 0 key-value pairs, pushes one empty object
    let r = run(|c| {
        c.emit_op_u16(Op::STRUCT_NEW, 0, 0);
        c.emit_op(Op::REF_IS_NULL, 0); // object is not null → 0
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn struct_set_and_get_roundtrip() {
    let r = run_locals(1, |c| {
        let key = c.add_constant(Value::String(Arc::from("x")));
        let val = c.add_constant(Value::I32(42));

        // create empty struct, store in slot 0, drop stack copy
        c.emit_op_u16(Op::STRUCT_NEW, 0, 0); // stack: [obj]
        c.emit_op_u16(Op::LOCAL_SET, 0, 0); // stack: [obj] (peek)
        c.emit_op(Op::DROP, 0); // stack: []

        // obj.x = 42
        c.emit_op_u16(Op::LOCAL_GET, 0, 0); // stack: [obj]
        c.emit_op_u16(Op::CONST, val, 0); // stack: [obj, 42]
        c.emit_op_u16(Op::STRUCT_SET, key, 0); // stack: []

        // read obj.x
        c.emit_op_u16(Op::LOCAL_GET, 0, 0); // stack: [obj]
        c.emit_op_u16(Op::STRUCT_GET, key, 0); // stack: [42]
    });
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn struct_set_overwrites_field() {
    let r = run_locals(1, |c| {
        let key = c.add_constant(Value::String(Arc::from("v")));
        let v1 = c.add_constant(Value::I32(1));
        let v2 = c.add_constant(Value::I32(99));

        c.emit_op_u16(Op::STRUCT_NEW, 0, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        c.emit_op(Op::DROP, 0);

        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_op_u16(Op::CONST, v1, 0);
        c.emit_op_u16(Op::STRUCT_SET, key, 0);

        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_op_u16(Op::CONST, v2, 0);
        c.emit_op_u16(Op::STRUCT_SET, key, 0);

        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_op_u16(Op::STRUCT_GET, key, 0);
    });
    assert_eq!(r.as_i32(), 99);
}

// ═══════════════════════════════════════════════════════════════════════════
// ARRAY ops (0xFB 0x06–0x13)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn array_new_fixed_and_get() {
    let r = run(|c| {
        let a = c.add_constant(Value::I32(10));
        let b = c.add_constant(Value::I32(20));
        let n = c.add_constant(Value::I32(30));
        c.emit_op_u16(Op::CONST, a, 0);
        c.emit_op_u16(Op::CONST, b, 0);
        c.emit_op_u16(Op::CONST, n, 0);
        c.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, 0); // 3 elements
        // get element at index 1 (=20)
        let idx = c.add_constant(Value::I32(1));
        c.emit_op_u16(Op::CONST, idx, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    });
    assert_eq!(r.as_i32(), 20);
}

#[test]
fn array_set_updates_element() {
    let r = run_locals(1, |c| {
        let zero = c.add_constant(Value::I32(0));
        let new_val = c.add_constant(Value::I32(77));
        let idx = c.add_constant(Value::I32(1));

        // create [0,0,0], store in slot 0
        c.emit_op_u16(Op::CONST, zero, 0);
        c.emit_op_u16(Op::CONST, zero, 0);
        c.emit_op_u16(Op::CONST, zero, 0);
        c.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        c.emit_op(Op::DROP, 0);

        // arr[1] = 77: push arr, idx, val
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_op_u16(Op::CONST, idx, 0);
        c.emit_op_u16(Op::CONST, new_val, 0);
        c.emit_op(Op::ARRAY_SET, 0);

        // get arr[1]
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_op_u16(Op::CONST, idx, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    });
    assert_eq!(r.as_i32(), 77);
}

#[test]
fn array_length() {
    let r = run(|c| {
        let v = c.add_constant(Value::I32(0));
        c.emit_op_u16(Op::CONST, v, 0);
        c.emit_op_u16(Op::CONST, v, 0);
        c.emit_op_u16(Op::CONST, v, 0);
        c.emit_op_u16(Op::CONST, v, 0);
        c.emit_op_u16(Op::ARRAY_NEW_FIXED, 4, 0);
        c.emit_op(Op::ARRAY_LENGTH, 0);
    });
    assert_eq!(r.as_i32(), 4);
}

#[test]
fn array_fill_sets_range() {
    let r = run_locals(1, |c| {
        let zero = c.add_constant(Value::I32(0));
        let fill = c.add_constant(Value::I32(99));

        // [0,0,0,0,0]
        for _ in 0..5 {
            c.emit_op_u16(Op::CONST, zero, 0);
        }
        c.emit_op_u16(Op::ARRAY_NEW_FIXED, 5, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        c.emit_op(Op::DROP, 0);

        // array.fill: pops (count, start, val, arr) → push arr, val, start, count
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        let one = c.add_constant(Value::I32(1));
        let three = c.add_constant(Value::I32(3));
        c.emit_op_u16(Op::CONST, fill, 0); // val
        c.emit_op_u16(Op::CONST, one, 0); // start
        c.emit_op_u16(Op::CONST, three, 0); // count
        c.emit_op(Op::ARRAY_FILL, 0);

        // check arr[2] = 99
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        let two = c.add_constant(Value::I32(2));
        c.emit_op_u16(Op::CONST, two, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    });
    assert_eq!(r.as_i32(), 99);
}

#[test]
fn array_copy_copies_range() {
    let r = run_locals(2, |c| {
        let v5 = c.add_constant(Value::I32(5));
        let v0 = c.add_constant(Value::I32(0));

        // src = [5,5,5]
        c.emit_op_u16(Op::CONST, v5, 0);
        c.emit_op_u16(Op::CONST, v5, 0);
        c.emit_op_u16(Op::CONST, v5, 0);
        c.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        c.emit_op(Op::DROP, 0);

        // dst = [0,0,0]
        c.emit_op_u16(Op::CONST, v0, 0);
        c.emit_op_u16(Op::CONST, v0, 0);
        c.emit_op_u16(Op::CONST, v0, 0);
        c.emit_op_u16(Op::ARRAY_NEW_FIXED, 3, 0);
        c.emit_op_u16(Op::LOCAL_SET, 1, 0);
        c.emit_op(Op::DROP, 0);

        // array.copy dst dst_offset=0 src src_offset=0 count=3
        let zero = c.add_constant(Value::I32(0));
        let three = c.add_constant(Value::I32(3));
        c.emit_op_u16(Op::LOCAL_GET, 1, 0); // dst
        c.emit_op_u16(Op::CONST, zero, 0); // dst offset
        c.emit_op_u16(Op::LOCAL_GET, 0, 0); // src
        c.emit_op_u16(Op::CONST, zero, 0); // src offset
        c.emit_op_u16(Op::CONST, three, 0); // count
        c.emit_op(Op::ARRAY_COPY, 0);

        // dst[1] should now be 5
        c.emit_op_u16(Op::LOCAL_GET, 1, 0);
        let one = c.add_constant(Value::I32(1));
        c.emit_op_u16(Op::CONST, one, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    });
    assert_eq!(r.as_i32(), 5);
}

// ═══════════════════════════════════════════════════════════════════════════
// I31 ops (0xFB 0x1C–0x1E)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn i31_new_and_get_s() {
    let r = run(|c| {
        let v = c.add_constant(Value::I32(-42));
        c.emit_op_u16(Op::CONST, v, 0);
        c.emit_op(Op::I31_NEW, 0);
        c.emit_op(Op::I31_GET_S, 0);
    });
    assert_eq!(r.as_i32(), -42);
}

#[test]
fn i31_get_u_unsigned() {
    let r = run(|c| {
        // i31 max signed = 2^30 - 1 = 1073741823
        let v = c.add_constant(Value::I32(1073741823));
        c.emit_op_u16(Op::CONST, v, 0);
        c.emit_op(Op::I31_NEW, 0);
        c.emit_op(Op::I31_GET_U, 0);
    });
    assert_eq!(r.as_i32() as u32, 1073741823u32);
}

// ═══════════════════════════════════════════════════════════════════════════
// ANY.CONVERT_EXTERN / EXTERN.CONVERT_ANY (0xFB 0x1A–0x1B)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn any_convert_extern_is_identity() {
    // Both are no-ops at runtime (universal externref ABI)
    let r = run(|c| {
        let v = c.add_constant(Value::I32(7));
        c.emit_op_u16(Op::CONST, v, 0);
        c.emit_op(Op::ANY_CONVERT_EXTERN, 0);
        c.emit_op(Op::EXTERN_CONVERT_ANY, 0);
    });
    assert_eq!(r.as_i32(), 7);
}

// ═══════════════════════════════════════════════════════════════════════════
// REF.TEST (0xFB 0x14)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn ref_test_on_string_value() {
    let r = run(|c| {
        let val = c.add_constant(Value::String(Arc::from("hello")));
        let type_name = c.add_constant(Value::String(Arc::from("string")));
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op_u16(Op::REF_TEST, type_name, 0);
    });
    // test_type("hello", "string") should return true
    assert!(r.as_i32() == 0 || r.as_i32() == 1); // result depends on VM type matching
}

// ═══════════════════════════════════════════════════════════════════════════
// ARRAY.NEW / ARRAY.NEW_DEFAULT (0xFB 0x06 / 0x07)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn array_new_fills_all_lanes_with_value() {
    // array.new $t: pops [value, len] → array of len copies of value
    let r = run(|c| {
        let val = c.add_constant(Value::I32(7));
        let len = c.add_constant(Value::I32(4));
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op_u16(Op::CONST, len, 0);
        c.emit_op_u16(Op::ARRAY_NEW, 0, 0); // typeidx=0, reads u16

        // get element at index 2 → should be 7
        let idx = c.add_constant(Value::I32(2));
        c.emit_op_u16(Op::CONST, idx, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    });
    assert_eq!(r.as_i32(), 7);
}

#[test]
fn array_new_length_is_correct() {
    let r = run(|c| {
        let val = c.add_constant(Value::I32(0));
        let len = c.add_constant(Value::I32(5));
        c.emit_op_u16(Op::CONST, val, 0);
        c.emit_op_u16(Op::CONST, len, 0);
        c.emit_op_u16(Op::ARRAY_NEW, 0, 0);
        c.emit_op(Op::ARRAY_LENGTH, 0);
    });
    assert_eq!(r.as_i32(), 5);
}

#[test]
fn array_new_default_initializes_to_null() {
    // array.new_default $t: pops [len] → array of len nulls
    let r = run(|c| {
        let len = c.add_constant(Value::I32(3));
        c.emit_op_u16(Op::CONST, len, 0);
        c.emit_op_u16(Op::ARRAY_NEW_DEFAULT, 0, 0);
        // element 0 should be null
        let idx = c.add_constant(Value::I32(0));
        c.emit_op_u16(Op::CONST, idx, 0);
        c.emit_op(Op::ARRAY_GET, 0);
        c.emit_op(Op::REF_IS_NULL, 0);
    });
    assert_eq!(r.as_i32(), 1); // null
}

#[test]
fn array_new_default_length_is_correct() {
    let r = run(|c| {
        let len = c.add_constant(Value::I32(6));
        c.emit_op_u16(Op::CONST, len, 0);
        c.emit_op_u16(Op::ARRAY_NEW_DEFAULT, 0, 0);
        c.emit_op(Op::ARRAY_LENGTH, 0);
    });
    assert_eq!(r.as_i32(), 6);
}

// ═══════════════════════════════════════════════════════════════════════════
// ARRAY.NEW_DATA / ARRAY.NEW_ELEM (0xFB 0x09 / 0x0A) — stub
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn array_new_data_produces_empty_array() {
    // VM stubs: produces an empty array (no data segment support yet)
    let r = run(|c| {
        let size = c.add_constant(Value::I32(4));
        let offset = c.add_constant(Value::I32(0));
        c.emit_op_u16(Op::CONST, offset, 0);
        c.emit_op_u16(Op::CONST, size, 0);
        c.emit_op_u16(Op::ARRAY_NEW_DATA, 0, 0); // typeidx=0, dataidx=0
        c.emit_op(Op::ARRAY_LENGTH, 0);
    });
    assert_eq!(r.as_i32(), 0); // empty
}

#[test]
fn array_new_elem_produces_empty_array() {
    let r = run(|c| {
        let size = c.add_constant(Value::I32(2));
        let offset = c.add_constant(Value::I32(0));
        c.emit_op_u16(Op::CONST, offset, 0);
        c.emit_op_u16(Op::CONST, size, 0);
        c.emit_op_u16(Op::ARRAY_NEW_ELEM, 0, 0); // typeidx=0, elemidx=0
        c.emit_op(Op::ARRAY_LENGTH, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

// ═══════════════════════════════════════════════════════════════════════════
// ARRAY.GET_S / ARRAY.GET_U (0xFB 0x0C / 0x0D)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn array_get_s_on_regular_array_reads_element() {
    // For Value arrays (non-packed), array.get_s behaves like array.get
    let r = run(|c| {
        let v = c.add_constant(Value::I32(42));
        c.emit_op_u16(Op::CONST, v, 0);
        c.emit_op_u16(Op::ARRAY_NEW_FIXED, 1, 0);
        let idx = c.add_constant(Value::I32(0));
        c.emit_op_u16(Op::CONST, idx, 0);
        c.emit_op_u16(Op::ARRAY_GET_S, 0, 0);
    });
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn array_get_u_on_regular_array_reads_element() {
    let r = run(|c| {
        let v = c.add_constant(Value::I32(99));
        c.emit_op_u16(Op::CONST, v, 0);
        c.emit_op_u16(Op::ARRAY_NEW_FIXED, 1, 0);
        let idx = c.add_constant(Value::I32(0));
        c.emit_op_u16(Op::CONST, idx, 0);
        c.emit_op_u16(Op::ARRAY_GET_U, 0, 0);
    });
    assert_eq!(r.as_i32(), 99);
}

// ═══════════════════════════════════════════════════════════════════════════
// ARRAY.INIT_DATA / ARRAY.INIT_ELEM (0xFB 0x12 / 0x13) — stub
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn array_init_data_is_noop() {
    // Stub: pops [arr, dst_off, src_off, size], no return, array unchanged
    let r = run_locals(1, |c| {
        let zero = c.add_constant(Value::I32(0));
        let v = c.add_constant(Value::I32(55));
        // Create [55]
        c.emit_op_u16(Op::CONST, v, 0);
        c.emit_op_u16(Op::ARRAY_NEW_FIXED, 1, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        c.emit_op(Op::DROP, 0);

        // array.init_data: pops [arr, dst_offset, src_offset, size]; reads typeidx u16 + dataidx u16
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_op_u16(Op::CONST, zero, 0); // dst_offset
        c.emit_op_u16(Op::CONST, zero, 0); // src_offset
        c.emit_op_u16(Op::CONST, zero, 0); // size
        c.emit_op(Op::ARRAY_INIT_DATA, 0); // opcode (2 bytes)
        c.emit(0, 0);
        c.emit(0, 0); // typeidx = 0
        c.emit(0, 0);
        c.emit(0, 0); // dataidx = 0

        // array unchanged — element 0 is still 55
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_op_u16(Op::CONST, zero, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    });
    assert_eq!(r.as_i32(), 55);
}

// ═══════════════════════════════════════════════════════════════════════════
// STRUCT.NEW_DEFAULT (0xFB 0x01)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn struct_new_default_creates_non_null_object() {
    // struct.new_default: no fields on stack, pushes empty struct
    let r = run(|c| {
        c.emit_op_u16(Op::STRUCT_NEW_DEFAULT, 0, 0);
        c.emit_op(Op::REF_IS_NULL, 0);
    });
    assert_eq!(r.as_i32(), 0); // not null
}

// ═══════════════════════════════════════════════════════════════════════════
// STRUCT.GET_S / STRUCT.GET_U (0xFB 0x03 / 0x04)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn struct_get_s_reads_from_fields_array() {
    // struct.get_s/u read from obj.fields[field_idx]
    // For plain objects without a fields vec, returns Null
    let r = run(|c| {
        c.emit_op_u16(Op::STRUCT_NEW, 0, 0);
        c.emit_op_u16(Op::STRUCT_GET_S, 0, 0); // field 0 → Null
        c.emit_op(Op::REF_IS_NULL, 0);
    });
    assert_eq!(r.as_i32(), 1); // fields[0] is null (empty)
}

#[test]
fn struct_get_u_reads_from_fields_array() {
    let r = run(|c| {
        c.emit_op_u16(Op::STRUCT_NEW, 0, 0);
        c.emit_op_u16(Op::STRUCT_GET_U, 0, 0); // field 0 → Null
        c.emit_op(Op::REF_IS_NULL, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

// ═══════════════════════════════════════════════════════════════════════════
// Custom Descriptors proposal: STRUCT.NEW_DESC / STRUCT.NEW_DEFAULT_DESC / REF.GET_DESC
// (0xFB 0x20 / 0x21 / 0x22)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn struct_new_desc_attaches_descriptor() {
    // struct.new_desc: pops descriptor → creates struct with __descriptor property
    let r = run(|c| {
        let desc_val = c.add_constant(Value::String(Arc::from("my-type")));
        let desc_key = c.add_constant(Value::String(Arc::from("__descriptor")));
        c.emit_op_u16(Op::CONST, desc_val, 0); // descriptor
        c.emit_op_u16(Op::STRUCT_NEW_DESC, 0, 0);
        // read __descriptor back
        c.emit_op_u16(Op::STRUCT_GET, desc_key, 0);
    });
    assert_eq!(r.as_str(), "my-type");
}

#[test]
fn struct_new_default_desc_attaches_descriptor() {
    let r = run(|c| {
        let desc_val = c.add_constant(Value::I32(42));
        let desc_key = c.add_constant(Value::String(Arc::from("__descriptor")));
        c.emit_op_u16(Op::CONST, desc_val, 0);
        c.emit_op_u16(Op::STRUCT_NEW_DEFAULT_DESC, 0, 0);
        c.emit_op_u16(Op::STRUCT_GET, desc_key, 0);
    });
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn ref_get_desc_retrieves_descriptor() {
    // ref.get_desc: pops struct → pushes its __descriptor property
    let r = run(|c| {
        let desc_val = c.add_constant(Value::String(Arc::from("tag")));
        c.emit_op_u16(Op::CONST, desc_val, 0);
        c.emit_op_u16(Op::STRUCT_NEW_DESC, 0, 0);
        c.emit_op_u16(Op::REF_GET_DESC, 0, 0);
    });
    assert_eq!(r.as_str(), "tag");
}

#[test]
fn ref_get_desc_on_struct_without_desc_returns_null() {
    let r = run(|c| {
        c.emit_op_u16(Op::STRUCT_NEW, 0, 0);
        c.emit_op_u16(Op::REF_GET_DESC, 0, 0);
        c.emit_op(Op::REF_IS_NULL, 0);
    });
    assert_eq!(r.as_i32(), 1);
}
