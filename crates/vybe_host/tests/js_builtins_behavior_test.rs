//! Phase B6 — `wasm:js-*` builtin behavioral tests.
//!
//! Exercises the **handler side** end-to-end: for each import, build
//! a tiny chunk that invokes `CALL_IMPORT`, run it through the VM
//! with all host fns registered, and verify the observable result.
//!
//! Companion to
//! `vybe_bytecode/tests/js_builtins_compliance_test.rs`
//! (byte/signature tests; pure emitter-side, no handler dispatch).
//!
//! See `dynamicruntime_support.md` Phase B6.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::register_all;

// ──────────────────────────────────────────────────────────────────────
// Harness: run a small chunk that calls a single host import.
// ──────────────────────────────────────────────────────────────────────

/// Build a chunk that:
///   1. Pushes each `pre_stack` value onto the stack (as constants).
///   2. Emits `CALL_IMPORT` targeting `(module, name)` with `argc`.
///   3. Returns the result.
///
/// The caller provides fully constructed `Value`s on the pre-stack;
/// we add them as constants and emit `CONST` for each.
fn call_import(module: &str, name: &str, pre_stack: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<test>");
    let import_idx = chunk.add_import(module, name);
    let argc = pre_stack.len() as u8;
    for v in pre_stack {
        let k = chunk.add_constant(v);
        chunk.emit_op_u16(Op::CONST, k, 0);
    }
    // CALL_IMPORT: u16 import_idx + u8 argc
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_all(&mut vm);
    vm.run(vec![chunk]).expect("VM run failed")
}

fn new_array(elements: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(elements))))
}

fn len_of(v: &Value) -> usize {
    if let Value::Object(o) = v {
        let l = o.lock().unwrap();
        if let ObjectKind::Array(ref elems) = l.kind {
            return elems.len();
        }
    }
    0
}

fn element_at(v: &Value, i: usize) -> Value {
    if let Value::Object(o) = v {
        let l = o.lock().unwrap();
        if let ObjectKind::Array(ref elems) = l.kind {
            return elems.get(i).cloned().unwrap_or(Value::Null);
        }
    }
    Value::Null
}

// ──────────────────────────────────────────────────────────────────────
// Array — mutation primitives
// ──────────────────────────────────────────────────────────────────────

#[test]
fn array_new_then_push_reports_new_length() {
    let arr = call_import("wasm:js-array", "new", vec![]);
    // push(arr, 42) → new length 1
    let r = call_import("wasm:js-array", "push", vec![arr.clone(), Value::I32(42)]);
    assert_eq!(r.as_i32(), 1, "push must return the new length per ECMA-262");
    assert_eq!(len_of(&arr), 1);
    assert_eq!(element_at(&arr, 0).as_i32(), 42);
}

#[test]
fn array_push_pop_roundtrip() {
    let arr = new_array(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    let popped = call_import("wasm:js-array", "pop", vec![arr.clone()]);
    assert_eq!(popped.as_i32(), 3);
    assert_eq!(len_of(&arr), 2);
}

#[test]
fn array_pop_on_empty_returns_undefined() {
    let arr = new_array(vec![]);
    let r = call_import("wasm:js-array", "pop", vec![arr]);
    assert!(matches!(r, Value::Undefined),
        "pop() on empty must return undefined per ECMA-262, got {:?}", r);
}

#[test]
fn array_shift_removes_from_front() {
    let arr = new_array(vec![Value::I32(10), Value::I32(20), Value::I32(30)]);
    let shifted = call_import("wasm:js-array", "shift", vec![arr.clone()]);
    assert_eq!(shifted.as_i32(), 10);
    assert_eq!(len_of(&arr), 2);
    assert_eq!(element_at(&arr, 0).as_i32(), 20);
    assert_eq!(element_at(&arr, 1).as_i32(), 30);
}

#[test]
fn array_unshift_prepends_and_returns_new_length() {
    let arr = new_array(vec![Value::I32(2), Value::I32(3)]);
    let len = call_import("wasm:js-array", "unshift", vec![arr.clone(), Value::I32(1)]);
    assert_eq!(len.as_i32(), 3);
    assert_eq!(element_at(&arr, 0).as_i32(), 1);
    assert_eq!(element_at(&arr, 1).as_i32(), 2);
}

#[test]
fn array_length_returns_current_length() {
    let arr = new_array(vec![Value::I32(1); 7]);
    let n = call_import("wasm:js-array", "length", vec![arr]);
    assert_eq!(n.as_i32(), 7);
}

#[test]
fn array_get_and_set_reflect_changes() {
    let arr = new_array(vec![Value::I32(10), Value::I32(20), Value::I32(30)]);
    let v = call_import("wasm:js-array", "get",
        vec![arr.clone(), Value::I32(1)]);
    assert_eq!(v.as_i32(), 20);

    call_import("wasm:js-array", "set",
        vec![arr.clone(), Value::I32(1), Value::I32(99)]);
    let after = call_import("wasm:js-array", "get",
        vec![arr, Value::I32(1)]);
    assert_eq!(after.as_i32(), 99);
}

#[test]
fn array_at_out_of_bounds_returns_undefined() {
    let arr = new_array(vec![Value::I32(1)]);
    let r = call_import("wasm:js-array", "at", vec![arr, Value::I32(5)]);
    assert!(matches!(r, Value::Undefined),
        "at(5) on 1-element array must be undefined per ECMA-262, got {:?}", r);
}

#[test]
fn array_slice_does_not_mutate_original() {
    let arr = new_array(vec![Value::I32(1), Value::I32(2), Value::I32(3), Value::I32(4)]);
    let sliced = call_import("wasm:js-array", "slice",
        vec![arr.clone(), Value::I32(1), Value::I32(3)]);
    assert_eq!(len_of(&sliced), 2);
    assert_eq!(element_at(&sliced, 0).as_i32(), 2);
    assert_eq!(element_at(&sliced, 1).as_i32(), 3);
    // Original unchanged
    assert_eq!(len_of(&arr), 4);
}

#[test]
fn array_concat_appends_second_into_first() {
    let a = new_array(vec![Value::I32(1), Value::I32(2)]);
    let b = new_array(vec![Value::I32(3), Value::I32(4)]);
    let out = call_import("wasm:js-array", "concat", vec![a, b]);
    assert_eq!(len_of(&out), 4);
    assert_eq!(element_at(&out, 3).as_i32(), 4);
}

#[test]
fn array_index_of_returns_first_match_or_minus_one() {
    let arr = new_array(vec![Value::I32(10), Value::I32(20), Value::I32(10)]);
    let i = call_import("wasm:js-array", "indexOf",
        vec![arr.clone(), Value::I32(10), Value::I32(0)]);
    assert_eq!(i.as_i32(), 0);

    let i2 = call_import("wasm:js-array", "indexOf",
        vec![arr, Value::I32(99), Value::I32(0)]);
    assert_eq!(i2.as_i32(), -1);
}

#[test]
fn array_includes_returns_01_boolean() {
    let arr = new_array(vec![Value::I32(1), Value::I32(2)]);
    let hit = call_import("wasm:js-array", "includes",
        vec![arr.clone(), Value::I32(1), Value::I32(0)]);
    assert_eq!(hit.as_i32(), 1);
    let miss = call_import("wasm:js-array", "includes",
        vec![arr, Value::I32(99), Value::I32(0)]);
    assert_eq!(miss.as_i32(), 0);
}

#[test]
fn array_join_uses_separator() {
    let arr = new_array(vec![
        Value::I32(1), Value::I32(2), Value::I32(3),
    ]);
    let joined = call_import("wasm:js-array", "join",
        vec![arr, Value::String(Arc::from(","))]);
    if let Value::String(s) = joined {
        assert_eq!(s.as_ref(), "1,2,3");
    } else {
        panic!("join must return a string");
    }
}

#[test]
fn array_is_array_recognizes_arrays() {
    let arr = new_array(vec![Value::I32(1)]);
    let r = call_import("wasm:js-array", "isArray", vec![arr]);
    assert_eq!(r.as_i32(), 1);
    let not_arr = call_import("wasm:js-array", "isArray", vec![Value::I32(42)]);
    assert_eq!(not_arr.as_i32(), 0);
}

#[test]
fn array_to_reversed_returns_new_without_mutating_original() {
    let arr = new_array(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    let rev = call_import("wasm:js-array", "toReversed", vec![arr.clone()]);
    assert_eq!(element_at(&rev, 0).as_i32(), 3);
    assert_eq!(element_at(&rev, 2).as_i32(), 1);
    // Original unchanged per ECMA-262 §23.1.3.33
    assert_eq!(element_at(&arr, 0).as_i32(), 1);
    assert_eq!(element_at(&arr, 2).as_i32(), 3);
}

// ──────────────────────────────────────────────────────────────────────
// Map
// ──────────────────────────────────────────────────────────────────────

#[test]
fn map_new_empty_has_size_zero() {
    let m = call_import("wasm:js-map", "new", vec![]);
    let n = call_import("wasm:js-map", "size", vec![m]);
    assert_eq!(n.as_i32(), 0);
}

#[test]
fn map_set_get_roundtrip() {
    let m = call_import("wasm:js-map", "new", vec![]);
    let key = Value::String(Arc::from("foo"));
    call_import("wasm:js-map", "set",
        vec![m.clone(), key.clone(), Value::I32(42)]);
    let got = call_import("wasm:js-map", "get", vec![m, key]);
    assert_eq!(got.as_i32(), 42);
}

#[test]
fn map_has_reports_presence() {
    let m = call_import("wasm:js-map", "new", vec![]);
    let key = Value::String(Arc::from("k"));
    call_import("wasm:js-map", "set",
        vec![m.clone(), key.clone(), Value::I32(1)]);
    let has = call_import("wasm:js-map", "has", vec![m.clone(), key.clone()]);
    assert_eq!(has.as_i32(), 1);
    let absent = call_import("wasm:js-map", "has",
        vec![m, Value::String(Arc::from("nope"))]);
    assert_eq!(absent.as_i32(), 0);
}

#[test]
fn map_delete_then_has_returns_false() {
    let m = call_import("wasm:js-map", "new", vec![]);
    let key = Value::String(Arc::from("k"));
    call_import("wasm:js-map", "set",
        vec![m.clone(), key.clone(), Value::I32(1)]);
    let d = call_import("wasm:js-map", "delete",
        vec![m.clone(), key.clone()]);
    assert_eq!(d.as_i32(), 1, "delete must return true-as-1 for present key");
    let h = call_import("wasm:js-map", "has", vec![m, key]);
    assert_eq!(h.as_i32(), 0);
}

#[test]
fn map_size_tracks_insertions() {
    let m = call_import("wasm:js-map", "new", vec![]);
    for i in 0..5 {
        call_import("wasm:js-map", "set",
            vec![m.clone(), Value::I32(i), Value::I32(i * 10)]);
    }
    let n = call_import("wasm:js-map", "size", vec![m]);
    assert_eq!(n.as_i32(), 5);
}

// ──────────────────────────────────────────────────────────────────────
// Set
// ──────────────────────────────────────────────────────────────────────

#[test]
fn set_add_is_idempotent() {
    let s = call_import("wasm:js-set", "new", vec![]);
    call_import("wasm:js-set", "add", vec![s.clone(), Value::I32(1)]);
    call_import("wasm:js-set", "add", vec![s.clone(), Value::I32(1)]);
    call_import("wasm:js-set", "add", vec![s.clone(), Value::I32(1)]);
    let n = call_import("wasm:js-set", "size", vec![s]);
    assert_eq!(n.as_i32(), 1, "Set must dedupe on value equality");
}

#[test]
fn set_has_and_delete() {
    let s = call_import("wasm:js-set", "new", vec![]);
    call_import("wasm:js-set", "add", vec![s.clone(), Value::I32(7)]);
    let has = call_import("wasm:js-set", "has", vec![s.clone(), Value::I32(7)]);
    assert_eq!(has.as_i32(), 1);
    let d = call_import("wasm:js-set", "delete", vec![s.clone(), Value::I32(7)]);
    assert_eq!(d.as_i32(), 1);
    let has2 = call_import("wasm:js-set", "has", vec![s, Value::I32(7)]);
    assert_eq!(has2.as_i32(), 0);
}

#[test]
fn set_union_contains_all_members() {
    let a = call_import("wasm:js-set", "new", vec![]);
    call_import("wasm:js-set", "add", vec![a.clone(), Value::I32(1)]);
    call_import("wasm:js-set", "add", vec![a.clone(), Value::I32(2)]);

    let b = call_import("wasm:js-set", "new", vec![]);
    call_import("wasm:js-set", "add", vec![b.clone(), Value::I32(2)]);
    call_import("wasm:js-set", "add", vec![b.clone(), Value::I32(3)]);

    let u = call_import("wasm:js-set", "union", vec![a, b]);
    let n = call_import("wasm:js-set", "size", vec![u]);
    assert_eq!(n.as_i32(), 3, "union of {{1,2}} and {{2,3}} must have 3 elements");
}

#[test]
fn set_intersection_keeps_only_common() {
    let a = call_import("wasm:js-set", "new", vec![]);
    for i in 1..=3 {
        call_import("wasm:js-set", "add", vec![a.clone(), Value::I32(i)]);
    }
    let b = call_import("wasm:js-set", "new", vec![]);
    for i in 2..=4 {
        call_import("wasm:js-set", "add", vec![b.clone(), Value::I32(i)]);
    }
    let int = call_import("wasm:js-set", "intersection", vec![a, b]);
    let n = call_import("wasm:js-set", "size", vec![int]);
    assert_eq!(n.as_i32(), 2);
}

// ──────────────────────────────────────────────────────────────────────
// Object
// ──────────────────────────────────────────────────────────────────────

#[test]
fn object_new_then_set_get() {
    let o = call_import("wasm:js-object", "new", vec![]);
    call_import("wasm:js-object", "set",
        vec![o.clone(), Value::String(Arc::from("x")), Value::I32(7)]);
    let v = call_import("wasm:js-object", "get",
        vec![o, Value::String(Arc::from("x"))]);
    assert_eq!(v.as_i32(), 7);
}

#[test]
fn object_get_absent_key_returns_undefined() {
    let o = call_import("wasm:js-object", "new", vec![]);
    let v = call_import("wasm:js-object", "get",
        vec![o, Value::String(Arc::from("missing"))]);
    assert!(matches!(v, Value::Undefined),
        "missing key must return undefined per spec, got {:?}", v);
}

#[test]
fn object_has_own_distinguishes_own_from_inherited() {
    // Create parent with prop, child with __proto__ = parent.
    let parent = call_import("wasm:js-object", "new", vec![]);
    call_import("wasm:js-object", "set",
        vec![parent.clone(), Value::String(Arc::from("p")), Value::I32(1)]);

    let child = call_import("wasm:js-object", "create", vec![parent]);

    // hasOwn on child for "p" — should be 0 (inherited, not own)
    let has_own = call_import("wasm:js-object", "hasOwn",
        vec![child.clone(), Value::String(Arc::from("p"))]);
    assert_eq!(has_own.as_i32(), 0,
        "hasOwn must return false for inherited properties");

    // has on child for "p" — should be 1 (walks prototype chain)
    let has = call_import("wasm:js-object", "has",
        vec![child, Value::String(Arc::from("p"))]);
    assert_eq!(has.as_i32(), 1,
        "has must return true for inherited properties");
}

#[test]
fn object_freeze_prevents_writes() {
    let o = call_import("wasm:js-object", "new", vec![]);
    call_import("wasm:js-object", "set",
        vec![o.clone(), Value::String(Arc::from("v")), Value::I32(1)]);
    call_import("wasm:js-object", "freeze", vec![o.clone()]);
    // Attempt overwrite — should silently fail (strict mode would throw,
    // but our MVP is non-strict).
    call_import("wasm:js-object", "set",
        vec![o.clone(), Value::String(Arc::from("v")), Value::I32(999)]);
    let v = call_import("wasm:js-object", "get",
        vec![o, Value::String(Arc::from("v"))]);
    assert_eq!(v.as_i32(), 1, "frozen object must reject the write");
}

#[test]
fn object_is_distinguishes_nan() {
    // Object.is(NaN, NaN) === true per spec (SameValue algorithm),
    // while NaN === NaN is false.
    let r = call_import("wasm:js-object", "is",
        vec![Value::F64(f64::NAN), Value::F64(f64::NAN)]);
    assert_eq!(r.as_i32(), 1,
        "Object.is(NaN, NaN) must be true per ECMA-262 §7.2.10");
}

#[test]
fn object_keys_returns_own_enumerable() {
    let o = call_import("wasm:js-object", "new", vec![]);
    call_import("wasm:js-object", "set",
        vec![o.clone(), Value::String(Arc::from("a")), Value::I32(1)]);
    call_import("wasm:js-object", "set",
        vec![o.clone(), Value::String(Arc::from("b")), Value::I32(2)]);
    let keys = call_import("wasm:js-object", "keys", vec![o]);
    assert_eq!(len_of(&keys), 2);
}

#[test]
fn object_php_append_auto_key_increments() {
    let o = call_import("wasm:js-object", "new", vec![]);
    let k0 = call_import("wasm:js-object", "appendAutoKey",
        vec![o.clone(), Value::I32(10)]);
    assert_eq!(k0.as_i32(), 0);
    let k1 = call_import("wasm:js-object", "appendAutoKey",
        vec![o.clone(), Value::I32(20)]);
    assert_eq!(k1.as_i32(), 1);
    // Reading back `$a[1]` (stringified index) returns the value.
    let v = call_import("wasm:js-object", "get",
        vec![o, Value::String(Arc::from("1"))]);
    assert_eq!(v.as_i32(), 20);
}

// ──────────────────────────────────────────────────────────────────────
// ArrayBuffer + DataView
// ──────────────────────────────────────────────────────────────────────

#[test]
fn arraybuffer_new_has_expected_byte_length() {
    let ab = call_import("wasm:js-arraybuffer", "new", vec![Value::I32(16)]);
    let n = call_import("wasm:js-arraybuffer", "byteLength", vec![ab]);
    assert_eq!(n.as_i32(), 16);
}

#[test]
fn arraybuffer_slice_copies_subrange() {
    let ab = call_import("wasm:js-arraybuffer", "new", vec![Value::I32(16)]);
    let slice = call_import("wasm:js-arraybuffer", "slice",
        vec![ab, Value::I32(4), Value::I32(12)]);
    let n = call_import("wasm:js-arraybuffer", "byteLength", vec![slice]);
    assert_eq!(n.as_i32(), 8);
}

#[test]
fn dataview_get_set_int32_roundtrip_little_endian() {
    let ab = call_import("wasm:js-arraybuffer", "new", vec![Value::I32(16)]);
    let dv = call_import("wasm:js-dataview", "new",
        vec![ab, Value::I32(0), Value::I32(-1)]);
    call_import("wasm:js-dataview", "setInt32",
        vec![dv.clone(), Value::I32(4), Value::I32(0x12345678), Value::I32(1)]);
    let got = call_import("wasm:js-dataview", "getInt32",
        vec![dv, Value::I32(4), Value::I32(1)]);
    assert_eq!(got.as_i32(), 0x12345678);
}

#[test]
fn dataview_get_set_int32_roundtrip_big_endian() {
    let ab = call_import("wasm:js-arraybuffer", "new", vec![Value::I32(16)]);
    let dv = call_import("wasm:js-dataview", "new",
        vec![ab, Value::I32(0), Value::I32(-1)]);
    call_import("wasm:js-dataview", "setInt32",
        vec![dv.clone(), Value::I32(0), Value::I32(0x7EAD_BEEFi32), Value::I32(0)]);
    let got = call_import("wasm:js-dataview", "getInt32",
        vec![dv, Value::I32(0), Value::I32(0)]);
    assert_eq!(got.as_i32(), 0x7EAD_BEEFi32);
}

#[test]
fn dataview_get_uint8_zero_extends() {
    let ab = call_import("wasm:js-arraybuffer", "new", vec![Value::I32(4)]);
    let dv = call_import("wasm:js-dataview", "new",
        vec![ab, Value::I32(0), Value::I32(-1)]);
    call_import("wasm:js-dataview", "setUint8",
        vec![dv.clone(), Value::I32(0), Value::I32(0xFF)]);
    let got = call_import("wasm:js-dataview", "getUint8",
        vec![dv, Value::I32(0)]);
    assert_eq!(got.as_i32(), 255,
        "getUint8 must zero-extend (0xFF → 255), not sign-extend");
}

#[test]
fn dataview_get_int8_sign_extends() {
    let ab = call_import("wasm:js-arraybuffer", "new", vec![Value::I32(4)]);
    let dv = call_import("wasm:js-dataview", "new",
        vec![ab, Value::I32(0), Value::I32(-1)]);
    call_import("wasm:js-dataview", "setInt8",
        vec![dv.clone(), Value::I32(0), Value::I32(0xFF)]);
    let got = call_import("wasm:js-dataview", "getInt8",
        vec![dv, Value::I32(0)]);
    assert_eq!(got.as_i32(), -1,
        "getInt8 must sign-extend (0xFF → -1), not zero-extend");
}

// ──────────────────────────────────────────────────────────────────────
// TypedArray
// ──────────────────────────────────────────────────────────────────────

#[test]
fn uint8array_new_with_length_zero_fills() {
    let arr = call_import("wasm:js-uint8array", "newWithLength", vec![Value::I32(5)]);
    let n = call_import("wasm:js-uint8array", "length", vec![arr.clone()]);
    assert_eq!(n.as_i32(), 5);
    // Byte 2 is 0 by default
    let b = call_import("wasm:js-uint8array", "get", vec![arr, Value::I32(2)]);
    assert_eq!(b.as_i32(), 0);
}

#[test]
fn uint8array_set_clamps_to_byte_range() {
    // Uint8Array coerces `300` to `300 & 0xFF = 44`, not 255.
    let arr = call_import("wasm:js-uint8array", "newWithLength", vec![Value::I32(4)]);
    call_import("wasm:js-uint8array", "set",
        vec![arr.clone(), Value::I32(0), Value::I32(300)]);
    let got = call_import("wasm:js-uint8array", "get",
        vec![arr, Value::I32(0)]);
    assert_eq!(got.as_i32(), 44, "Uint8Array.set must truncate to u8 via & 0xFF");
}

#[test]
fn uint8_clamped_array_saturates_out_of_range_writes() {
    // Uint8ClampedArray clamps instead of truncating.
    let arr = call_import("wasm:js-uint8clamped", "newWithLength", vec![Value::I32(4)]);
    call_import("wasm:js-uint8clamped", "set",
        vec![arr.clone(), Value::I32(0), Value::I32(300)]);
    let over = call_import("wasm:js-uint8clamped", "get",
        vec![arr.clone(), Value::I32(0)]);
    assert_eq!(over.as_i32(), 255,
        "Uint8ClampedArray must clamp 300 to 255 per ECMA-262 §23.2.3");

    call_import("wasm:js-uint8clamped", "set",
        vec![arr.clone(), Value::I32(0), Value::I32(-5)]);
    let under = call_import("wasm:js-uint8clamped", "get",
        vec![arr, Value::I32(0)]);
    assert_eq!(under.as_i32(), 0,
        "Uint8ClampedArray must clamp -5 to 0 per ECMA-262 §23.2.3");
}

#[test]
fn int8array_get_sign_extends() {
    let arr = call_import("wasm:js-int8array", "newWithLength", vec![Value::I32(4)]);
    call_import("wasm:js-int8array", "set",
        vec![arr.clone(), Value::I32(0), Value::I32(0xFF)]);
    let got = call_import("wasm:js-int8array", "get",
        vec![arr, Value::I32(0)]);
    assert_eq!(got.as_i32(), -1,
        "Int8Array.get must sign-extend 0xFF to -1");
}

#[test]
fn float64array_preserves_full_precision() {
    let arr = call_import("wasm:js-float64array", "newWithLength", vec![Value::I32(2)]);
    let target = 3.141592653589793_f64;
    call_import("wasm:js-float64array", "set",
        vec![arr.clone(), Value::I32(0), Value::F64(target)]);
    let got = call_import("wasm:js-float64array", "get",
        vec![arr, Value::I32(0)]);
    if let Value::F64(f) = got {
        assert_eq!(f, target);
    } else {
        panic!("Float64Array.get must return F64 per convention, got {:?}", got);
    }
}

#[test]
fn int32array_length_and_byte_length_consistent() {
    let arr = call_import("wasm:js-int32array", "newWithLength", vec![Value::I32(4)]);
    let len = call_import("wasm:js-int32array", "length", vec![arr.clone()]);
    assert_eq!(len.as_i32(), 4);
    let bl = call_import("wasm:js-int32array", "byteLength", vec![arr]);
    assert_eq!(bl.as_i32(), 16, "Int32Array(4).byteLength == 4 * 4 bytes");
}

// ──────────────────────────────────────────────────────────────────────
// WeakMap — object-key semantics
// ──────────────────────────────────────────────────────────────────────

#[test]
fn weakmap_rejects_primitive_keys() {
    let wm = call_import("wasm:js-weakmap", "new", vec![]);
    // Per spec: setting with a non-object key throws TypeError. MVP
    // returns the map without inserting.
    call_import("wasm:js-weakmap", "set",
        vec![wm.clone(), Value::I32(42), Value::I32(1)]);
    let has = call_import("wasm:js-weakmap", "has",
        vec![wm, Value::I32(42)]);
    assert_eq!(has.as_i32(), 0,
        "WeakMap must not store entries keyed by primitives");
}

#[test]
fn weakmap_object_key_roundtrip() {
    let wm = call_import("wasm:js-weakmap", "new", vec![]);
    let key = Value::Object(Arc::new(Mutex::new(Object::new())));
    call_import("wasm:js-weakmap", "set",
        vec![wm.clone(), key.clone(), Value::I32(99)]);
    let got = call_import("wasm:js-weakmap", "get",
        vec![wm.clone(), key.clone()]);
    assert_eq!(got.as_i32(), 99);
    let has = call_import("wasm:js-weakmap", "has", vec![wm, key]);
    assert_eq!(has.as_i32(), 1);
}

// ──────────────────────────────────────────────────────────────────────
// Identity preservation — the interop contract
// ──────────────────────────────────────────────────────────────────────

#[test]
fn externref_identity_preserved_through_push_and_at() {
    // An externref-wrapped Object round-trips through Array.push + at
    // as the SAME backing Arc — i.e. pointer identity.
    let inner = Arc::new(Mutex::new(Object::new()));
    let value = Value::Object(inner.clone());
    let arr = new_array(vec![]);
    call_import("wasm:js-array", "push", vec![arr.clone(), value.clone()]);
    let retrieved = call_import("wasm:js-array", "at",
        vec![arr, Value::I32(0)]);
    match retrieved {
        Value::Object(out) => {
            assert!(Arc::ptr_eq(&inner, &out),
                "push + at must preserve externref identity (Arc pointer equality)");
        }
        other => panic!("expected Object, got {:?}", other),
    }
}

#[test]
fn externref_identity_preserved_through_map_set_and_get() {
    let inner = Arc::new(Mutex::new(Object::new()));
    let value = Value::Object(inner.clone());
    let m = call_import("wasm:js-map", "new", vec![]);
    let key = Value::String(Arc::from("k"));
    call_import("wasm:js-map", "set",
        vec![m.clone(), key.clone(), value.clone()]);
    let got = call_import("wasm:js-map", "get", vec![m, key]);
    match got {
        Value::Object(out) => {
            assert!(Arc::ptr_eq(&inner, &out),
                "Map.set + get must preserve externref identity");
        }
        other => panic!("expected Object, got {:?}", other),
    }
}

// ──────────────────────────────────────────────────────────────────────
// Cross-view buffer sharing — the core Phase B4 invariant
// ──────────────────────────────────────────────────────────────────────
//
// ECMA-262 §23.2 requires that writes through any view of an
// `ArrayBuffer` are immediately observable through every other view
// on the same buffer. Our `Arc<Mutex<Vec<u8>>>` backing inside
// `ObjectKind::ArrayBuffer` + `ObjectKind::TypedArray` makes this
// free — but it's the **whole point** of the packed-bytes refactor,
// so it deserves explicit coverage.

#[test]
fn typedarray_write_observable_via_dataview_on_same_buffer() {
    // Build an ArrayBuffer, wrap it in both a Uint8Array and a DataView,
    // write through the Uint8Array, read back through the DataView.
    let ab = call_import("wasm:js-arraybuffer", "new", vec![Value::I32(16)]);
    let u8a = call_import("wasm:js-uint8array", "newFromBuffer",
        vec![ab.clone(), Value::I32(0), Value::I32(-1)]);
    let dv = call_import("wasm:js-dataview", "new",
        vec![ab, Value::I32(0), Value::I32(-1)]);

    // Write byte 0xAB at offset 3 via the Uint8Array view.
    call_import("wasm:js-uint8array", "set",
        vec![u8a, Value::I32(3), Value::I32(0xAB)]);

    // Read the same byte back through the DataView.
    let got = call_import("wasm:js-dataview", "getUint8",
        vec![dv, Value::I32(3)]);
    assert_eq!(got.as_i32(), 0xAB,
        "TypedArray write must be observable via DataView on the same ArrayBuffer");
}

#[test]
fn dataview_write_observable_via_typedarray_on_same_buffer() {
    // Opposite direction: write via DataView, read via TypedArray.
    let ab = call_import("wasm:js-arraybuffer", "new", vec![Value::I32(8)]);
    let dv = call_import("wasm:js-dataview", "new",
        vec![ab.clone(), Value::I32(0), Value::I32(-1)]);
    let i32a = call_import("wasm:js-int32array", "newFromBuffer",
        vec![ab, Value::I32(0), Value::I32(-1)]);

    // setInt32(offset=0, value=0x12345678, littleEndian=1)
    call_import("wasm:js-dataview", "setInt32",
        vec![dv, Value::I32(0), Value::I32(0x1234_5678), Value::I32(1)]);

    // Read element 0 from the Int32Array.
    let got = call_import("wasm:js-int32array", "get",
        vec![i32a, Value::I32(0)]);
    assert_eq!(got.as_i32(), 0x1234_5678,
        "DataView write must be observable via Int32Array on the same ArrayBuffer");
}

#[test]
fn subarray_shares_storage_with_parent() {
    // Subarray is a view over the same bytes — writes visible both ways.
    let src = call_import("wasm:js-uint8array", "newWithLength", vec![Value::I32(10)]);
    for i in 0..10 {
        call_import("wasm:js-uint8array", "set",
            vec![src.clone(), Value::I32(i), Value::I32(i * 10)]);
    }
    // subarray(2, 6) — elements 2..6, length 4
    let sub = call_import("wasm:js-uint8array", "subarray",
        vec![src.clone(), Value::I32(2), Value::I32(6)]);
    let sub_len = call_import("wasm:js-uint8array", "length", vec![sub.clone()]);
    assert_eq!(sub_len.as_i32(), 4);

    // Read element 0 of subarray — corresponds to element 2 of src (= 20)
    let v = call_import("wasm:js-uint8array", "get",
        vec![sub.clone(), Value::I32(0)]);
    assert_eq!(v.as_i32(), 20);

    // Write through the subarray — src sees it.
    call_import("wasm:js-uint8array", "set",
        vec![sub, Value::I32(0), Value::I32(99)]);
    let src_elem = call_import("wasm:js-uint8array", "get",
        vec![src, Value::I32(2)]);
    assert_eq!(src_elem.as_i32(), 99,
        "Write through subarray must be visible in the parent (shared buffer)");
}

#[test]
fn slice_does_NOT_share_storage_with_parent() {
    // Per ECMA-262 §23.2.3.24, slice() copies bytes into a new
    // buffer — writes through the slice do NOT affect the parent.
    let src = call_import("wasm:js-uint8array", "newWithLength", vec![Value::I32(5)]);
    for i in 0..5 {
        call_import("wasm:js-uint8array", "set",
            vec![src.clone(), Value::I32(i), Value::I32((i + 1) * 10)]);
    }
    let sliced = call_import("wasm:js-uint8array", "slice",
        vec![src.clone(), Value::I32(0), Value::I32(3)]);

    // Mutate the slice's element 0 — src must be unaffected.
    call_import("wasm:js-uint8array", "set",
        vec![sliced, Value::I32(0), Value::I32(0xFF)]);

    let src0 = call_import("wasm:js-uint8array", "get",
        vec![src, Value::I32(0)]);
    assert_eq!(src0.as_i32(), 10,
        "slice() must copy, not share — parent must be unchanged after slice write");
}

#[test]
fn typedarray_buffer_returns_the_underlying_arraybuffer() {
    // Creating a view and calling `.buffer` on it must return the
    // same externref the caller constructed it from.
    let ab = call_import("wasm:js-arraybuffer", "new", vec![Value::I32(16)]);
    let u8a = call_import("wasm:js-uint8array", "newFromBuffer",
        vec![ab.clone(), Value::I32(0), Value::I32(-1)]);
    let got = call_import("wasm:js-uint8array", "buffer", vec![u8a]);
    if let (Value::Object(ab_arc), Value::Object(got_arc)) = (&ab, &got) {
        assert!(Arc::ptr_eq(ab_arc, got_arc),
            "TypedArray.buffer must return the same ArrayBuffer externref it was built from");
    } else {
        panic!("expected Object values for ab and .buffer result");
    }
}

#[test]
fn two_typedarray_views_on_same_buffer_see_each_others_writes() {
    // Uint8Array and Int16Array on the same ArrayBuffer — write via
    // Uint8, read via Int16 (reinterpreted).
    let ab = call_import("wasm:js-arraybuffer", "new", vec![Value::I32(8)]);
    let u8a = call_import("wasm:js-uint8array", "newFromBuffer",
        vec![ab.clone(), Value::I32(0), Value::I32(-1)]);
    let i16a = call_import("wasm:js-int16array", "newFromBuffer",
        vec![ab, Value::I32(0), Value::I32(-1)]);

    // Write the two bytes of a little-endian 0x1234 via Uint8Array.
    call_import("wasm:js-uint8array", "set",
        vec![u8a.clone(), Value::I32(0), Value::I32(0x34)]);
    call_import("wasm:js-uint8array", "set",
        vec![u8a, Value::I32(1), Value::I32(0x12)]);

    // Read element 0 of Int16Array — should be 0x1234.
    let got = call_import("wasm:js-int16array", "get",
        vec![i16a, Value::I32(0)]);
    assert_eq!(got.as_i32(), 0x1234,
        "Int16Array.get should see the bytes written via Uint8Array on the same buffer");
}
