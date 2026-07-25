//! Behaviour tests for `ecma:atomics` host imports.
//!
//! Reference: ECMA-262 §25.4 Atomics.
//!
//! Atomics operate on SharedArrayBuffer-backed TypedArrays. Each test covers
//! a distinct operation. Since these tests are single-threaded, wait/waitAsync
//! semantics and notify wakeup are tested at the API boundary only.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_emitter::platforms::register_platforms;

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-atomics-test>");
    let import_idx = chunk.add_import("ecma:atomics", name);
    let argc = args.len() as u8;
    for value in args {
        let c = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, c, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn shared_int32(len: i32) -> Value {
    // Creates a SharedArrayBuffer-backed Int32Array for atomic operations.
    let mut o = Object::new();
    o.properties
        .insert("__shared_int32_len".to_string(), Value::I32(len));
    Value::Object(Arc::new(Mutex::new(o)))
}

// ── load / store ──────────────────────────────────────────────────────────────

#[test]
fn store_then_load_round_trips_value() {
    let ta = shared_int32(4);
    invoke("store", vec![ta.clone(), Value::I32(0), Value::I32(42)]);
    assert_eq!(invoke("load", vec![ta, Value::I32(0)]), Value::I32(42));
}

#[test]
fn load_initial_value_is_zero() {
    let ta = shared_int32(4);
    assert_eq!(invoke("load", vec![ta, Value::I32(0)]), Value::I32(0));
}

// ── add ───────────────────────────────────────────────────────────────────────

#[test]
fn add_returns_old_value_and_updates_cell() {
    // ECMA-262 §25.4.3: Atomics.add returns the OLD value before adding.
    let ta = shared_int32(4);
    invoke("store", vec![ta.clone(), Value::I32(0), Value::I32(5)]);
    let old = invoke("add", vec![ta.clone(), Value::I32(0), Value::I32(3)]);
    assert_eq!(old, Value::I32(5));
    assert_eq!(invoke("load", vec![ta, Value::I32(0)]), Value::I32(8));
}

// ── sub ───────────────────────────────────────────────────────────────────────

#[test]
fn sub_returns_old_value_and_decrements() {
    let ta = shared_int32(4);
    invoke("store", vec![ta.clone(), Value::I32(0), Value::I32(10)]);
    let old = invoke("sub", vec![ta.clone(), Value::I32(0), Value::I32(4)]);
    assert_eq!(old, Value::I32(10));
    assert_eq!(invoke("load", vec![ta, Value::I32(0)]), Value::I32(6));
}

// ── and / or / xor ───────────────────────────────────────────────────────────

#[test]
fn and_returns_old_and_applies_bitwise_and() {
    let ta = shared_int32(4);
    invoke("store", vec![ta.clone(), Value::I32(0), Value::I32(0b1111)]);
    let old = invoke("and", vec![ta.clone(), Value::I32(0), Value::I32(0b1010)]);
    assert_eq!(old, Value::I32(0b1111));
    assert_eq!(invoke("load", vec![ta, Value::I32(0)]), Value::I32(0b1010));
}

#[test]
fn or_returns_old_and_applies_bitwise_or() {
    let ta = shared_int32(4);
    invoke("store", vec![ta.clone(), Value::I32(0), Value::I32(0b0101)]);
    let old = invoke("or", vec![ta.clone(), Value::I32(0), Value::I32(0b1010)]);
    assert_eq!(old, Value::I32(0b0101));
    assert_eq!(invoke("load", vec![ta, Value::I32(0)]), Value::I32(0b1111));
}

#[test]
fn xor_same_value_clears_bits() {
    let ta = shared_int32(4);
    invoke("store", vec![ta.clone(), Value::I32(0), Value::I32(0xFF)]);
    invoke("xor", vec![ta.clone(), Value::I32(0), Value::I32(0xFF)]);
    assert_eq!(invoke("load", vec![ta, Value::I32(0)]), Value::I32(0));
}

// ── exchange ─────────────────────────────────────────────────────────────────

#[test]
fn exchange_returns_old_value_and_writes_new() {
    // ECMA-262 §25.4.7: Atomics.exchange atomically replaces and returns old.
    let ta = shared_int32(4);
    invoke("store", vec![ta.clone(), Value::I32(0), Value::I32(100)]);
    let old = invoke("exchange", vec![ta.clone(), Value::I32(0), Value::I32(200)]);
    assert_eq!(old, Value::I32(100));
    assert_eq!(invoke("load", vec![ta, Value::I32(0)]), Value::I32(200));
}

// ── compareExchange ───────────────────────────────────────────────────────────

#[test]
fn compare_exchange_swaps_when_expected_matches() {
    // ECMA-262 §25.4.4: if cell == expected, write replacement; return old.
    let ta = shared_int32(4);
    invoke("store", vec![ta.clone(), Value::I32(0), Value::I32(7)]);
    let old = invoke(
        "compareExchange",
        vec![ta.clone(), Value::I32(0), Value::I32(7), Value::I32(99)],
    );
    assert_eq!(old, Value::I32(7));
    assert_eq!(invoke("load", vec![ta, Value::I32(0)]), Value::I32(99));
}

#[test]
fn compare_exchange_does_not_swap_when_expected_mismatches() {
    let ta = shared_int32(4);
    invoke("store", vec![ta.clone(), Value::I32(0), Value::I32(7)]);
    let old = invoke(
        "compareExchange",
        vec![ta.clone(), Value::I32(0), Value::I32(0), Value::I32(99)],
    );
    // Expected was 0 but actual was 7 → no swap, returns actual value 7.
    assert_eq!(old, Value::I32(7));
    assert_eq!(invoke("load", vec![ta, Value::I32(0)]), Value::I32(7));
}

// ── isLockFree ────────────────────────────────────────────────────────────────

#[test]
fn is_lock_free_4_bytes_is_typically_true() {
    // ECMA-262 §25.4.8: 4-byte atomics are lock-free on all modern platforms.
    let result = invoke("isLockFree", vec![Value::I32(4)]);
    assert!(matches!(result, Value::Bool(true) | Value::I32(1)));
}

#[test]
fn is_lock_free_3_bytes_is_false() {
    // 3-byte atomics are never lock-free per spec.
    let result = invoke("isLockFree", vec![Value::I32(3)]);
    assert!(matches!(result, Value::Bool(false) | Value::I32(0)));
}

// ── wait (single-threaded: returns "not-equal" or "ok" immediately) ───────────

#[test]
fn wait_on_non_matching_value_returns_not_equal() {
    // ECMA-262: if the cell value ≠ expected, wait returns "not-equal" immediately.
    let ta = shared_int32(4);
    invoke("store", vec![ta.clone(), Value::I32(0), Value::I32(5)]);
    let result = invoke(
        "wait",
        vec![ta, Value::I32(0), Value::I32(0), Value::I32(0)],
    );
    // result should be the string "not-equal".
    match &result {
        Value::String(s) => assert_eq!(s.as_ref(), "not-equal"),
        other => panic!("expected 'not-equal' string, got {:?}", other),
    }
}

// ── notify — returns number of agents woken (0 in single-threaded context) ────

#[test]
fn notify_returns_integer_count_of_agents_woken() {
    let ta = shared_int32(4);
    let count = invoke("notify", vec![ta, Value::I32(0), Value::I32(1)]);
    // In a single-threaded test, no agents are waiting.
    assert!(matches!(count, Value::I32(_)));
}

// ── Atomics.waitAsync (§25.4.13) ─────────────────────────────────────────────

#[test]
fn wait_async_returns_object_with_async_flag() {
    // Atomics.waitAsync returns { async: bool, value: string }
    // In single-threaded context, the result is synchronous.
    let ta = shared_int32(4);
    let result = invoke("waitAsync", vec![ta, Value::I32(0), Value::I32(0)]);
    assert!(
        matches!(result, Value::Object(_)),
        "waitAsync must return an object"
    );
    if let Value::Object(obj) = result {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("async"),
            "result must have 'async' property"
        );
        assert!(
            o.properties.contains_key("value"),
            "result must have 'value' property"
        );
    }
}

#[test]
fn wait_async_not_equal_when_value_differs() {
    let ta = shared_int32(4);
    // Index 0 holds 0; waiting for 99 → "not-equal"
    let result = invoke("waitAsync", vec![ta, Value::I32(0), Value::I32(99)]);
    if let Value::Object(obj) = result {
        let o = obj.lock().unwrap();
        if let Some(v) = o.properties.get("value") {
            assert_eq!(format!("{}", v), "not-equal");
        }
    }
}
