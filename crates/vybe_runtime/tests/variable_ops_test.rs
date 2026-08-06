//! Tests for WASM variable instructions (§5.3):
//! local.get (0x20), local.set (0x21), local.tee (0x22),
//! global.get (0x23), global.set (0x24).

use std::sync::Arc;
use vybe_runtime::chunk::GlobalInit;
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

fn run(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut c = Chunk::new("<script>");
    c.local_count = 4; // enough for all tests
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    VM::new().run(vec![c]).expect("run failed")
}

fn push_i32(c: &mut Chunk, v: i32) {
    c.emit_i32_const(v, 0);
}

// ── local.get / local.set ────────────────────────────────────────────────

#[test]
fn local_get_returns_stored_value() {
    let r = run(|c| {
        push_i32(c, 42);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    });
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn local_set_overwrites_previous() {
    let r = run(|c| {
        push_i32(c, 1);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        push_i32(c, 99);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    });
    assert_eq!(r.as_i32(), 99);
}

#[test]
fn two_distinct_local_slots() {
    let r = run(|c| {
        push_i32(c, 10);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        push_i32(c, 20);
        c.emit_op_u16(Op::LOCAL_SET, 1, 0);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_op_u16(Op::LOCAL_GET, 1, 0);
        c.emit_op(Op::I32_ADD, 0);
    });
    assert_eq!(r.as_i32(), 30);
}

#[test]
fn unset_local_get_returns_null() {
    let r = run(|c| {
        c.emit_op_u16(Op::LOCAL_GET, 3, 0);
    });
    assert!(matches!(r, Value::Null));
}

#[test]
fn local_get_out_of_range_traps_without_panicking() {
    let result = std::panic::catch_unwind(|| {
        let mut c = Chunk::new("<script>");
        c.local_count = 1;
        c.emit_op_u16(Op::LOCAL_GET, 2, 0);
        c.emit_op(Op::RETURN, 0);
        VM::new().run(vec![c])
    });

    assert!(
        matches!(result, Ok(Err(_))),
        "local.get with an out-of-range local index must trap, not panic or produce a value"
    );
}

#[test]
fn local_set_out_of_range_traps_without_panicking() {
    let result = std::panic::catch_unwind(|| {
        let mut c = Chunk::new("<script>");
        c.local_count = 1;
        push_i32(&mut c, 1);
        c.emit_op_u16(Op::LOCAL_SET, 2, 0);
        c.emit_op(Op::RETURN, 0);
        VM::new().run(vec![c])
    });

    assert!(
        matches!(result, Ok(Err(_))),
        "local.set with an out-of-range local index must trap, not panic or write outside locals"
    );
}

// ── local.tee ────────────────────────────────────────────────────────────

#[test]
fn local_tee_stores_and_leaves_on_stack() {
    let r = run(|c| {
        // init slot 0
        push_i32(c, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        // tee: stores 77, leaves it on stack
        push_i32(c, 77);
        c.emit_op_u16(Op::LOCAL_TEE, 0, 0);
        // value is still on stack
    });
    assert_eq!(r.as_i32(), 77);
}

#[test]
fn local_tee_slot_holds_value_after_pop() {
    let r = run(|c| {
        push_i32(c, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        push_i32(c, 55);
        c.emit_op_u16(Op::LOCAL_TEE, 0, 0);
        c.emit_op(Op::DROP, 0); // drop the stack copy
        c.emit_op_u16(Op::LOCAL_GET, 0, 0); // retrieve from slot
    });
    assert_eq!(r.as_i32(), 55);
}

// ── global.get / global.set ──────────────────────────────────────────────

#[test]
fn global_set_and_get_roundtrip() {
    let mut c = Chunk::new("<script>");
    let name_k = c.add_constant(Value::String(Arc::from("__x")));
    push_i32(&mut c, 42);
    c.emit_op_u16(Op::GLOBAL_SET, name_k, 0);
    c.emit_op_u16(Op::GLOBAL_GET, name_k, 0);
    c.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![c]).expect("run failed");
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn global_initialized_via_global_init() {
    let mut c = Chunk::new("<script>");
    c.global_inits.push(GlobalInit {
        name: "__g".to_string(),
        init: vybe_runtime::chunk::ConstExpr::Value(Value::I32(99)),
    });
    let name_k = c.add_constant(Value::String(Arc::from("__g")));
    c.emit_op_u16(Op::GLOBAL_GET, name_k, 0);
    c.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![c]).expect("run failed");
    assert_eq!(r.as_i32(), 99);
}

#[test]
fn global_set_overwrites_init_value() {
    let mut c = Chunk::new("<script>");
    c.global_inits.push(GlobalInit {
        name: "__h".to_string(),
        init: vybe_runtime::chunk::ConstExpr::Value(Value::I32(1)),
    });
    let name_k = c.add_constant(Value::String(Arc::from("__h")));
    push_i32(&mut c, 100);
    c.emit_op_u16(Op::GLOBAL_SET, name_k, 0);
    c.emit_op_u16(Op::GLOBAL_GET, name_k, 0);
    c.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![c]).expect("run failed");
    assert_eq!(r.as_i32(), 100);
}

#[test]
fn missing_global_get_returns_undefined() {
    let mut c = Chunk::new("<script>");
    let name_k = c.add_constant(Value::String(Arc::from("__missing")));
    c.emit_op_u16(Op::GLOBAL_GET, name_k, 0);
    c.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![c]).expect("run failed");
    assert!(matches!(r, Value::Undefined));
}

#[test]
fn global_set_consumes_value_and_stores() {
    // WASM §4.4.5.3: global.set pops; verify value is stored correctly.
    let mut c = Chunk::new("<script>");
    let name_k = c.add_constant(Value::String(Arc::from("__stack")));
    push_i32(&mut c, 12);
    c.emit_op_u16(Op::GLOBAL_SET, name_k, 0);
    c.emit_op_u16(Op::GLOBAL_GET, name_k, 0);
    c.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![c]).expect("run failed");
    assert_eq!(r.as_i32(), 12);
}

// ── WASM spec compliance: stack effects ─────────────────────────────────
//
// WASM §4.4.5.4: local.set CONSUMES (pops) the value from the stack.
// WASM §4.4.5.5: local.tee PEEKS — stores AND leaves the value on stack.
// WASM §4.4.5.3: global.set CONSUMES the value.

#[test]
fn wasm_local_set_consumes_value_from_stack() {
    // WASM §4.4.5.4: local.set pops the value.
    // push 10 → push 20 → local.set 0 → return
    // set pops 20, stores in local[0], stack has [10] → returns 10.
    let r = run(|c| {
        push_i32(c, 10);
        push_i32(c, 20);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
    });
    assert_eq!(
        r.as_i32(),
        10,
        "local.set must pop the value (WASM §4.4.5.4); the underlying 10 should remain"
    );
}

#[test]
fn wasm_local_set_roundtrip_without_drop() {
    // After local.set, the value is consumed. local.get retrieves it.
    let r = run(|c| {
        push_i32(c, 42);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    });
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn wasm_local_tee_keeps_value_on_stack() {
    // WASM §4.4.5.5: local.tee stores AND leaves the value.
    let r = run(|c| {
        push_i32(c, 42);
        c.emit_op_u16(Op::LOCAL_TEE, 0, 0);
    });
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn wasm_global_set_consumes_value_from_stack() {
    // WASM §4.4.5.3: global.set pops the value.
    // push 10 → push 20 → global.set → return
    // set pops 20, stack has [10] → returns 10.
    let mut c = Chunk::new("<script>");
    c.local_count = 4;
    let name_k = c.add_constant(Value::String(Arc::from("__g_pop_test")));
    push_i32(&mut c, 10);
    push_i32(&mut c, 20);
    c.emit_op_u16(Op::GLOBAL_SET, name_k, 0);
    c.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![c]).expect("run failed");
    assert_eq!(
        r.as_i32(),
        10,
        "global.set must pop the value (WASM §4.4.5.3); the underlying 10 should remain"
    );
}
