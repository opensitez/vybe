//! Behaviour tests for `ecma:weakset` host imports.
//!
//! Reference: ECMA-262 §24.4 WeakSet.
//!
//! WeakSet holds object references weakly (no size, no iteration).
//! Each test covers a distinct behaviour.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-weakset-test>");
    let import_idx = chunk.add_import("ecma:weakset", name);
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

fn obj() -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new())))
}

// ── has false before add ──────────────────────────────────────────────────────

#[test]
fn has_false_before_add() {
    let ws = invoke("new", vec![]);
    assert_eq!(invoke("has", vec![ws, obj()]), Value::Bool(false));
}

// ── add / has ─────────────────────────────────────────────────────────────────

#[test]
fn has_true_after_add() {
    let ws = invoke("new", vec![]);
    let o = obj();
    invoke("add", vec![ws.clone(), o.clone()]);
    assert_eq!(invoke("has", vec![ws, o]), Value::Bool(true));
}

#[test]
fn add_returns_the_weakset_itself() {
    // ECMA-262: WeakSet.prototype.add returns the WeakSet (chainable).
    let ws = invoke("new", vec![]);
    let result = invoke("add", vec![ws.clone(), obj()]);
    let ws_ptr = match &ws {
        Value::Object(a) => Arc::as_ptr(a) as usize,
        _ => 0,
    };
    let res_ptr = match &result {
        Value::Object(a) => Arc::as_ptr(a) as usize,
        _ => 1,
    };
    assert_eq!(ws_ptr, res_ptr);
}

// ── delete ────────────────────────────────────────────────────────────────────

#[test]
fn delete_returns_true_for_existing_member() {
    let ws = invoke("new", vec![]);
    let o = obj();
    invoke("add", vec![ws.clone(), o.clone()]);
    assert_eq!(invoke("delete", vec![ws, o]), Value::Bool(true));
}

#[test]
fn delete_returns_false_for_non_member() {
    let ws = invoke("new", vec![]);
    assert_eq!(invoke("delete", vec![ws, obj()]), Value::Bool(false));
}

#[test]
fn has_false_after_delete() {
    let ws = invoke("new", vec![]);
    let o = obj();
    invoke("add", vec![ws.clone(), o.clone()]);
    invoke("delete", vec![ws.clone(), o.clone()]);
    assert_eq!(invoke("has", vec![ws, o]), Value::Bool(false));
}

// ── two distinct objects are independent members ───────────────────────────────

#[test]
fn two_distinct_objects_tracked_independently() {
    let ws = invoke("new", vec![]);
    let a = obj();
    let b = obj();
    invoke("add", vec![ws.clone(), a.clone()]);
    assert_eq!(invoke("has", vec![ws.clone(), a]), Value::Bool(true));
    assert_eq!(invoke("has", vec![ws, b]), Value::Bool(false));
}

// ── WeakSet has no size and no iteration (spec: no .size, .forEach, .values) ──

#[test]
fn weakset_has_no_size_property() {
    // ECMA-262: WeakSet intentionally has no .size — exposing it would
    // prevent GC from collecting entries. Querying it must return Undefined.
    let ws = invoke("new", vec![]);
    assert_eq!(invoke("size", vec![ws]), Value::Undefined);
}

// ── Adding the same object twice is idempotent ────────────────────────────────

#[test]
fn adding_same_object_twice_is_idempotent() {
    let ws = invoke("new", vec![]);
    let o = obj();
    invoke("add", vec![ws.clone(), o.clone()]);
    invoke("add", vec![ws.clone(), o.clone()]);
    // Still present, and deleting once removes it.
    assert_eq!(
        invoke("delete", vec![ws.clone(), o.clone()]),
        Value::Bool(true)
    );
    assert_eq!(invoke("has", vec![ws, o]), Value::Bool(false));
}
