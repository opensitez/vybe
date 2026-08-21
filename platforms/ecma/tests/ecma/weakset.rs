//! Behaviour tests for `ecma:weakset` host imports.
//!
//! Reference: ECMA-262 §24.4 WeakSet.
//!
//! WeakSet holds object references weakly (no size, no iteration).
//! Each test covers a distinct behaviour.

use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{Object, Value};
use vybe_runtime::{Chunk, Op, VM};

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn push_arg(vm: &mut VM, chunk: &mut Chunk, value: Value) {
    match value {
        Value::I32(n) => chunk.emit_i32_const(n, 0),
        Value::I64(n) => chunk.emit_i64_const(n, 0),
        Value::F32(f) => chunk.emit_f32_const(f, 0),
        Value::F64(f) => chunk.emit_f64_const(f, 0),
        Value::Bool(b) => chunk.emit_bool_const(b, 0),
        Value::String(s) => chunk.emit_string_const(&s, 0),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
        other => {
            let global = format!(
                "__test_arg_{}",
                TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
            );
            vm.set_global_owned(global.clone(), other);
            let ci = chunk.intern_string_constant(&global);
            chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
        }
    }
}

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut chunk = Chunk::new("<ecma-weakset-test>");
    let import_idx = chunk.add_import("ecma:weakset", name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
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
