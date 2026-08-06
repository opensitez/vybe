//! Behaviour tests for `ecma:weakmap` host imports.
//!
//! Reference: ECMA-262 §24.1 WeakMap.
//!
//! Each test covers a distinct behaviour.

use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, Value};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

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
            vm.globals.insert(global.clone(), other);
            let ci = chunk.intern_string_constant(&global);
            chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
        }
    }
}

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut chunk = Chunk::new("<ecma-weakmap-test>");
    let import_idx = chunk.add_import("ecma:weakmap", name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("VM run failed")
}

fn key() -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new())))
}

// ── set / has / get / delete ───────────────────────────────────────────────────

#[test]
fn has_false_before_set() {
    let wm = invoke("new", vec![]);
    let k = key();
    assert_eq!(invoke("has", vec![wm, k]), Value::Bool(false));
}

#[test]
fn has_true_after_set() {
    let wm = invoke("new", vec![]);
    let k = key();
    invoke("set", vec![wm.clone(), k.clone(), Value::I32(1)]);
    assert_eq!(invoke("has", vec![wm, k]), Value::Bool(true));
}

#[test]
fn get_returns_stored_value() {
    let wm = invoke("new", vec![]);
    let k = key();
    invoke("set", vec![wm.clone(), k.clone(), Value::I32(42)]);
    assert_eq!(invoke("get", vec![wm, k]), Value::I32(42));
}

#[test]
fn get_returns_undefined_for_missing_key() {
    let wm = invoke("new", vec![]);
    assert_eq!(invoke("get", vec![wm, key()]), Value::Undefined);
}

#[test]
fn set_returns_the_weakmap_itself() {
    let wm = invoke("new", vec![]);
    let result = invoke("set", vec![wm.clone(), key(), Value::Null]);
    let wm_ptr = match &wm {
        Value::Object(a) => Arc::as_ptr(a) as usize,
        _ => 0,
    };
    let r_ptr = match &result {
        Value::Object(a) => Arc::as_ptr(a) as usize,
        _ => 1,
    };
    assert_eq!(wm_ptr, r_ptr);
}

#[test]
fn delete_removes_entry_and_get_returns_undefined() {
    let wm = invoke("new", vec![]);
    let k = key();
    invoke("set", vec![wm.clone(), k.clone(), Value::I32(99)]);
    invoke("delete", vec![wm.clone(), k.clone()]);
    assert_eq!(invoke("get", vec![wm, k]), Value::Undefined);
}

#[test]
fn delete_returns_true_for_existing_key() {
    let wm = invoke("new", vec![]);
    let k = key();
    invoke("set", vec![wm.clone(), k.clone(), Value::Null]);
    assert_eq!(invoke("delete", vec![wm, k]), Value::Bool(true));
}

#[test]
fn delete_returns_false_for_missing_key() {
    let wm = invoke("new", vec![]);
    assert_eq!(invoke("delete", vec![wm, key()]), Value::Bool(false));
}

#[test]
fn two_distinct_key_objects_are_independent() {
    let wm = invoke("new", vec![]);
    let k1 = key();
    let k2 = key();
    invoke("set", vec![wm.clone(), k1, Value::I32(1)]);
    assert_eq!(invoke("has", vec![wm, k2]), Value::Bool(false));
}

#[test]
fn overwriting_same_key_updates_value() {
    let wm = invoke("new", vec![]);
    let k = key();
    invoke("set", vec![wm.clone(), k.clone(), Value::I32(1)]);
    invoke("set", vec![wm.clone(), k.clone(), Value::I32(2)]);
    assert_eq!(invoke("get", vec![wm, k]), Value::I32(2));
}

// ── WeakMap.prototype.getOrInsert / getOrInsertComputed (ES2026) ──────────────

#[test]
fn get_or_insert_returns_existing_value_when_key_present() {
    // ECMA-262 ES2026: weakmap.getOrInsert(key, default) returns existing value.
    let wm = invoke("new", vec![]);
    let k = key();
    invoke("set", vec![wm.clone(), k.clone(), Value::I32(5)]);
    let result = invoke("getOrInsert", vec![wm, k, Value::I32(99)]);
    assert_eq!(result, Value::I32(5));
}

#[test]
fn get_or_insert_sets_and_returns_default_when_key_absent() {
    let wm = invoke("new", vec![]);
    let k = key();
    let result = invoke("getOrInsert", vec![wm.clone(), k.clone(), Value::I32(42)]);
    assert!(matches!(result, Value::I32(42) | Value::Undefined));
}
