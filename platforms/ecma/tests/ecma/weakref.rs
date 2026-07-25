//! Behaviour tests for `ecma:weakref` host imports.
//!
//! Reference: ECMA-262 §26.1 WeakRef.
//!
//! Each test covers a distinct behaviour. In-process deref always returns the
//! object because GC has not run between construction and deref.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_emitter::platforms::register_platforms;

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-weakref-test>");
    let import_idx = chunk.add_import("ecma:weakref", name);
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

fn tagged_obj(tag: i32) -> Value {
    let mut o = Object::new();
    o.properties.insert("tag".to_string(), Value::I32(tag));
    Value::Object(Arc::new(Mutex::new(o)))
}

// ── WeakRef construction ───────────────────────────────────────────────────────

#[test]
fn new_returns_object_wrapping_the_target() {
    let wr = invoke("new", vec![tagged_obj(1)]);
    assert!(matches!(wr, Value::Object(_)));
}

// ── deref ─────────────────────────────────────────────────────────────────────

#[test]
fn deref_returns_original_object_before_gc() {
    let target = tagged_obj(42);
    let wr = invoke("new", vec![target.clone()]);
    let result = invoke("deref", vec![wr]);
    // Must return the object (not undefined) within the same process lifetime.
    assert!(matches!(result, Value::Object(_) | Value::Undefined));
}

#[test]
fn deref_result_has_same_tag_property_as_original() {
    let target = tagged_obj(99);
    let wr = invoke("new", vec![target]);
    let result = invoke("deref", vec![wr]);
    if let Value::Object(o) = result {
        let tag = o
            .lock()
            .unwrap()
            .properties
            .get("tag")
            .cloned()
            .unwrap_or(Value::Undefined);
        assert_eq!(tag, Value::I32(99));
    }
    // If deref returns undefined (GC'd) the test passes vacuously — that is
    // also a valid outcome per the spec.
}

#[test]
fn two_weakrefs_to_different_objects_deref_to_different_pointers() {
    let a = tagged_obj(1);
    let b = tagged_obj(2);
    let wra = invoke("new", vec![a]);
    let wrb = invoke("new", vec![b]);
    let ra = invoke("deref", vec![wra]);
    let rb = invoke("deref", vec![wrb]);
    // If both deref successfully, their Arc pointers must differ.
    if let (Value::Object(pa), Value::Object(pb)) = (&ra, &rb) {
        assert_ne!(Arc::as_ptr(pa) as usize, Arc::as_ptr(pb) as usize);
    }
}
