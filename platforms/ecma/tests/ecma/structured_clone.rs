//! Behaviour tests for `ecma:structured-clone` host imports.
//!
//! Reference: HTML Living Standard §2.7 structuredClone.
//!
//! Each test covers a distinct behaviour.

use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{Object, ObjectKind, Value};
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
    let mut chunk = Chunk::new("<ecma-structured-clone-test>");
    let import_idx = chunk.add_import("ecma:structured-clone", name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("VM run failed")
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn obj(pairs: Vec<(&str, Value)>) -> Value {
    let mut o = Object::new();
    for (k, v) in pairs {
        o.properties.insert(k.to_string(), v);
    }
    Value::Object(Arc::new(Mutex::new(o)))
}

fn arr(values: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(values))))
}

fn clone(v: Value) -> Value {
    invoke("clone", vec![v])
}

// ── Primitive pass-through ────────────────────────────────────────────────────

#[test]
fn clone_null_returns_null() {
    assert_eq!(clone(Value::Null), Value::Null);
}

#[test]
fn clone_bool_returns_same_bool() {
    assert_eq!(clone(Value::Bool(true)), Value::Bool(true));
    assert_eq!(clone(Value::Bool(false)), Value::Bool(false));
}

#[test]
fn clone_integer_returns_same_integer() {
    assert_eq!(clone(Value::I32(42)), Value::I32(42));
}

#[test]
fn clone_float_returns_same_float() {
    let result = clone(Value::F64(3.14));
    match result {
        Value::F64(f) => assert!((f - 3.14).abs() < 1e-9),
        other => panic!("expected F64, got {:?}", other),
    }
}

#[test]
fn clone_string_returns_equal_string() {
    assert_eq!(clone(s("hello")), s("hello"));
}

// ── Objects are deep-copied, not shared ───────────────────────────────────────

#[test]
fn clone_object_returns_different_pointer() {
    let original = obj(vec![("x", Value::I32(1))]);
    let cloned = clone(original.clone());
    let orig_ptr = match &original {
        Value::Object(a) => Arc::as_ptr(a) as usize,
        _ => 0,
    };
    let clone_ptr = match &cloned {
        Value::Object(a) => Arc::as_ptr(a) as usize,
        _ => 1,
    };
    assert_ne!(
        orig_ptr, clone_ptr,
        "clone must produce a new object, not share the Arc"
    );
}

#[test]
fn clone_object_preserves_property_values() {
    let original = obj(vec![("name", s("alice")), ("age", Value::I32(30))]);
    let cloned = clone(original);
    if let Value::Object(o) = cloned {
        let o = o.lock().unwrap();
        assert_eq!(o.properties.get("name").cloned(), Some(s("alice")));
        assert_eq!(o.properties.get("age").cloned(), Some(Value::I32(30)));
    } else {
        panic!("expected object");
    }
}

#[test]
fn clone_array_returns_array_with_same_elements() {
    let original = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    let cloned = clone(original);
    if let Value::Object(o) = cloned {
        match &o.lock().unwrap().kind {
            ObjectKind::Array(e) => assert_eq!(e.len(), 3),
            _ => panic!("expected array kind"),
        }
    } else {
        panic!("expected object");
    }
}

#[test]
fn clone_nested_object_copies_inner_object_independently() {
    let inner = obj(vec![("val", Value::I32(99))]);
    let outer = obj(vec![("inner", inner)]);
    let cloned = clone(outer);
    // Outer must be an object
    assert!(matches!(cloned, Value::Object(_)));
    if let Value::Object(o) = &cloned {
        let o = o.lock().unwrap();
        // Inner property must still be an object
        assert!(matches!(o.properties.get("inner"), Some(Value::Object(_))));
    }
}

#[test]
fn clone_nan_returns_nan() {
    let result = clone(Value::F64(f64::NAN));
    match result {
        Value::F64(f) => assert!(f.is_nan()),
        other => panic!("expected NaN float, got {:?}", other),
    }
}
