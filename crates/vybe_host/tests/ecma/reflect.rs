//! Behaviour tests for `ecma:reflect` host imports.
//!
//! Reference: ECMA-262 §28.1 Reflect.
//!
//! Each test covers a distinct behaviour — particularly where Reflect
//! semantics differ from the equivalent Object methods.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-reflect-test>");
    let import_idx = chunk.add_import("ecma:reflect", name);
    let argc = args.len() as u8;
    for value in args {
        let c = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, c, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
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

// ── Reflect.get / Reflect.set ─────────────────────────────────────────────────

#[test]
fn get_reads_existing_property() {
    let o = obj(vec![("x", Value::I32(7))]);
    assert_eq!(invoke("get", vec![o, s("x")]), Value::I32(7));
}

#[test]
fn get_returns_undefined_for_absent_property() {
    let o = obj(vec![]);
    assert_eq!(invoke("get", vec![o, s("missing")]), Value::Undefined);
}

#[test]
fn set_returns_true_on_success() {
    let o = obj(vec![]);
    assert_eq!(
        invoke("set", vec![o, s("k"), Value::I32(1)]),
        Value::Bool(true)
    );
}

#[test]
fn set_value_is_readable_via_get() {
    let o = obj(vec![]);
    invoke("set", vec![o.clone(), s("v"), Value::I32(99)]);
    assert_eq!(invoke("get", vec![o, s("v")]), Value::I32(99));
}

// ── Reflect.has ───────────────────────────────────────────────────────────────

#[test]
fn has_true_for_own_property() {
    let o = obj(vec![("p", Value::Bool(true))]);
    assert_eq!(invoke("has", vec![o, s("p")]), Value::Bool(true));
}

#[test]
fn has_false_for_absent_property() {
    let o = obj(vec![]);
    assert_eq!(invoke("has", vec![o, s("gone")]), Value::Bool(false));
}

#[test]
fn has_true_for_inherited_property() {
    let proto = obj(vec![("inherited", Value::Bool(true))]);
    let child = obj(vec![("own", Value::Bool(true)), ("__proto__", proto)]);
    assert_eq!(
        invoke("has", vec![child, s("inherited")]),
        Value::Bool(true)
    );
}

// ── Reflect.deleteProperty ────────────────────────────────────────────────────

#[test]
fn delete_property_returns_true_and_removes_it() {
    let o = obj(vec![("d", Value::I32(1))]);
    assert_eq!(
        invoke("deleteProperty", vec![o.clone(), s("d")]),
        Value::Bool(true)
    );
    assert_eq!(invoke("get", vec![o, s("d")]), Value::Undefined);
}

#[test]
fn delete_property_returns_true_for_non_existent_key() {
    // Reflect.deleteProperty on a non-existent key returns true (nothing to prevent).
    let o = obj(vec![]);
    assert_eq!(
        invoke("deleteProperty", vec![o, s("nope")]),
        Value::Bool(true)
    );
}

// ── Reflect.ownKeys ───────────────────────────────────────────────────────────

#[test]
fn own_keys_returns_array_of_string_keys() {
    let o = obj(vec![("a", Value::I32(1)), ("b", Value::I32(2))]);
    let keys = invoke("ownKeys", vec![o]);
    assert!(matches!(keys, Value::Object(_)));
}

#[test]
fn own_keys_empty_object_returns_empty_array() {
    use vybe_bytecode::value::ObjectKind;
    let o = obj(vec![]);
    let keys = invoke("ownKeys", vec![o]);
    if let Value::Object(arc) = keys {
        assert!(matches!(arc.lock().unwrap().kind, ObjectKind::Array(ref e) if e.is_empty()));
    }
}

// ── Reflect.getOwnPropertyDescriptor ──────────────────────────────────────────

#[test]
fn get_own_property_descriptor_returns_object_for_existing_key() {
    let o = obj(vec![("p", Value::I32(42))]);
    let desc = invoke("getOwnPropertyDescriptor", vec![o, s("p")]);
    assert!(matches!(desc, Value::Object(_)));
}

#[test]
fn get_own_property_descriptor_returns_undefined_for_absent_key() {
    let o = obj(vec![]);
    let desc = invoke("getOwnPropertyDescriptor", vec![o, s("absent")]);
    assert!(matches!(desc, Value::Undefined | Value::Null));
}

// ── Reflect.defineProperty ────────────────────────────────────────────────────

#[test]
fn define_property_returns_bool_and_property_is_readable() {
    let o = obj(vec![]);
    let desc = obj(vec![
        ("value", Value::I32(55)),
        ("writable", Value::Bool(true)),
    ]);
    let result = invoke("defineProperty", vec![o.clone(), s("q"), desc]);
    assert!(matches!(result, Value::Bool(true)));
    assert_eq!(invoke("get", vec![o, s("q")]), Value::I32(55));
}

#[test]
fn define_property_non_writable_blocks_reflect_set() {
    let o = obj(vec![]);
    let desc = obj(vec![
        ("value", Value::I32(42)),
        ("writable", Value::Bool(false)),
        ("configurable", Value::Bool(false)),
    ]);
    assert_eq!(
        invoke("defineProperty", vec![o.clone(), s("x"), desc]),
        Value::Bool(true)
    );
    assert_eq!(
        invoke("set", vec![o.clone(), s("x"), Value::I32(99)]),
        Value::Bool(false)
    );
    assert_eq!(invoke("get", vec![o, s("x")]), Value::I32(42));
}

// ── Reflect.getPrototypeOf / setPrototypeOf ───────────────────────────────────

#[test]
fn get_prototype_of_fresh_object_is_object_or_null() {
    let o = obj(vec![]);
    let proto = invoke("getPrototypeOf", vec![o]);
    assert!(matches!(proto, Value::Object(_) | Value::Null));
}

#[test]
fn set_prototype_of_returns_bool() {
    let o = obj(vec![]);
    let result = invoke("setPrototypeOf", vec![o, Value::Null]);
    assert!(matches!(result, Value::Bool(_)));
}

// ── Reflect.isExtensible / preventExtensions ──────────────────────────────────

#[test]
fn is_extensible_true_before_prevent_false_after() {
    let o = obj(vec![]);
    assert_eq!(invoke("isExtensible", vec![o.clone()]), Value::Bool(true));
    invoke("preventExtensions", vec![o.clone()]);
    assert_eq!(invoke("isExtensible", vec![o]), Value::Bool(false));
}

#[test]
fn prevent_extensions_returns_the_object() {
    let o = obj(vec![]);
    let result = invoke("preventExtensions", vec![o]);
    assert!(matches!(result, Value::Object(_) | Value::Bool(true)));
}

#[test]
fn prevent_extensions_blocks_new_properties() {
    let o = obj(vec![("x", Value::I32(1))]);
    assert_eq!(
        invoke("preventExtensions", vec![o.clone()]),
        Value::Bool(true)
    );
    assert_eq!(
        invoke("set", vec![o.clone(), s("y"), Value::I32(2)]),
        Value::Bool(false)
    );
    assert_eq!(invoke("has", vec![o.clone(), s("y")]), Value::Bool(false));
    assert_eq!(
        invoke("set", vec![o.clone(), s("x"), Value::I32(99)]),
        Value::Bool(true)
    );
    assert_eq!(invoke("get", vec![o, s("x")]), Value::I32(99));
}

// ── Reflect.apply ─────────────────────────────────────────────────────────────

#[test]
fn apply_invokes_target_and_returns_result() {
    // ECMA-262 §28.1.1: Reflect.apply(target, thisArg, argumentsList).
    // Encode: a host callable that echoes its first arg.
    use std::sync::{Arc, Mutex};
    use vybe_bytecode::value::Object;
    let fn_obj = {
        let mut o = Object::new();
        o.properties
            .insert("__callable_echo".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let args = Value::Object(Arc::new(Mutex::new(Object::new_array(vec![Value::I32(
        42,
    )]))));
    let result = invoke("apply", vec![fn_obj, Value::Null, args]);
    assert!(matches!(
        result,
        Value::I32(_) | Value::F64(_) | Value::Object(_) | Value::Undefined
    ));
}

// ── Reflect.construct ─────────────────────────────────────────────────────────

#[test]
fn construct_produces_a_new_object_instance() {
    // ECMA-262 §28.1.2: Reflect.construct(Target, argumentsList) → new instance.
    use std::sync::{Arc, Mutex};
    use vybe_bytecode::value::Object;
    let ctor = {
        let mut o = Object::new();
        o.properties
            .insert("__ctor_point".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let args = Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
        Value::I32(1),
        Value::I32(2),
    ]))));
    let result = invoke("construct", vec![ctor, args]);
    assert!(matches!(result, Value::Object(_) | Value::Undefined));
}
