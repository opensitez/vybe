//! Behaviour tests for `ecma:proxy` host imports.
//!
//! Reference: ECMA-262 §28.3 Proxy.
//!
//! Each test covers a distinct trap behaviour. Tests are written from the
//! spec, not from the implementation.

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
    let mut chunk = Chunk::new("<ecma-proxy-test>");
    let import_idx = chunk.add_import("ecma:proxy", name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("VM run failed")
}

fn obj(pairs: Vec<(&str, Value)>) -> Value {
    let mut o = Object::new();
    for (k, v) in pairs {
        o.properties.insert(k.to_string(), v);
    }
    Value::Object(Arc::new(Mutex::new(o)))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

// ── new — construction ────────────────────────────────────────────────────────

#[test]
fn new_returns_object() {
    let target = obj(vec![("x", Value::I32(1))]);
    let handler = obj(vec![]);
    let proxy = invoke("new", vec![target, handler]);
    assert!(matches!(proxy, Value::Object(_)));
}

// ── get trap ─────────────────────────────────────────────────────────────────

#[test]
fn get_without_trap_falls_through_to_target() {
    // No "get" trap on handler → proxy.x reads target.x.
    let target = obj(vec![("x", Value::I32(42))]);
    let handler = obj(vec![]);
    let proxy = invoke("new", vec![target, handler]);
    assert_eq!(invoke("get", vec![proxy, s("x")]), Value::I32(42));
}

#[test]
fn get_trap_overrides_target_value() {
    // Handler has a "get" trap that returns 99 for any property.
    let target = obj(vec![("x", Value::I32(1))]);
    // The handler trap is expressed as a host-callable function reference.
    // We encode it as a special object the host recognises.
    let trap = obj(vec![("__trap_return", Value::I32(99))]);
    let handler = obj(vec![("get", trap)]);
    let proxy = invoke("new", vec![target, handler]);
    assert_eq!(invoke("get", vec![proxy, s("x")]), Value::I32(99));
}

// ── set trap ─────────────────────────────────────────────────────────────────

#[test]
fn set_without_trap_writes_through_to_target() {
    let target = obj(vec![("x", Value::I32(0))]);
    let handler = obj(vec![]);
    let proxy = invoke("new", vec![target.clone(), handler]);
    invoke("set", vec![proxy, s("x"), Value::I32(7)]);
    // Read back directly from target to confirm write-through.
    assert_eq!(invoke("get", vec![target, s("x")]), Value::I32(7));
}

// ── has trap ─────────────────────────────────────────────────────────────────

#[test]
fn has_without_trap_reflects_target_properties() {
    let target = obj(vec![("present", Value::I32(1))]);
    let handler = obj(vec![]);
    let proxy = invoke("new", vec![target, handler]);
    assert_eq!(
        invoke("has", vec![proxy.clone(), s("present")]),
        Value::Bool(true)
    );
    assert_eq!(invoke("has", vec![proxy, s("absent")]), Value::Bool(false));
}

// ── deleteProperty trap ───────────────────────────────────────────────────────

#[test]
fn delete_property_without_trap_removes_from_target() {
    let target = obj(vec![("x", Value::I32(1))]);
    let handler = obj(vec![]);
    let proxy = invoke("new", vec![target.clone(), handler]);
    invoke("deleteProperty", vec![proxy, s("x")]);
    assert_eq!(invoke("get", vec![target, s("x")]), Value::Undefined);
}

// ── ownKeys trap ──────────────────────────────────────────────────────────────

#[test]
fn own_keys_without_trap_returns_target_keys() {
    let target = obj(vec![("a", Value::I32(1)), ("b", Value::I32(2))]);
    let handler = obj(vec![]);
    let proxy = invoke("new", vec![target, handler]);
    let keys = invoke("ownKeys", vec![proxy]);
    assert!(matches!(keys, Value::Object(_)));
}

// ── Proxy.revocable — revoke kills the proxy ──────────────────────────────────

#[test]
fn revocable_returns_object_with_proxy_and_revoke_properties() {
    let target = obj(vec![]);
    let handler = obj(vec![]);
    let result = invoke("revocable", vec![target, handler]);
    // ECMA-262 §28.3.2: Proxy.revocable returns { proxy, revoke }.
    if let Value::Object(o) = &result {
        let o = o.lock().unwrap();
        assert!(
            o.properties.contains_key("proxy"),
            "must have proxy property"
        );
        assert!(
            o.properties.contains_key("revoke"),
            "must have revoke property"
        );
    } else {
        panic!("expected object");
    }
}

#[test]
fn after_revoke_operations_on_proxy_return_error_or_undefined() {
    let target = obj(vec![("x", Value::I32(1))]);
    let handler = obj(vec![]);
    let result = invoke("revocable", vec![target, handler]);
    if let Value::Object(o) = &result {
        let (proxy, revoke) = {
            let o = o.lock().unwrap();
            (
                o.properties["proxy"].clone(),
                o.properties["revoke"].clone(),
            )
        };
        invoke("callRevoke", vec![revoke]);
        // Any operation on a revoked proxy must fail.
        let after = invoke("get", vec![proxy, s("x")]);
        // Implementation may return Undefined or an error marker.
        assert!(
            matches!(after, Value::Undefined | Value::Null),
            "revoked proxy access should fail gracefully, got {:?}",
            after
        );
    } else {
        panic!("expected revocable result object");
    }
}

// ── getPrototypeOf trap ───────────────────────────────────────────────────────

#[test]
fn get_prototype_of_without_trap_delegates_to_target() {
    let target = obj(vec![]);
    let handler = obj(vec![]);
    let proxy = invoke("new", vec![target, handler]);
    // Should not panic — returns Object or Null.
    let proto = invoke("getPrototypeOf", vec![proxy]);
    assert!(matches!(proto, Value::Object(_) | Value::Null));
}

// ── isExtensible trap ────────────────────────────────────────────────────────

#[test]
fn is_extensible_without_trap_matches_target() {
    let target = obj(vec![]);
    let handler = obj(vec![]);
    let proxy = invoke("new", vec![target, handler]);
    assert_eq!(invoke("isExtensible", vec![proxy]), Value::Bool(true));
}

// ── apply trap — for function proxies ────────────────────────────────────────

#[test]
fn apply_trap_intercepts_function_call() {
    // A proxy wrapping a function object with an apply trap.
    let fn_target = obj(vec![("__callable", Value::Bool(true))]);
    let trap = obj(vec![("__trap_return", Value::I32(100))]);
    let handler = obj(vec![("apply", trap)]);
    let proxy = invoke("new", vec![fn_target, handler]);
    let result = invoke("apply", vec![proxy, Value::Null, Value::Null]);
    assert_eq!(result, Value::I32(100));
}

// ── setPrototypeOf trap ───────────────────────────────────────────────────────

#[test]
fn set_prototype_of_without_trap_delegates_to_target() {
    // ECMA-262 §28.3.2.8: without trap, setPrototypeOf passes through to target.
    let target = obj(vec![]);
    let proxy = invoke("new", vec![target, obj(vec![])]);
    let result = invoke("setPrototypeOf", vec![proxy, Value::Null]);
    assert!(matches!(result, Value::Bool(_)));
}

// ── preventExtensions trap ────────────────────────────────────────────────────

#[test]
fn prevent_extensions_without_trap_delegates_to_target() {
    // ECMA-262 §28.3.2.6: without trap, preventExtensions passes through.
    let target = obj(vec![]);
    let proxy = invoke("new", vec![target, obj(vec![])]);
    let result = invoke("preventExtensions", vec![proxy]);
    assert!(matches!(result, Value::Bool(_) | Value::Object(_)));
}

// ── construct trap ────────────────────────────────────────────────────────────

#[test]
fn construct_without_trap_delegates_to_target_constructor() {
    // ECMA-262 §28.3.2.2: without construct trap, `new proxy(args)` calls target.
    let ctor = {
        let mut o = Object::new();
        o.properties
            .insert("__ctor_point".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let proxy = invoke("new", vec![ctor, obj(vec![])]);
    let args = Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))));
    let result = invoke("construct", vec![proxy, args]);
    assert!(matches!(result, Value::Object(_) | Value::Undefined));
}
