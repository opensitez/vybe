//! ECMA-262 §20.3 — Boolean.
//!
//! `Boolean(v)` performs ToBoolean (§7.1.2):
//!
//!   undefined / null / +0 / -0 / NaN / "" / false → false
//!   everything else → true
//!
//! Vybe never wraps the result in a Boolean object — there is no demand
//! for the exotic Boolean wrapper and JS code rarely relies on the
//! `typeof new Boolean(x) === "object"` distinction. The compiler emits
//! direct boolean coercion inline where possible; this host fn is the
//! dynamic-dispatch fallback.

use std::sync::{Arc, Mutex, OnceLock};
use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

static BOOLEAN_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();

pub fn shared_boolean_prototype() -> Value {
    Value::Object(
        BOOLEAN_PROTOTYPE
            .get_or_init(|| vybe_runtime::heap::alloc(Object::new()))
            .clone(),
    )
}

pub fn boxed_boolean(value: bool) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Boolean")));
    obj.properties
        .insert("__primitive".into(), Value::Bool(value));
    obj.properties
        .insert("__proto__".into(), shared_boolean_prototype());
    Value::Object(vybe_runtime::heap::alloc(obj))
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:boolean",
        "Boolean",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(to_boolean(args.first().unwrap_or(&Value::Undefined)))
        }),
    );
    vm.register_host_fn(
        "ecma:boolean",
        "new",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            boxed_boolean(to_boolean(args.first().unwrap_or(&Value::Undefined)))
        }),
    );
    vm.register_host_fn(
        "ecma:boolean",
        "toString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::String(Arc::from(
                if boolean_value(args.first().unwrap_or(&Value::Undefined)) {
                    "true"
                } else {
                    "false"
                },
            ))
        }),
    );
    vm.register_host_fn(
        "ecma:boolean",
        "valueOf",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(boolean_value(args.first().unwrap_or(&Value::Undefined)))
        }),
    );
    // Alias used by the compiler for dynamic coercion.
    vm.register_host_fn(
        "ecma:boolean",
        "toBoolean",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::Bool(to_boolean(args.first().unwrap_or(&Value::Undefined)))
        }),
    );
}

fn boolean_value(value: &Value) -> bool {
    if let Value::Object(obj) = value {
        let primitive = {
            let locked = obj.lock().unwrap();
            if matches!(locked.properties.get("__type"), Some(Value::String(tag)) if tag.as_ref() == "Boolean")
            {
                locked.properties.get("__primitive").cloned()
            } else {
                None
            }
        };
        if let Some(Value::Bool(value)) = primitive {
            return value;
        }
    }
    to_boolean(value)
}

pub fn to_boolean(v: &Value) -> bool {
    match v {
        Value::Null | Value::Undefined => false,
        Value::Bool(b) => *b,
        Value::F64(n) => *n != 0.0 && !n.is_nan(),
        Value::I32(n) => *n != 0,
        Value::I64(n) => *n != 0,
        Value::BigInt(n) => !n.is_zero(),
        Value::String(s) => !s.is_empty(),
        _ => true,
    }
}
