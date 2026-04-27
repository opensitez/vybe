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

use vybe_bytecode::{VM, Value, HostContext};

pub fn register(vm: &mut VM) {
    vm.register_host_fn("ecma:boolean", "Boolean", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::Bool(to_boolean(args.first().unwrap_or(&Value::Undefined)))
    }));
}

fn to_boolean(v: &Value) -> bool {
    match v {
        Value::Null | Value::Undefined => false,
        Value::Bool(b) => *b,
        Value::F64(n) => *n != 0.0 && !n.is_nan(),
        Value::I32(n) => *n != 0,
        Value::I64(n) => *n != 0,
        Value::String(s) => !s.is_empty(),
        _ => true,
    }
}
