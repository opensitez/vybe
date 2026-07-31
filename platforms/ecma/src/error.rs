//! ECMA-262 §20.5 — Error and the seven native error subclasses.
//!
//!   §20.5.1  Error
//!   §20.5.5.1 EvalError
//!   §20.5.5.2 RangeError
//!   §20.5.5.3 ReferenceError
//!   §20.5.5.4 SyntaxError
//!   §20.5.5.5 TypeError
//!   §20.5.5.6 URIError
//!   §20.5.7   AggregateError (ES2021)
//!
//! Each constructor stamps the result Object with `__type=<Name>`,
//! `name=<Name>`, `message=<arg0>`, `stack=<Name>: <message>`. The `__type`
//! tag drives `instanceof` dispatch and try/catch matching elsewhere in
//! the VM. AggregateError additionally takes an iterable of errors as
//! its first arg (message becomes arg1).

use std::sync::Arc;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{HostContext, VM, Value};

pub fn register(vm: &mut VM) {
    for kind in [
        "Error",
        "EvalError",
        "RangeError",
        "ReferenceError",
        "SyntaxError",
        "TypeError",
        "URIError",
    ] {
        let kind_owned = kind.to_string();
        vm.register_host_fn(
            "ecma:error",
            kind,
            Box::new(move |ctx: &mut HostContext, args: &[Value]| {
                make_error(ctx, &kind_owned, args)
            }),
        );
    }

    // isError(v) → bool — checks if v is an Error-stamped object.
    vm.register_host_fn(
        "ecma:error",
        "isError",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                let is_err = o.properties.get("__exception_type").is_some()
                    || matches!(o.properties.get("__types"), Some(Value::Object(_)));
                return Value::Bool(is_err);
            }
            Value::Bool(false)
        }),
    );

    // ErrorWithCause(message, options?) → error object with optional cause.
    vm.register_host_fn(
        "ecma:error",
        "ErrorWithCause",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let message = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let cause = options_cause(args.get(1));
            let mut obj = Object::new();
            stamp_error_object(&mut obj, "Error", &message, cause);
            link_error_prototype(ctx, &mut obj, "Error");
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    vm.register_host_fn(
        "ecma:error",
        "AggregateError",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // AggregateError(errors, message?, options?)
            let this = args.first().cloned().unwrap_or(Value::Null);
            let errors = args
                .get(1)
                .cloned()
                .unwrap_or_else(|| Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![]))));
            let message = args.get(2).map(|v| format!("{}", v)).unwrap_or_default();
            let cause = options_cause(args.get(3));
            if let Value::Object(ref obj) = this {
                let mut o = obj.lock().unwrap();
                o.properties
                    .insert("__type".into(), Value::String(Arc::from("AggregateError")));
                o.properties.insert(
                    "__exception_type".into(),
                    Value::String(Arc::from("AggregateError")),
                );
                o.properties
                    .insert("name".into(), Value::String(Arc::from("AggregateError")));
                o.properties
                    .insert("message".into(), Value::String(Arc::from(message.as_str())));
                o.properties.insert(
                    "stack".into(),
                    Value::String(Arc::from(format!("AggregateError: {}", message).as_str())),
                );
                if let Some(c) = cause {
                    o.properties.insert("cause".into(), c);
                }
                // Wrap a plain iterable in an array if needed.
                if let Value::Object(ref earr) = errors {
                    let inner = earr.lock().unwrap();
                    if matches!(inner.kind, ObjectKind::Array(_)) {
                        drop(inner);
                        o.properties.insert("errors".into(), errors);
                    } else {
                        drop(inner);
                        o.properties.insert(
                            "errors".into(),
                            Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![errors]))),
                        );
                    }
                } else {
                    o.properties.insert(
                        "errors".into(),
                        Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![errors]))),
                    );
                }
                link_error_prototype(ctx, &mut o, "AggregateError");
            }
            this
        }),
    );

    vm.register_host_fn(
        "ecma:error",
        "SuppressedError",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            // SuppressedError(error, suppressed, message?)
            let this = args.first().cloned().unwrap_or(Value::Null);
            let error = args.get(1).cloned().unwrap_or(Value::Undefined);
            let suppressed = args.get(2).cloned().unwrap_or(Value::Undefined);
            let message = args.get(3).map(|v| format!("{}", v)).unwrap_or_default();
            if let Value::Object(ref obj) = this {
                let mut o = obj.lock().unwrap();
                stamp_error_object(&mut o, "SuppressedError", &message, None);
                o.properties.insert("error".into(), error);
                o.properties.insert("suppressed".into(), suppressed);
                link_error_prototype(ctx, &mut o, "SuppressedError");
            }
            this
        }),
    );

    // Error.prototype.toString() — §20.5.3.4
    // "name: message" (omit ": message" if message is empty; omit name if "Error")
    vm.register_host_fn(
        "ecma:error",
        "toString",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.lock().unwrap();
                let name = o
                    .properties
                    .get("name")
                    .map(|v| format!("{}", v))
                    .unwrap_or_else(|| "Error".to_string());
                let message = o
                    .properties
                    .get("message")
                    .map(|v| format!("{}", v))
                    .unwrap_or_default();
                let result = if message.is_empty() {
                    name
                } else {
                    format!("{}: {}", name, message)
                };
                return Value::String(Arc::from(result.as_str()));
            }
            Value::String(Arc::from("Error"))
        }),
    );
}

/// Link a host-minted error to the same per-VM prototype objects the JS
/// prelude wires onto the canonical `__ctor_<Kind>` constructors — host
/// errors and compiled errors share one prototype chain. No-op when the
/// running language declared no error prototypes.
fn link_error_prototype(ctx: &HostContext, obj: &mut Object, kind: &str) {
    if !obj.properties.contains_key("__proto__") {
        let Value::Object(ctor) = ctx.get_global(&format!("__ctor_{kind}")) else {
            return;
        };
        let Some(proto @ Value::Object(_)) =
            ctor.lock().unwrap().properties.get("prototype").cloned()
        else {
            return;
        };
        obj.properties.insert("__proto__".into(), proto);
    }
    // A chain is present (pre-linked by the compiled new-dispatch, or just
    // wired above): `name` resolves through the prototype (§20.5.3.2 — it
    // is NOT an own property: `new Error("x").hasOwnProperty("name")` is
    // false). Languages without wired prototypes keep the own stamp.
    obj.properties.shift_remove("name");
}

fn stamp_error_object(obj: &mut Object, kind: &str, message: &str, cause: Option<Value>) {
    obj.properties
        .insert("__type".into(), Value::String(Arc::from(kind)));
    obj.properties
        .insert("__exception_type".into(), Value::String(Arc::from(kind)));
    obj.properties
        .insert("name".into(), Value::String(Arc::from(kind)));
    obj.properties
        .insert("message".into(), Value::String(Arc::from(message)));
    obj.properties.insert(
        "stack".into(),
        Value::String(Arc::from(format!("{}: {}", kind, message).as_str())),
    );
    if let Some(c) = cause {
        obj.properties.insert("cause".into(), c);
    }
    let chain: Vec<Value> = error_ancestors(kind)
        .iter()
        .map(|n| Value::String(Arc::from(*n)))
        .collect();
    let chain_arr = vybe_runtime::value::Object::new_array(chain);
    obj.properties.insert(
        "__types".into(),
        Value::Object(std::sync::Arc::new(std::sync::Mutex::new(chain_arr))),
    );
}

/// §20.5 one-population constructor: every host-thrown error links the
/// SAME per-VM `__ctor_<Kind>.prototype` chain compiled `new TypeError()`
/// uses, so host-minted and compiled errors are indistinguishable
/// (`instanceof`, prototype `name`, §20.5.3.4 toString). Stamps are still
/// dual-written for the legacy readers until the Phase-5 sweep.
pub fn new_error(ctx: &HostContext, kind: &str, message: &str) -> Value {
    let mut obj = Object::new();
    stamp_error_object(&mut obj, kind, message, None);
    link_error_prototype(ctx, &mut obj, kind);
    Value::Object(vybe_runtime::heap::alloc(obj))
}

/// Stamps-only error for the rare helper with NO HostContext in reach
/// (deep recursion internals). Every throw-site constructor must use
/// `new_error` — an unlinked error is the two-populations bug.
pub fn new_error_flat(kind: &str, message: &str) -> Value {
    let mut obj = Object::new();
    stamp_error_object(&mut obj, kind, message, None);
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn make_error(ctx: &HostContext, kind: &str, args: &[Value]) -> Value {
    // Two call patterns:
    //   Compiler: args[0] = this (Object), args[1] = message, args[2] = options
    //   Direct:   args[0] = message (non-Object), args[1] = options
    if let Some(Value::Object(obj)) = args.first() {
        let message = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        let cause = options_cause(args.get(2));
        let mut o = obj.lock().unwrap();
        stamp_error_object(&mut o, kind, &message, cause);
        link_error_prototype(ctx, &mut o, kind);
        return Value::Object(obj.clone());
    }
    let message = args.first().map(|v| format!("{}", v)).unwrap_or_default();
    let cause = options_cause(args.get(1));
    let mut obj = Object::new();
    stamp_error_object(&mut obj, kind, &message, cause);
    link_error_prototype(ctx, &mut obj, kind);
    Value::Object(vybe_runtime::heap::alloc(obj))
}

/// JS error class hierarchy per ECMA-262 §20.5.5. Returned in
/// most-specific-first order (e.g. TypeError → Error).
fn error_ancestors(kind: &str) -> &'static [&'static str] {
    match kind {
        "Error" => &["Error"],
        "TypeError" => &["TypeError", "Error"],
        "RangeError" => &["RangeError", "Error"],
        "SyntaxError" => &["SyntaxError", "Error"],
        "ReferenceError" => &["ReferenceError", "Error"],
        "URIError" => &["URIError", "Error"],
        "EvalError" => &["EvalError", "Error"],
        "AggregateError" => &["AggregateError", "Error"],
        "SuppressedError" => &["SuppressedError", "Error"],
        _ => &[],
    }
}

/// ES2022 Error options: `new Error(msg, { cause })`. Returns `Some` when
/// `options` is an Object with own `cause`, regardless of its value.
fn options_cause(options: Option<&Value>) -> Option<Value> {
    let Value::Object(obj) = options? else {
        return None;
    };
    let o = obj.lock().unwrap();
    o.properties.get("cause").cloned()
}
