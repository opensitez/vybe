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

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

pub fn register(vm: &mut VM) {
    for kind in ["Error", "EvalError", "RangeError", "ReferenceError",
                 "SyntaxError", "TypeError", "URIError"] {
        let kind_owned = kind.to_string();
        vm.register_host_fn("ecma:error", kind, Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            make_error(&kind_owned, args)
        }));
    }

    vm.register_host_fn("ecma:error", "AggregateError", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        // AggregateError(errors, message?, options?)
        let this = args.first().cloned().unwrap_or(Value::Null);
        let errors = args.get(1).cloned().unwrap_or_else(|| {
            Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))))
        });
        let message = args.get(2).map(|v| format!("{}", v)).unwrap_or_default();
        let cause = options_cause(args.get(3));
        if let Value::Object(ref obj) = this {
            let mut o = obj.lock().unwrap();
            o.properties.insert("__type".into(), Value::String(Arc::from("AggregateError")));
            o.properties.insert("__exception_type".into(), Value::String(Arc::from("AggregateError")));
            o.properties.insert("name".into(), Value::String(Arc::from("AggregateError")));
            o.properties.insert("message".into(), Value::String(Arc::from(message.as_str())));
            o.properties.insert("stack".into(), Value::String(
                Arc::from(format!("AggregateError: {}", message).as_str())
            ));
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
                    o.properties.insert("errors".into(), Value::Object(
                        Arc::new(Mutex::new(Object::new_array(vec![errors])))
                    ));
                }
            } else {
                o.properties.insert("errors".into(), Value::Object(
                    Arc::new(Mutex::new(Object::new_array(vec![errors])))
                ));
            }
        }
        this
    }));
}

fn stamp_error_object(obj: &mut Object, kind: &str, message: &str, cause: Option<Value>) {
    obj.properties.insert("__type".into(), Value::String(Arc::from(kind)));
    obj.properties.insert("__exception_type".into(), Value::String(Arc::from(kind)));
    obj.properties.insert("name".into(), Value::String(Arc::from(kind)));
    obj.properties.insert("message".into(), Value::String(Arc::from(message)));
    obj.properties.insert("stack".into(), Value::String(
        Arc::from(format!("{}: {}", kind, message).as_str())
    ));
    if let Some(c) = cause {
        obj.properties.insert("cause".into(), c);
    }
    let chain: Vec<Value> = error_ancestors(kind)
        .iter()
        .map(|n| Value::String(Arc::from(*n)))
        .collect();
    let chain_arr = vybe_bytecode::value::Object::new_array(chain);
    obj.properties.insert("__types".into(),
        Value::Object(std::sync::Arc::new(std::sync::Mutex::new(chain_arr))));
}

pub(crate) fn new_error(kind: &str, message: &str) -> Value {
    let mut obj = Object::new();
    stamp_error_object(&mut obj, kind, message, None);
    Value::Object(Arc::new(Mutex::new(obj)))
}

fn make_error(kind: &str, args: &[Value]) -> Value {
    // args[0] = this (from `new`), args[1] = message, args[2] = options (ES2022)
    let this = args.first().cloned().unwrap_or(Value::Null);
    let message = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
    let cause = options_cause(args.get(2));
    if let Value::Object(ref obj) = this {
        let mut o = obj.lock().unwrap();
        stamp_error_object(&mut o, kind, &message, cause);
    }
    this
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
        _ => &[],
    }
}

/// ES2022 Error options: `new Error(msg, { cause })`. Returns `Some` when
/// `options` is an Object with own `cause`, regardless of its value.
fn options_cause(options: Option<&Value>) -> Option<Value> {
    let Value::Object(obj) = options? else { return None; };
    let o = obj.lock().unwrap();
    o.properties.get("cause").cloned()
}
