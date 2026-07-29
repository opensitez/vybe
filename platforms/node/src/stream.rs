//! `node:stream` — Node.js streams module.
//!
//! Reference: <https://nodejs.org/api/stream.html>.

use vybe_runtime::VM;
use vybe_runtime::value::{Object, ObjectKind, Value};

fn ee_methods() -> Vec<&'static str> {
    vec![
        "on",
        "once",
        "off",
        "emit",
        "addListener",
        "removeListener",
        "removeAllListeners",
        "listeners",
        "rawListeners",
        "listenerCount",
        "eventNames",
    ]
}

fn make_readable() -> Value {
    let mut o = Object::new();
    o.properties
        .insert("__isReadable".into(), Value::Bool(true));
    o.properties.insert("readable".into(), Value::Bool(true));
    o.properties.insert("destroyed".into(), Value::Bool(false));
    o.properties
        .insert("readableEnded".into(), Value::Bool(false));
    for m in ee_methods() {
        o.properties.insert(m.into(), Value::Undefined);
    }
    for m in [
        "pipe",
        "unpipe",
        "destroy",
        "pause",
        "resume",
        "read",
        "push",
        "setEncoding",
        "unshift",
        "wrap",
    ] {
        o.properties.insert(m.into(), Value::Undefined);
    }
    Value::Object(vybe_runtime::heap::alloc(o))
}

fn make_writable() -> Value {
    let mut o = Object::new();
    o.properties
        .insert("__isWritable".into(), Value::Bool(true));
    o.properties.insert("writable".into(), Value::Bool(true));
    o.properties
        .insert("writableEnded".into(), Value::Bool(false));
    o.properties.insert("destroyed".into(), Value::Bool(false));
    for m in ee_methods() {
        o.properties.insert(m.into(), Value::Undefined);
    }
    for m in [
        "write",
        "end",
        "destroy",
        "cork",
        "uncork",
        "setDefaultEncoding",
    ] {
        o.properties.insert(m.into(), Value::Undefined);
    }
    Value::Object(vybe_runtime::heap::alloc(o))
}

fn make_duplex() -> Value {
    let mut o = Object::new();
    o.properties
        .insert("__isReadable".into(), Value::Bool(true));
    o.properties
        .insert("__isWritable".into(), Value::Bool(true));
    o.properties.insert("readable".into(), Value::Bool(true));
    o.properties.insert("writable".into(), Value::Bool(true));
    o.properties.insert("destroyed".into(), Value::Bool(false));
    o.properties
        .insert("readableEnded".into(), Value::Bool(false));
    o.properties
        .insert("writableEnded".into(), Value::Bool(false));
    for m in ee_methods() {
        o.properties.insert(m.into(), Value::Undefined);
    }
    for m in [
        "pipe",
        "unpipe",
        "read",
        "push",
        "write",
        "end",
        "destroy",
        "pause",
        "resume",
        "setEncoding",
        "cork",
        "uncork",
    ] {
        o.properties.insert(m.into(), Value::Undefined);
    }
    Value::Object(vybe_runtime::heap::alloc(o))
}

fn is_readable_val(v: &Value) -> bool {
    if let Value::Object(o) = v {
        let o = o.lock().unwrap();
        matches!(o.properties.get("__isReadable"), Some(Value::Bool(true)))
    } else {
        false
    }
}

fn is_writable_val(v: &Value) -> bool {
    if let Value::Object(o) = v {
        let o = o.lock().unwrap();
        matches!(o.properties.get("__isWritable"), Some(Value::Bool(true)))
    } else {
        false
    }
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "node:stream",
        "Readable",
        Box::new(|_ctx, _args| make_readable()),
    );
    vm.register_host_fn(
        "node:stream",
        "Writable",
        Box::new(|_ctx, _args| make_writable()),
    );
    vm.register_host_fn(
        "node:stream",
        "Transform",
        Box::new(|_ctx, _args| make_duplex()),
    );
    vm.register_host_fn(
        "node:stream",
        "Duplex",
        Box::new(|_ctx, _args| make_duplex()),
    );
    vm.register_host_fn(
        "node:stream",
        "PassThrough",
        Box::new(|_ctx, _args| make_duplex()),
    );
    vm.register_host_fn(
        "node:stream",
        "Stream",
        Box::new(|_ctx, _args| {
            let mut o = Object::new();
            for m in ee_methods() {
                o.properties.insert(m.into(), Value::Undefined);
            }
            for m in ["pipe", "destroy", "pause", "resume"] {
                o.properties.insert(m.into(), Value::Undefined);
            }
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    vm.register_host_fn(
        "node:stream",
        "isReadable",
        Box::new(|_ctx, args| {
            Value::Bool(is_readable_val(args.first().unwrap_or(&Value::Undefined)))
        }),
    );

    vm.register_host_fn(
        "node:stream",
        "isWritable",
        Box::new(|_ctx, args| {
            Value::Bool(is_writable_val(args.first().unwrap_or(&Value::Undefined)))
        }),
    );

    vm.register_host_fn(
        "node:stream",
        "isDisturbed",
        Box::new(|_ctx, _args| Value::Bool(false)),
    );

    vm.register_host_fn(
        "node:stream",
        "readableFrom",
        Box::new(|_ctx, args| {
            let mut s = Object::new();
            s.properties
                .insert("__isReadable".into(), Value::Bool(true));
            s.properties.insert("readable".into(), Value::Bool(true));
            s.properties.insert("destroyed".into(), Value::Bool(false));
            // Store the source data
            if let Some(src) = args.first() {
                s.properties.insert("__source".into(), src.clone());
            }
            for m in ee_methods() {
                s.properties.insert(m.into(), Value::Undefined);
            }
            for m in ["pipe", "read", "push", "destroy", "pause", "resume"] {
                s.properties.insert(m.into(), Value::Undefined);
            }
            Value::Object(vybe_runtime::heap::alloc(s))
        }),
    );

    vm.register_host_fn(
        "node:stream",
        "pipeline",
        Box::new(|_ctx, _args| Value::Undefined),
    );
    vm.register_host_fn(
        "node:stream",
        "finished",
        Box::new(|_ctx, _args| Value::Undefined),
    );
    vm.register_host_fn(
        "node:stream",
        "addAbortSignal",
        Box::new(|_ctx, _args| Value::Undefined),
    );
    vm.register_host_fn(
        "node:stream",
        "compose",
        Box::new(|_ctx, _args| Value::Object(vybe_runtime::heap::alloc(Object::new()))),
    );
}

#[allow(dead_code)]
fn _suppress(_: ObjectKind) {}
