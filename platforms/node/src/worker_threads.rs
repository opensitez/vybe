//! `node:worker_threads` — Node.js worker threads module.
//!
//! Reference: <https://nodejs.org/api/worker_threads.html>.
//!
//! In a synchronous VM, workers can't actually run. This module provides
//! the correct surface so code that imports it doesn't crash.

use std::sync::atomic::{AtomicI32, Ordering};
use std::sync::Arc;
use vybe_runtime::VM;
use vybe_runtime::value::{Object, Value};

static NEXT_WORKER_ID: AtomicI32 = AtomicI32::new(1);

fn make_port() -> Value {
    let mut o = Object::new();
    for m in [
        "postMessage",
        "close",
        "start",
        "ref",
        "unref",
        "hasRef",
        "on",
        "once",
        "off",
        "emit",
        "addListener",
        "removeListener",
        "removeAllListeners",
        "addEventListener",
        "removeEventListener",
    ] {
        o.properties.insert(m.into(), Value::Undefined);
    }
    Value::Object(vybe_runtime::heap::alloc(o))
}

fn make_resource_limits() -> Value {
    let mut o = Object::new();
    o.properties
        .insert("maxOldGenerationSizeMb".into(), Value::I32(0));
    o.properties
        .insert("maxYoungGenerationSizeMb".into(), Value::I32(0));
    o.properties.insert("codeRangeSizeMb".into(), Value::I32(0));
    o.properties.insert("stackSizeMb".into(), Value::I32(4));
    Value::Object(vybe_runtime::heap::alloc(o))
}

fn null_stream() -> Value {
    let mut o = Object::new();
    for m in ["write", "read", "on", "once", "off", "pipe", "destroy"] {
        o.properties.insert(m.into(), Value::Undefined);
    }
    Value::Object(vybe_runtime::heap::alloc(o))
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "node:worker_threads",
        "isMainThread",
        Box::new(|_ctx, _args| Value::Bool(true)),
    );

    vm.register_host_fn(
        "node:worker_threads",
        "threadId",
        Box::new(|_ctx, _args| Value::I32(0)),
    );

    vm.register_host_fn(
        "node:worker_threads",
        "workerData",
        Box::new(|_ctx, _args| Value::Null),
    );

    vm.register_host_fn(
        "node:worker_threads",
        "parentPort",
        Box::new(|_ctx, _args| Value::Null),
    );

    vm.register_host_fn(
        "node:worker_threads",
        "resourceLimits",
        Box::new(|_ctx, _args| make_resource_limits()),
    );

    vm.register_host_fn(
        "node:worker_threads",
        "SHARE_ENV",
        Box::new(|_ctx, _args| Value::String(Arc::from("Symbol(SHARE_ENV)"))),
    );

    vm.register_host_fn(
        "node:worker_threads",
        "Worker",
        Box::new(|_ctx, _args| {
            let id = NEXT_WORKER_ID.fetch_add(1, Ordering::Relaxed);
            let mut o = Object::new();
            o.properties.insert("threadId".into(), Value::I32(id));
            o.properties.insert("exitCode".into(), Value::Null);
            o.properties.insert("stdin".into(), null_stream());
            o.properties.insert("stdout".into(), null_stream());
            o.properties.insert("stderr".into(), null_stream());
            o.properties
                .insert("resourceLimits".into(), make_resource_limits());
            for m in [
                "postMessage",
                "terminate",
                "ref",
                "unref",
                "getHeapSnapshot",
                "on",
                "once",
                "off",
                "emit",
                "addListener",
                "removeListener",
                "removeAllListeners",
                "listenerCount",
            ] {
                o.properties.insert(m.into(), Value::Undefined);
            }
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    vm.register_host_fn(
        "node:worker_threads",
        "MessageChannel",
        Box::new(|_ctx, _args| {
            let mut o = Object::new();
            o.properties.insert("port1".into(), make_port());
            o.properties.insert("port2".into(), make_port());
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    vm.register_host_fn(
        "node:worker_threads",
        "MessagePort",
        Box::new(|_ctx, _args| make_port()),
    );

    vm.register_host_fn(
        "node:worker_threads",
        "BroadcastChannel",
        Box::new(|_ctx, args| {
            let name = match args.first() {
                Some(Value::String(s)) => s.clone(),
                _ => Arc::from(""),
            };
            let mut o = Object::new();
            o.properties.insert("name".into(), Value::String(name));
            o.properties.insert("onmessage".into(), Value::Null);
            o.properties.insert("onmessageerror".into(), Value::Null);
            for m in ["postMessage", "close"] {
                o.properties.insert(m.into(), Value::Undefined);
            }
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    vm.register_host_fn(
        "node:worker_threads",
        "receiveMessageOnPort",
        Box::new(|_ctx, _args| Value::Undefined),
    );

    vm.register_host_fn(
        "node:worker_threads",
        "moveMessagePortToContext",
        Box::new(|_ctx, _args| Value::Undefined),
    );

    vm.register_host_fn(
        "node:worker_threads",
        "markAsUntransferable",
        Box::new(|_ctx, _args| Value::Undefined),
    );

    vm.register_host_fn(
        "node:worker_threads",
        "isMarkedAsUntransferable",
        Box::new(|_ctx, _args| Value::Bool(false)),
    );

    vm.register_host_fn(
        "node:worker_threads",
        "setEnvironmentData",
        Box::new(|_ctx, _args| Value::Undefined),
    );

    vm.register_host_fn(
        "node:worker_threads",
        "getEnvironmentData",
        Box::new(|_ctx, _args| Value::Undefined),
    );
}
