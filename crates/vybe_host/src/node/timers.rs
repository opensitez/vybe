//! `node:timers` — Node.js timer functions.
//!
//! Reference: <https://nodejs.org/api/timers.html>.
//!
//! In a synchronous VM, timers cannot actually fire after a delay.
//! This module returns handle objects with the required shape and
//! no-ops for the clear* functions.

use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, Value};
use vybe_bytecode::VM;

static NEXT_TIMER_ID: AtomicU64 = AtomicU64::new(1);

fn make_handle() -> Value {
    let id = NEXT_TIMER_ID.fetch_add(1, Ordering::Relaxed);
    let mut o = Object::new();
    o.properties.insert("_id".into(), Value::I64(id as i64));
    // Node timer handle methods (stubs)
    o.properties.insert("ref".into(), Value::Undefined);
    o.properties.insert("unref".into(), Value::Undefined);
    o.properties.insert("hasRef".into(), Value::Undefined);
    o.properties.insert("refresh".into(), Value::Undefined);
    Value::Object(Arc::new(Mutex::new(o)))
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn("node:timers", "setTimeout", Box::new(|_ctx, _args| {
        make_handle()
    }));

    vm.register_host_fn("node:timers", "clearTimeout", Box::new(|_ctx, _args| {
        Value::Undefined
    }));

    vm.register_host_fn("node:timers", "setInterval", Box::new(|_ctx, _args| {
        make_handle()
    }));

    vm.register_host_fn("node:timers", "clearInterval", Box::new(|_ctx, _args| {
        Value::Undefined
    }));

    vm.register_host_fn("node:timers", "setImmediate", Box::new(|_ctx, _args| {
        make_handle()
    }));

    vm.register_host_fn("node:timers", "clearImmediate", Box::new(|_ctx, _args| {
        Value::Undefined
    }));

    vm.register_host_fn("node:timers", "queueMicrotask", Box::new(|_ctx, _args| {
        Value::Undefined
    }));
}
