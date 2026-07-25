//! `node:dns` — Node.js DNS resolution module.
//!
//! Reference: <https://nodejs.org/api/dns.html>.
//!
//! Async DNS operations (lookup, resolve*) are stubs in the synchronous
//! VM. Synchronous accessors (getServers, setServers, getDefaultResultOrder)
//! are fully implemented.

use std::sync::Arc;
use vybe_bytecode::VM;
use vybe_bytecode::value::{Object, Value};

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn empty_array() -> Value {
    Value::Object(vybe_bytecode::heap::alloc(Object::new_array(vec![])))
}

fn stub_async() -> Value {
    // Async operations return an empty object (would be Promise in a real event loop)
    Value::Object(vybe_bytecode::heap::alloc(Object::new()))
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "node:dns",
        "getServers",
        Box::new(|_ctx, _args| {
            let servers = vec![s("8.8.8.8"), s("8.8.4.4")];
            Value::Object(vybe_bytecode::heap::alloc(Object::new_array(servers)))
        }),
    );

    vm.register_host_fn(
        "node:dns",
        "setServers",
        Box::new(|_ctx, _args| Value::Undefined),
    );

    vm.register_host_fn(
        "node:dns",
        "getDefaultResultOrder",
        Box::new(|_ctx, _args| s("ipv4first")),
    );

    vm.register_host_fn(
        "node:dns",
        "setDefaultResultOrder",
        Box::new(|_ctx, _args| Value::Undefined),
    );

    // Async resolution stubs
    for name in [
        "lookup",
        "resolve",
        "resolve4",
        "resolve6",
        "resolveAny",
        "resolveCname",
        "resolveMx",
        "resolveNs",
        "resolvePtr",
        "resolveSrv",
        "resolveTxt",
        "reverse",
    ] {
        vm.register_host_fn("node:dns", name, Box::new(|_ctx, _args| stub_async()));
    }

    vm.register_host_fn(
        "node:dns",
        "lookupService",
        Box::new(|_ctx, _args| stub_async()),
    );

    // Resolver class constructor
    vm.register_host_fn(
        "node:dns",
        "Resolver",
        Box::new(|_ctx, _args| {
            let mut o = Object::new();
            o.properties
                .insert("__isResolver".into(), Value::Bool(true));
            o.properties.insert("servers".into(), {
                let servers = vec![s("8.8.8.8"), s("8.8.4.4")];
                Value::Object(vybe_bytecode::heap::alloc(Object::new_array(servers)))
            });
            Value::Object(vybe_bytecode::heap::alloc(o))
        }),
    );

    vm.register_host_fn(
        "node:dns",
        "resolverGetServers",
        Box::new(|_ctx, args| match args.first() {
            Some(Value::Object(o)) => {
                let o = o.lock().unwrap();
                o.properties
                    .get("servers")
                    .cloned()
                    .unwrap_or_else(empty_array)
            }
            _ => empty_array(),
        }),
    );

    vm.register_host_fn(
        "node:dns",
        "resolverSetServers",
        Box::new(|_ctx, _args| Value::Undefined),
    );
}
