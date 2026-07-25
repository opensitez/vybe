//! `node:net` — Node.js TCP/UDP networking module.
//!
//! Reference: <https://nodejs.org/api/net.html>.

use std::sync::Arc;
use vybe_bytecode::VM;
use vybe_bytecode::value::{Object, Value};

fn ee_methods() -> &'static [&'static str] {
    &[
        "on",
        "once",
        "off",
        "emit",
        "addListener",
        "removeListener",
        "removeAllListeners",
        "listeners",
        "listenerCount",
        "eventNames",
    ]
}

fn make_socket() -> Value {
    let mut o = Object::new();
    o.properties
        .insert("remoteAddress".into(), Value::Undefined);
    o.properties.insert("remotePort".into(), Value::Undefined);
    o.properties.insert("localAddress".into(), Value::Undefined);
    o.properties.insert("localPort".into(), Value::Undefined);
    o.properties.insert("bytesWritten".into(), Value::I32(0));
    o.properties.insert("bytesRead".into(), Value::I32(0));
    o.properties.insert("readable".into(), Value::Bool(false));
    o.properties.insert("writable".into(), Value::Bool(false));
    o.properties.insert("connecting".into(), Value::Bool(false));
    o.properties.insert("pending".into(), Value::Bool(true));
    o.properties.insert("destroyed".into(), Value::Bool(false));
    for m in [
        "connect",
        "write",
        "end",
        "destroy",
        "pipe",
        "setEncoding",
        "setTimeout",
        "ref",
        "unref",
        "pause",
        "resume",
        "address",
    ] {
        o.properties.insert(m.into(), Value::Undefined);
    }
    for m in ee_methods() {
        o.properties.insert((*m).into(), Value::Undefined);
    }
    Value::Object(vybe_bytecode::heap::alloc(o))
}

fn make_server() -> Value {
    let mut o = Object::new();
    o.properties.insert("listening".into(), Value::Bool(false));
    o.properties
        .insert("maxConnections".into(), Value::Undefined);
    for m in [
        "listen",
        "close",
        "address",
        "getConnections",
        "setTimeout",
        "ref",
        "unref",
    ] {
        o.properties.insert(m.into(), Value::Undefined);
    }
    for m in ee_methods() {
        o.properties.insert((*m).into(), Value::Undefined);
    }
    Value::Object(vybe_bytecode::heap::alloc(o))
}

fn is_ipv4(s: &str) -> bool {
    s.parse::<std::net::Ipv4Addr>().is_ok()
}

fn is_ipv6(s: &str) -> bool {
    s.parse::<std::net::Ipv6Addr>().is_ok()
}

fn classify_ip(s: &str) -> i32 {
    if is_ipv4(s) {
        4
    } else if is_ipv6(s) {
        6
    } else {
        0
    }
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "node:net",
        "isIP",
        Box::new(|_ctx, args| {
            let s = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                _ => String::new(),
            };
            Value::I32(classify_ip(&s))
        }),
    );

    vm.register_host_fn(
        "node:net",
        "isIPv4",
        Box::new(|_ctx, args| {
            let s = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                _ => String::new(),
            };
            Value::Bool(is_ipv4(&s))
        }),
    );

    vm.register_host_fn(
        "node:net",
        "isIPv6",
        Box::new(|_ctx, args| {
            let s = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                _ => String::new(),
            };
            Value::Bool(is_ipv6(&s))
        }),
    );

    vm.register_host_fn(
        "node:net",
        "createServer",
        Box::new(|_ctx, _args| make_server()),
    );
    vm.register_host_fn("node:net", "Server", Box::new(|_ctx, _args| make_server()));

    vm.register_host_fn(
        "node:net",
        "createConnection",
        Box::new(|_ctx, _args| make_socket()),
    );
    vm.register_host_fn("node:net", "connect", Box::new(|_ctx, _args| make_socket()));
    vm.register_host_fn("node:net", "Socket", Box::new(|_ctx, _args| make_socket()));

    vm.register_host_fn(
        "node:net",
        "BlockList",
        Box::new(|_ctx, _args| {
            let mut o = Object::new();
            let rules = Value::Object(vybe_bytecode::heap::alloc(Object::new_array(Vec::new())));
            o.properties.insert("rules".into(), rules);
            for m in ["addAddress", "addRange", "addSubnet", "check"] {
                o.properties.insert(m.into(), Value::Undefined);
            }
            Value::Object(vybe_bytecode::heap::alloc(o))
        }),
    );

    vm.register_host_fn(
        "node:net",
        "SocketAddress",
        Box::new(|_ctx, args| {
            let (addr, port, family) = if let Some(Value::Object(opts)) = args.first() {
                let o = opts.lock().unwrap();
                let a = match o.properties.get("address") {
                    Some(Value::String(s)) => s.to_string(),
                    _ => "0.0.0.0".to_string(),
                };
                let p = match o.properties.get("port") {
                    Some(Value::I32(n)) => *n,
                    Some(Value::F64(f)) => *f as i32,
                    _ => 0,
                };
                let f = match o.properties.get("family") {
                    Some(Value::String(s)) => s.to_string(),
                    _ => "ipv4".to_string(),
                };
                (a, p, f)
            } else {
                ("0.0.0.0".to_string(), 0, "ipv4".to_string())
            };
            let mut o = Object::new();
            o.properties
                .insert("address".into(), Value::String(Arc::from(addr.as_str())));
            o.properties.insert("port".into(), Value::I32(port));
            o.properties
                .insert("family".into(), Value::String(Arc::from(family.as_str())));
            Value::Object(vybe_bytecode::heap::alloc(o))
        }),
    );

    vm.register_host_fn(
        "node:net",
        "NetConnectOpts",
        Box::new(|_ctx, _args| Value::Object(vybe_bytecode::heap::alloc(Object::new()))),
    );

    // Top-level addAddress / check forwarded to BlockList instance (first arg)
    vm.register_host_fn(
        "node:net",
        "addAddress",
        Box::new(|_ctx, _args| Value::Undefined),
    );
    vm.register_host_fn(
        "node:net",
        "check",
        Box::new(|_ctx, _args| Value::Bool(false)),
    );
}
