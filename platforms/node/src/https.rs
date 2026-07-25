//! `node:https` — Node.js HTTPS module.
//!
//! Reference: <https://nodejs.org/api/https.html>.
//! Surface-level stubs — TLS not actually performed.

use std::sync::Arc;
use vybe_bytecode::VM;
use vybe_bytecode::value::{Object, Value};

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn stub() -> Value {
    s("")
}

fn make_empty_obj() -> Value {
    Value::Object(vybe_bytecode::heap::alloc(Object::new()))
}

fn make_client_request(method: &str, host: &str, path: &str) -> Value {
    let mut o = Object::new();
    o.properties.insert("method".into(), s(method));
    o.properties.insert("path".into(), s(path));
    o.properties.insert("host".into(), s(host));
    o.properties.insert("protocol".into(), s("https:"));
    o.properties.insert("finished".into(), Value::Bool(false));
    o.properties.insert("destroyed".into(), Value::Bool(false));
    for m in [
        "end",
        "write",
        "setHeader",
        "getHeader",
        "removeHeader",
        "destroy",
        "setTimeout",
        "abort",
        "on",
        "once",
        "emit",
        "pipe",
    ] {
        o.properties.insert(m.into(), stub());
    }
    Value::Object(vybe_bytecode::heap::alloc(o))
}

fn make_server() -> Value {
    let mut o = Object::new();
    o.properties.insert("listening".into(), Value::Bool(false));
    o.properties.insert("timeout".into(), Value::I32(0));
    o.properties
        .insert("keepAliveTimeout".into(), Value::I32(5000));
    for m in [
        "listen",
        "close",
        "address",
        "setTimeout",
        "setSecureContext",
        "getConnections",
        "on",
        "once",
        "emit",
        "ref",
        "unref",
    ] {
        o.properties.insert(m.into(), stub());
    }
    Value::Object(vybe_bytecode::heap::alloc(o))
}

fn make_agent(opts: Option<&Value>) -> Value {
    let mut o = Object::new();
    // https Agent.maxSockets defaults to Infinity (unlike http which defaults to 256)
    o.properties
        .insert("maxSockets".into(), Value::F64(f64::INFINITY));
    o.properties
        .insert("maxFreeSockets".into(), Value::I32(256));
    o.properties
        .insert("maxTotalSockets".into(), Value::F64(f64::INFINITY));
    o.properties
        .insert("maxCachedSessions".into(), Value::I32(100));
    o.properties.insert("sockets".into(), make_empty_obj());
    o.properties.insert("freeSockets".into(), make_empty_obj());
    o.properties.insert("requests".into(), make_empty_obj());
    // TLS options echoed back
    if let Some(Value::Object(opts_obj)) = opts {
        let opts_obj = opts_obj.lock().unwrap();
        for key in [
            "ca",
            "cert",
            "key",
            "rejectUnauthorized",
            "checkServerIdentity",
        ] {
            if let Some(v) = opts_obj.properties.get(key) {
                o.properties.insert(key.into(), v.clone());
            }
        }
    }
    for m in [
        "destroy",
        "getName",
        "createConnection",
        "addRequest",
        "reuseSocket",
        "removeSocket",
    ] {
        o.properties.insert(m.into(), stub());
    }
    Value::Object(vybe_bytecode::heap::alloc(o))
}

fn parse_url(url: &str) -> (String, String) {
    let stripped = url
        .trim_start_matches("https://")
        .trim_start_matches("http://");
    if let Some(pos) = stripped.find('/') {
        (stripped[..pos].to_string(), stripped[pos..].to_string())
    } else {
        (stripped.to_string(), "/".to_string())
    }
}

fn get_str_prop(args: &[Value], idx: usize, key: &str, default: &str) -> String {
    match args.get(idx) {
        Some(Value::Object(opts)) => {
            let o = opts.lock().unwrap();
            match o.properties.get(key) {
                Some(Value::String(s)) => s.to_string(),
                _ => default.to_string(),
            }
        }
        _ => default.to_string(),
    }
}

pub fn register(vm: &mut VM) {
    // createServer([opts, cb]) → Server
    vm.register_host_fn(
        "node:https",
        "createServer",
        Box::new(|_ctx, _args| make_server()),
    );

    // Server class alias
    vm.register_host_fn(
        "node:https",
        "Server",
        Box::new(|_ctx, _args| make_server()),
    );

    // request(options[, cb]) → ClientRequest
    vm.register_host_fn(
        "node:https",
        "request",
        Box::new(|_ctx, args| {
            let host = get_str_prop(args, 0, "host", "localhost");
            let path = get_str_prop(args, 0, "path", "/");
            let method = get_str_prop(args, 0, "method", "GET");
            make_client_request(&method, &host, &path)
        }),
    );

    // get(url[, opts, cb]) → ClientRequest (GET)
    vm.register_host_fn(
        "node:https",
        "get",
        Box::new(|_ctx, args| {
            let (host, path) = match args.first() {
                Some(Value::String(url)) => parse_url(url),
                _ => {
                    let host = get_str_prop(args, 0, "host", "localhost");
                    let path = get_str_prop(args, 0, "path", "/");
                    (host, path)
                }
            };
            make_client_request("GET", &host, &path)
        }),
    );

    // Agent([opts]) → Agent
    vm.register_host_fn(
        "node:https",
        "Agent",
        Box::new(|_ctx, args| make_agent(args.first())),
    );

    // globalAgent
    vm.register_host_fn(
        "node:https",
        "globalAgent",
        Box::new(|_ctx, _args| make_agent(None)),
    );
}
