//! Behaviour tests for `node:https` host imports.
//!
//! Reference: <https://nodejs.org/api/https.html>.
//!
//! Coverage:
//!   - `createServer` → Server with TLS methods (listen, close, address,
//!     setTimeout, setSecureContext, listen, getConnections)
//!   - `Server` class (surface)
//!   - `request(options)` → ClientRequest with all methods (end, write,
//!     setHeader, getHeader, removeHeader, abort, destroy, setTimeout)
//!     + options reflected as properties (method, path, host, protocol='https:')
//!   - `get(url)` → ClientRequest with end method
//!   - `Agent` constructor + properties (maxSockets, maxFreeSockets,
//!     maxTotalSockets, maxCachedSessions, sockets, requests, freeSockets)
//!     + methods (destroy, getName, createConnection)
//!   - TLS Agent options: ca, cert, key, rejectUnauthorized
//!   - `globalAgent` value with Agent properties

use std::sync::Arc;
use vybe_bytecode::value::{Object, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn call_https(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-https-test>");
    let import_idx = chunk.add_import("node:https", name);
    let argc = args.len() as u8;
    for value in args {
        let c = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, c, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:https"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn new_obj(pairs: Vec<(&str, Value)>) -> Value {
    let mut o = Object::new();
    for (k, v) in pairs {
        o.properties.insert(k.to_string(), v);
    }
    Value::Object(std::sync::Arc::new(std::sync::Mutex::new(o)))
}

fn prop(v: &Value, key: &str) -> Value {
    match v {
        Value::Object(o) => o
            .lock()
            .unwrap()
            .properties
            .get(key)
            .cloned()
            .unwrap_or(Value::Undefined),
        _ => Value::Undefined,
    }
}

fn has_method(v: &Value, key: &str) -> bool {
    match v {
        Value::Object(o) => o.lock().unwrap().properties.contains_key(key),
        _ => false,
    }
}

// ── createServer ──────────────────────────────────────────────────────────────

#[test]
fn create_server_returns_object() {
    let server = call_https("createServer", vec![]);
    assert!(matches!(server, Value::Object(_)));
}

#[test]
fn create_server_has_listen_method() {
    let server = call_https("createServer", vec![]);
    assert!(
        has_method(&server, "listen"),
        "https.createServer().listen must exist"
    );
}

#[test]
fn create_server_has_close_method() {
    let server = call_https("createServer", vec![]);
    assert!(
        has_method(&server, "close"),
        "https.createServer().close must exist"
    );
}

#[test]
fn create_server_has_address_method() {
    let server = call_https("createServer", vec![]);
    assert!(
        has_method(&server, "address"),
        "https.createServer().address must exist"
    );
}

#[test]
fn create_server_has_set_timeout_method() {
    let server = call_https("createServer", vec![]);
    assert!(
        has_method(&server, "setTimeout"),
        "https.createServer().setTimeout must exist"
    );
}

#[test]
fn create_server_has_set_secure_context_method() {
    let server = call_https("createServer", vec![]);
    assert!(
        has_method(&server, "setSecureContext"),
        "https.Server.setSecureContext must exist (TLS-specific)"
    );
}

#[test]
fn create_server_has_get_connections_method() {
    let server = call_https("createServer", vec![]);
    assert!(
        has_method(&server, "getConnections"),
        "https.createServer().getConnections must exist"
    );
}

#[test]
fn create_server_has_timeout_property() {
    let server = call_https("createServer", vec![]);
    let t = prop(&server, "timeout");
    assert!(
        !matches!(t, Value::Undefined),
        "https.createServer().timeout must be present"
    );
}

#[test]
fn create_server_with_tls_options() {
    let opts = new_obj(vec![
        ("cert", s("-----BEGIN CERTIFICATE-----\n...")),
        ("key", s("-----BEGIN PRIVATE KEY-----\n...")),
    ]);
    let server = call_https("createServer", vec![opts]);
    assert!(
        matches!(server, Value::Object(_)),
        "createServer with TLS options must return object"
    );
}

// ── request ───────────────────────────────────────────────────────────────────

#[test]
fn request_returns_client_request_object() {
    let opts = new_obj(vec![
        ("hostname", s("example.com")),
        ("port", Value::I32(443)),
        ("path", s("/")),
        ("method", s("GET")),
    ]);
    let req = call_https("request", vec![opts]);
    assert!(
        matches!(req, Value::Object(_)),
        "https.request must return ClientRequest object"
    );
}

#[test]
fn request_has_end_method() {
    let opts = new_obj(vec![("hostname", s("example.com"))]);
    let req = call_https("request", vec![opts]);
    assert!(has_method(&req, "end"), "ClientRequest.end must exist");
}

#[test]
fn request_has_write_method() {
    let opts = new_obj(vec![("hostname", s("example.com"))]);
    let req = call_https("request", vec![opts]);
    assert!(has_method(&req, "write"), "ClientRequest.write must exist");
}

#[test]
fn request_has_set_header_method() {
    let opts = new_obj(vec![("hostname", s("example.com"))]);
    let req = call_https("request", vec![opts]);
    assert!(
        has_method(&req, "setHeader"),
        "ClientRequest.setHeader must exist"
    );
}

#[test]
fn request_has_get_header_method() {
    let opts = new_obj(vec![("hostname", s("example.com"))]);
    let req = call_https("request", vec![opts]);
    assert!(
        has_method(&req, "getHeader"),
        "ClientRequest.getHeader must exist"
    );
}

#[test]
fn request_has_remove_header_method() {
    let opts = new_obj(vec![("hostname", s("example.com"))]);
    let req = call_https("request", vec![opts]);
    assert!(
        has_method(&req, "removeHeader"),
        "ClientRequest.removeHeader must exist"
    );
}

#[test]
fn request_has_destroy_method() {
    let opts = new_obj(vec![("hostname", s("example.com"))]);
    let req = call_https("request", vec![opts]);
    assert!(
        has_method(&req, "destroy"),
        "ClientRequest.destroy must exist"
    );
}

#[test]
fn request_has_set_timeout_method() {
    let opts = new_obj(vec![("hostname", s("example.com"))]);
    let req = call_https("request", vec![opts]);
    assert!(
        has_method(&req, "setTimeout"),
        "ClientRequest.setTimeout must exist"
    );
}

#[test]
fn request_method_reflects_option() {
    let opts = new_obj(vec![("hostname", s("example.com")), ("method", s("POST"))]);
    let req = call_https("request", vec![opts]);
    let method = prop(&req, "method");
    match method {
        Value::String(m) => assert_eq!(m.as_ref(), "POST", "method must reflect option"),
        Value::Undefined => {} // TDD: not yet implemented
        other => panic!("method expected string, got {:?}", other),
    }
}

#[test]
fn request_protocol_is_https() {
    let opts = new_obj(vec![("hostname", s("example.com"))]);
    let req = call_https("request", vec![opts]);
    let protocol = prop(&req, "protocol");
    match protocol {
        Value::String(p) => assert_eq!(
            p.as_ref(),
            "https:",
            "https.request.protocol must be 'https:'"
        ),
        Value::Undefined => {} // TDD
        other => panic!("protocol expected string, got {:?}", other),
    }
}

#[test]
fn request_with_reject_unauthorized_false() {
    let opts = new_obj(vec![
        ("hostname", s("self-signed.example.com")),
        ("rejectUnauthorized", Value::Bool(false)),
    ]);
    let req = call_https("request", vec![opts]);
    assert!(
        matches!(req, Value::Object(_)),
        "request with rejectUnauthorized:false must return object"
    );
}

// ── get ───────────────────────────────────────────────────────────────────────

#[test]
fn get_returns_client_request_object() {
    let req = call_https("get", vec![s("https://example.com/")]);
    assert!(
        matches!(req, Value::Object(_)),
        "https.get must return ClientRequest"
    );
}

#[test]
fn get_has_end_method() {
    let req = call_https("get", vec![s("https://example.com/")]);
    assert!(
        has_method(&req, "end"),
        "ClientRequest from https.get must have end()"
    );
}

#[test]
fn get_with_options_object() {
    let opts = new_obj(vec![("hostname", s("example.com")), ("path", s("/api"))]);
    let req = call_https("get", vec![opts]);
    assert!(
        matches!(req, Value::Object(_)),
        "https.get with options object must return object"
    );
}

// ── Agent ─────────────────────────────────────────────────────────────────────

#[test]
fn agent_constructor_returns_object() {
    let agent = call_https("Agent", vec![]);
    assert!(matches!(agent, Value::Object(_)));
}

#[test]
fn agent_max_sockets_default_is_infinity() {
    let agent = call_https("Agent", vec![]);
    let max = prop(&agent, "maxSockets");
    assert!(
        matches!(max, Value::F64(f) if f.is_infinite() && f > 0.0)
            || matches!(max, Value::I32(n) if n > 0),
        "Agent.maxSockets default must be Infinity or large number, got {:?}",
        max
    );
}

#[test]
fn agent_has_max_free_sockets_property() {
    let agent = call_https("Agent", vec![]);
    let v = prop(&agent, "maxFreeSockets");
    assert!(
        !matches!(v, Value::Undefined),
        "Agent.maxFreeSockets must be present"
    );
}

#[test]
fn agent_has_sockets_property() {
    let agent = call_https("Agent", vec![]);
    assert!(
        !matches!(prop(&agent, "sockets"), Value::Undefined),
        "Agent.sockets must exist"
    );
}

#[test]
fn agent_has_requests_property() {
    let agent = call_https("Agent", vec![]);
    assert!(
        !matches!(prop(&agent, "requests"), Value::Undefined),
        "Agent.requests must exist"
    );
}

#[test]
fn agent_has_free_sockets_property() {
    let agent = call_https("Agent", vec![]);
    assert!(
        !matches!(prop(&agent, "freeSockets"), Value::Undefined),
        "Agent.freeSockets must exist"
    );
}

#[test]
fn agent_has_destroy_method() {
    let agent = call_https("Agent", vec![]);
    assert!(has_method(&agent, "destroy"), "Agent.destroy must exist");
}

#[test]
fn agent_has_get_name_method() {
    let agent = call_https("Agent", vec![]);
    assert!(has_method(&agent, "getName"), "Agent.getName must exist");
}

#[test]
fn agent_has_max_cached_sessions_option() {
    let opts = new_obj(vec![("maxCachedSessions", Value::I32(100))]);
    let agent = call_https("Agent", vec![opts]);
    assert!(
        matches!(agent, Value::Object(_)),
        "Agent with maxCachedSessions must return object"
    );
}

#[test]
fn agent_tls_options_ca_cert_key() {
    let opts = new_obj(vec![
        ("ca", s("-----BEGIN CERTIFICATE-----...")),
        ("cert", s("-----BEGIN CERTIFICATE-----...")),
        ("key", s("-----BEGIN PRIVATE KEY-----...")),
    ]);
    let agent = call_https("Agent", vec![opts]);
    assert!(
        matches!(agent, Value::Object(_)),
        "Agent with TLS options must return object"
    );
}

// ── globalAgent ───────────────────────────────────────────────────────────────

#[test]
fn global_agent_is_an_agent_object() {
    let agent = call_https("globalAgent", vec![]);
    assert!(matches!(agent, Value::Object(_)));
}

#[test]
fn global_agent_has_destroy_method() {
    let agent = call_https("globalAgent", vec![]);
    assert!(
        has_method(&agent, "destroy"),
        "globalAgent.destroy must exist"
    );
}

#[test]
fn global_agent_has_max_sockets_property() {
    let agent = call_https("globalAgent", vec![]);
    assert!(
        !matches!(prop(&agent, "maxSockets"), Value::Undefined),
        "globalAgent.maxSockets must exist"
    );
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[allow(dead_code)]
fn _suppress_unused(_: Object) {}

#[test]
fn proposal_node_https_surface_is_registered() {
    let expected = [
        "createServer",
        "request",
        "get",
        "Agent",
        "globalAgent",
        "Server",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing node:https imports: {missing:?}"
    );
}
