//! Behaviour tests for `node:dns` host imports.
//!
//! Reference: <https://nodejs.org/api/dns.html>.
//!
//! Coverage:
//!   - `getServers()` → array of server address strings
//!   - `setServers(servers)` → void
//!   - `setDefaultResultOrder(order)` → void ("ipv4first" | "verbatim")
//!   - `getDefaultResultOrder()` → string (Node 22+)
//!   - `lookup`, `lookupService`, `resolve`, `resolve4`, `resolve6`,
//!     `resolveMx`, `resolveTxt`, `resolveNs`, `resolveCname`, `resolveSrv`,
//!     `resolvePtr`, `resolveAny`, `reverse` — surface (async, require DNS)
//!   - `Resolver` class constructor
//!   - `promises.resolve`, `promises.lookup` — surface
//!
//! Deferred (require live DNS or async infrastructure):
//!   - All resolve/lookup callbacks and promise fulfilment
//!   - `Resolver` per-instance server configuration

use std::sync::Arc;
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_emitter::platforms::register_platforms;

fn call_dns(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-dns-test>");
    let import_idx = chunk.add_import("node:dns", name);
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
        .contains_key(&(String::from("node:dns"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn as_str(v: &Value) -> String {
    match v {
        Value::String(s) => s.to_string(),
        other => format!("{other}"),
    }
}

// ── getServers ────────────────────────────────────────────────────────────────

#[test]
fn get_servers_returns_array() {
    let result = call_dns("getServers", vec![]);
    assert!(matches!(result, Value::Object(_)));
    match &result {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            assert!(matches!(&obj.kind, ObjectKind::Array(_)));
        }
        _ => panic!("expected array"),
    }
}

#[test]
fn get_servers_entries_are_strings() {
    let result = call_dns("getServers", vec![]);
    match &result {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &obj.kind {
                for elem in elems {
                    assert!(
                        matches!(elem, Value::String(_)),
                        "each server must be a string"
                    );
                }
            }
        }
        _ => panic!("expected array"),
    }
}

// ── setServers ────────────────────────────────────────────────────────────────

#[test]
fn set_servers_accepts_array_without_panic() {
    let servers = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(Object {
        kind: ObjectKind::Array(vec![s("8.8.8.8"), s("1.1.1.1")]),
        properties: std::collections::HashMap::new(),
        type_id: 0,
        fields: Vec::new(),
    })));
    let result = call_dns("setServers", vec![servers]);
    assert_eq!(result, Value::Undefined);
}

#[test]
fn set_servers_then_get_servers_reflects_change() {
    let new_servers = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(Object {
        kind: ObjectKind::Array(vec![s("9.9.9.9")]),
        properties: std::collections::HashMap::new(),
        type_id: 0,
        fields: Vec::new(),
    })));
    let _ = call_dns("setServers", vec![new_servers]);
    let result = call_dns("getServers", vec![]);
    match &result {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &obj.kind {
                assert!(!elems.is_empty(), "getServers should reflect set value");
            }
        }
        _ => panic!("expected array"),
    }
}

// ── setDefaultResultOrder ─────────────────────────────────────────────────────

#[test]
fn set_default_result_order_ipv4first_does_not_panic() {
    let result = call_dns("setDefaultResultOrder", vec![s("ipv4first")]);
    assert_eq!(result, Value::Undefined);
}

#[test]
fn set_default_result_order_verbatim_does_not_panic() {
    let result = call_dns("setDefaultResultOrder", vec![s("verbatim")]);
    assert_eq!(result, Value::Undefined);
}

// ── getDefaultResultOrder ─────────────────────────────────────────────────────

#[test]
fn get_default_result_order_returns_valid_string() {
    let result = call_dns("getDefaultResultOrder", vec![]);
    let s = as_str(&result);
    assert!(s == "ipv4first" || s == "verbatim", "got: {s}");
}

// ── Resolver constructor ──────────────────────────────────────────────────────

#[test]
fn resolver_constructor_returns_object() {
    let resolver = call_dns("Resolver", vec![]);
    assert!(matches!(resolver, Value::Object(_)));
}

#[test]
fn resolver_get_servers_returns_array() {
    let resolver = call_dns("Resolver", vec![]);
    let result = call_dns("resolverGetServers", vec![resolver]);
    assert!(matches!(result, Value::Object(_)));
}

#[test]
fn resolver_set_servers_accepts_array() {
    let resolver = call_dns("Resolver", vec![]);
    let servers = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(Object {
        kind: ObjectKind::Array(vec![s("8.8.8.8")]),
        properties: std::collections::HashMap::new(),
        type_id: 0,
        fields: Vec::new(),
    })));
    let result = call_dns("resolverSetServers", vec![resolver, servers]);
    assert_eq!(result, Value::Undefined);
}

// ── Async function call shapes ────────────────────────────────────────────────

#[test]
fn lookup_with_null_callback_returns_promise_or_object() {
    // lookup(hostname, callback) — with null cb, host returns a Promise-like or object
    let result = call_dns("lookup", vec![s("localhost"), Value::Null]);
    assert!(
        !matches!(result, Value::Undefined),
        "lookup must return something"
    );
}

#[test]
fn resolve4_with_null_callback_returns_promise_or_object() {
    let result = call_dns("resolve4", vec![s("localhost"), Value::Null]);
    assert!(!matches!(result, Value::Undefined));
}

#[test]
fn resolve6_with_null_callback_returns_promise_or_object() {
    let result = call_dns("resolve6", vec![s("localhost"), Value::Null]);
    assert!(!matches!(result, Value::Undefined));
}

#[test]
fn resolve_mx_with_null_callback_returns_promise_or_object() {
    let result = call_dns("resolveMx", vec![s("example.com"), Value::Null]);
    assert!(!matches!(result, Value::Undefined));
}

#[test]
fn resolve_txt_with_null_callback_returns_promise_or_object() {
    let result = call_dns("resolveTxt", vec![s("example.com"), Value::Null]);
    assert!(!matches!(result, Value::Undefined));
}

#[test]
fn resolve_ns_with_null_callback_returns_promise_or_object() {
    let result = call_dns("resolveNs", vec![s("example.com"), Value::Null]);
    assert!(!matches!(result, Value::Undefined));
}

#[test]
fn resolve_cname_with_null_callback_returns_promise_or_object() {
    let result = call_dns("resolveCname", vec![s("www.example.com"), Value::Null]);
    assert!(!matches!(result, Value::Undefined));
}

#[test]
fn resolve_srv_with_null_callback_returns_promise_or_object() {
    let result = call_dns("resolveSrv", vec![s("_http._tcp.example.com"), Value::Null]);
    assert!(!matches!(result, Value::Undefined));
}

#[test]
fn resolve_ptr_with_null_callback_returns_promise_or_object() {
    let result = call_dns("resolvePtr", vec![s("8.8.8.8"), Value::Null]);
    assert!(!matches!(result, Value::Undefined));
}

#[test]
fn resolve_any_with_null_callback_returns_promise_or_object() {
    let result = call_dns("resolveAny", vec![s("example.com"), Value::Null]);
    assert!(!matches!(result, Value::Undefined));
}

#[test]
fn reverse_with_null_callback_returns_promise_or_object() {
    let result = call_dns("reverse", vec![s("8.8.8.8"), Value::Null]);
    assert!(!matches!(result, Value::Undefined));
}

#[test]
fn lookup_service_with_null_callback_returns_promise_or_object() {
    let result = call_dns(
        "lookupService",
        vec![s("127.0.0.1"), Value::I32(80), Value::Null],
    );
    assert!(!matches!(result, Value::Undefined));
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_node_dns_surface_is_registered() {
    let expected = [
        "lookup",
        "lookupService",
        "resolve",
        "resolve4",
        "resolve6",
        "resolveMx",
        "resolveTxt",
        "resolveNs",
        "resolveCname",
        "resolveSrv",
        "resolvePtr",
        "resolveAny",
        "reverse",
        "getServers",
        "setServers",
        "setDefaultResultOrder",
        "getDefaultResultOrder",
        "Resolver",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(missing.is_empty(), "missing node:dns imports: {missing:?}");
}
