//! Behaviour tests for Vybe's internal HTTP host layer (`vybe:http`).
//!
//! This is the shared HTTP primitive layer used by PHP, JS, and other
//! languages. It is NOT the Node.js `node:http` module. Tests here cover
//! the raw request/response context primitives that all Vybe languages
//! forward to.
//!
//! Reference: Vybe HTTP host module (`crates/vybe_host/src/modules/http.rs`).

use vybe_runtime::value::{ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn call_http(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<vybe-http-test>");
    let import_idx = chunk.add_import("node:http", name);
    let argc = args.len() as u8;
    for value in args {
        let constant = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, constant, 0);
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
        .contains_key(&(String::from("node:http"), name.to_string()))
}

fn array_len(value: &Value) -> usize {
    let Value::Object(array) = value else {
        return 0;
    };
    let array = array.lock().unwrap();
    let ObjectKind::Array(values) = &array.kind else {
        return 0;
    };
    values.len()
}

fn as_string(value: &Value) -> String {
    match value {
        Value::String(text) => text.to_string(),
        other => format!("{}", other),
    }
}

// ── Runtime mode ──────────────────────────────────────────────────────────────

// ── Request accessors — sentinel values without a live request ────────────────

#[test]
fn method_returns_empty_string_without_request() {
    assert_eq!(as_string(&call_http("method", vec![])), "");
}

#[test]
fn header_returns_null_without_request() {
    assert_eq!(
        call_http("header", vec![Value::String("x-test".into())]),
        Value::Null
    );
}

#[test]
fn headers_returns_empty_array_without_request() {
    assert_eq!(array_len(&call_http("headers", vec![])), 0);
}

#[test]
fn uri_returns_string_without_request() {
    let v = call_http("uri", vec![]);
    assert!(matches!(
        v,
        Value::String(_) | Value::Null | Value::Undefined
    ));
}

#[test]
fn remote_addr_returns_string_without_request() {
    let v = call_http("remote_addr", vec![]);
    assert!(matches!(
        v,
        Value::String(_) | Value::Null | Value::Undefined
    ));
}

// ── Response accessors — sentinel values without a live response ───────────────

#[test]
fn status_returns_zero_without_response() {
    assert_eq!(call_http("status", vec![]), Value::F64(0.0));
}

#[test]
fn headers_sent_is_false_without_response() {
    assert_eq!(call_http("headers_sent", vec![]), Value::Bool(false));
}

#[test]
fn has_header_is_false_without_response() {
    assert_eq!(
        call_http("has_header", vec![Value::String("x-test".into())]),
        Value::Bool(false)
    );
}

#[test]
fn set_status_does_not_panic() {
    assert!(matches!(
        call_http("set_status", vec![Value::F64(404.0)]),
        Value::Null | Value::Undefined
    ));
}

#[test]
fn set_header_does_not_panic() {
    assert!(matches!(
        call_http(
            "set_header",
            vec![Value::String("X-Foo".into()), Value::String("bar".into())]
        ),
        Value::Null | Value::Undefined
    ));
}

#[test]
fn add_header_does_not_panic() {
    assert!(matches!(
        call_http(
            "add_header",
            vec![Value::String("X-Foo".into()), Value::String("bar".into())]
        ),
        Value::Null | Value::Undefined
    ));
}

#[test]
fn remove_header_does_not_panic() {
    assert!(matches!(
        call_http("remove_header", vec![Value::String("X-Foo".into())]),
        Value::Null | Value::Undefined
    ));
}

#[test]
fn write_does_not_panic() {
    assert!(matches!(
        call_http("write", vec![Value::String("body".into())]),
        Value::Null | Value::Undefined
    ));
}

#[test]
fn end_does_not_panic() {
    assert!(matches!(
        call_http("end", vec![]),
        Value::Null | Value::Undefined
    ));
}

#[test]
fn flush_does_not_panic() {
    assert!(matches!(
        call_http("flush", vec![]),
        Value::Null | Value::Undefined
    ));
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_vybe_http_surface_is_registered() {
    let expected = [
        "method",
        "uri",
        "remote_addr",
        "remote_port",
        "header",
        "header_all",
        "headers",
        "set_status",
        "status",
        "set_header",
        "add_header",
        "remove_header",
        "has_header",
        "headers_sent",
        "write",
        "end",
        "flush",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(missing.is_empty(), "missing vybe:http imports: {missing:?}");
}
