//! `vybex --serve` publishes each request through the `wasi:http` spec surface.
//!
//! Step 2 of `documentation/httpserver.md` §4a.10: the server stops being the
//! only thing that knows what a request is. It builds `incoming-request` /
//! `response-outparam` handles and exposes them as globals, so request-shaping
//! PRIMITIVES — and through them PHP superglobals, WSGI environ, Rack env —
//! read one language-neutral source.
//!
//! These assert the MAPPING, which is where the spec is easy to get wrong:
//! `path-with-query` joins path and query; `scheme`/`authority` are
//! `option<…>` so empty means absent, not `""`.

use std::sync::Arc;

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};
use vybex::server::script::{
    publish_wasi_request, WASI_REQUEST_GLOBAL, WASI_RESPONSE_OUT_GLOBAL,
};

fn vm() -> VM {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm
}

/// Call a `wasi:http/types` function with an already-built handle.
fn call(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<serve-wasi-http-test>");
    let import_idx = chunk.add_import("wasi:http/types", name);
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

fn str_of(value: &Value) -> Option<String> {
    match value {
        Value::String(text) => Some(text.to_string()),
        _ => None,
    }
}

fn publish(
    vm: &mut VM,
    method: &str,
    path: &str,
    query: &str,
    scheme: &str,
    host: &str,
) -> Value {
    publish_wasi_request(
        vm,
        method,
        path,
        query,
        scheme,
        host,
        vec![("accept".to_string(), b"text/html".to_vec())],
        b"body-bytes".to_vec(),
    );
    vm.globals
        .get(WASI_REQUEST_GLOBAL)
        .cloned()
        .expect("request global missing")
}

#[test]
fn serving_a_request_publishes_both_wasi_handles() {
    let mut vm = vm();
    publish(&mut vm, "GET", "/", "", "http", "localhost:8080");

    assert!(
        vm.globals.contains_key(WASI_REQUEST_GLOBAL),
        "incoming-request must be published for the request-shaping primitives"
    );
    assert!(
        vm.globals.contains_key(WASI_RESPONSE_OUT_GLOBAL),
        "response-outparam must be published so the guest can answer"
    );
}

#[test]
fn method_and_headers_survive_the_mapping() {
    let mut vm = vm();
    let request = publish(&mut vm, "PATCH", "/items", "", "https", "api.test");

    let method = call("[method]incoming-request.method", vec![request.clone()]);
    assert_eq!(str_of(&method).as_deref(), Some("PATCH"), "got {method:?}");

    let headers = call("[method]incoming-request.headers", vec![request]);
    let has = call(
        "[method]fields.has",
        vec![headers, Value::String(Arc::from("accept"))],
    );
    assert!(
        !matches!(has, Value::Null),
        "request headers should reach the fields resource, got {has:?}"
    );
}

#[test]
fn path_with_query_joins_path_and_query() {
    // §incoming-request.path-with-query — "the path with query parameters".
    // The server holds them separately, so this join is ours to get right.
    let mut vm = vm();
    let request = publish(&mut vm, "GET", "/search", "q=cats&page=2", "https", "x.test");

    let got = call("[method]incoming-request.path-with-query", vec![request]);
    assert_eq!(
        str_of(&got).as_deref(),
        Some("/search?q=cats&page=2"),
        "got {got:?}"
    );
}

#[test]
fn path_with_query_omits_the_separator_when_there_is_no_query() {
    let mut vm = vm();
    let request = publish(&mut vm, "GET", "/plain", "", "https", "x.test");

    let got = call("[method]incoming-request.path-with-query", vec![request]);
    assert_eq!(str_of(&got).as_deref(), Some("/plain"), "got {got:?}");
}

#[test]
fn empty_scheme_and_host_map_to_none_not_empty_string() {
    // Both are `option<…>` in the WIT. `""` would be a present-but-blank value,
    // which is a different thing and would confuse every adapter downstream.
    let mut vm = vm();
    let request = publish(&mut vm, "GET", "/", "", "", "");

    for name in [
        "[method]incoming-request.scheme",
        "[method]incoming-request.authority",
    ] {
        let got = call(name, vec![request.clone()]);
        assert!(
            matches!(got, Value::Null),
            "{name} should be none when the server has no value, got {got:?}"
        );
    }
}

#[test]
fn scheme_and_authority_are_carried_when_present() {
    let mut vm = vm();
    let request = publish(&mut vm, "GET", "/", "", "https", "example.test:8443");

    let scheme = call("[method]incoming-request.scheme", vec![request.clone()]);
    assert_eq!(str_of(&scheme).as_deref(), Some("https"), "got {scheme:?}");

    let authority = call("[method]incoming-request.authority", vec![request]);
    assert_eq!(
        str_of(&authority).as_deref(),
        Some("example.test:8443"),
        "got {authority:?}"
    );
}

#[test]
fn each_request_gets_distinct_handles() {
    // Handles are per-request; two requests must not share resource ids or one
    // request could consume another's body.
    let mut vm_a = vm();
    let (req_a, param_a) = publish_wasi_request(
        &mut vm_a, "GET", "/a", "", "http", "h", Vec::new(), Vec::new(),
    );
    let mut vm_b = vm();
    let (req_b, param_b) = publish_wasi_request(
        &mut vm_b, "GET", "/b", "", "http", "h", Vec::new(), Vec::new(),
    );

    assert_ne!(req_a, req_b, "incoming-request ids must be per-request");
    assert_ne!(param_a, param_b, "response-outparam ids must be per-request");
}
