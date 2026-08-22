//! `vybex --serve` publishes each request through the `wasi:http` spec surface.
//!
//! Step 2 of `documentation/httpserver.md` §4a.10: the server stops being the
//! only thing that knows what a request is. It builds the request and
//! response-outparam handles and exposes them as globals, so request-shaping
//! PRIMITIVES — and through them PHP superglobals, WSGI environ, Rack env —
//! read one language-neutral source.
//!
//! Read through the names `wasi:http@0.3.1` declares — `request.get-*` — and
//! NOT through 0.2's `incoming-request.*`, because those are the names
//! `primitives/http_request_env.rs` emits. This file used to call the 0.2
//! spelling, which meant it stayed green while covering a surface the compiler
//! no longer reached: the one test guarding `--serve` was testing the wrong
//! function names.
//!
//! These assert the MAPPING, which is where the spec is easy to get wrong:
//! `path-with-query` joins path and query; `scheme`/`authority` are
//! `option<…>` so empty means absent, not `""`.

use std::sync::Arc;

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};
use vybex::server::script::{WASI_REQUEST_GLOBAL, WASI_RESPONSE_GLOBAL, publish_wasi_request};

fn vm() -> VM {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm
}

/// Call a `wasi:http/types` function with an already-built handle.
static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn call(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<serve-wasi-http-test>");
    let import_idx = chunk.add_import("wasi:http/types", name);
    let argc = args.len() as u8;
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    for value in args {
        match value {
            Value::I32(n) => chunk.emit_i32_const(n, 0),
            Value::I64(n) => chunk.emit_i64_const(n, 0),
            Value::F32(f) => chunk.emit_f32_const(f, 0),
            Value::F64(f) => chunk.emit_f64_const(f, 0),
            Value::Bool(b) => chunk.emit_bool_const(b, 0),
            Value::String(text) => chunk.emit_string_const(&text, 0),
            Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
            other => {
                let global_name = format!(
                    "__test_arg_{}",
                    TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
                );
                vm.set_global_owned(global_name.clone(), other);
                let ci = chunk.intern_string_constant(&global_name);
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
            }
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.run(vec![chunk]).expect("VM run failed")
}

fn str_of(value: &Value) -> Option<String> {
    match value {
        Value::String(text) => Some(text.to_string()),
        _ => None,
    }
}

fn publish(vm: &mut VM, method: &str, path: &str, query: &str, scheme: &str, host: &str) -> Value {
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
    vm.global(WASI_REQUEST_GLOBAL)
        .cloned()
        .expect("request global missing")
}

#[test]
fn serving_a_request_publishes_both_wasi_handles() {
    let mut vm = vm();
    publish(&mut vm, "GET", "/", "", "http", "localhost:8080");

    assert!(
        vm.has_global(WASI_REQUEST_GLOBAL),
        "the request handle must be published for the request-shaping primitives"
    );
    assert!(
        vm.has_global(WASI_RESPONSE_GLOBAL),
        "the response handle must be published so the guest can answer — 0.3.1 \
         deleted `response-outparam`, `handler.handle` RETURNS its response"
    );
}

#[test]
fn method_and_headers_survive_the_mapping() {
    let mut vm = vm();
    let request = publish(&mut vm, "PATCH", "/items", "", "https", "api.test");

    let method = call("[method]request.get-method", vec![request.clone()]);
    assert_eq!(str_of(&method).as_deref(), Some("PATCH"), "got {method:?}");

    let headers = call("[method]request.get-headers", vec![request]);
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
    // §request.get-path-with-query — "the path with query parameters".
    // The server holds them separately, so this join is ours to get right.
    let mut vm = vm();
    let request = publish(
        &mut vm,
        "GET",
        "/search",
        "q=cats&page=2",
        "https",
        "x.test",
    );

    let got = call("[method]request.get-path-with-query", vec![request]);
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

    let got = call("[method]request.get-path-with-query", vec![request]);
    assert_eq!(str_of(&got).as_deref(), Some("/plain"), "got {got:?}");
}

#[test]
fn empty_scheme_and_host_map_to_none_not_empty_string() {
    // Both are `option<…>` in the WIT. `""` would be a present-but-blank value,
    // which is a different thing and would confuse every adapter downstream.
    let mut vm = vm();
    let request = publish(&mut vm, "GET", "/", "", "", "");

    for name in [
        "[method]request.get-scheme",
        "[method]request.get-authority",
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

    let scheme = call("[method]request.get-scheme", vec![request.clone()]);
    assert_eq!(str_of(&scheme).as_deref(), Some("https"), "got {scheme:?}");

    let authority = call("[method]request.get-authority", vec![request]);
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
        &mut vm_a,
        "GET",
        "/a",
        "",
        "http",
        "h",
        Vec::new(),
        Vec::new(),
    );
    let mut vm_b = vm();
    let (req_b, param_b) = publish_wasi_request(
        &mut vm_b,
        "GET",
        "/b",
        "",
        "http",
        "h",
        Vec::new(),
        Vec::new(),
    );

    assert_ne!(req_a, req_b, "incoming-request ids must be per-request");
    assert_ne!(
        param_a, param_b,
        "response-outparam ids must be per-request"
    );
}
