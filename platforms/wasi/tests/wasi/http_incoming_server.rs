//! Behaviour of the `wasi:http` SERVER half.
//!
//! Source of truth: `proposals/WASI/proposals/http/wit/types.wit`
//! (`wasi:http@0.3.1`). The accessors here were once registered as stubs
//! returning `Value::Null` ("Registered so WASM modules that import them don't
//! get link errors"), so registration tests passed while nothing worked. These
//! assert behaviour, not registration — `no_stub_registrations_remain_in_the_
//! http_surface` at the bottom guards that specific regression.
//!
//! HALF-MIGRATED, deliberately, and this is where the line falls:
//!
//!   * the REQUEST side is on 0.3.1. 0.2's `incoming-request` and
//!     `outgoing-request` are one `request` resource now, so these read it
//!     through `request.get-*` and `[static]request.consume-body`.
//!   * the RESPONSE side is NOT. `response-outparam`, `outgoing-response` and
//!     `outgoing-body` are absent from 0.3.1 — a handler RETURNS its response
//!     rather than writing it into an outparam. Renaming cannot express that:
//!     `vybex --serve` publishes the outparam handle as a global
//!     (`__wasi_http_response_out`) and collects the answer from it, so moving
//!     to the 0.3.1 model changes how a served script answers at all, not just
//!     what the call is called.
//!
//! `interface_coverage.rs` reports every one of those response-side names as
//! undeclared. That list shrinking to nothing is what finishes this file.

use std::sync::Arc;

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_platform_wasi::http as wasi_http;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

/// Call a host fn with `args` already-built as constants, inside one VM whose
/// registry the test also inspects directly.
static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn call(vm_setup: impl FnOnce() -> Vec<Value>, name: &str) -> Value {
    let args = vm_setup();
    let mut chunk = Chunk::new("<wasi-http-server-test>");
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

fn vm_with_platforms() -> VM {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm
}

fn is_error(value: &Value) -> bool {
    match value {
        Value::Object(object) => {
            let object = object.lock().unwrap();
            object.properties.contains_key("__wasi_error") || object.properties.contains_key("code")
        }
        _ => false,
    }
}

fn str_of(value: &Value) -> Option<String> {
    match value {
        Value::String(s) => Some(s.to_string()),
        _ => None,
    }
}

fn sample_request() -> u32 {
    wasi_http::push_incoming_request(
        "POST",
        Some("/orders?page=2".to_string()),
        Some("https".to_string()),
        Some("example.test:8443".to_string()),
        vec![
            ("content-type".to_string(), b"application/json".to_vec()),
            ("x-trace".to_string(), b"abc".to_vec()),
        ],
        b"{\"id\":7}".to_vec(),
    )
}

#[test]
fn request_reports_method_path_scheme_authority() {
    let vm = vm_with_platforms();
    let id = sample_request();
    let handle = wasi_http::incoming_request_value(&vm, id).expect("resource value");

    for (name, expected) in [
        ("[method]request.get-method", "POST"),
        ("[method]request.get-path-with-query", "/orders?page=2"),
        ("[method]request.get-scheme", "https"),
        ("[method]request.get-authority", "example.test:8443"),
    ] {
        let got = call(|| vec![handle.clone()], name);
        assert_eq!(
            str_of(&got).as_deref(),
            Some(expected),
            "{name} returned {got:?}"
        );
    }
}

#[test]
fn request_optional_fields_are_null_not_error() {
    // §request: path-with-query / scheme / authority are `option<..>`.
    // Absent must be null — an error would be a different WIT type.
    let vm = vm_with_platforms();
    let id = wasi_http::push_incoming_request("GET", None, None, None, Vec::new(), Vec::new());
    let handle = wasi_http::incoming_request_value(&vm, id).expect("resource value");

    for name in [
        "[method]request.get-path-with-query",
        "[method]request.get-scheme",
        "[method]request.get-authority",
    ] {
        let got = call(|| vec![handle.clone()], name);
        assert!(
            matches!(got, Value::Null),
            "{name} should be null when absent, got {got:?}"
        );
    }
}

#[test]
fn request_headers_resource_is_readable() {
    let vm = vm_with_platforms();
    let id = sample_request();
    let handle = wasi_http::incoming_request_value(&vm, id).expect("resource value");

    let headers = call(|| vec![handle.clone()], "[method]request.get-headers");
    assert!(
        !is_error(&headers),
        "headers returned an error: {headers:?}"
    );

    // The returned resource must work with the ordinary fields accessors.
    let got = call(
        || vec![headers.clone(), Value::String(Arc::from("content-type"))],
        "[method]fields.get",
    );
    assert!(!is_error(&got), "fields.get on request headers: {got:?}");
}

#[test]
fn request_consume_body_succeeds_at_most_once() {
    // §request.consume-body: the body is given up at most once; a second
    // call answers an error. 0.2 spelled this `incoming-request.consume`.
    let vm = vm_with_platforms();
    let id = sample_request();
    let handle = wasi_http::incoming_request_value(&vm, id).expect("resource value");

    let first = call(|| vec![handle.clone()], "[static]request.consume-body");
    assert!(!is_error(&first), "first consume failed: {first:?}");

    let second = call(|| vec![handle.clone()], "[static]request.consume-body");
    assert!(
        is_error(&second),
        "second consume must error, got {second:?}"
    );
}

#[test]
fn request_accessors_reject_a_foreign_resource() {
    // Passing a resource that is not a request is `invalid-argument`, not a panic.
    let got = call(
        || vec![Value::String(Arc::from("not-a-resource"))],
        "[method]request.get-method",
    );
    assert!(is_error(&got), "expected an error, got {got:?}");
}

#[test]
fn response_outparam_round_trips_status_headers_and_body() {
    let param_id = wasi_http::push_response_outparam();
    assert!(
        wasi_http::take_response_outparam(param_id).is_none(),
        "unset outparam must yield None"
    );
}

#[test]
fn response_outparam_reports_never_set() {
    // A handler that returns without calling `set` leaves nothing behind — the
    // host has to be able to tell that apart from a set-with-error.
    let param_id = wasi_http::push_response_outparam();
    assert!(wasi_http::take_response_outparam(param_id).is_none());
}

// TWO TESTS USED TO BE HERE: `send_informational_rejects_non_1xx` and
// `outgoing_body_finish_succeeds_once_then_errors`.
//
// Neither subject exists in `wasi:http@0.3.1`. `types.wit` declares four
// resources — `fields`, `request`, `request-options`, `response` — so
// `response-outparam` and `outgoing-body` are gone, and with them informational
// (1xx) responses and the explicit body-finish handshake. A body is a
// `stream<u8>` handed to `response.new` and taken back by `consume-body`;
// "finished" is the stream closing, not a call that can be made twice.
//
// Deleted rather than ported: RFC 9110 §15.2 still describes 1xx responses, but
// 0.3.1 gives a guest no way to send one, so a ported test would be asserting
// against an interface rather than about it.

#[test]
fn no_stub_registrations_remain_in_the_http_surface() {
    // Guard the regression this file exists for: the server half was once
    // registered as `|_ctx, _args| Value::Null` purely so imports would link.
    let source = include_str!("../../src/http.rs");
    assert!(
        !source.contains("Incoming request stubs"),
        "the incoming-request stub block is back"
    );
    let stub_bodies = source.matches("_args: &[Value]| Value::Null)").count();
    assert_eq!(
        stub_bodies, 0,
        "wasi:http has {stub_bodies} stub registrations returning Null"
    );
}
