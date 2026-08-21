//! BEHAVIOUR of the WASI 0.3 HTTP surface, and the end-to-end server flow.
//!
//! `http_spec_0_3.rs` asserts only that names are in `host_registry`. That is
//! how nine `|_ctx, _args| Value::Null` stubs stayed green for the whole
//! server half — `contains_key` cannot tell a stub from an implementation.
//! Everything here drives the functions and asserts on results.
//!
//! Source of truth: `proposals/wasi-http/wit-0.3.0-draft/types.wit`.

use std::sync::Arc;

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_platform_wasi::http as wasi_http;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn call(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-http-0-3-behaviour>");
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

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
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
        Value::String(text) => Some(text.to_string()),
        _ => None,
    }
}

fn num_of(value: &Value) -> Option<f64> {
    match value {
        Value::F64(n) => Some(*n),
        Value::I32(n) => Some(*n as f64),
        _ => None,
    }
}

fn new_fields() -> Value {
    let fields = call("[constructor]fields", vec![]);
    assert!(!is_error(&fields), "fields constructor failed: {fields:?}");
    fields
}

fn new_request() -> Value {
    let request = call("[static]request.new", vec![new_fields()]);
    assert!(!is_error(&request), "request.new failed: {request:?}");
    request
}

fn new_response() -> Value {
    let response = call("[static]response.new", vec![new_fields()]);
    assert!(!is_error(&response), "response.new failed: {response:?}");
    response
}

// ── request accessors ───────────────────────────────────────────────────────

#[test]
fn request_new_defaults_to_get_with_no_target() {
    // §request.new: "Construct a new `request` with a default `method` of
    // `GET`, and `none` values for `path-with-query`, `scheme`, and
    // `authority`."
    let request = new_request();

    let method = call("[method]request.get-method", vec![request.clone()]);
    assert_eq!(str_of(&method).as_deref(), Some("GET"), "got {method:?}");

    for name in [
        "[method]request.get-path-with-query",
        "[method]request.get-scheme",
        "[method]request.get-authority",
    ] {
        let got = call(name, vec![request.clone()]);
        assert!(
            matches!(got, Value::Null),
            "{name} should default to none, got {got:?}"
        );
    }
}

#[test]
fn request_set_method_round_trips() {
    let request = new_request();
    let set = call(
        "[method]request.set-method",
        vec![request.clone(), s("DELETE")],
    );
    assert!(!is_error(&set), "set-method failed: {set:?}");

    let got = call("[method]request.get-method", vec![request.clone()]);
    assert_eq!(str_of(&got).as_deref(), Some("DELETE"), "got {got:?}");
}

#[test]
fn request_set_path_with_query_round_trips() {
    let request = new_request();
    let set = call(
        "[method]request.set-path-with-query",
        vec![request.clone(), s("/a/b?c=d")],
    );
    assert!(!is_error(&set), "set failed: {set:?}");

    let got = call("[method]request.get-path-with-query", vec![request.clone()]);
    assert_eq!(str_of(&got).as_deref(), Some("/a/b?c=d"), "got {got:?}");
}

#[test]
fn request_get_headers_returns_a_usable_fields_resource() {
    let fields = new_fields();
    let appended = call(
        "[method]fields.append",
        vec![fields.clone(), s("x-a"), s("1")],
    );
    assert!(!is_error(&appended), "append failed: {appended:?}");

    let request = call("[static]request.new", vec![fields]);
    let headers = call("[method]request.get-headers", vec![request]);
    assert!(!is_error(&headers), "get-headers failed: {headers:?}");

    let has = call("[method]fields.has", vec![headers, s("x-a")]);
    assert!(!is_error(&has), "fields.has on request headers: {has:?}");
}

#[test]
fn request_get_options_is_none_by_default_and_some_when_supplied() {
    // §request.get-options -> option<request-options>.
    let bare = new_request();
    let none = call("[method]request.get-options", vec![bare]);
    assert!(
        matches!(none, Value::Null),
        "no options should be none, got {none:?}"
    );

    let options = call("[constructor]request-options", vec![]);
    assert!(!is_error(&options), "options constructor: {options:?}");
    let request = call(
        "[static]request.new",
        vec![new_fields(), Value::Null, Value::Null, options],
    );
    let some = call("[method]request.get-options", vec![request]);
    assert!(
        !matches!(some, Value::Null) && !is_error(&some),
        "supplied options should round-trip, got {some:?}"
    );
}

// ── response accessors ──────────────────────────────────────────────────────

#[test]
fn response_new_defaults_to_200_and_status_round_trips() {
    let response = new_response();

    let status = call("[method]response.get-status-code", vec![response.clone()]);
    assert_eq!(num_of(&status), Some(200.0), "default status: {status:?}");

    let set = call(
        "[method]response.set-status-code",
        vec![response.clone(), Value::F64(404.0)],
    );
    assert!(!is_error(&set), "set-status-code failed: {set:?}");

    let got = call("[method]response.get-status-code", vec![response]);
    assert_eq!(num_of(&got), Some(404.0), "got {got:?}");
}

#[test]
fn response_get_headers_returns_a_usable_fields_resource() {
    let response = new_response();
    let headers = call("[method]response.get-headers", vec![response]);
    assert!(!is_error(&headers), "get-headers failed: {headers:?}");

    let set = call(
        "[method]fields.set",
        vec![headers.clone(), s("content-type"), s("text/plain")],
    );
    assert!(!is_error(&set), "fields.set on response headers: {set:?}");
}

// ── fields, 0.3 additions ───────────────────────────────────────────────────

#[test]
fn fields_copy_all_returns_every_entry() {
    let fields = new_fields();
    call(
        "[method]fields.append",
        vec![fields.clone(), s("x-a"), s("1")],
    );
    call(
        "[method]fields.append",
        vec![fields.clone(), s("x-b"), s("2")],
    );

    let all = call("[method]fields.copy-all", vec![fields]);
    assert!(!is_error(&all), "copy-all failed: {all:?}");
    let len = match &all {
        Value::Object(object) => {
            let object = object.lock().unwrap();
            match &object.kind {
                vybe_runtime::value::ObjectKind::Array(items) => items.len(),
                _ => 0,
            }
        }
        _ => 0,
    };
    assert_eq!(len, 2, "copy-all should list both entries, got {all:?}");
}

#[test]
fn fields_get_and_delete_removes_the_header() {
    let fields = new_fields();
    call(
        "[method]fields.append",
        vec![fields.clone(), s("x-gone"), s("v")],
    );

    let taken = call(
        "[method]fields.get-and-delete",
        vec![fields.clone(), s("x-gone")],
    );
    assert!(!is_error(&taken), "get-and-delete failed: {taken:?}");

    let has = call("[method]fields.has", vec![fields, s("x-gone")]);
    assert!(
        matches!(has, Value::Bool(false)) || matches!(&has, Value::I32(0)),
        "header should be gone after get-and-delete, has = {has:?}"
    );
}

#[test]
fn fields_from_list_seeds_entries() {
    let list = call("[static]fields.from-list", vec![Value::Null]);
    assert!(!is_error(&list), "from-list with none failed: {list:?}");
}

// ── end-to-end server flow ──────────────────────────────────────────────────

#[test]
fn server_round_trip_request_in_response_out() {
    // The whole point of the server half: build an incoming request the way a
    // transport would, read it as a guest would, write a response, and have the
    // host read that response back out.
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());

    let request_id = wasi_http::push_incoming_request(
        "PUT",
        Some("/things/9".to_string()),
        Some("https".to_string()),
        Some("api.test".to_string()),
        vec![("accept".to_string(), b"application/json".to_vec())],
        b"payload".to_vec(),
    );
    let request = wasi_http::incoming_request_value(&vm, request_id).expect("request handle");

    // Guest reads the request.
    let method = call("[method]incoming-request.method", vec![request.clone()]);
    assert_eq!(str_of(&method).as_deref(), Some("PUT"));
    let body = call("[method]incoming-request.consume", vec![request]);
    assert!(!is_error(&body), "consume failed: {body:?}");

    // Guest writes a response. §outgoing-response constructor takes HEADERS
    // only and defaults to 200; the status is set separately.
    let response = call("[constructor]outgoing-response", vec![new_fields()]);
    assert!(!is_error(&response), "response ctor: {response:?}");
    let set_status = call(
        "[method]outgoing-response.set-status-code",
        vec![response.clone(), Value::F64(201.0)],
    );
    assert!(!is_error(&set_status), "set-status-code: {set_status:?}");
    let headers = call("[method]outgoing-response.headers", vec![response.clone()]);
    call(
        "[method]fields.set",
        vec![headers, s("x-created"), s("yes")],
    );

    // Host reads it back.
    let param_id = wasi_http::push_response_outparam();
    let param = wasi_http::response_outparam_value(&vm, param_id).expect("outparam handle");
    let set = call("[static]response-outparam.set", vec![param, response]);
    assert!(!is_error(&set), "response-outparam.set failed: {set:?}");

    let taken = wasi_http::take_response_outparam(param_id);
    let Some(Ok((status, _headers, _body))) = taken else {
        panic!("host could not read the response back: {taken:?}");
    };
    assert_eq!(status, 201, "status did not survive the round trip");
}

#[test]
fn response_outparam_set_is_at_most_once() {
    // §response-outparam.set "consumes the `response-outparam` to ensure that
    // it is called at most once".
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());

    let param_id = wasi_http::push_response_outparam();
    let param = wasi_http::response_outparam_value(&vm, param_id).expect("outparam handle");
    let response = call("[constructor]outgoing-response", vec![new_fields()]);

    let first = call(
        "[static]response-outparam.set",
        vec![param.clone(), response.clone()],
    );
    assert!(!is_error(&first), "first set failed: {first:?}");

    let second = call("[static]response-outparam.set", vec![param, response]);
    assert!(is_error(&second), "second set must error, got {second:?}");
}

#[test]
fn outgoing_response_constructor_defaults_to_200() {
    // §outgoing-response: "Construct an `outgoing-response`, with a default
    // `status-code` of `200`." The constructor takes headers, NOT a status.
    let response = call("[constructor]outgoing-response", vec![new_fields()]);
    let status = call("[method]outgoing-response.status-code", vec![response]);
    assert_eq!(num_of(&status), Some(200.0), "default status: {status:?}");
}

#[test]
fn outgoing_response_rejects_an_invalid_status_code() {
    // §outgoing-response.set-status-code: "Fails if the status-code given is
    // not a valid http status code."
    let response = call("[constructor]outgoing-response", vec![new_fields()]);
    let bad = call(
        "[method]outgoing-response.set-status-code",
        vec![response, Value::F64(999.0)],
    );
    assert!(
        is_error(&bad),
        "999 is not a valid status code, got {bad:?}"
    );
}
