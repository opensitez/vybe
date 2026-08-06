//! Behaviour tests for the real outgoing WASI HTTP client slice.

use std::io::{Read, Write};
use std::net::TcpListener;
use std::thread;

use vybe_runtime::value::{ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-http-test>");
    let import_idx = chunk.add_import(module, name);
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
                vm.globals.insert(global_name.clone(), other);
                let ci = chunk.intern_string_constant(&global_name);
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
            }
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(module: &str, name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(module.to_string(), name.to_string()))
}

fn types(name: &str, args: Vec<Value>) -> Value {
    invoke("wasi:http/types", name, args)
}

fn outgoing(name: &str, args: Vec<Value>) -> Value {
    invoke("wasi:http/outgoing-handler", name, args)
}

fn legacy(name: &str, args: Vec<Value>) -> Value {
    invoke("wasi:http", name, args)
}

fn s(text: &str) -> Value {
    Value::String(std::sync::Arc::from(text))
}

fn is_error(value: &Value) -> Option<String> {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        if let Some(Value::String(text)) = object.properties.get("__wasi_error") {
            return Some(text.to_string());
        }
    }
    None
}

fn prop(value: &Value, key: &str) -> Value {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        if let Some(value) = object.properties.get(key) {
            return value.clone();
        }
    }
    Value::Null
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

fn byte_array_as_string(value: &Value) -> Option<String> {
    let Value::Object(bytes) = value else {
        return None;
    };
    let bytes = bytes.lock().unwrap();
    let ObjectKind::Array(byte_values) = &bytes.kind else {
        return None;
    };
    let decoded = byte_values
        .iter()
        .filter_map(|value| match value {
            Value::F64(byte) => Some(*byte as u8),
            Value::I32(byte) => Some(*byte as u8),
            _ => None,
        })
        .collect::<Vec<_>>();
    String::from_utf8(decoded).ok()
}

fn all_header_values_as_strings(value: &Value) -> Vec<String> {
    let Value::Object(array) = value else {
        return Vec::new();
    };
    let array = array.lock().unwrap();
    let ObjectKind::Array(values) = &array.kind else {
        return Vec::new();
    };
    values.iter().filter_map(byte_array_as_string).collect()
}

fn header_entries_as_strings(value: &Value) -> Vec<(String, String)> {
    let Value::Object(array) = value else {
        return Vec::new();
    };
    let array = array.lock().unwrap();
    let ObjectKind::Array(entries) = &array.kind else {
        return Vec::new();
    };
    entries
        .iter()
        .filter_map(|entry| {
            let Value::Object(pair) = entry else {
                return None;
            };
            let pair = pair.lock().unwrap();
            let ObjectKind::Array(pair_values) = &pair.kind else {
                return None;
            };
            let name = match pair_values.first() {
                Some(Value::String(text)) => text.to_string(),
                _ => return None,
            };
            let value = pair_values.get(1).and_then(byte_array_as_string)?;
            Some((name, value))
        })
        .collect()
}

fn first_header_value_as_string(value: &Value) -> Option<String> {
    let Value::Object(array) = value else {
        return None;
    };
    let array = array.lock().unwrap();
    let ObjectKind::Array(values) = &array.kind else {
        return None;
    };
    let first = values.first()?;
    let Value::Object(bytes) = first else {
        return None;
    };
    let bytes = bytes.lock().unwrap();
    let ObjectKind::Array(byte_values) = &bytes.kind else {
        return None;
    };
    let decoded = byte_values
        .iter()
        .filter_map(|value| match value {
            Value::F64(byte) => Some(*byte as u8),
            Value::I32(byte) => Some(*byte as u8),
            _ => None,
        })
        .collect::<Vec<_>>();
    String::from_utf8(decoded).ok()
}

fn start_server(status_line: &str, extra_headers: &[(&str, &str)], body: &str) -> String {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind test server");
    let address = listener.local_addr().expect("server addr");
    let status_line = status_line.to_string();
    let headers = extra_headers
        .iter()
        .map(|(name, value)| format!("{}: {}\r\n", name, value))
        .collect::<Vec<_>>()
        .join("");
    let body = body.to_string();

    thread::spawn(move || {
        let (mut stream, _) = listener.accept().expect("accept request");
        let mut buffer = [0u8; 1024];
        let _ = stream.read(&mut buffer);
        let response = format!(
            "HTTP/1.1 {}\r\nContent-Length: {}\r\n{}\r\n{}",
            status_line,
            body.len(),
            headers,
            body,
        );
        stream
            .write_all(response.as_bytes())
            .expect("write response");
    });

    format!("{}", address)
}

#[test]
fn outgoing_handler_round_trip_returns_status_and_headers() {
    let authority = start_server("201 Created", &[("Content-Type", "text/plain")], "ok");

    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(
        is_error(&request).is_none(),
        "request constructor should succeed"
    );

    assert!(
        is_error(&types(
            "[method]outgoing-request.set-method",
            vec![request.clone(), s("POST")]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-scheme",
            vec![request.clone(), s("http")]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-authority",
            vec![request.clone(), s(&authority)]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-path-with-query",
            vec![request.clone(), s("/submit?ok=1")]
        ))
        .is_none()
    );

    let future = outgoing("handle", vec![request, Value::Null]);
    assert!(
        is_error(&future).is_none(),
        "handle should return a future resource"
    );

    let response = types("[method]future-incoming-response.get", vec![future]);
    assert!(
        is_error(&response).is_none(),
        "future get should resolve to an incoming response"
    );
    assert_eq!(
        types("[method]incoming-response.status", vec![response.clone()]),
        Value::F64(201.0)
    );

    let response_headers = types("[method]incoming-response.headers", vec![response]);
    assert_eq!(
        types(
            "[method]fields.has",
            vec![response_headers.clone(), s("content-type")]
        ),
        Value::Bool(true)
    );
    let content_type = types(
        "[method]fields.get",
        vec![response_headers, s("Content-Type")],
    );
    assert_eq!(
        first_header_value_as_string(&content_type).as_deref(),
        Some("text/plain")
    );
}

#[test]
fn outgoing_handler_reports_transport_errors_through_future() {
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-scheme",
            vec![request.clone(), s("http")]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-authority",
            vec![request.clone(), s("127.0.0.1:1")]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-path-with-query",
            vec![request.clone(), s("/")]
        ))
        .is_none()
    );

    let future = outgoing("handle", vec![request, Value::Null]);
    assert!(
        is_error(&future).is_none(),
        "transport failures should still produce a future"
    );

    let result = types("[method]future-incoming-response.get", vec![future]);
    assert_eq!(is_error(&result).as_deref(), Some("connection-refused"));
}

#[test]
fn fields_entries_and_get_preserve_duplicate_header_values() {
    let authority = start_server(
        "200 OK",
        &[
            ("Set-Cookie", "a=1"),
            ("Set-Cookie", "b=2"),
            ("X-Trace", "ok"),
        ],
        "body",
    );

    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-scheme",
            vec![request.clone(), s("http")]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-authority",
            vec![request.clone(), s(&authority)]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-path-with-query",
            vec![request.clone(), s("/")]
        ))
        .is_none()
    );

    let future = outgoing("handle", vec![request, Value::Null]);
    let response = types("[method]future-incoming-response.get", vec![future]);
    let response_headers = types("[method]incoming-response.headers", vec![response]);

    let entries = types("[method]fields.entries", vec![response_headers.clone()]);
    let decoded = header_entries_as_strings(&entries);
    assert!(decoded.contains(&(String::from("Set-Cookie"), String::from("a=1"))));
    assert!(decoded.contains(&(String::from("Set-Cookie"), String::from("b=2"))));
    assert!(decoded.contains(&(String::from("X-Trace"), String::from("ok"))));

    let values = types(
        "[method]fields.get",
        vec![response_headers, s("set-cookie")],
    );
    assert_eq!(
        all_header_values_as_strings(&values),
        vec![String::from("a=1"), String::from("b=2")]
    );
}

#[test]
fn fields_get_missing_header_returns_empty_array() {
    let headers = types("[constructor]fields", vec![]);
    let missing = types("[method]fields.get", vec![headers, s("missing")]);
    assert_eq!(array_len(&missing), 0);
}

#[test]
fn fields_entries_on_new_headers_returns_empty_array() {
    let headers = types("[constructor]fields", vec![]);
    let entries = types("[method]fields.entries", vec![headers]);
    assert_eq!(array_len(&entries), 0);
}

#[test]
fn fields_get_with_null_name_returns_empty_array() {
    let headers = types("[constructor]fields", vec![]);
    let result = types("[method]fields.get", vec![headers, Value::Null]);
    assert_eq!(array_len(&result), 0);
}

#[test]
fn fields_get_rejects_invalid_handle() {
    let result = types("[method]fields.get", vec![s("not-fields"), s("x")]);
    assert_eq!(is_error(&result).as_deref(), Some("invalid-argument"));
}

#[test]
fn fields_has_is_case_insensitive_on_response_headers() {
    let authority = start_server("200 OK", &[("X-Trace", "ok")], "body");

    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-scheme",
            vec![request.clone(), s("http")]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-authority",
            vec![request.clone(), s(&authority)]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-path-with-query",
            vec![request.clone(), s("/")]
        ))
        .is_none()
    );

    let future = outgoing("handle", vec![request, Value::Null]);
    let response = types("[method]future-incoming-response.get", vec![future]);
    let response_headers = types("[method]incoming-response.headers", vec![response]);

    assert_eq!(
        types(
            "[method]fields.has",
            vec![response_headers.clone(), s("x-trace")]
        ),
        Value::Bool(true)
    );
    assert_eq!(
        types("[method]fields.has", vec![response_headers, s("X-TRACE")]),
        Value::Bool(true)
    );
}

#[test]
fn fields_has_returns_false_for_missing_header() {
    let headers = types("[constructor]fields", vec![]);
    assert_eq!(
        types("[method]fields.has", vec![headers, s("missing")]),
        Value::Bool(false)
    );
}

#[test]
fn fields_has_with_null_name_returns_false() {
    let headers = types("[constructor]fields", vec![]);
    assert_eq!(
        types("[method]fields.has", vec![headers, Value::Null]),
        Value::Bool(false)
    );
}

#[test]
fn fields_has_rejects_invalid_handle() {
    let result = types("[method]fields.has", vec![s("not-fields"), s("x")]);
    assert_eq!(is_error(&result).as_deref(), Some("invalid-argument"));
}

#[test]
fn outgoing_request_headers_returns_resource_handle() {
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    let request_headers = types("[method]outgoing-request.headers", vec![request]);
    let entries = types("[method]fields.entries", vec![request_headers]);
    assert_eq!(array_len(&entries), 0);
}

#[test]
fn outgoing_request_constructor_rejects_invalid_headers_handle() {
    let result = types("[constructor]outgoing-request", vec![s("not-headers")]);
    assert_eq!(is_error(&result).as_deref(), Some("invalid-argument"));
}

#[test]
fn fields_entries_reject_invalid_handle() {
    let result = types("[method]fields.entries", vec![s("not-fields")]);
    assert_eq!(is_error(&result).as_deref(), Some("invalid-argument"));
}

#[test]
fn outgoing_handler_accepts_request_options_resource() {
    let authority = start_server("204 No Content", &[], "");

    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    let options = types("[constructor]request-options", vec![]);
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-scheme",
            vec![request.clone(), s("http")]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-authority",
            vec![request.clone(), s(&authority)]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-path-with-query",
            vec![request.clone(), s("/")]
        ))
        .is_none()
    );

    let future = outgoing("handle", vec![request, options]);
    let response = types("[method]future-incoming-response.get", vec![future]);
    assert_eq!(
        types("[method]incoming-response.status", vec![response]),
        Value::F64(204.0)
    );
}

#[test]
fn request_options_constructor_returns_distinct_resources() {
    let first = types("[constructor]request-options", vec![]);
    let second = types("[constructor]request-options", vec![]);
    assert!(matches!(first, Value::Object(_)));
    assert!(matches!(second, Value::Object(_)));
    assert_ne!(first, second);
}

#[test]
fn outgoing_handler_rejects_invalid_request_handle() {
    let result = outgoing("handle", vec![s("not-a-request"), Value::Null]);
    assert_eq!(is_error(&result).as_deref(), Some("HTTP-request-denied"));
}

#[test]
fn outgoing_handler_rejects_non_request_options_second_argument() {
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    let result = outgoing("handle", vec![request, s("not-options")]);
    assert_eq!(is_error(&result).as_deref(), Some("HTTP-request-denied"));
}

#[test]
fn outgoing_request_set_method_rejects_blank_methods() {
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    let result = types(
        "[method]outgoing-request.set-method",
        vec![request, s("   ")],
    );
    assert_eq!(
        is_error(&result).as_deref(),
        Some("HTTP-request-method-invalid")
    );
}

#[test]
fn outgoing_request_set_method_rejects_invalid_request_handle() {
    let result = types(
        "[method]outgoing-request.set-method",
        vec![s("not-a-request"), s("GET")],
    );
    assert_eq!(is_error(&result).as_deref(), Some("invalid-argument"));
}

#[test]
fn outgoing_request_set_scheme_rejects_invalid_request_handle() {
    let result = types(
        "[method]outgoing-request.set-scheme",
        vec![s("not-a-request"), s("http")],
    );
    assert_eq!(is_error(&result).as_deref(), Some("invalid-argument"));
}

#[test]
fn outgoing_request_set_path_with_query_rejects_invalid_request_handle() {
    let result = types(
        "[method]outgoing-request.set-path-with-query",
        vec![s("not-a-request"), s("/")],
    );
    assert_eq!(is_error(&result).as_deref(), Some("invalid-argument"));
}

#[test]
fn outgoing_request_set_authority_rejects_invalid_request_handle() {
    let result = types(
        "[method]outgoing-request.set-authority",
        vec![s("not-a-request"), s("example.com")],
    );
    assert_eq!(is_error(&result).as_deref(), Some("invalid-argument"));
}

#[test]
fn outgoing_request_headers_reject_invalid_request_handle() {
    let result = types("[method]outgoing-request.headers", vec![s("not-a-request")]);
    assert_eq!(is_error(&result).as_deref(), Some("invalid-argument"));
}

#[test]
fn outgoing_handler_rejects_invalid_scheme() {
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-scheme",
            vec![request.clone(), s("https")]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-authority",
            vec![request.clone(), s("example.com")]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-path-with-query",
            vec![request.clone(), s("/")]
        ))
        .is_none()
    );

    let result = outgoing("handle", vec![request, Value::Null]);
    assert_eq!(
        is_error(&result).as_deref(),
        Some("HTTP-request-URI-invalid")
    );
}

#[test]
fn null_path_uses_default_slash() {
    let authority = start_server("200 OK", &[], "ok");
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-scheme",
            vec![request.clone(), s("http")]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-authority",
            vec![request.clone(), s(&authority)]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-path-with-query",
            vec![request.clone(), Value::Null]
        ))
        .is_none()
    );

    let future = outgoing("handle", vec![request, Value::Null]);
    let response = types("[method]future-incoming-response.get", vec![future]);
    assert_eq!(
        types("[method]incoming-response.status", vec![response]),
        Value::F64(200.0)
    );
}

#[test]
fn future_incoming_response_get_errors_when_consumed_twice() {
    let authority = start_server("200 OK", &[], "ok");

    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-scheme",
            vec![request.clone(), s("http")]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-authority",
            vec![request.clone(), s(&authority)]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-path-with-query",
            vec![request.clone(), s("/")]
        ))
        .is_none()
    );

    let future = outgoing("handle", vec![request, Value::Null]);
    let first = types("[method]future-incoming-response.get", vec![future.clone()]);
    assert!(is_error(&first).is_none());

    let second = types("[method]future-incoming-response.get", vec![future]);
    assert_eq!(is_error(&second).as_deref(), Some("already-consumed"));
}

#[test]
fn future_incoming_response_get_rejects_invalid_handle() {
    let result = types(
        "[method]future-incoming-response.get",
        vec![s("not-a-future")],
    );
    assert_eq!(is_error(&result).as_deref(), Some("invalid-argument"));
}

#[test]
fn incoming_response_status_rejects_invalid_handle() {
    let result = types(
        "[method]incoming-response.status",
        vec![s("not-a-response")],
    );
    assert_eq!(is_error(&result).as_deref(), Some("invalid-argument"));
}

#[test]
fn incoming_response_headers_rejects_invalid_handle() {
    let result = types(
        "[method]incoming-response.headers",
        vec![s("not-a-response")],
    );
    assert_eq!(is_error(&result).as_deref(), Some("invalid-argument"));
}

#[test]
fn incoming_response_headers_expose_server_headers() {
    let authority = start_server(
        "203 Non-Authoritative Information",
        &[("Content-Type", "text/plain")],
        "body",
    );

    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-scheme",
            vec![request.clone(), s("http")]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-authority",
            vec![request.clone(), s(&authority)]
        ))
        .is_none()
    );
    assert!(
        is_error(&types(
            "[method]outgoing-request.set-path-with-query",
            vec![request.clone(), s("/")]
        ))
        .is_none()
    );

    let future = outgoing("handle", vec![request, Value::Null]);
    let response = types("[method]future-incoming-response.get", vec![future]);
    assert_eq!(
        types("[method]incoming-response.status", vec![response.clone()]),
        Value::F64(203.0)
    );

    let response_headers = types("[method]incoming-response.headers", vec![response]);
    let content_type = types(
        "[method]fields.get",
        vec![response_headers, s("content-type")],
    );
    assert_eq!(
        first_header_value_as_string(&content_type).as_deref(),
        Some("text/plain")
    );
}

#[test]
fn proposal_http_outgoing_handler_is_registered() {
    assert!(
        has_import("wasi:http/outgoing-handler", "handle"),
        "wasi:http/outgoing-handler.handle should be registered"
    );
}

#[test]
fn proposal_http_fields_surface_is_registered() {
    let expected = [
        "[constructor]fields",
        "[static]fields.from-list",
        "[method]fields.get",
        "[method]fields.has",
        "[method]fields.set",
        "[method]fields.delete",
        "[method]fields.append",
        "[method]fields.clone",
        "[method]fields.entries",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:http/types", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing http fields imports: {missing:?}"
    );
}

#[test]
fn proposal_http_outgoing_request_surface_is_registered() {
    let expected = [
        "[constructor]outgoing-request",
        "[method]outgoing-request.method",
        "[method]outgoing-request.set-method",
        "[method]outgoing-request.path-with-query",
        "[method]outgoing-request.set-path-with-query",
        "[method]outgoing-request.scheme",
        "[method]outgoing-request.set-scheme",
        "[method]outgoing-request.authority",
        "[method]outgoing-request.set-authority",
        "[method]outgoing-request.headers",
        "[method]outgoing-request.body",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:http/types", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing http outgoing-request imports: {missing:?}"
    );
}

#[test]
fn proposal_http_request_options_surface_is_registered() {
    let expected = [
        "[constructor]request-options",
        "[method]request-options.connect-timeout",
        "[method]request-options.set-connect-timeout",
        "[method]request-options.first-byte-timeout",
        "[method]request-options.set-first-byte-timeout",
        "[method]request-options.between-bytes-timeout",
        "[method]request-options.set-between-bytes-timeout",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:http/types", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing http request-options imports: {missing:?}"
    );
}

#[test]
fn proposal_http_response_surface_is_registered() {
    let expected = [
        "[static]response.new",
        "[static]response.consume-body",
        "[method]incoming-response.status",
        "[method]incoming-response.headers",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:http/types", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing http response imports: {missing:?}"
    );
}
