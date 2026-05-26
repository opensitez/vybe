use std::io::{Read, Write};
use std::net::TcpListener;
use std::sync::{Arc, Mutex};
use std::thread;

use vybe_bytecode::value::{ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-http-lifecycle-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for value in args {
        let constant = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, constant, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
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
    let Value::Object(array) = value else { return 0 };
    let array = array.lock().unwrap();
    let ObjectKind::Array(values) = &array.kind else { return 0 };
    values.len()
}

fn capture_server(status_line: &str, body: &str) -> (String, Arc<Mutex<Option<String>>>) {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind test server");
    let address = listener.local_addr().expect("server addr");
    let status_line = status_line.to_string();
    let body = body.to_string();
    let request_line = Arc::new(Mutex::new(None));
    let request_line_clone = request_line.clone();

    thread::spawn(move || {
        let (mut stream, _) = listener.accept().expect("accept request");
        let mut buffer = [0u8; 2048];
        let read = stream.read(&mut buffer).expect("read request");
        let raw = String::from_utf8_lossy(&buffer[..read]).to_string();
        let first_line = raw.lines().next().unwrap_or_default().trim_end().to_string();
        *request_line_clone.lock().unwrap() = Some(first_line);
        let response = format!(
            "HTTP/1.1 {}\r\nContent-Length: {}\r\nContent-Type: text/plain\r\n\r\n{}",
            status_line,
            body.len(),
            body,
        );
        stream.write_all(response.as_bytes()).expect("write response");
    });

    (format!("{}", address), request_line)
}

#[test]
fn set_method_trims_and_uppercases_request_line() {
    let (authority, request_line) = capture_server("200 OK", "ok");
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(is_error(&types("[method]outgoing-request.set-method", vec![request.clone(), s(" post ")])).is_none());
    assert!(is_error(&types("[method]outgoing-request.set-authority", vec![request.clone(), s(&authority)])).is_none());
    assert!(is_error(&types("[method]outgoing-request.set-path-with-query", vec![request.clone(), s("submit")])).is_none());
    let future = outgoing("handle", vec![request, Value::Null]);
    let response = types("[method]future-incoming-response.get", vec![future]);
    assert!(is_error(&response).is_none());
    assert_eq!(request_line.lock().unwrap().clone().as_deref(), Some("POST /submit HTTP/1.1"));
}

#[test]
fn default_method_is_get_when_unset() {
    let (authority, request_line) = capture_server("200 OK", "ok");
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(is_error(&types("[method]outgoing-request.set-authority", vec![request.clone(), s(&authority)])).is_none());
    let future = outgoing("handle", vec![request, Value::Null]);
    let response = types("[method]future-incoming-response.get", vec![future]);
    assert!(is_error(&response).is_none());
    assert_eq!(request_line.lock().unwrap().clone().as_deref(), Some("GET / HTTP/1.1"));
}

#[test]
fn set_scheme_lowercases_mixed_case_http_and_succeeds() {
    let (authority, request_line) = capture_server("200 OK", "ok");
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(is_error(&types("[method]outgoing-request.set-scheme", vec![request.clone(), s("HtTp")])).is_none());
    assert!(is_error(&types("[method]outgoing-request.set-authority", vec![request.clone(), s(&authority)])).is_none());
    let future = outgoing("handle", vec![request, Value::Null]);
    let response = types("[method]future-incoming-response.get", vec![future]);
    assert!(is_error(&response).is_none());
    assert_eq!(request_line.lock().unwrap().clone().as_deref(), Some("GET / HTTP/1.1"));
}

#[test]
fn path_without_leading_slash_is_normalized() {
    let (authority, request_line) = capture_server("200 OK", "ok");
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(is_error(&types("[method]outgoing-request.set-authority", vec![request.clone(), s(&authority)])).is_none());
    assert!(is_error(&types("[method]outgoing-request.set-path-with-query", vec![request.clone(), s("api/items?x=1")])).is_none());
    let future = outgoing("handle", vec![request, Value::Null]);
    let response = types("[method]future-incoming-response.get", vec![future]);
    assert!(is_error(&response).is_none());
    assert_eq!(request_line.lock().unwrap().clone().as_deref(), Some("GET /api/items?x=1 HTTP/1.1"));
}

#[test]
fn blank_authority_is_rejected() {
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(is_error(&types("[method]outgoing-request.set-authority", vec![request.clone(), s("   ")])).is_none());
    let result = outgoing("handle", vec![request, Value::Null]);
    assert_eq!(is_error(&result).as_deref(), Some("HTTP-request-URI-invalid"));
}

#[test]
fn non_string_authority_is_stringified_and_can_fail_at_transport_time() {
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(is_error(&types("[method]outgoing-request.set-authority", vec![request.clone(), Value::I32(1)])).is_none());
    let future = outgoing("handle", vec![request, Value::Null]);
    assert!(is_error(&future).is_none(), "transport errors should still surface through the future");
    assert!(is_error(&types("[method]future-incoming-response.get", vec![future])).is_some());
}

#[test]
fn non_string_method_is_stringified_into_request_line() {
    let (authority, request_line) = capture_server("200 OK", "ok");
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(is_error(&types("[method]outgoing-request.set-method", vec![request.clone(), Value::I32(9)])).is_none());
    assert!(is_error(&types("[method]outgoing-request.set-authority", vec![request.clone(), s(&authority)])).is_none());
    let future = outgoing("handle", vec![request, Value::Null]);
    let response = types("[method]future-incoming-response.get", vec![future]);
    assert!(is_error(&response).is_none());
    assert_eq!(request_line.lock().unwrap().clone().as_deref(), Some("9 / HTTP/1.1"));
}

#[test]
fn non_string_path_uses_default_slash() {
    let (authority, request_line) = capture_server("200 OK", "ok");
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(is_error(&types("[method]outgoing-request.set-authority", vec![request.clone(), s(&authority)])).is_none());
    assert!(is_error(&types("[method]outgoing-request.set-path-with-query", vec![request.clone(), Value::I32(4)])).is_none());
    let future = outgoing("handle", vec![request, Value::Null]);
    let response = types("[method]future-incoming-response.get", vec![future]);
    assert!(is_error(&response).is_none());
    assert_eq!(request_line.lock().unwrap().clone().as_deref(), Some("GET /4 HTTP/1.1"));
}

#[test]
fn non_string_scheme_is_stringified_and_rejected_as_invalid_uri() {
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(is_error(&types("[method]outgoing-request.set-scheme", vec![request.clone(), Value::I32(7)])).is_none());
    assert!(is_error(&types("[method]outgoing-request.set-authority", vec![request.clone(), s("example.com")])).is_none());
    let result = outgoing("handle", vec![request, Value::Null]);
    assert_eq!(is_error(&result).as_deref(), Some("HTTP-request-URI-invalid"));
}

#[test]
fn transport_error_future_returns_error_then_already_consumed() {
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(is_error(&types("[method]outgoing-request.set-authority", vec![request.clone(), s("127.0.0.1:1")])).is_none());
    let future = outgoing("handle", vec![request, Value::Null]);
    let first = types("[method]future-incoming-response.get", vec![future.clone()]);
    assert_eq!(is_error(&first).as_deref(), Some("connection-refused"));
    let second = types("[method]future-incoming-response.get", vec![future]);
    assert_eq!(is_error(&second).as_deref(), Some("already-consumed"));
}

#[test]
fn response_status_tracks_non_ok_status_code() {
    let (authority, _) = capture_server("418 I'm a teapot", "short and stout");
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(is_error(&types("[method]outgoing-request.set-authority", vec![request.clone(), s(&authority)])).is_none());
    let future = outgoing("handle", vec![request, Value::Null]);
    let response = types("[method]future-incoming-response.get", vec![future]);
    assert_eq!(types("[method]incoming-response.status", vec![response]), Value::F64(418.0));
}

#[test]
fn legacy_fetch_marks_non_2xx_status_as_not_ok() {
    let (authority, _) = capture_server("404 Not Found", "missing");
    let result = legacy("fetch", vec![s(&format!("http://{}/missing", authority)), s("GET"), Value::Null]);
    assert_eq!(prop(&result, "status"), Value::F64(404.0));
    assert_eq!(prop(&result, "ok"), Value::Bool(false));
    assert_eq!(prop(&result, "body"), s("missing"));
}

#[test]
fn legacy_get_returns_error_string_for_unreachable_endpoint() {
    let result = legacy("get", vec![s("http://127.0.0.1:1/fail")]);
    let Value::String(text) = result else { panic!("legacy get should return string") };
    assert!(text.starts_with("Error: "));
}

#[test]
fn legacy_post_returns_error_string_for_unreachable_endpoint() {
    let result = legacy("post", vec![s("http://127.0.0.1:1/fail"), s("payload=1")]);
    let Value::String(text) = result else { panic!("legacy post should return string") };
    assert!(text.starts_with("Error: "));
}

#[test]
fn response_headers_resource_can_be_queried_multiple_times() {
    let (authority, _) = capture_server("200 OK", "ok");
    let headers = types("[constructor]fields", vec![]);
    let request = types("[constructor]outgoing-request", vec![headers]);
    assert!(is_error(&types("[method]outgoing-request.set-authority", vec![request.clone(), s(&authority)])).is_none());
    let future = outgoing("handle", vec![request, Value::Null]);
    let response = types("[method]future-incoming-response.get", vec![future]);
    let response_headers = types("[method]incoming-response.headers", vec![response]);
    let first = types("[method]fields.get", vec![response_headers.clone(), s("content-type")]);
    let second = types("[method]fields.get", vec![response_headers, s("content-type")]);
    assert_eq!(array_len(&first), 1);
    assert_eq!(array_len(&second), 1);
}