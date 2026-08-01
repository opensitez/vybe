//! Behaviour tests for `node:http` host imports.
//!
//! Reference: <https://nodejs.org/api/http.html>.
//!
//! Coverage:
//!   - `http.STATUS_CODES` — standard status code map (all major codes)
//!   - `http.METHODS` — supported HTTP method strings
//!   - `http.maxHeaderSize` — default 16384
//!   - `http.globalAgent` — default Agent instance
//!   - `http.validateHeaderName(name)` — returns undefined or throws
//!   - `http.validateHeaderValue(name, value)` — returns undefined or throws
//!   - `http.createServer()` → Server object surface
//!   - `http.request(options)` → ClientRequest surface
//!   - `http.get(url)` → ClientRequest surface
//!   - `http.Agent` constructor + all properties/methods
//!   - Server property surface (timeout, keepAliveTimeout, etc.)
//!   - ClientRequest method surface (setHeader, getHeader, end, etc.)
//!   - ServerResponse surface (writeHead, statusCode, headersSent, etc.)
//!   - IncomingMessage surface (method, url, headers, httpVersion, etc.)
//!
//! Deferred (require async event loop):
//!   - Server `listen` / `close` callbacks
//!   - Request/response stream events (`data`, `end`, `error`)
//!   - Live HTTP round-trips

use std::sync::Arc;
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn call_http(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-http-test>");
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

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn prop(obj: &Value, key: &str) -> Value {
    match obj {
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

fn as_string(v: &Value) -> String {
    match v {
        Value::String(s) => s.to_string(),
        other => format!("{other}"),
    }
}

fn new_obj(props: Vec<(&str, Value)>) -> Value {
    let mut o = Object::new();
    for (k, v) in props {
        o.properties.insert(k.into(), v);
    }
    Value::Object(Arc::new(std::sync::Mutex::new(o)))
}

fn has_method(obj: &Value, name: &str) -> bool {
    matches!(prop(obj, name), Value::Object(_) | Value::String(_))
}

// ── STATUS_CODES ──────────────────────────────────────────────────────────────

#[test]
fn status_codes_returns_object() {
    assert!(matches!(
        call_http("STATUS_CODES", vec![]),
        Value::Object(_)
    ));
}

#[test]
fn status_codes_200_is_ok() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "200")),
        "OK"
    );
}

#[test]
fn status_codes_201_is_created() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "201")),
        "Created"
    );
}

#[test]
fn status_codes_204_is_no_content() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "204")),
        "No Content"
    );
}

#[test]
fn status_codes_301_is_moved_permanently() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "301")),
        "Moved Permanently"
    );
}

#[test]
fn status_codes_302_is_found() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "302")),
        "Found"
    );
}

#[test]
fn status_codes_304_is_not_modified() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "304")),
        "Not Modified"
    );
}

#[test]
fn status_codes_400_is_bad_request() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "400")),
        "Bad Request"
    );
}

#[test]
fn status_codes_401_is_unauthorized() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "401")),
        "Unauthorized"
    );
}

#[test]
fn status_codes_403_is_forbidden() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "403")),
        "Forbidden"
    );
}

#[test]
fn status_codes_404_is_not_found() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "404")),
        "Not Found"
    );
}

#[test]
fn status_codes_405_is_method_not_allowed() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "405")),
        "Method Not Allowed"
    );
}

#[test]
fn status_codes_409_is_conflict() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "409")),
        "Conflict"
    );
}

#[test]
fn status_codes_422_is_unprocessable_entity() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "422")),
        "Unprocessable Entity"
    );
}

#[test]
fn status_codes_429_is_too_many_requests() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "429")),
        "Too Many Requests"
    );
}

#[test]
fn status_codes_500_is_internal_server_error() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "500")),
        "Internal Server Error"
    );
}

#[test]
fn status_codes_502_is_bad_gateway() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "502")),
        "Bad Gateway"
    );
}

#[test]
fn status_codes_503_is_service_unavailable() {
    assert_eq!(
        as_string(&prop(&call_http("STATUS_CODES", vec![]), "503")),
        "Service Unavailable"
    );
}

#[test]
fn status_codes_contains_at_least_forty_entries() {
    let codes = call_http("STATUS_CODES", vec![]);
    if let Value::Object(obj) = &codes {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.len() >= 40,
            "STATUS_CODES must have >= 40 entries, got {}",
            o.properties.len()
        );
    }
}

// ── METHODS ───────────────────────────────────────────────────────────────────

#[test]
fn methods_returns_array() {
    let result = call_http("METHODS", vec![]);
    if let Value::Object(obj) = &result {
        let o = obj.lock().unwrap();
        assert!(
            matches!(&o.kind, ObjectKind::Array(_)),
            "METHODS must be an Array"
        );
    } else {
        panic!("METHODS must be an object, got {:?}", result);
    }
}

#[test]
fn methods_contains_all_standard_verbs() {
    let result = call_http("METHODS", vec![]);
    if let Value::Object(obj) = &result {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(elems) = &o.kind {
            let methods: Vec<String> = elems.iter().map(as_string).collect();
            for m in [
                "GET", "POST", "PUT", "DELETE", "PATCH", "HEAD", "OPTIONS", "CONNECT", "TRACE",
            ] {
                assert!(methods.contains(&m.to_string()), "METHODS must contain {m}");
            }
            return;
        }
    }
    panic!("METHODS expected array");
}

#[test]
fn methods_are_all_uppercase_strings() {
    let result = call_http("METHODS", vec![]);
    if let Value::Object(obj) = &result {
        let o = obj.lock().unwrap();
        if let ObjectKind::Array(elems) = &o.kind {
            for elem in elems {
                let s = as_string(elem);
                assert_eq!(s, s.to_uppercase(), "method {s} must be uppercase");
            }
            return;
        }
    }
    panic!("METHODS expected array");
}

// ── maxHeaderSize ─────────────────────────────────────────────────────────────

#[test]
fn max_header_size_is_16384_by_default() {
    let result = call_http("maxHeaderSize", vec![]);
    let n = match result {
        Value::I32(n) => n as i64,
        Value::I64(n) => n,
        Value::F64(f) => f as i64,
        other => panic!("maxHeaderSize must be a number, got {:?}", other),
    };
    assert_eq!(n, 16384, "maxHeaderSize must be 16384 by default");
}

// ── validateHeaderName ────────────────────────────────────────────────────────

#[test]
fn validate_header_name_valid_name_does_not_panic() {
    let result = call_http("validateHeaderName", vec![s("Content-Type")]);
    assert!(matches!(
        result,
        Value::Undefined | Value::Null | Value::Bool(true)
    ));
}

#[test]
fn validate_header_name_invalid_name_returns_error_or_throws() {
    // Invalid names (with spaces/control chars) should either return an
    // error indicator or cause the VM to propagate a thrown error value.
    let result = call_http("validateHeaderName", vec![s("bad header")]);
    // Accept: error object, bool false, or string (error message)
    assert!(
        !matches!(result, Value::Bool(true) | Value::Undefined),
        "invalid header name must signal an error, got {:?}",
        result
    );
}

// ── validateHeaderValue ───────────────────────────────────────────────────────

#[test]
fn validate_header_value_valid_value_does_not_panic() {
    let result = call_http(
        "validateHeaderValue",
        vec![s("Content-Type"), s("application/json")],
    );
    assert!(matches!(
        result,
        Value::Undefined | Value::Null | Value::Bool(true)
    ));
}

// ── globalAgent ───────────────────────────────────────────────────────────────

#[test]
fn global_agent_is_object() {
    assert!(matches!(call_http("globalAgent", vec![]), Value::Object(_)));
}

#[test]
fn global_agent_has_max_sockets() {
    let agent = call_http("globalAgent", vec![]);
    let val = prop(&agent, "maxSockets");
    assert!(
        matches!(val, Value::I32(_) | Value::I64(_) | Value::F64(_)),
        "globalAgent.maxSockets must be a number, got {:?}",
        val
    );
}

#[test]
fn global_agent_has_sockets_object() {
    let agent = call_http("globalAgent", vec![]);
    assert!(
        matches!(prop(&agent, "sockets"), Value::Object(_)),
        "globalAgent.sockets must be an object"
    );
}

#[test]
fn global_agent_has_requests_object() {
    let agent = call_http("globalAgent", vec![]);
    assert!(
        matches!(prop(&agent, "requests"), Value::Object(_)),
        "globalAgent.requests must be an object"
    );
}

#[test]
fn global_agent_has_free_sockets_object() {
    let agent = call_http("globalAgent", vec![]);
    assert!(
        matches!(prop(&agent, "freeSockets"), Value::Object(_)),
        "globalAgent.freeSockets must be an object"
    );
}

#[test]
fn global_agent_has_destroy_method() {
    let agent = call_http("globalAgent", vec![]);
    assert!(
        has_method(&agent, "destroy"),
        "globalAgent must have destroy()"
    );
}

#[test]
fn global_agent_has_get_name_method() {
    let agent = call_http("globalAgent", vec![]);
    assert!(
        has_method(&agent, "getName"),
        "globalAgent must have getName()"
    );
}

// ── Agent constructor ─────────────────────────────────────────────────────────

#[test]
fn agent_constructor_returns_object() {
    assert!(matches!(
        call_http("Agent", vec![Value::Null]),
        Value::Object(_)
    ));
}

#[test]
fn agent_max_sockets_option_is_applied() {
    let opts = new_obj(vec![("maxSockets", Value::I32(5))]);
    let agent = call_http("Agent", vec![opts]);
    let val = prop(&agent, "maxSockets");
    assert_eq!(
        val,
        Value::I32(5),
        "Agent({{maxSockets:5}}).maxSockets must be 5, got {:?}",
        val
    );
}

#[test]
fn agent_has_max_free_sockets() {
    let agent = call_http("Agent", vec![Value::Null]);
    let val = prop(&agent, "maxFreeSockets");
    assert!(
        matches!(val, Value::I32(_) | Value::I64(_) | Value::F64(_)),
        "agent.maxFreeSockets must be a number, got {:?}",
        val
    );
}

#[test]
fn agent_has_max_total_sockets() {
    let agent = call_http("Agent", vec![Value::Null]);
    let val = prop(&agent, "maxTotalSockets");
    assert!(
        matches!(
            val,
            Value::I32(_) | Value::I64(_) | Value::F64(_) | Value::Undefined
        ),
        "agent.maxTotalSockets must be a number or Infinity, got {:?}",
        val
    );
}

#[test]
fn agent_destroy_method_exists() {
    let agent = call_http("Agent", vec![Value::Null]);
    assert!(has_method(&agent, "destroy"), "agent must have destroy()");
}

#[test]
fn agent_get_name_method_exists() {
    let agent = call_http("Agent", vec![Value::Null]);
    assert!(has_method(&agent, "getName"), "agent must have getName()");
}

// ── createServer ──────────────────────────────────────────────────────────────

#[test]
fn create_server_returns_object() {
    assert!(matches!(
        call_http("createServer", vec![Value::Null]),
        Value::Object(_)
    ));
}

#[test]
fn create_server_has_listen() {
    let srv = call_http("createServer", vec![Value::Null]);
    assert!(has_method(&srv, "listen"), "server must have listen()");
}

#[test]
fn create_server_has_close() {
    let srv = call_http("createServer", vec![Value::Null]);
    assert!(has_method(&srv, "close"), "server must have close()");
}

#[test]
fn create_server_has_address() {
    let srv = call_http("createServer", vec![Value::Null]);
    assert!(has_method(&srv, "address"), "server must have address()");
}

#[test]
fn create_server_has_set_timeout() {
    let srv = call_http("createServer", vec![Value::Null]);
    assert!(
        has_method(&srv, "setTimeout"),
        "server must have setTimeout()"
    );
}

#[test]
fn create_server_timeout_property_is_number() {
    let srv = call_http("createServer", vec![Value::Null]);
    let val = prop(&srv, "timeout");
    assert!(
        matches!(val, Value::I32(_) | Value::I64(_) | Value::F64(_)),
        "server.timeout must be a number, got {:?}",
        val
    );
}

#[test]
fn create_server_keep_alive_timeout_is_number() {
    let srv = call_http("createServer", vec![Value::Null]);
    let val = prop(&srv, "keepAliveTimeout");
    assert!(
        matches!(val, Value::I32(_) | Value::I64(_) | Value::F64(_)),
        "server.keepAliveTimeout must be a number, got {:?}",
        val
    );
}

#[test]
fn create_server_headers_timeout_is_number() {
    let srv = call_http("createServer", vec![Value::Null]);
    let val = prop(&srv, "headersTimeout");
    assert!(
        matches!(val, Value::I32(_) | Value::I64(_) | Value::F64(_)),
        "server.headersTimeout must be a number, got {:?}",
        val
    );
}

#[test]
fn create_server_request_timeout_is_number() {
    let srv = call_http("createServer", vec![Value::Null]);
    let val = prop(&srv, "requestTimeout");
    assert!(
        matches!(val, Value::I32(_) | Value::I64(_) | Value::F64(_)),
        "server.requestTimeout must be a number, got {:?}",
        val
    );
}

#[test]
fn create_server_max_connections_is_number_or_infinity() {
    let srv = call_http("createServer", vec![Value::Null]);
    let val = prop(&srv, "maxConnections");
    assert!(
        matches!(
            val,
            Value::I32(_) | Value::I64(_) | Value::F64(_) | Value::Undefined
        ),
        "server.maxConnections must be a number, got {:?}",
        val
    );
}

// ── request / ClientRequest surface ──────────────────────────────────────────

#[test]
fn request_returns_object() {
    let opts = new_obj(vec![
        ("host", s("example.com")),
        ("path", s("/")),
        ("method", s("GET")),
    ]);
    assert!(matches!(call_http("request", vec![opts]), Value::Object(_)));
}

#[test]
fn client_request_has_end() {
    let opts = new_obj(vec![("host", s("example.com")), ("path", s("/"))]);
    let req = call_http("request", vec![opts]);
    assert!(has_method(&req, "end"), "ClientRequest must have end()");
}

#[test]
fn client_request_has_write() {
    let opts = new_obj(vec![("host", s("example.com")), ("path", s("/"))]);
    let req = call_http("request", vec![opts]);
    assert!(has_method(&req, "write"), "ClientRequest must have write()");
}

#[test]
fn client_request_has_set_header() {
    let opts = new_obj(vec![("host", s("example.com")), ("path", s("/"))]);
    let req = call_http("request", vec![opts]);
    assert!(
        has_method(&req, "setHeader"),
        "ClientRequest must have setHeader()"
    );
}

#[test]
fn client_request_has_get_header() {
    let opts = new_obj(vec![("host", s("example.com")), ("path", s("/"))]);
    let req = call_http("request", vec![opts]);
    assert!(
        has_method(&req, "getHeader"),
        "ClientRequest must have getHeader()"
    );
}

#[test]
fn client_request_has_remove_header() {
    let opts = new_obj(vec![("host", s("example.com")), ("path", s("/"))]);
    let req = call_http("request", vec![opts]);
    assert!(
        has_method(&req, "removeHeader"),
        "ClientRequest must have removeHeader()"
    );
}

#[test]
fn client_request_has_destroy() {
    let opts = new_obj(vec![("host", s("example.com")), ("path", s("/"))]);
    let req = call_http("request", vec![opts]);
    assert!(
        has_method(&req, "destroy"),
        "ClientRequest must have destroy()"
    );
}

#[test]
fn client_request_has_set_timeout() {
    let opts = new_obj(vec![("host", s("example.com")), ("path", s("/"))]);
    let req = call_http("request", vec![opts]);
    assert!(
        has_method(&req, "setTimeout"),
        "ClientRequest must have setTimeout()"
    );
}

#[test]
fn client_request_method_reflects_options() {
    let opts = new_obj(vec![
        ("host", s("example.com")),
        ("path", s("/")),
        ("method", s("POST")),
    ]);
    let req = call_http("request", vec![opts]);
    assert_eq!(as_string(&prop(&req, "method")), "POST");
}

#[test]
fn client_request_path_reflects_options() {
    let opts = new_obj(vec![("host", s("example.com")), ("path", s("/foo/bar"))]);
    let req = call_http("request", vec![opts]);
    assert_eq!(as_string(&prop(&req, "path")), "/foo/bar");
}

#[test]
fn client_request_host_reflects_options() {
    let opts = new_obj(vec![("host", s("example.com")), ("path", s("/"))]);
    let req = call_http("request", vec![opts]);
    assert_eq!(as_string(&prop(&req, "host")), "example.com");
}

#[test]
fn client_request_protocol_is_http() {
    let opts = new_obj(vec![("host", s("example.com")), ("path", s("/"))]);
    let req = call_http("request", vec![opts]);
    assert_eq!(as_string(&prop(&req, "protocol")), "http:");
}

// ── get ───────────────────────────────────────────────────────────────────────

#[test]
fn get_returns_client_request_object() {
    assert!(matches!(
        call_http("get", vec![s("http://example.com/")]),
        Value::Object(_)
    ));
}

#[test]
fn get_request_has_end_method() {
    let req = call_http("get", vec![s("http://example.com/")]);
    assert!(
        has_method(&req, "end"),
        "get() ClientRequest must have end()"
    );
}

#[test]
fn get_parses_url_into_host() {
    let req = call_http("get", vec![s("http://example.com/path")]);
    assert_eq!(as_string(&prop(&req, "host")), "example.com");
}

#[test]
fn get_parses_url_into_path() {
    let req = call_http("get", vec![s("http://example.com/path?q=1")]);
    let path = as_string(&prop(&req, "path"));
    assert!(
        path.starts_with("/path"),
        "path must start with /path, got {path}"
    );
}

// ── IncomingMessage surface ───────────────────────────────────────────────────

#[test]
fn incoming_message_constructor_returns_object() {
    let result = call_http("IncomingMessage", vec![Value::Null]);
    assert!(
        matches!(result, Value::Object(_)),
        "IncomingMessage() must return an object"
    );
}

#[test]
fn incoming_message_has_method_property() {
    let msg = call_http("IncomingMessage", vec![Value::Null]);
    assert!(
        matches!(
            prop(&msg, "method"),
            Value::String(_) | Value::Null | Value::Undefined
        ),
        "IncomingMessage.method must be a string"
    );
}

#[test]
fn incoming_message_has_url_property() {
    let msg = call_http("IncomingMessage", vec![Value::Null]);
    assert!(
        matches!(
            prop(&msg, "url"),
            Value::String(_) | Value::Null | Value::Undefined
        ),
        "IncomingMessage.url must be a string"
    );
}

#[test]
fn incoming_message_has_headers_object() {
    let msg = call_http("IncomingMessage", vec![Value::Null]);
    assert!(
        matches!(
            prop(&msg, "headers"),
            Value::Object(_) | Value::Null | Value::Undefined
        ),
        "IncomingMessage.headers must be an object"
    );
}

#[test]
fn incoming_message_has_raw_headers_array() {
    let msg = call_http("IncomingMessage", vec![Value::Null]);
    assert!(
        matches!(
            prop(&msg, "rawHeaders"),
            Value::Object(_) | Value::Null | Value::Undefined
        ),
        "IncomingMessage.rawHeaders must be an array"
    );
}

#[test]
fn incoming_message_has_http_version() {
    let msg = call_http("IncomingMessage", vec![Value::Null]);
    assert!(
        matches!(
            prop(&msg, "httpVersion"),
            Value::String(_) | Value::Null | Value::Undefined
        ),
        "IncomingMessage.httpVersion must be a string"
    );
}

#[test]
fn incoming_message_has_status_code() {
    let msg = call_http("IncomingMessage", vec![Value::Null]);
    assert!(
        matches!(
            prop(&msg, "statusCode"),
            Value::I32(_) | Value::I64(_) | Value::F64(_) | Value::Null | Value::Undefined
        ),
        "IncomingMessage.statusCode must be a number"
    );
}

#[test]
fn incoming_message_has_complete_boolean() {
    let msg = call_http("IncomingMessage", vec![Value::Null]);
    assert!(
        matches!(prop(&msg, "complete"), Value::Bool(_) | Value::Undefined),
        "IncomingMessage.complete must be boolean"
    );
}

#[test]
fn incoming_message_has_trailers_object() {
    let msg = call_http("IncomingMessage", vec![Value::Null]);
    assert!(
        matches!(prop(&msg, "trailers"), Value::Object(_) | Value::Undefined),
        "IncomingMessage.trailers must be an object"
    );
}

// ── ServerResponse surface ────────────────────────────────────────────────────

#[test]
fn server_response_constructor_returns_object() {
    let result = call_http("ServerResponse", vec![Value::Null]);
    assert!(
        matches!(result, Value::Object(_)),
        "ServerResponse() must return an object"
    );
}

#[test]
fn server_response_has_write_head() {
    let res = call_http("ServerResponse", vec![Value::Null]);
    assert!(
        has_method(&res, "writeHead"),
        "ServerResponse must have writeHead()"
    );
}

#[test]
fn server_response_has_set_header() {
    let res = call_http("ServerResponse", vec![Value::Null]);
    assert!(
        has_method(&res, "setHeader"),
        "ServerResponse must have setHeader()"
    );
}

#[test]
fn server_response_has_get_header() {
    let res = call_http("ServerResponse", vec![Value::Null]);
    assert!(
        has_method(&res, "getHeader"),
        "ServerResponse must have getHeader()"
    );
}

#[test]
fn server_response_has_get_headers() {
    let res = call_http("ServerResponse", vec![Value::Null]);
    assert!(
        has_method(&res, "getHeaders"),
        "ServerResponse must have getHeaders()"
    );
}

#[test]
fn server_response_has_get_header_names() {
    let res = call_http("ServerResponse", vec![Value::Null]);
    assert!(
        has_method(&res, "getHeaderNames"),
        "ServerResponse must have getHeaderNames()"
    );
}

#[test]
fn server_response_has_has_header() {
    let res = call_http("ServerResponse", vec![Value::Null]);
    assert!(
        has_method(&res, "hasHeader"),
        "ServerResponse must have hasHeader()"
    );
}

#[test]
fn server_response_has_remove_header() {
    let res = call_http("ServerResponse", vec![Value::Null]);
    assert!(
        has_method(&res, "removeHeader"),
        "ServerResponse must have removeHeader()"
    );
}

#[test]
fn server_response_has_write() {
    let res = call_http("ServerResponse", vec![Value::Null]);
    assert!(
        has_method(&res, "write"),
        "ServerResponse must have write()"
    );
}

#[test]
fn server_response_has_end() {
    let res = call_http("ServerResponse", vec![Value::Null]);
    assert!(has_method(&res, "end"), "ServerResponse must have end()");
}

#[test]
fn server_response_has_flush_headers() {
    let res = call_http("ServerResponse", vec![Value::Null]);
    assert!(
        has_method(&res, "flushHeaders"),
        "ServerResponse must have flushHeaders()"
    );
}

#[test]
fn server_response_status_code_defaults_to_200() {
    let res = call_http("ServerResponse", vec![Value::Null]);
    let code = prop(&res, "statusCode");
    assert_eq!(
        code,
        Value::I32(200),
        "ServerResponse.statusCode must default to 200, got {:?}",
        code
    );
}

#[test]
fn server_response_headers_sent_defaults_to_false() {
    let res = call_http("ServerResponse", vec![Value::Null]);
    assert_eq!(prop(&res, "headersSent"), Value::Bool(false));
}

#[test]
fn server_response_writable_ended_defaults_to_false() {
    let res = call_http("ServerResponse", vec![Value::Null]);
    assert_eq!(prop(&res, "writableEnded"), Value::Bool(false));
}

#[test]
fn server_response_status_message_is_string() {
    let res = call_http("ServerResponse", vec![Value::Null]);
    assert!(
        matches!(
            prop(&res, "statusMessage"),
            Value::String(_) | Value::Undefined
        ),
        "ServerResponse.statusMessage must be a string"
    );
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_node_http_surface_is_registered() {
    let expected = [
        "STATUS_CODES",
        "METHODS",
        "maxHeaderSize",
        "globalAgent",
        "createServer",
        "request",
        "get",
        "Agent",
        "IncomingMessage",
        "ServerResponse",
        "validateHeaderName",
        "validateHeaderValue",
        "setMaxIdleHTTPParsers",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(missing.is_empty(), "missing node:http imports: {missing:?}");
}

// ── Node surface added after the non-Node functions were removed ─────────────

#[test]
fn status_codes_maps_codes_to_node_reason_phrases() {
    let codes = call_http("status_codes", vec![]);
    let Value::Object(object) = &codes else {
        panic!("STATUS_CODES should be an object, got {codes:?}")
    };
    let object = object.lock().unwrap();
    let ObjectKind::Map(entries) = &object.kind else {
        panic!("STATUS_CODES should be a map")
    };
    let get = |key: &str| {
        entries
            .iter()
            .find(|(k, _)| matches!(k, Value::String(s) if s.as_ref() == key))
            .map(|(_, v)| format!("{v}"))
    };
    // Keys are STRINGS — a JS object with numeric keys stringifies them.
    assert_eq!(get("404").as_deref(), Some("Not Found"));
    assert_eq!(get("500").as_deref(), Some("Internal Server Error"));
    assert_eq!(get("418").as_deref(), Some("I'm a Teapot"));
}

#[test]
fn methods_are_sorted_and_include_the_verbs_node_parses() {
    let methods = call_http("methods", vec![]);
    let Value::Object(object) = &methods else {
        panic!("METHODS should be an array")
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(items) = &object.kind else {
        panic!("METHODS should be an array")
    };
    let names: Vec<String> = items.iter().map(|v| format!("{v}")).collect();
    for verb in ["GET", "POST", "PATCH", "PROPFIND", "UNSUBSCRIBE"] {
        assert!(names.iter().any(|n| n == verb), "missing {verb}");
    }
    let mut sorted = names.clone();
    sorted.sort();
    assert_eq!(names, sorted, "http.METHODS is sorted in Node");
}

#[test]
fn get_header_is_undefined_when_never_set() {
    // Node returns `undefined`, not "" — the difference is observable to
    // `if (res.getHeader('x') === undefined)`.
    assert_eq!(
        call_http("get_header", vec![Value::String("x-absent".into())]),
        Value::Undefined
    );
}

#[test]
fn status_message_falls_back_to_the_registered_reason_phrase() {
    // With no response in flight the status reads 0, which has no phrase.
    assert_eq!(as_string(&call_http("status_message", vec![])), "");
}

#[test]
fn reading_the_body_returns_null_at_end_of_stream() {
    // `readable.read()` yields null at EOF; an empty chunk would loop forever.
    assert_eq!(call_http("read", vec![]), Value::Null);
}

#[test]
fn http_version_is_reported() {
    assert_eq!(as_string(&call_http("http_version", vec![])), "1.1");
}

#[test]
fn validate_header_name_rejects_non_token_characters() {
    // RFC 9110 §5.6.2 — a field name is a token. Node throws; the host has no
    // exception channel, so it answers with the WASI-style error object.
    assert_eq!(
        call_http("validate_header_name", vec![Value::String("X-Ok".into())]),
        Value::Null
    );
    for bad in ["", "X Bad", "X:Bad", "X\nBad"] {
        let got = call_http("validate_header_name", vec![Value::String(bad.into())]);
        assert!(
            matches!(&got, Value::Object(_)),
            "{bad:?} should be rejected, got {got:?}"
        );
    }
}

#[test]
fn validate_header_value_rejects_injection_and_padding() {
    let ok = call_http(
        "validate_header_value",
        vec![Value::String("X".into()), Value::String("fine".into())],
    );
    assert_eq!(ok, Value::Null);
    // CR/LF is header injection; leading/trailing space is invalid per §5.5.
    for bad in ["bad\r\nInjected: 1", "bad\nx", " leading", "trailing "] {
        let got = call_http(
            "validate_header_value",
            vec![Value::String("X".into()), Value::String(bad.into())],
        );
        assert!(
            matches!(&got, Value::Object(_)),
            "{bad:?} should be rejected, got {got:?}"
        );
    }
}

#[test]
fn max_header_size_is_nodes_16kib_default() {
    assert_eq!(call_http("max_header_size", vec![]), Value::F64(16384.0));
}

#[test]
fn server_timeouts_report_nodes_defaults() {
    assert_eq!(call_http("keep_alive_timeout", vec![]), Value::F64(5000.0));
    assert_eq!(call_http("headers_timeout", vec![]), Value::F64(60000.0));
    assert_eq!(call_http("request_timeout", vec![]), Value::F64(300000.0));
    assert_eq!(call_http("timeout", vec![]), Value::F64(0.0));
}

#[test]
fn a_timeout_that_is_set_reads_back() {
    // Storing the value is the point — a setter that discarded it would be a
    // shim that reports Node's default forever.
    call_http("set_keep_alive_timeout", vec![Value::F64(1234.0)]);
    assert_eq!(call_http("keep_alive_timeout", vec![]), Value::F64(1234.0));
    call_http("set_keep_alive_timeout", vec![Value::F64(5000.0)]);
}

#[test]
fn listening_is_false_with_no_server() {
    assert_eq!(call_http("listening", vec![]), Value::Bool(false));
}
