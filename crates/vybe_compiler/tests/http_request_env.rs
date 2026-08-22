//! `primitives/http_request_env` reads the request from `wasi:http`, not from any
//! language-specific context.
//!
//! `documentation/httpserver.md` §4a: PHP `$_SERVER`/`$_GET`, WSGI `environ`
//! and Rack `env` are the same data renamed. These tests drive the primitive
//! directly, so a regression shows up here rather than in one language's suite.

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_compiler::primitives::{dispatch, http_request_env};
use vybe_platform_wasi::http as wasi_http;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

/// Publish a request the way `vybex --serve` does, then run `common:<op>` and
/// return what it produced.
fn eval(op: &str, method: &str, path_with_query: &str) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());

    let request_id = wasi_http::push_incoming_request(
        method,
        Some(path_with_query.to_string()),
        Some("https".to_string()),
        Some("api.test:443".to_string()),
        vec![("x-demo".to_string(), b"1".to_vec())],
        Vec::new(),
    );
    let handle = wasi_http::incoming_request_value(&vm, request_id).expect("request handle");
    vm.set_global_owned(http_request_env::REQUEST_GLOBAL.to_string(), handle);

    let mut chunks = vec![Chunk::new("<request-env-test>")];
    let handled = dispatch::emit_common(op, &mut chunks, 0, 0, 0);
    assert!(handled, "common:{op} was not dispatched");
    chunks[0].emit_op(Op::RETURN, 0);

    vm.run(chunks).expect("VM run failed")
}

fn str_of(value: &Value) -> Option<String> {
    match value {
        Value::String(text) => Some(text.to_string()),
        _ => None,
    }
}

#[test]
fn method_comes_from_the_wasi_request() {
    let got = eval("http_request.method", "DELETE", "/x");
    assert_eq!(str_of(&got).as_deref(), Some("DELETE"), "got {got:?}");
}

#[test]
fn path_with_query_is_reported_whole() {
    let got = eval("http_request.path_with_query", "GET", "/search?q=1&p=2");
    assert_eq!(
        str_of(&got).as_deref(),
        Some("/search?q=1&p=2"),
        "got {got:?}"
    );
}

#[test]
fn path_drops_the_query_string() {
    // `wasi:http` reports one `path-with-query`; PATH_INFO / $_SERVER want the
    // path alone.
    let got = eval("http_request.path", "GET", "/search?q=1&p=2");
    assert_eq!(str_of(&got).as_deref(), Some("/search"), "got {got:?}");
}

#[test]
fn path_is_the_whole_value_when_there_is_no_query() {
    let got = eval("http_request.path", "GET", "/plain");
    assert_eq!(str_of(&got).as_deref(), Some("/plain"), "got {got:?}");
}

#[test]
fn query_string_is_the_part_after_the_question_mark() {
    let got = eval("http_request.query_string", "GET", "/search?q=1&p=2");
    assert_eq!(str_of(&got).as_deref(), Some("q=1&p=2"), "got {got:?}");
}

#[test]
fn query_string_is_empty_when_absent() {
    // QUERY_STRING is "" for a query-less request, not null — every CGI-shaped
    // surface (PHP, WSGI, Rack) expects a string here.
    let got = eval("http_request.query_string", "GET", "/plain");
    assert_eq!(str_of(&got).as_deref(), Some(""), "got {got:?}");
}

#[test]
fn query_string_is_empty_for_a_trailing_question_mark() {
    let got = eval("http_request.query_string", "GET", "/plain?");
    assert_eq!(str_of(&got).as_deref(), Some(""), "got {got:?}");
}

#[test]
fn scheme_and_authority_come_through() {
    let scheme = eval("http_request.scheme", "GET", "/");
    assert_eq!(str_of(&scheme).as_deref(), Some("https"), "got {scheme:?}");

    let authority = eval("http_request.authority", "GET", "/");
    assert_eq!(
        str_of(&authority).as_deref(),
        Some("api.test:443"),
        "got {authority:?}"
    );
}

// ── CGI environ ─────────────────────────────────────────────────────────────

/// Publish a request plus the server's deployment metadata, then build environ.
fn environ(method: &str, path_with_query: &str, scheme: &str) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());

    let request_id = wasi_http::push_incoming_request(
        method,
        Some(path_with_query.to_string()),
        Some(scheme.to_string()),
        Some("api.test:443".to_string()),
        vec![
            ("content-type".to_string(), b"application/json".to_vec()),
            ("content-length".to_string(), b"17".to_vec()),
        ],
        Vec::new(),
    );
    let handle = wasi_http::incoming_request_value(&vm, request_id).expect("request handle");
    vm.set_global_owned(http_request_env::REQUEST_GLOBAL.to_string(), handle);

    let mut chunks = vec![Chunk::new("<environ-test>")];
    assert!(dispatch::emit_common(
        "http_request.environ",
        &mut chunks,
        0,
        0,
        0
    ));
    chunks[0].emit_op(Op::RETURN, 0);
    vm.run(chunks).expect("VM run failed")
}

fn key(map: &Value, name: &str) -> Option<String> {
    let Value::Object(object) = map else {
        return None;
    };
    let object = object.lock().unwrap();
    match &object.kind {
        vybe_runtime::value::ObjectKind::Map(entries) => entries
            .iter()
            .find(|(k, _)| matches!(k, Value::String(s) if s.as_ref() == name))
            .and_then(|(_, v)| str_of(v)),
        _ => None,
    }
}

#[test]
fn environ_carries_the_cgi_request_keys() {
    // Symfony's `Request::createFromGlobals()` INDEXES these rather than
    // iterating, so a missing key silently yields a wrong route, not an error.
    let env = environ("POST", "/orders?page=2", "https");

    assert_eq!(key(&env, "REQUEST_METHOD").as_deref(), Some("POST"));
    assert_eq!(key(&env, "REQUEST_URI").as_deref(), Some("/orders?page=2"));
    assert_eq!(key(&env, "PATH_INFO").as_deref(), Some("/orders"));
    assert_eq!(key(&env, "QUERY_STRING").as_deref(), Some("page=2"));
    assert_eq!(key(&env, "HTTP_HOST").as_deref(), Some("api.test:443"));
    assert_eq!(key(&env, "SERVER_NAME").as_deref(), Some("api.test:443"));
    assert_eq!(key(&env, "REQUEST_SCHEME").as_deref(), Some("https"));
    assert_eq!(
        key(&env, "CONTENT_TYPE").as_deref(),
        Some("application/json")
    );
    assert_eq!(key(&env, "CONTENT_LENGTH").as_deref(), Some("17"));
}

#[test]
fn https_is_on_for_a_secure_request() {
    // `Request::isSecure()` tests `$_SERVER['HTTPS']` for a non-"off" value.
    let env = environ("GET", "/", "https");
    assert_eq!(key(&env, "HTTPS").as_deref(), Some("on"));
}

#[test]
fn https_is_absent_for_a_plain_request() {
    // CGI omits HTTPS entirely over http; a present "off" would also be
    // acceptable to Symfony, but absent is what servers actually send.
    let env = environ("GET", "/", "http");
    assert_eq!(key(&env, "HTTPS"), None, "HTTPS must not be set over http");
}

#[test]
fn every_header_appears_under_its_cgi_http_name() {
    // CGI §4.1.18: `Content-Type` → `HTTP_CONTENT_TYPE`, uppercased with `-`
    // replaced by `_`. Symfony rebuilds its header bag from these, and
    // WSGI/Rack use the identical convention.
    let env = environ("GET", "/", "https");
    assert_eq!(
        key(&env, "HTTP_CONTENT_TYPE").as_deref(),
        Some("application/json"),
        "header should also appear under its HTTP_ name"
    );
    assert_eq!(key(&env, "HTTP_CONTENT_LENGTH").as_deref(), Some("17"));
}

/// Every request-shaped op, run with NO request published at all.
///
/// A CLI script is compiled by the same pipeline as a served one, so a PHP
/// file that merely mentions `$_SERVER` reaches these ops with
/// `REQUEST_GLOBAL` unset. Trapping here would kill every non-served script
/// that touches a superglobal — the ops must degrade to an empty map, the way
/// real PHP gives `$_GET === []` on the command line.
#[test]
fn every_request_op_survives_with_no_request() {
    for op in [
        "http_request.environ",
        "http_request.query_params",
        "http_request.request_params",
        "http_form.fields",
        "http_form.files",
        "http_cookie.request_cookies",
    ] {
        let mut vm = VM::new();
        register_platforms(&mut vm, &Capabilities::all());

        let mut chunks = vec![Chunk::new("<no-request-test>")];
        assert!(
            dispatch::emit_common(op, &mut chunks, 0, 0, 0),
            "common:{op} was not dispatched"
        );
        chunks[0].emit_op(Op::RETURN, 0);

        let got = vm
            .run(chunks)
            .unwrap_or_else(|e| panic!("common:{op} trapped with no request: {e}"));
        assert!(
            matches!(&got, Value::Object(_)),
            "common:{op} should give an empty map with no request, got {got:?}"
        );
    }
}
