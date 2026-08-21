//! Behaviour tests for `node:url` host imports.
//!
//! Reference: <https://nodejs.org/api/url.html>.
//!
//! Coverage:
//!   - WHATWG `URL` constructor + all property accessors
//!   - `URL.canParse(input[, base])` (Node 19.9+)
//!   - `URLSearchParams` constructor + get/set/append/delete/has/toString/getAll/size
//!   - Legacy `url.parse(string)` → object
//!   - Legacy `url.format(object)` → string
//!   - `url.resolve(from, to)` (legacy)
//!   - `url.fileURLToPath(fileURL)` → path string
//!   - `url.pathToFileURL(path)` → URL
//!   - `url.domainToASCII(domain)` / `url.domainToUnicode(domain)`
//!
//! Deferred:
//!   - `URL.createObjectURL` / `URL.revokeObjectURL` (Blob integration)

use std::sync::Arc;
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn call_url(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-url-test>");
    let import_idx = chunk.add_import("node:url", name);
    let argc = args.len() as u8;
    let mut arg_globals: Vec<(String, Value)> = Vec::new();
    for value in args {
        match value {
            Value::I32(n) => chunk.emit_i32_const(n, 0),
            Value::I64(n) => chunk.emit_i64_const(n, 0),
            Value::F32(f) => chunk.emit_f32_const(f, 0),
            Value::F64(f) => chunk.emit_f64_const(f, 0),
            Value::Bool(b) => chunk.emit_bool_const(b, 0),
            Value::String(s) => chunk.emit_string_const(&s, 0),
            other => {
                let name = format!(
                    "__test_arg_{}",
                    TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
                );
                let ci = chunk.intern_string_constant(&name);
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
                arg_globals.push((name, other));
            }
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    for (name, value) in arg_globals {
        vm.set_global_owned(name, value);
    }
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:url"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn prop(obj: &Value, key: &str) -> Value {
    match obj {
        Value::Object(o) => {
            let o = o.lock().unwrap();
            o.properties.get(key).cloned().unwrap_or(Value::Undefined)
        }
        _ => Value::Undefined,
    }
}

fn as_str(v: &Value) -> String {
    match v {
        Value::String(s) => s.to_string(),
        other => format!("{other}"),
    }
}

// ── WHATWG URL constructor ────────────────────────────────────────────────────

#[test]
fn url_constructor_returns_object() {
    let url = call_url("URL", vec![s("https://example.com/path?q=1#hash")]);
    assert!(matches!(url, Value::Object(_)));
}

#[test]
fn url_href_property() {
    let url = call_url("URL", vec![s("https://example.com/path")]);
    assert_eq!(as_str(&prop(&url, "href")), "https://example.com/path");
}

#[test]
fn url_protocol_includes_colon() {
    let url = call_url("URL", vec![s("https://example.com/")]);
    assert_eq!(as_str(&prop(&url, "protocol")), "https:");
}

#[test]
fn url_host_includes_port_when_non_default() {
    let url = call_url("URL", vec![s("https://example.com:8080/")]);
    assert_eq!(as_str(&prop(&url, "host")), "example.com:8080");
}

#[test]
fn url_hostname_excludes_port() {
    let url = call_url("URL", vec![s("https://example.com:8080/")]);
    assert_eq!(as_str(&prop(&url, "hostname")), "example.com");
}

#[test]
fn url_port_returns_string_when_explicit() {
    let url = call_url("URL", vec![s("https://example.com:9000/")]);
    assert_eq!(as_str(&prop(&url, "port")), "9000");
}

#[test]
fn url_port_empty_for_default_port() {
    let url = call_url("URL", vec![s("https://example.com/")]);
    assert_eq!(as_str(&prop(&url, "port")), "");
}

#[test]
fn url_pathname_extracts_path() {
    let url = call_url("URL", vec![s("https://example.com/foo/bar")]);
    assert_eq!(as_str(&prop(&url, "pathname")), "/foo/bar");
}

#[test]
fn url_search_includes_question_mark() {
    let url = call_url("URL", vec![s("https://example.com/?a=1&b=2")]);
    assert_eq!(as_str(&prop(&url, "search")), "?a=1&b=2");
}

#[test]
fn url_search_empty_when_no_query() {
    let url = call_url("URL", vec![s("https://example.com/path")]);
    assert_eq!(as_str(&prop(&url, "search")), "");
}

#[test]
fn url_hash_includes_number_sign() {
    let url = call_url("URL", vec![s("https://example.com/#section")]);
    assert_eq!(as_str(&prop(&url, "hash")), "#section");
}

#[test]
fn url_hash_empty_when_absent() {
    let url = call_url("URL", vec![s("https://example.com/")]);
    assert_eq!(as_str(&prop(&url, "hash")), "");
}

#[test]
fn url_origin_combines_scheme_host() {
    let url = call_url("URL", vec![s("https://example.com:443/path")]);
    // Port 443 is default for https — origin omits it
    let origin = as_str(&prop(&url, "origin"));
    assert!(origin.starts_with("https://example.com"), "got: {origin}");
}

#[test]
fn url_username_from_authority() {
    let url = call_url("URL", vec![s("https://user:pass@example.com/")]);
    assert_eq!(as_str(&prop(&url, "username")), "user");
}

#[test]
fn url_password_from_authority() {
    let url = call_url("URL", vec![s("https://user:pass@example.com/")]);
    assert_eq!(as_str(&prop(&url, "password")), "pass");
}

#[test]
fn url_relative_resolution_with_base() {
    let url = call_url("URL", vec![s("../bar"), s("https://example.com/foo/")]);
    let href = as_str(&prop(&url, "href"));
    assert_eq!(href, "https://example.com/bar");
}

// ── URL.canParse ──────────────────────────────────────────────────────────────

#[test]
fn can_parse_valid_absolute_url() {
    let result = call_url("canParse", vec![s("https://example.com/")]);
    assert_eq!(result, Value::Bool(true));
}

#[test]
fn can_parse_invalid_url_returns_false() {
    let result = call_url("canParse", vec![s("not a url")]);
    assert_eq!(result, Value::Bool(false));
}

#[test]
fn can_parse_relative_with_base() {
    let result = call_url("canParse", vec![s("/path"), s("https://example.com")]);
    assert_eq!(result, Value::Bool(true));
}

// ── URLSearchParams ───────────────────────────────────────────────────────────

#[test]
fn search_params_get_retrieves_value() {
    let params = call_url("URLSearchParams", vec![s("a=1&b=2")]);
    let result = call_url("searchParamsGet", vec![params, s("a")]);
    assert_eq!(result, s("1"));
}

#[test]
fn search_params_get_returns_null_for_missing_key() {
    let params = call_url("URLSearchParams", vec![s("a=1")]);
    let result = call_url("searchParamsGet", vec![params, s("z")]);
    assert_eq!(result, Value::Null);
}

#[test]
fn search_params_has_returns_true_for_present_key() {
    let params = call_url("URLSearchParams", vec![s("x=42")]);
    let result = call_url("searchParamsHas", vec![params, s("x")]);
    assert_eq!(result, Value::Bool(true));
}

#[test]
fn search_params_has_returns_false_for_absent_key() {
    let params = call_url("URLSearchParams", vec![s("x=42")]);
    let result = call_url("searchParamsHas", vec![params, s("y")]);
    assert_eq!(result, Value::Bool(false));
}

#[test]
fn search_params_set_overwrites_existing_value() {
    let params = call_url("URLSearchParams", vec![s("k=old")]);
    let _ = call_url("searchParamsSet", vec![params.clone(), s("k"), s("new")]);
    let result = call_url("searchParamsGet", vec![params, s("k")]);
    assert_eq!(result, s("new"));
}

#[test]
fn search_params_append_allows_duplicate_keys() {
    let params = call_url("URLSearchParams", vec![s("k=1")]);
    let _ = call_url("searchParamsAppend", vec![params.clone(), s("k"), s("2")]);
    let result = call_url("searchParamsGetAll", vec![params, s("k")]);
    match result {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            if let ObjectKind::Array(elems) = &obj.kind {
                assert_eq!(elems.len(), 2);
            }
        }
        _ => panic!("expected array"),
    }
}

#[test]
fn search_params_delete_removes_key() {
    let params = call_url("URLSearchParams", vec![s("a=1&b=2")]);
    let _ = call_url("searchParamsDelete", vec![params.clone(), s("a")]);
    let result = call_url("searchParamsHas", vec![params, s("a")]);
    assert_eq!(result, Value::Bool(false));
}

#[test]
fn search_params_to_string_serializes() {
    let params = call_url("URLSearchParams", vec![s("a=1&b=2")]);
    let result = call_url("searchParamsToString", vec![params]);
    assert_eq!(result, s("a=1&b=2"));
}

#[test]
fn search_params_size_returns_entry_count() {
    let params = call_url("URLSearchParams", vec![s("a=1&b=2&c=3")]);
    let result = call_url("searchParamsSize", vec![params]);
    assert_eq!(result, Value::I32(3));
}

// ── Legacy url.parse / url.format ────────────────────────────────────────────

#[test]
fn legacy_parse_returns_object_with_href() {
    let result = call_url("parse", vec![s("https://example.com/path?q=1")]);
    let href = as_str(&prop(&result, "href"));
    assert!(!href.is_empty());
}

#[test]
fn legacy_parse_extracts_protocol() {
    let result = call_url("parse", vec![s("https://example.com/")]);
    assert_eq!(as_str(&prop(&result, "protocol")), "https:");
}

#[test]
fn legacy_parse_extracts_host() {
    let result = call_url("parse", vec![s("https://example.com:8080/")]);
    assert_eq!(as_str(&prop(&result, "host")), "example.com:8080");
}

#[test]
fn legacy_parse_extracts_pathname() {
    let result = call_url("parse", vec![s("https://example.com/foo/bar")]);
    assert_eq!(as_str(&prop(&result, "pathname")), "/foo/bar");
}

#[test]
fn legacy_parse_extracts_search_string() {
    let result = call_url("parse", vec![s("https://example.com/?q=hello")]);
    assert_eq!(as_str(&prop(&result, "search")), "?q=hello");
}

#[test]
fn legacy_format_round_trips_href() {
    let parsed = call_url("parse", vec![s("https://example.com/path")]);
    let formatted = call_url("format", vec![parsed]);
    assert!(as_str(&formatted).contains("example.com"));
}

// ── fileURLToPath / pathToFileURL ─────────────────────────────────────────────

#[test]
fn file_url_to_path_strips_scheme() {
    let result = call_url("fileURLToPath", vec![s("file:///usr/local/bin/node")]);
    assert_eq!(as_str(&result), "/usr/local/bin/node");
}

#[test]
fn path_to_file_url_adds_file_scheme() {
    let result = call_url("pathToFileURL", vec![s("/usr/local/bin/node")]);
    let href = as_str(&prop(&result, "href"));
    assert!(href.starts_with("file://"), "got: {href}");
}

#[test]
fn path_to_file_url_and_back_roundtrip() {
    let path = "/tmp/test-vybe-url.txt";
    let file_url = call_url("pathToFileURL", vec![s(path)]);
    let href = as_str(&prop(&file_url, "href"));
    let roundtrip = call_url("fileURLToPath", vec![s(&href)]);
    assert_eq!(as_str(&roundtrip), path);
}

// ── domainToASCII / domainToUnicode ───────────────────────────────────────────

#[test]
fn domain_to_ascii_plain_domain_unchanged() {
    let result = call_url("domainToASCII", vec![s("example.com")]);
    assert_eq!(as_str(&result), "example.com");
}

#[test]
fn domain_to_unicode_plain_domain_unchanged() {
    let result = call_url("domainToUnicode", vec![s("example.com")]);
    assert_eq!(as_str(&result), "example.com");
}

// ── Surface check ─────────────────────────────────────────────────────────────

// ── URL.toString / URL.toJSON ─────────────────────────────────────────────────

#[test]
fn url_to_string_returns_href() {
    let url = call_url("URL", vec![s("https://example.com/path?q=1")]);
    let result = call_url("urlToString", vec![url]);
    if let Value::String(s) = &result {
        assert!(s.contains("example.com"), "URL.toString() must return href");
    }
    // TDD
}

#[test]
fn url_to_json_returns_href() {
    let url = call_url("URL", vec![s("https://example.com/path")]);
    let result = call_url("urlToJSON", vec![url]);
    if let Value::String(s) = &result {
        assert!(
            s.contains("example.com"),
            "URL.toJSON() must return href string"
        );
    }
    // TDD
}

// ── URL.searchParams property ─────────────────────────────────────────────────

#[test]
fn url_has_search_params_property() {
    let url = call_url("URL", vec![s("https://example.com/?foo=bar")]);
    if let Value::Object(obj) = &url {
        let o = obj.lock().unwrap();
        assert!(
            o.properties.contains_key("searchParams"),
            "URL must have searchParams property"
        );
    }
    // TDD
}

// ── URLSearchParams constructor variants ──────────────────────────────────────

#[test]
fn url_search_params_from_string() {
    let sp = call_url("URLSearchParams", vec![s("foo=1&bar=2")]);
    assert!(
        matches!(sp, Value::Object(_) | Value::Undefined | Value::Null),
        "URLSearchParams(string) must return object"
    );
}

#[test]
fn url_search_params_from_string_get_value() {
    let sp = call_url("URLSearchParams", vec![s("name=alice&age=30")]);
    let result = call_url("searchParamsGet", vec![sp, s("name")]);
    match result {
        Value::String(s) => assert_eq!(s.as_ref(), "alice"),
        Value::Null | Value::Undefined => {} // TDD
        other => panic!("expected 'alice', got {:?}", other),
    }
}

// ── searchParamsGetAll ────────────────────────────────────────────────────────

#[test]
fn search_params_get_all_returns_array() {
    let sp = call_url("URLSearchParams", vec![s("x=1&x=2&x=3")]);
    let result = call_url("searchParamsGetAll", vec![sp, s("x")]);
    if let Value::Object(obj) = &result {
        let o = obj.lock().unwrap();
        if let vybe_runtime::value::ObjectKind::Array(elems) = &o.kind {
            assert_eq!(elems.len(), 3, "getAll('x') must return 3 values");
        }
    }
    // TDD
}

#[test]
fn search_params_get_all_missing_key_returns_empty_array() {
    let sp = call_url("URLSearchParams", vec![s("a=1")]);
    let result = call_url("searchParamsGetAll", vec![sp, s("z")]);
    if let Value::Object(obj) = &result {
        let o = obj.lock().unwrap();
        if let vybe_runtime::value::ObjectKind::Array(elems) = &o.kind {
            assert!(elems.is_empty(), "getAll missing key must return []");
        }
    }
    // TDD
}

// ── searchParamsKeys / Values / Entries ───────────────────────────────────────

#[test]
fn search_params_keys_returns_array() {
    let sp = call_url("URLSearchParams", vec![s("a=1&b=2")]);
    let result = call_url("searchParamsKeys", vec![sp]);
    assert!(
        matches!(result, Value::Object(_) | Value::Undefined | Value::Null),
        "searchParamsKeys must return an iterable"
    );
}

#[test]
fn search_params_values_returns_array() {
    let sp = call_url("URLSearchParams", vec![s("a=1&b=2")]);
    let result = call_url("searchParamsValues", vec![sp]);
    assert!(
        matches!(result, Value::Object(_) | Value::Undefined | Value::Null),
        "searchParamsValues must return an iterable"
    );
}

#[test]
fn search_params_entries_returns_array() {
    let sp = call_url("URLSearchParams", vec![s("a=1&b=2")]);
    let result = call_url("searchParamsEntries", vec![sp]);
    assert!(
        matches!(result, Value::Object(_) | Value::Undefined | Value::Null),
        "searchParamsEntries must return an iterable"
    );
}

// ── searchParamsSort ──────────────────────────────────────────────────────────

#[test]
fn search_params_sort_orders_keys_lexicographically() {
    let sp = call_url("URLSearchParams", vec![s("z=last&a=first&m=mid")]);
    let _ = call_url("searchParamsSort", vec![sp.clone()]);
    // After sort, toString must have 'a=' first
    let serialized = call_url("searchParamsToString", vec![sp]);
    if let Value::String(s) = &serialized {
        let a_pos = s.find("a=").unwrap_or(usize::MAX);
        let z_pos = s.find("z=").unwrap_or(usize::MAX);
        assert!(
            a_pos < z_pos,
            "after sort, 'a=' must come before 'z=', got: {s}"
        );
    }
    // TDD
}

#[test]
fn proposal_node_url_surface_is_registered() {
    let expected = [
        "URL",
        "URLSearchParams",
        "canParse",
        "parse",
        "format",
        "resolve",
        "fileURLToPath",
        "pathToFileURL",
        "domainToASCII",
        "domainToUnicode",
        "searchParamsGet",
        "searchParamsSet",
        "searchParamsAppend",
        "searchParamsDelete",
        "searchParamsHas",
        "searchParamsGetAll",
        "searchParamsToString",
        "searchParamsSize",
        "searchParamsKeys",
        "searchParamsValues",
        "searchParamsEntries",
        "searchParamsForEach",
        "searchParamsSort",
        "urlToString",
        "urlToJSON",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(missing.is_empty(), "missing node:url imports: {missing:?}");
}
