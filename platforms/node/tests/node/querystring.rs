//! Behaviour tests for `node:querystring` host imports.
//!
//! Reference: <https://nodejs.org/api/querystring.html>.
//!
//! Coverage:
//!   - `parse(str[, sep[, eq[, options]]])` → object
//!   - `stringify(obj[, sep[, eq[, options]]])` → string
//!   - `escape(str)` → percent-encoded string
//!   - `unescape(str)` → decoded string
//!
//! Note: querystring is legacy — the WHATWG URLSearchParams API in `node:url`
//! is preferred for new code, but querystring remains in LTS Node.js.

use std::sync::Arc;
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_emitter::platforms::register_platforms;

fn call_qs(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-querystring-test>");
    let import_idx = chunk.add_import("node:querystring", name);
    let argc = args.len() as u8;
    for value in args {
        let c = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, c, 0);
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
        .contains_key(&(String::from("node:querystring"), name.to_string()))
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

// ── parse ─────────────────────────────────────────────────────────────────────

#[test]
fn parse_simple_key_value_pairs() {
    let result = call_qs("parse", vec![s("a=1&b=2")]);
    assert_eq!(as_str(&prop(&result, "a")), "1");
    assert_eq!(as_str(&prop(&result, "b")), "2");
}

#[test]
fn parse_empty_string_returns_empty_object() {
    let result = call_qs("parse", vec![s("")]);
    assert!(matches!(result, Value::Object(_)));
    match &result {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            assert!(obj.properties.is_empty());
        }
        _ => panic!("expected object"),
    }
}

#[test]
fn parse_repeated_key_produces_array_value() {
    let result = call_qs("parse", vec![s("a=1&a=2")]);
    let val = prop(&result, "a");
    assert!(matches!(val, Value::Object(_)));
    match &val {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            assert!(matches!(&obj.kind, ObjectKind::Array(elems) if elems.len() == 2));
        }
        _ => panic!("expected array for repeated key"),
    }
}

#[test]
fn parse_percent_decodes_values() {
    let result = call_qs("parse", vec![s("name=hello%20world")]);
    assert_eq!(as_str(&prop(&result, "name")), "hello world");
}

#[test]
fn parse_plus_sign_decoded_as_space() {
    let result = call_qs("parse", vec![s("name=hello+world")]);
    assert_eq!(as_str(&prop(&result, "name")), "hello world");
}

#[test]
fn parse_key_without_value_defaults_to_empty_string() {
    let result = call_qs("parse", vec![s("bare")]);
    assert_eq!(as_str(&prop(&result, "bare")), "");
}

#[test]
fn parse_custom_separator_semicolon() {
    let result = call_qs("parse", vec![s("a=1;b=2"), s(";")]);
    assert_eq!(as_str(&prop(&result, "a")), "1");
    assert_eq!(as_str(&prop(&result, "b")), "2");
}

#[test]
fn parse_custom_equals_colon() {
    let result = call_qs("parse", vec![s("a:1&b:2"), s("&"), s(":")]);
    assert_eq!(as_str(&prop(&result, "a")), "1");
    assert_eq!(as_str(&prop(&result, "b")), "2");
}

#[test]
fn parse_unicode_percent_encoded_key() {
    // "café" percent-encoded key
    let result = call_qs("parse", vec![s("caf%C3%A9=yes")]);
    assert_eq!(as_str(&prop(&result, "café")), "yes");
}

// ── stringify ─────────────────────────────────────────────────────────────────

#[test]
fn stringify_simple_object() {
    let mut obj = Object::new();
    obj.properties.insert("a".to_string(), s("1"));
    obj.properties.insert("b".to_string(), s("2"));
    let input = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(obj)));
    let result = call_qs("stringify", vec![input]);
    let out = as_str(&result);
    assert!(out.contains("a=1"), "got: {out}");
    assert!(out.contains("b=2"), "got: {out}");
    assert!(out.contains('&'), "got: {out}");
}

#[test]
fn stringify_array_value_produces_repeated_key() {
    let mut obj = Object::new();
    let arr = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(Object {
        kind: ObjectKind::Array(vec![s("x"), s("y")]),
        properties: std::collections::HashMap::new(),
        type_id: 0,
        fields: Vec::new(),
    })));
    obj.properties.insert("k".to_string(), arr);
    let input = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(obj)));
    let result = call_qs("stringify", vec![input]);
    let out = as_str(&result);
    assert!(out.contains("k=x"), "got: {out}");
    assert!(out.contains("k=y"), "got: {out}");
}

#[test]
fn stringify_encodes_special_characters() {
    let mut obj = Object::new();
    obj.properties.insert("q".to_string(), s("hello world"));
    let input = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(obj)));
    let result = call_qs("stringify", vec![input]);
    let out = as_str(&result);
    // spaces encoded as %20 or + depending on impl
    assert!(out.contains("q=") && (out.contains("%20") || out.contains('+')));
}

#[test]
fn stringify_empty_object_returns_empty_string() {
    let input = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(Object::new())));
    let result = call_qs("stringify", vec![input]);
    assert_eq!(as_str(&result), "");
}

#[test]
fn stringify_custom_separator() {
    let mut obj = Object::new();
    obj.properties.insert("a".to_string(), s("1"));
    obj.properties.insert("b".to_string(), s("2"));
    let input = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(obj)));
    let result = call_qs("stringify", vec![input, s(";")]);
    let out = as_str(&result);
    assert!(out.contains(';'), "got: {out}");
}

// ── escape / unescape ─────────────────────────────────────────────────────────

#[test]
fn escape_encodes_space_as_percent20() {
    let result = call_qs("escape", vec![s("hello world")]);
    assert!(as_str(&result).contains("%20") || as_str(&result).contains('+'));
}

#[test]
fn escape_leaves_alphanumeric_unchanged() {
    let result = call_qs("escape", vec![s("abc123")]);
    assert_eq!(as_str(&result), "abc123");
}

#[test]
fn escape_encodes_ampersand() {
    let result = call_qs("escape", vec![s("a&b")]);
    let out = as_str(&result);
    assert!(out.contains("%26"), "got: {out}");
}

#[test]
fn unescape_decodes_percent20_as_space() {
    let result = call_qs("unescape", vec![s("hello%20world")]);
    assert_eq!(as_str(&result), "hello world");
}

#[test]
fn unescape_decodes_plus_as_space() {
    let result = call_qs("unescape", vec![s("hello+world")]);
    assert_eq!(as_str(&result), "hello world");
}

#[test]
fn unescape_leaves_already_plain_string_unchanged() {
    let result = call_qs("unescape", vec![s("abc")]);
    assert_eq!(as_str(&result), "abc");
}

#[test]
fn unescape_decodes_percent_encoded_ampersand() {
    let result = call_qs("unescape", vec![s("a%26b")]);
    assert_eq!(as_str(&result), "a&b");
}

// ── Round-trip ────────────────────────────────────────────────────────────────

#[test]
fn escape_then_unescape_roundtrips() {
    let original = "hello world & goodbye!";
    let escaped = call_qs("escape", vec![s(original)]);
    let restored = call_qs("unescape", vec![escaped]);
    assert_eq!(as_str(&restored), original);
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_node_querystring_surface_is_registered() {
    let expected = ["parse", "stringify", "escape", "unescape"];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing node:querystring imports: {missing:?}"
    );
}
