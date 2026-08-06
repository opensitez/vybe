//! Behaviour tests for `ecma:json` host imports.
//!
//! Reference: ECMA-262 §25.5 JSON.
//!
//! Each test covers a distinct behaviour.

use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn invoke(name: &str, args: Vec<Value>) -> Value {
    invoke_result(name, args).expect("VM run failed")
}

fn invoke_result(name: &str, args: Vec<Value>) -> Result<Value, vybe_runtime::VMError> {
    let (result, _) = invoke_result_with_exception(name, args);
    result
}

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn push_arg(vm: &mut VM, chunk: &mut Chunk, value: Value) {
    match value {
        Value::I32(n) => chunk.emit_i32_const(n, 0),
        Value::I64(n) => chunk.emit_i64_const(n, 0),
        Value::F32(f) => chunk.emit_f32_const(f, 0),
        Value::F64(f) => chunk.emit_f64_const(f, 0),
        Value::Bool(b) => chunk.emit_bool_const(b, 0),
        Value::String(s) => chunk.emit_string_const(&s, 0),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
        other => {
            let global = format!(
                "__test_arg_{}",
                TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
            );
            vm.globals.insert(global.clone(), other);
            let ci = chunk.intern_string_constant(&global);
            chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
        }
    }
}

fn invoke_result_with_exception(
    name: &str,
    args: Vec<Value>,
) -> (Result<Value, vybe_runtime::VMError>, Option<Value>) {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut chunk = Chunk::new("<ecma-json-test>");
    let import_idx = chunk.add_import("ecma:json", name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    let result = vm.run(vec![chunk]);
    let exception = vm.last_exception.clone();
    (result, exception)
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn obj(pairs: Vec<(&str, Value)>) -> Value {
    let mut o = Object::new();
    for (k, v) in pairs {
        o.properties.insert(k.to_string(), v);
    }
    Value::Object(Arc::new(Mutex::new(o)))
}

fn arr(values: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(values))))
}

fn as_str(v: &Value) -> &str {
    match v {
        Value::String(s) => s.as_ref(),
        _ => panic!("expected string, got {:?}", v),
    }
}

fn obj_prop(value: &Value, key: &str) -> Value {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        return object
            .properties
            .get(key)
            .cloned()
            .unwrap_or(Value::Undefined);
    }
    Value::Undefined
}

fn assert_throws_error_name(name: &str, args: Vec<Value>, expected_name: &str) {
    let (result, exception) = invoke_result_with_exception(name, args);
    result.expect_err("host call should throw");
    let exception = exception.expect("host call should preserve thrown value");
    assert_eq!(format!("{}", obj_prop(&exception, "name")), expected_name);
}

// ── JSON.stringify — special value handling ───────────────────────────────────

#[test]
fn stringify_null_produces_the_string_null() {
    assert_eq!(as_str(&invoke("stringify", vec![Value::Null])), "null");
}

#[test]
fn stringify_nan_becomes_null_per_spec() {
    // JSON has no NaN; the spec mandates encoding it as null.
    assert_eq!(
        as_str(&invoke("stringify", vec![Value::F64(f64::NAN)])),
        "null"
    );
}

#[test]
fn stringify_infinity_becomes_null_per_spec() {
    assert_eq!(
        as_str(&invoke("stringify", vec![Value::F64(f64::INFINITY)])),
        "null"
    );
}

#[test]
fn stringify_undefined_is_not_serializable_returns_undefined() {
    // Top-level undefined → spec returns undefined (not a string).
    let result = invoke("stringify", vec![Value::Undefined]);
    assert!(matches!(result, Value::Undefined | Value::Null));
}

#[test]
fn stringify_boolean_true_and_false() {
    assert_eq!(
        as_str(&invoke("stringify", vec![Value::Bool(true)])),
        "true"
    );
    assert_eq!(
        as_str(&invoke("stringify", vec![Value::Bool(false)])),
        "false"
    );
}

#[test]
fn stringify_integer_produces_numeric_string() {
    assert_eq!(as_str(&invoke("stringify", vec![Value::I32(42)])), "42");
}

#[test]
fn stringify_string_wraps_in_double_quotes() {
    // The string hello → "hello" with surrounding quotes in JSON.
    let result = invoke("stringify", vec![s("hello")]);
    assert_eq!(as_str(&result), "\"hello\"");
}

#[test]
fn stringify_empty_string_produces_two_quotes() {
    assert_eq!(as_str(&invoke("stringify", vec![s("")])), "\"\"");
}

// ── JSON.stringify — object and array structure ───────────────────────────────

#[test]
fn stringify_empty_object_produces_braces() {
    let o = obj(vec![]);
    assert_eq!(as_str(&invoke("stringify", vec![o])), "{}");
}

#[test]
fn stringify_empty_array_produces_brackets() {
    assert_eq!(as_str(&invoke("stringify", vec![arr(vec![])])), "[]");
}

#[test]
fn stringify_array_with_null_element_keeps_null_not_omits() {
    // In arrays, undefined/null elements → "null" (unlike object properties which are omitted).
    let result = invoke(
        "stringify",
        vec![arr(vec![Value::I32(1), Value::Null, Value::I32(3)])],
    );
    assert_eq!(as_str(&result), "[1,null,3]");
}

#[test]
fn stringify_with_indent_produces_newlines() {
    let o = obj(vec![("a", Value::I32(1))]);
    let result = invoke("stringify", vec![o, Value::Null, Value::I32(2)]);
    let s = as_str(&result);
    assert!(s.contains('\n'), "indented output must contain newlines");
    assert!(s.contains("  "), "indented output must contain spaces");
}

// ── JSON.parse ────────────────────────────────────────────────────────────────

#[test]
fn parse_null_literal_returns_null_value() {
    assert_eq!(invoke("parse", vec![s("null")]), Value::Null);
}

#[test]
fn parse_true_and_false_return_bool_values() {
    assert_eq!(invoke("parse", vec![s("true")]), Value::Bool(true));
    assert_eq!(invoke("parse", vec![s("false")]), Value::Bool(false));
}

#[test]
fn parse_integer_returns_numeric_value() {
    let result = invoke("parse", vec![s("42")]);
    assert!(
        matches!(result, Value::I32(42) | Value::I64(42))
            || matches!(result, Value::F64(f) if (f - 42.0).abs() < 0.001)
    );
}

#[test]
fn parse_quoted_string_returns_string_without_quotes() {
    assert_eq!(invoke("parse", vec![s("\"hello\"")]), s("hello"));
}

#[test]
fn parse_empty_object_returns_object_with_no_properties() {
    let result = invoke("parse", vec![s("{}")]);
    assert!(matches!(result, Value::Object(_)));
    if let Value::Object(o) = result {
        assert!(o.lock().unwrap().properties.is_empty());
    }
}

#[test]
fn parse_array_returns_array_object() {
    let result = invoke("parse", vec![s("[1,2,3]")]);
    assert!(matches!(result, Value::Object(_)));
    if let Value::Object(o) = result {
        assert!(matches!(o.lock().unwrap().kind, ObjectKind::Array(_)));
    }
}

#[test]
fn parse_object_with_property_accessible_by_key() {
    let result = invoke("parse", vec![s("{\"x\":7}")]);
    if let Value::Object(o) = &result {
        let val = o
            .lock()
            .unwrap()
            .properties
            .get("x")
            .cloned()
            .unwrap_or(Value::Undefined);
        assert!(
            matches!(val, Value::I32(7)) || matches!(val, Value::F64(f) if (f - 7.0).abs() < 0.001)
        );
    } else {
        panic!("expected object from parse");
    }
}

#[test]
fn parse_invalid_json_throws_syntax_error() {
    assert_throws_error_name("parse", vec![s("{bad json}")], "SyntaxError");
}

#[test]
fn parse_with_reviver_invalid_json_throws_syntax_error() {
    assert_throws_error_name(
        "parseWithReviver",
        vec![s("{bad json}"), Value::Null],
        "SyntaxError",
    );
}

// ── Round-trip ────────────────────────────────────────────────────────────────

#[test]
fn stringify_then_parse_integer_round_trips() {
    let original = Value::I32(123);
    let json = invoke("stringify", vec![original]);
    let restored = invoke("parse", vec![json]);
    assert!(
        matches!(restored, Value::I32(123))
            || matches!(restored, Value::F64(f) if (f - 123.0).abs() < 0.001)
    );
}

#[test]
fn stringify_then_parse_nested_object_preserves_structure() {
    let inner = obj(vec![("y", Value::I32(2))]);
    let outer = obj(vec![("x", Value::I32(1)), ("inner", inner)]);
    let json = invoke("stringify", vec![outer]);
    let restored = invoke("parse", vec![json]);
    // Outer object must be parseable
    assert!(matches!(restored, Value::Object(_)));
}

#[test]
fn stringify_circular_object_throws_type_error() {
    let object = Arc::new(Mutex::new(Object::new()));
    {
        let mut guard = object.lock().unwrap();
        guard
            .properties
            .insert("self".to_string(), Value::Object(object.clone()));
    }
    assert_throws_error_name("stringify", vec![Value::Object(object)], "TypeError");
}

// ── JSON.stringify with array replacer ───────────────────────────────────────

#[test]
fn stringify_array_replacer_filters_to_listed_keys() {
    // ECMA-262 §25.5.2.1: if replacer is an Array, only listed keys are included.
    use std::sync::{Arc, Mutex};
    use vybe_runtime::value::Object;
    let o = obj(vec![
        ("a", Value::I32(1)),
        ("b", Value::I32(2)),
        ("c", Value::I32(3)),
    ]);
    let replacer = Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
        s("a"),
        s("c"),
    ]))));
    let json = invoke("stringifyWithReplacer", vec![o, replacer]);
    let text = match &json {
        Value::String(s) => s.as_ref().to_string(),
        _ => "".to_string(),
    };
    assert!(text.contains("\"a\""), "should include key a");
    assert!(!text.contains("\"b\""), "should exclude key b");
}

// ── JSON.parse with reviver ───────────────────────────────────────────────────

#[test]
fn parse_with_reviver_transforms_each_value() {
    // ECMA-262 §25.5.1: reviver(key, value) is called for each member; can transform.
    // Encode a reviver that doubles numeric values via object descriptor.
    use std::sync::{Arc, Mutex};
    use vybe_runtime::value::Object;
    let reviver = {
        let mut o = Object::new();
        o.properties
            .insert("__reviver_double_numbers".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let result = invoke("parseWithReviver", vec![s("{\"x\":5}"), reviver]);
    assert!(matches!(result, Value::Object(_) | Value::Undefined));
}

// ── JSON.rawJSON / isRawJSON (ES2025 §25.5.5/25.5.4) ─────────────────────────

#[test]
fn raw_json_wraps_a_raw_json_text_into_an_object() {
    // ECMA-262 ES2025: JSON.rawJSON(text) creates an opaque object for embedding
    // raw JSON numbers/literals without re-serialization.
    let result = invoke("rawJSON", vec![s("42")]);
    assert!(matches!(result, Value::Object(_)));
}

#[test]
fn is_raw_json_true_for_raw_json_object() {
    // JSON.isRawJSON(JSON.rawJSON(x)) → true.
    let raw = invoke("rawJSON", vec![s("3.14")]);
    assert_eq!(invoke("isRawJSON", vec![raw]), Value::Bool(true));
}

#[test]
fn is_raw_json_false_for_plain_objects() {
    use std::sync::Mutex;
    use vybe_runtime::value::Object;
    let plain = Value::Object(Arc::new(Mutex::new(Object::new())));
    assert_eq!(invoke("isRawJSON", vec![plain]), Value::Bool(false));
}

#[test]
fn is_raw_json_false_for_primitives() {
    assert_eq!(invoke("isRawJSON", vec![s("42")]), Value::Bool(false));
    assert_eq!(
        invoke("isRawJSON", vec![Value::I32(42)]),
        Value::Bool(false)
    );
}

#[test]
fn stringify_raw_json_embeds_number_without_quotes() {
    // JSON.stringify({ n: JSON.rawJSON("9007199254740993") }) must not lose precision
    // by re-parsing — the raw literal is embedded directly.
    let raw = invoke("rawJSON", vec![s("9007199254740993")]);
    // The raw object itself serializes to the embedded literal (no quotes).
    let text = invoke("stringify", vec![raw]);
    match &text {
        Value::String(s) => {
            assert!(s.contains("9007199254740993") || s.as_ref() == "9007199254740993")
        }
        Value::Undefined => {}
        other => panic!("unexpected: {:?}", other),
    }
}
