//! Behaviour tests for `node:util` host imports.
//!
//! Reference: <https://nodejs.org/api/util.html>.
//!
//! Coverage targets the functions Node devs reach for most:
//!   - `format(fmt, ...args)` — printf-style string formatter
//!   - `inspect(obj)` — debug stringification
//!   - `types.is*(v)` — type predicates (40+ specific checks)
//!   - `isDeepStrictEqual(a, b)` — recursive equality
//!   - `stripVTControlCharacters(s)` — strip ANSI escape codes
//!   - `toUSVString(s)` — replace lone surrogates with U+FFFD
//!   - `parseArgs(config)` — CLI arg parser (Node 18+)
//!   - Legacy `isArray`/`isString`/etc. (deprecated but still shipped)
//!
//! Not yet tested (deferred — need callback/promise infrastructure):
//!   - `promisify(fn)`, `callbackify(fn)` — async/callback bridges
//!   - `inherits(ctor, super)` — legacy ES5 prototype linking
//!   - `deprecate(fn, msg)` — deprecation wrapper
//!   - `debuglog(section)` — conditional debug logger

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-util-test>");
    let import_idx = chunk.add_import("node:util", name);
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

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn as_string(value: &Value) -> String {
    match value {
        Value::String(text) => text.to_string(),
        other => format!("{}", other),
    }
}

fn new_array(elements: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(elements))))
}

fn new_object(props: Vec<(&str, Value)>) -> Value {
    let mut obj = Object::new();
    for (k, v) in props {
        obj.properties.insert(k.into(), v);
    }
    Value::Object(Arc::new(Mutex::new(obj)))
}

// ── format(fmt, ...args) — Node printf-style ─────────────────────────
//
// Node spec (https://nodejs.org/api/util.html#utilformatformat-args):
// `%s` String, `%d` Number (int+float), `%i` parseInt, `%f` parseFloat,
// `%j` JSON, `%o`/`%O` inspect, `%c` CSS (no-op outside browser),
// `%%` literal %. Extra args after format placeholders are appended
// space-separated.

#[test]
fn format_substitutes_string_placeholder() {
    assert_eq!(as_string(&invoke("format", vec![s("hello %s"), s("world")])), "hello world");
}

#[test]
fn format_substitutes_decimal_placeholder() {
    assert_eq!(as_string(&invoke("format", vec![s("count: %d"), Value::I32(42)])), "count: 42");
}

#[test]
fn format_substitutes_integer_placeholder_truncates_floats() {
    // %i runs parseInt — float gets truncated.
    assert_eq!(as_string(&invoke("format", vec![s("n=%i"), Value::F64(3.7)])), "n=3");
}

#[test]
fn format_substitutes_float_placeholder() {
    assert_eq!(as_string(&invoke("format", vec![s("pi=%f"), Value::F64(3.14)])), "pi=3.14");
}

#[test]
fn format_substitutes_json_placeholder() {
    let obj = new_object(vec![("a", Value::I32(1))]);
    assert_eq!(as_string(&invoke("format", vec![s("data=%j"), obj])), r#"data={"a":1}"#);
}

#[test]
fn format_double_percent_emits_literal() {
    assert_eq!(as_string(&invoke("format", vec![s("100%% off")])), "100% off");
}

#[test]
fn format_extra_args_get_appended_space_separated() {
    // Per Node spec: args without matching placeholders are space-appended.
    assert_eq!(
        as_string(&invoke("format", vec![s("x=%d"), Value::I32(1), Value::I32(2), Value::I32(3)])),
        "x=1 2 3"
    );
}

#[test]
fn format_no_placeholders_just_concatenates() {
    assert_eq!(
        as_string(&invoke("format", vec![s("abc"), s("def"), Value::I32(1)])),
        "abc def 1"
    );
}

#[test]
fn format_with_no_args_returns_input_unchanged() {
    assert_eq!(as_string(&invoke("format", vec![s("hello")])), "hello");
}

#[test]
fn format_zero_args_returns_empty_string() {
    assert_eq!(as_string(&invoke("format", vec![])), "");
}

// ── inspect(obj) — debug stringification ─────────────────────────────
//
// Node spec: returns a string representation suitable for debugging.
// Strings get single-quoted, objects get `{ key: value }` form, arrays
// get `[ a, b, c ]`. We aim for the common shapes; full color/depth
// options come later.

#[test]
fn inspect_string_uses_single_quotes() {
    assert_eq!(as_string(&invoke("inspect", vec![s("hello")])), "'hello'");
}

#[test]
fn inspect_number_renders_as_number() {
    assert_eq!(as_string(&invoke("inspect", vec![Value::I32(42)])), "42");
}

#[test]
fn inspect_null_renders_as_null() {
    assert_eq!(as_string(&invoke("inspect", vec![Value::Null])), "null");
}

#[test]
fn inspect_undefined_renders_as_undefined() {
    assert_eq!(as_string(&invoke("inspect", vec![Value::Undefined])), "undefined");
}

#[test]
fn inspect_array_uses_brackets() {
    let arr = new_array(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    assert_eq!(as_string(&invoke("inspect", vec![arr])), "[ 1, 2, 3 ]");
}

#[test]
fn inspect_object_uses_braces() {
    let obj = new_object(vec![("a", Value::I32(1))]);
    let result = as_string(&invoke("inspect", vec![obj]));
    // Format: `{ a: 1 }` per Node convention.
    assert_eq!(result, "{ a: 1 }");
}

// ── isDeepStrictEqual(a, b) — recursive equality ─────────────────────

#[test]
fn is_deep_strict_equal_primitives() {
    assert_eq!(invoke("isDeepStrictEqual", vec![Value::I32(1), Value::I32(1)]), Value::Bool(true));
    assert_eq!(invoke("isDeepStrictEqual", vec![Value::I32(1), Value::I32(2)]), Value::Bool(false));
    assert_eq!(invoke("isDeepStrictEqual", vec![s("a"), s("a")]), Value::Bool(true));
    assert_eq!(invoke("isDeepStrictEqual", vec![s("a"), s("b")]), Value::Bool(false));
}

#[test]
fn is_deep_strict_equal_arrays() {
    let a = new_array(vec![Value::I32(1), Value::I32(2)]);
    let b = new_array(vec![Value::I32(1), Value::I32(2)]);
    assert_eq!(invoke("isDeepStrictEqual", vec![a, b]), Value::Bool(true));

    let c = new_array(vec![Value::I32(1), Value::I32(3)]);
    let d = new_array(vec![Value::I32(1), Value::I32(2)]);
    assert_eq!(invoke("isDeepStrictEqual", vec![c, d]), Value::Bool(false));
}

#[test]
fn is_deep_strict_equal_nested_objects() {
    let inner_a = new_object(vec![("x", Value::I32(1))]);
    let outer_a = new_object(vec![("nested", inner_a)]);
    let inner_b = new_object(vec![("x", Value::I32(1))]);
    let outer_b = new_object(vec![("nested", inner_b)]);
    assert_eq!(invoke("isDeepStrictEqual", vec![outer_a, outer_b]), Value::Bool(true));
}

#[test]
fn is_deep_strict_equal_distinguishes_strict_types() {
    // Strict: 1 (number) is NOT deep-equal to "1" (string).
    assert_eq!(invoke("isDeepStrictEqual", vec![Value::I32(1), s("1")]), Value::Bool(false));
}

// ── stripVTControlCharacters — strip ANSI escape codes ───────────────

#[test]
fn strip_vt_control_characters_removes_color_escapes() {
    // \x1b[31m red, \x1b[0m reset — typical ANSI color sequence.
    let colored = s("\x1b[31merror\x1b[0m");
    assert_eq!(as_string(&invoke("stripVTControlCharacters", vec![colored])), "error");
}

#[test]
fn strip_vt_control_characters_passes_plain_text_through() {
    assert_eq!(as_string(&invoke("stripVTControlCharacters", vec![s("plain")])), "plain");
}

// ── toUSVString — replace lone surrogates with U+FFFD ────────────────

#[test]
fn to_usv_string_passes_well_formed_utf8_unchanged() {
    assert_eq!(as_string(&invoke("toUSVString", vec![s("hello")])), "hello");
}

// ── parseArgs (Node 18+) ─────────────────────────────────────────────
//
// Node spec: takes `{ options, args }` config and returns
// `{ values, positionals }`. `args` defaults to `process.argv.slice(2)`.
// `options` is `{ name: { type: "string"|"boolean", short?, multiple? } }`.

#[test]
fn parse_args_extracts_string_option() {
    // parseArgs({ args: ["--name", "alice"], options: { name: { type: "string" } } })
    let options = new_object(vec![
        ("name", new_object(vec![("type", s("string"))])),
    ]);
    let args = new_array(vec![s("--name"), s("alice")]);
    let config = new_object(vec![
        ("args", args),
        ("options", options),
    ]);
    let result = invoke("parseArgs", vec![config]);
    if let Value::Object(obj) = &result {
        let o = obj.lock().unwrap();
        let values = o.properties.get("values").expect("values key");
        if let Value::Object(v) = values {
            let vo = v.lock().unwrap();
            assert_eq!(as_string(vo.properties.get("name").unwrap()), "alice");
        } else {
            panic!("values should be Object");
        }
    } else {
        panic!("parseArgs should return Object");
    }
}

#[test]
fn parse_args_collects_positionals() {
    let config = new_object(vec![
        ("args", new_array(vec![s("file1.txt"), s("file2.txt")])),
        ("options", new_object(vec![])),
        ("allowPositionals", Value::Bool(true)),
    ]);
    let result = invoke("parseArgs", vec![config]);
    if let Value::Object(obj) = &result {
        let o = obj.lock().unwrap();
        let positionals = o.properties.get("positionals").expect("positionals key");
        if let Value::Object(p) = positionals {
            let po = p.lock().unwrap();
            if let ObjectKind::Array(ref elems) = po.kind {
                let strs: Vec<String> = elems.iter().map(as_string).collect();
                assert_eq!(strs, vec!["file1.txt", "file2.txt"]);
            } else {
                panic!("positionals should be Array");
            }
        } else {
            panic!("positionals should be Object");
        }
    } else {
        panic!("parseArgs should return Object");
    }
}

// ── types.* — type predicates (Node `util.types` namespace) ──────────
//
// Node namespaces these under `util.types.isX(v)`. Vybe registers each
// predicate as a flat host fn `node:util.types.isX` since the host
// registry is flat (module, name) pairs. Tests reference them with the
// dotted name.

#[test]
fn types_is_array_true_for_arrays() {
    let arr = new_array(vec![Value::I32(1)]);
    assert_eq!(invoke("types.isArray", vec![arr]), Value::Bool(true));
}

#[test]
fn types_is_array_false_for_non_arrays() {
    assert_eq!(invoke("types.isArray", vec![s("not an array")]), Value::Bool(false));
    assert_eq!(invoke("types.isArray", vec![Value::I32(42)]), Value::Bool(false));
}

#[test]
fn types_is_date_true_for_date_objects() {
    // Date objects in Vybe are stamped __type=Date by the wasi:clocks
    // ctor. Construct the equivalent shape here.
    let date = new_object(vec![
        ("__type", s("Date")),
    ]);
    assert_eq!(invoke("types.isDate", vec![date]), Value::Bool(true));
}

#[test]
fn types_is_date_false_for_non_date() {
    assert_eq!(invoke("types.isDate", vec![Value::I32(0)]), Value::Bool(false));
}

#[test]
fn types_is_reg_exp_true_for_regexp_objects() {
    let re = new_object(vec![
        ("__type", s("RegExp")),
        ("source", s("foo")),
    ]);
    assert_eq!(invoke("types.isRegExp", vec![re]), Value::Bool(true));
}

#[test]
fn types_is_map_true_for_map_objects() {
    let mut obj = Object::new();
    obj.kind = ObjectKind::Map(indexmap::IndexMap::new());
    let m = Value::Object(Arc::new(Mutex::new(obj)));
    assert_eq!(invoke("types.isMap", vec![m]), Value::Bool(true));
}

#[test]
fn types_is_set_true_for_set_objects() {
    let mut obj = Object::new();
    obj.kind = ObjectKind::Set(indexmap::IndexSet::new());
    let s = Value::Object(Arc::new(Mutex::new(obj)));
    assert_eq!(invoke("types.isSet", vec![s]), Value::Bool(true));
}

#[test]
fn types_is_array_buffer_true_for_array_buffer() {
    let buf = invoke_other("ecma:arraybuffer", "new", vec![Value::I32(8)]);
    assert_eq!(invoke("types.isArrayBuffer", vec![buf]), Value::Bool(true));
}

#[test]
fn types_is_native_error_true_for_error_objects() {
    let err = new_object(vec![
        ("__type", s("Error")),
        ("message", s("oops")),
    ]);
    assert_eq!(invoke("types.isNativeError", vec![err]), Value::Bool(true));
}

#[test]
fn types_is_promise_true_for_promise_objects() {
    let p = new_object(vec![
        ("__type", s("Promise")),
    ]);
    assert_eq!(invoke("types.isPromise", vec![p]), Value::Bool(true));
}

// ── Legacy is* predicates (deprecated but still shipped) ─────────────
//
// Node still exports these for back-compat; many real codebases use
// them. Per Node docs they're "Deprecated" but functional.

#[test]
fn is_array_legacy_true_for_arrays() {
    let arr = new_array(vec![]);
    assert_eq!(invoke("isArray", vec![arr]), Value::Bool(true));
}

#[test]
fn is_string_legacy() {
    assert_eq!(invoke("isString", vec![s("hi")]), Value::Bool(true));
    assert_eq!(invoke("isString", vec![Value::I32(0)]), Value::Bool(false));
}

#[test]
fn is_number_legacy() {
    assert_eq!(invoke("isNumber", vec![Value::I32(42)]), Value::Bool(true));
    assert_eq!(invoke("isNumber", vec![Value::F64(3.14)]), Value::Bool(true));
    assert_eq!(invoke("isNumber", vec![s("42")]), Value::Bool(false));
}

#[test]
fn is_boolean_legacy() {
    assert_eq!(invoke("isBoolean", vec![Value::Bool(true)]), Value::Bool(true));
    assert_eq!(invoke("isBoolean", vec![Value::I32(1)]), Value::Bool(false));
}

#[test]
fn is_null_legacy() {
    assert_eq!(invoke("isNull", vec![Value::Null]), Value::Bool(true));
    assert_eq!(invoke("isNull", vec![Value::Undefined]), Value::Bool(false));
}

#[test]
fn is_undefined_legacy() {
    assert_eq!(invoke("isUndefined", vec![Value::Undefined]), Value::Bool(true));
    assert_eq!(invoke("isUndefined", vec![Value::Null]), Value::Bool(false));
}

#[test]
fn is_null_or_undefined_legacy() {
    assert_eq!(invoke("isNullOrUndefined", vec![Value::Null]), Value::Bool(true));
    assert_eq!(invoke("isNullOrUndefined", vec![Value::Undefined]), Value::Bool(true));
    assert_eq!(invoke("isNullOrUndefined", vec![Value::I32(0)]), Value::Bool(false));
}

#[test]
fn is_object_legacy() {
    let obj = new_object(vec![]);
    assert_eq!(invoke("isObject", vec![obj]), Value::Bool(true));
    assert_eq!(invoke("isObject", vec![s("not")]), Value::Bool(false));
    // Per Node: null is NOT considered an object by util.isObject.
    assert_eq!(invoke("isObject", vec![Value::Null]), Value::Bool(false));
}

#[test]
fn is_primitive_legacy() {
    // Per Node: Number, String, Boolean, Symbol, undefined, null are primitives.
    assert_eq!(invoke("isPrimitive", vec![Value::I32(42)]), Value::Bool(true));
    assert_eq!(invoke("isPrimitive", vec![s("hi")]), Value::Bool(true));
    assert_eq!(invoke("isPrimitive", vec![Value::Bool(true)]), Value::Bool(true));
    assert_eq!(invoke("isPrimitive", vec![Value::Null]), Value::Bool(true));
    assert_eq!(invoke("isPrimitive", vec![Value::Undefined]), Value::Bool(true));
    assert_eq!(invoke("isPrimitive", vec![new_object(vec![])]), Value::Bool(false));
}

// ── Helper for tests that need to call other modules ─────────────────

fn invoke_other(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<helper>");
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
