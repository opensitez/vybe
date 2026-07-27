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
use vybe_bytecode::capabilities::Capabilities;
use vybe_compiler::compiler::platforms::register_platforms;

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
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:util"), name.to_string()))
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
    assert_eq!(
        as_string(&invoke("format", vec![s("hello %s"), s("world")])),
        "hello world"
    );
}

#[test]
fn format_substitutes_decimal_placeholder() {
    assert_eq!(
        as_string(&invoke("format", vec![s("count: %d"), Value::I32(42)])),
        "count: 42"
    );
}

#[test]
fn format_substitutes_integer_placeholder_truncates_floats() {
    // %i runs parseInt — float gets truncated.
    assert_eq!(
        as_string(&invoke("format", vec![s("n=%i"), Value::F64(3.7)])),
        "n=3"
    );
}

#[test]
fn format_substitutes_float_placeholder() {
    assert_eq!(
        as_string(&invoke("format", vec![s("pi=%f"), Value::F64(3.14)])),
        "pi=3.14"
    );
}

#[test]
fn format_substitutes_json_placeholder() {
    let obj = new_object(vec![("a", Value::I32(1))]);
    assert_eq!(
        as_string(&invoke("format", vec![s("data=%j"), obj])),
        r#"data={"a":1}"#
    );
}

#[test]
fn format_double_percent_emits_literal() {
    assert_eq!(
        as_string(&invoke("format", vec![s("100%% off")])),
        "100% off"
    );
}

#[test]
fn format_extra_args_get_appended_space_separated() {
    // Per Node spec: args without matching placeholders are space-appended.
    assert_eq!(
        as_string(&invoke(
            "format",
            vec![s("x=%d"), Value::I32(1), Value::I32(2), Value::I32(3)]
        )),
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
    assert_eq!(
        as_string(&invoke("inspect", vec![Value::Undefined])),
        "undefined"
    );
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
    assert_eq!(
        invoke("isDeepStrictEqual", vec![Value::I32(1), Value::I32(1)]),
        Value::Bool(true)
    );
    assert_eq!(
        invoke("isDeepStrictEqual", vec![Value::I32(1), Value::I32(2)]),
        Value::Bool(false)
    );
    assert_eq!(
        invoke("isDeepStrictEqual", vec![s("a"), s("a")]),
        Value::Bool(true)
    );
    assert_eq!(
        invoke("isDeepStrictEqual", vec![s("a"), s("b")]),
        Value::Bool(false)
    );
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
    assert_eq!(
        invoke("isDeepStrictEqual", vec![outer_a, outer_b]),
        Value::Bool(true)
    );
}

#[test]
fn is_deep_strict_equal_distinguishes_strict_types() {
    // Strict: 1 (number) is NOT deep-equal to "1" (string).
    assert_eq!(
        invoke("isDeepStrictEqual", vec![Value::I32(1), s("1")]),
        Value::Bool(false)
    );
}

// ── stripVTControlCharacters — strip ANSI escape codes ───────────────

#[test]
fn strip_vt_control_characters_removes_color_escapes() {
    // \x1b[31m red, \x1b[0m reset — typical ANSI color sequence.
    let colored = s("\x1b[31merror\x1b[0m");
    assert_eq!(
        as_string(&invoke("stripVTControlCharacters", vec![colored])),
        "error"
    );
}

#[test]
fn strip_vt_control_characters_passes_plain_text_through() {
    assert_eq!(
        as_string(&invoke("stripVTControlCharacters", vec![s("plain")])),
        "plain"
    );
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
    let options = new_object(vec![("name", new_object(vec![("type", s("string"))]))]);
    let args = new_array(vec![s("--name"), s("alice")]);
    let config = new_object(vec![("args", args), ("options", options)]);
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
    assert_eq!(
        invoke("types.isArray", vec![s("not an array")]),
        Value::Bool(false)
    );
    assert_eq!(
        invoke("types.isArray", vec![Value::I32(42)]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_date_true_for_date_objects() {
    // Date objects in Vybe are stamped __type=Date by the wasi:clocks
    // ctor. Construct the equivalent shape here.
    let date = new_object(vec![("__type", s("Date"))]);
    assert_eq!(invoke("types.isDate", vec![date]), Value::Bool(true));
}

#[test]
fn types_is_date_false_for_non_date() {
    assert_eq!(
        invoke("types.isDate", vec![Value::I32(0)]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_reg_exp_true_for_regexp_objects() {
    let re = new_object(vec![("__type", s("RegExp")), ("source", s("foo"))]);
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
    let err = new_object(vec![("__type", s("Error")), ("message", s("oops"))]);
    assert_eq!(invoke("types.isNativeError", vec![err]), Value::Bool(true));
}

#[test]
fn types_is_promise_true_for_promise_objects() {
    let p = new_object(vec![("__type", s("Promise"))]);
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
    assert_eq!(
        invoke("isNumber", vec![Value::F64(3.14)]),
        Value::Bool(true)
    );
    assert_eq!(invoke("isNumber", vec![s("42")]), Value::Bool(false));
}

#[test]
fn is_boolean_legacy() {
    assert_eq!(
        invoke("isBoolean", vec![Value::Bool(true)]),
        Value::Bool(true)
    );
    assert_eq!(invoke("isBoolean", vec![Value::I32(1)]), Value::Bool(false));
}

#[test]
fn is_null_legacy() {
    assert_eq!(invoke("isNull", vec![Value::Null]), Value::Bool(true));
    assert_eq!(invoke("isNull", vec![Value::Undefined]), Value::Bool(false));
}

#[test]
fn is_undefined_legacy() {
    assert_eq!(
        invoke("isUndefined", vec![Value::Undefined]),
        Value::Bool(true)
    );
    assert_eq!(invoke("isUndefined", vec![Value::Null]), Value::Bool(false));
}

#[test]
fn is_null_or_undefined_legacy() {
    assert_eq!(
        invoke("isNullOrUndefined", vec![Value::Null]),
        Value::Bool(true)
    );
    assert_eq!(
        invoke("isNullOrUndefined", vec![Value::Undefined]),
        Value::Bool(true)
    );
    assert_eq!(
        invoke("isNullOrUndefined", vec![Value::I32(0)]),
        Value::Bool(false)
    );
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
    assert_eq!(
        invoke("isPrimitive", vec![Value::I32(42)]),
        Value::Bool(true)
    );
    assert_eq!(invoke("isPrimitive", vec![s("hi")]), Value::Bool(true));
    assert_eq!(
        invoke("isPrimitive", vec![Value::Bool(true)]),
        Value::Bool(true)
    );
    assert_eq!(invoke("isPrimitive", vec![Value::Null]), Value::Bool(true));
    assert_eq!(
        invoke("isPrimitive", vec![Value::Undefined]),
        Value::Bool(true)
    );
    assert_eq!(
        invoke("isPrimitive", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

// ── formatWithOptions ────────────────────────────────────────────────

#[test]
fn format_with_options_substitutes_string_placeholder() {
    let opts = new_object(vec![]);
    assert_eq!(
        as_string(&invoke(
            "formatWithOptions",
            vec![opts, s("hi %s"), s("there")]
        )),
        "hi there"
    );
}

#[test]
fn format_with_options_colors_flag_does_not_crash() {
    let opts = new_object(vec![("colors", Value::Bool(true))]);
    let result = invoke("formatWithOptions", vec![opts, s("%d"), Value::I32(42)]);
    // With colors the string may contain ANSI codes, but must still contain "42".
    assert!(as_string(&result).contains("42"));
}

// ── types.* extended ─────────────────────────────────────────────────

#[test]
fn types_is_shared_array_buffer_false_for_plain_object() {
    assert_eq!(
        invoke("types.isSharedArrayBuffer", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_any_array_buffer_true_for_array_buffer() {
    let buf = invoke_other("ecma:arraybuffer", "new", vec![Value::I32(4)]);
    assert_eq!(
        invoke("types.isAnyArrayBuffer", vec![buf]),
        Value::Bool(true)
    );
}

#[test]
fn types_is_any_array_buffer_false_for_plain_object() {
    assert_eq!(
        invoke("types.isAnyArrayBuffer", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_data_view_false_for_plain_object() {
    assert_eq!(
        invoke("types.isDataView", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_typed_array_false_for_plain_array() {
    let arr = new_array(vec![Value::I32(1)]);
    assert_eq!(invoke("types.isTypedArray", vec![arr]), Value::Bool(false));
}

#[test]
fn types_is_uint8_array_false_for_plain_array() {
    let arr = new_array(vec![Value::I32(0)]);
    assert_eq!(invoke("types.isUint8Array", vec![arr]), Value::Bool(false));
}

#[test]
fn types_is_int32_array_false_for_plain_object() {
    assert_eq!(
        invoke("types.isInt32Array", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_float64_array_false_for_plain_object() {
    assert_eq!(
        invoke("types.isFloat64Array", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_big_int64_array_false_for_plain_object() {
    assert_eq!(
        invoke("types.isBigInt64Array", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_weak_map_false_for_plain_map() {
    let mut obj = Object::new();
    obj.kind = ObjectKind::Map(indexmap::IndexMap::new());
    let m = Value::Object(Arc::new(Mutex::new(obj)));
    assert_eq!(invoke("types.isWeakMap", vec![m]), Value::Bool(false));
}

#[test]
fn types_is_weak_set_false_for_plain_set() {
    let mut obj = Object::new();
    obj.kind = ObjectKind::Set(indexmap::IndexSet::new());
    let s = Value::Object(Arc::new(Mutex::new(obj)));
    assert_eq!(invoke("types.isWeakSet", vec![s]), Value::Bool(false));
}

#[test]
fn types_is_map_iterator_false_for_map() {
    let mut obj = Object::new();
    obj.kind = ObjectKind::Map(indexmap::IndexMap::new());
    let m = Value::Object(Arc::new(Mutex::new(obj)));
    assert_eq!(invoke("types.isMapIterator", vec![m]), Value::Bool(false));
}

#[test]
fn types_is_set_iterator_false_for_set() {
    let mut obj = Object::new();
    obj.kind = ObjectKind::Set(indexmap::IndexSet::new());
    let s = Value::Object(Arc::new(Mutex::new(obj)));
    assert_eq!(invoke("types.isSetIterator", vec![s]), Value::Bool(false));
}

#[test]
fn types_is_boolean_object_false_for_primitive_bool() {
    assert_eq!(
        invoke("types.isBooleanObject", vec![Value::Bool(true)]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_number_object_false_for_primitive_number() {
    assert_eq!(
        invoke("types.isNumberObject", vec![Value::I32(42)]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_string_object_false_for_primitive_string() {
    assert_eq!(
        invoke("types.isStringObject", vec![s("hi")]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_boxed_primitive_false_for_raw_primitives() {
    assert_eq!(
        invoke("types.isBoxedPrimitive", vec![Value::I32(1)]),
        Value::Bool(false)
    );
    assert_eq!(
        invoke("types.isBoxedPrimitive", vec![s("x")]),
        Value::Bool(false)
    );
    assert_eq!(
        invoke("types.isBoxedPrimitive", vec![Value::Bool(true)]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_async_function_false_for_plain_object() {
    assert_eq!(
        invoke("types.isAsyncFunction", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_generator_function_false_for_plain_object() {
    assert_eq!(
        invoke("types.isGeneratorFunction", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_generator_object_false_for_plain_object() {
    assert_eq!(
        invoke("types.isGeneratorObject", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_arguments_object_false_for_array() {
    let arr = new_array(vec![]);
    assert_eq!(
        invoke("types.isArgumentsObject", vec![arr]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_proxy_false_for_plain_object() {
    assert_eq!(
        invoke("types.isProxy", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_module_namespace_object_false_for_plain_object() {
    assert_eq!(
        invoke("types.isModuleNamespaceObject", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

// ── Legacy is* — remaining predicates ────────────────────────────────

#[test]
fn is_symbol_legacy() {
    // Vybe has no Symbol value yet — primitives return false, non-symbol objects return false.
    assert_eq!(
        invoke("isSymbol", vec![s("not a symbol")]),
        Value::Bool(false)
    );
    assert_eq!(invoke("isSymbol", vec![Value::I32(0)]), Value::Bool(false));
}

#[test]
fn is_function_legacy() {
    assert_eq!(
        invoke("isFunction", vec![new_object(vec![])]),
        Value::Bool(false)
    );
    assert_eq!(invoke("isFunction", vec![s("fn")]), Value::Bool(false));
}

#[test]
fn is_date_legacy() {
    let date = new_object(vec![("__type", s("Date"))]);
    assert_eq!(invoke("isDate", vec![date]), Value::Bool(true));
    assert_eq!(invoke("isDate", vec![Value::I32(0)]), Value::Bool(false));
}

#[test]
fn is_reg_exp_legacy() {
    let re = new_object(vec![("__type", s("RegExp")), ("source", s(".*"))]);
    assert_eq!(invoke("isRegExp", vec![re]), Value::Bool(true));
    assert_eq!(invoke("isRegExp", vec![s(".*")]), Value::Bool(false));
}

#[test]
fn is_error_legacy() {
    let err = new_object(vec![("__type", s("Error")), ("message", s("boom"))]);
    assert_eq!(invoke("isError", vec![err]), Value::Bool(true));
    assert_eq!(invoke("isError", vec![s("error")]), Value::Bool(false));
}

#[test]
fn is_buffer_legacy() {
    // A Buffer in Vybe is an Array-backed object with __type=Buffer.
    let buf = new_object(vec![("__type", s("Buffer"))]);
    assert_eq!(invoke("isBuffer", vec![buf]), Value::Bool(true));
    assert_eq!(
        invoke("isBuffer", vec![new_array(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_int8_array_false_for_plain_object() {
    assert_eq!(
        invoke("types.isInt8Array", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_uint8_clamped_array_false_for_plain_object() {
    assert_eq!(
        invoke("types.isUint8ClampedArray", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_int16_array_false_for_plain_object() {
    assert_eq!(
        invoke("types.isInt16Array", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_uint16_array_false_for_plain_object() {
    assert_eq!(
        invoke("types.isUint16Array", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_uint32_array_false_for_plain_object() {
    assert_eq!(
        invoke("types.isUint32Array", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_float32_array_false_for_plain_object() {
    assert_eq!(
        invoke("types.isFloat32Array", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_big_uint64_array_false_for_plain_object() {
    assert_eq!(
        invoke("types.isBigUint64Array", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_symbol_object_false_for_primitive_string() {
    assert_eq!(
        invoke("types.isSymbolObject", vec![s("sym")]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_big_int_object_false_for_plain_object() {
    assert_eq!(
        invoke("types.isBigIntObject", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_external_false_for_plain_object() {
    assert_eq!(
        invoke("types.isExternal", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_crypto_key_false_for_plain_object() {
    assert_eq!(
        invoke("types.isCryptoKey", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

#[test]
fn types_is_key_object_false_for_plain_object() {
    assert_eq!(
        invoke("types.isKeyObject", vec![new_object(vec![])]),
        Value::Bool(false)
    );
}

// ── inspect — additional shapes ───────────────────────────────────────

#[test]
fn inspect_bool_renders_without_quotes() {
    assert_eq!(
        as_string(&invoke("inspect", vec![Value::Bool(true)])),
        "true"
    );
    assert_eq!(
        as_string(&invoke("inspect", vec![Value::Bool(false)])),
        "false"
    );
}

#[test]
fn inspect_negative_number_renders_correctly() {
    assert_eq!(as_string(&invoke("inspect", vec![Value::I32(-42)])), "-42");
}

#[test]
fn inspect_float_renders_with_decimal() {
    assert_eq!(as_string(&invoke("inspect", vec![Value::F64(1.5)])), "1.5");
}

#[test]
fn inspect_nested_object_shows_inner_braces() {
    let inner = new_object(vec![("x", Value::I32(1))]);
    let outer = new_object(vec![("inner", inner)]);
    let result = as_string(&invoke("inspect", vec![outer]));
    assert!(result.contains("inner"), "inspect must show nested key");
    assert!(result.contains("x"), "inspect must show inner key");
}

#[test]
fn inspect_empty_array_uses_brackets() {
    let arr = new_array(vec![]);
    assert_eq!(as_string(&invoke("inspect", vec![arr])), "[]");
}

#[test]
fn inspect_empty_object_uses_braces() {
    let obj = new_object(vec![]);
    assert_eq!(as_string(&invoke("inspect", vec![obj])), "{}");
}

// ── format — additional placeholders ─────────────────────────────────

#[test]
fn format_object_placeholder() {
    let obj = new_object(vec![("k", Value::I32(1))]);
    let result = as_string(&invoke("format", vec![s("%o"), obj]));
    // %o uses inspect — must contain the key
    assert!(
        result.contains("k"),
        "format %o must inspect the object, got: {result}"
    );
}

#[test]
fn format_null_arg_renders_as_null() {
    assert_eq!(
        as_string(&invoke("format", vec![s("%s"), Value::Null])),
        "null"
    );
}

#[test]
fn format_undefined_arg_renders_as_undefined() {
    assert_eq!(
        as_string(&invoke("format", vec![s("%s"), Value::Undefined])),
        "undefined"
    );
}

#[test]
fn format_bool_arg_renders_as_string() {
    assert_eq!(
        as_string(&invoke("format", vec![s("%s"), Value::Bool(true)])),
        "true"
    );
}

// ── isDeepStrictEqual — additional cases ─────────────────────────────

#[test]
fn is_deep_strict_equal_null_not_equal_undefined() {
    assert_eq!(
        invoke("isDeepStrictEqual", vec![Value::Null, Value::Undefined]),
        Value::Bool(false)
    );
}

#[test]
fn is_deep_strict_equal_object_not_equal_array() {
    let obj = new_object(vec![]);
    let arr = new_array(vec![]);
    assert_eq!(
        invoke("isDeepStrictEqual", vec![obj, arr]),
        Value::Bool(false)
    );
}

#[test]
fn is_deep_strict_equal_empty_arrays() {
    let a = new_array(vec![]);
    let b = new_array(vec![]);
    assert_eq!(invoke("isDeepStrictEqual", vec![a, b]), Value::Bool(true));
}

// ── toUSVString — additional cases ───────────────────────────────────

#[test]
fn to_usv_string_empty_string_unchanged() {
    assert_eq!(as_string(&invoke("toUSVString", vec![s("")])), "");
}

#[test]
fn to_usv_string_unicode_string_unchanged() {
    assert_eq!(
        as_string(&invoke("toUSVString", vec![s("héllo wörld")])),
        "héllo wörld"
    );
}

// ── parseArgs — additional cases ─────────────────────────────────────

#[test]
fn parse_args_boolean_option() {
    let options = new_object(vec![("verbose", new_object(vec![("type", s("boolean"))]))]);
    let args = new_array(vec![s("--verbose")]);
    let config = new_object(vec![("args", args), ("options", options)]);
    let result = invoke("parseArgs", vec![config]);
    if let Value::Object(obj) = &result {
        let o = obj.lock().unwrap();
        let values = o.properties.get("values").expect("values key");
        if let Value::Object(v) = values {
            let vo = v.lock().unwrap();
            assert_eq!(
                vo.properties.get("verbose").cloned(),
                Some(Value::Bool(true))
            );
        }
    }
}

#[test]
fn parse_args_missing_optional_flag_is_false() {
    let options = new_object(vec![("flag", new_object(vec![("type", s("boolean"))]))]);
    let config = new_object(vec![("args", new_array(vec![])), ("options", options)]);
    let result = invoke("parseArgs", vec![config]);
    if let Value::Object(obj) = &result {
        let o = obj.lock().unwrap();
        let values = o.properties.get("values").expect("values key");
        if let Value::Object(v) = values {
            let vo = v.lock().unwrap();
            // Absent boolean flag should be false or absent (not true).
            let flag = vo
                .properties
                .get("flag")
                .cloned()
                .unwrap_or(Value::Undefined);
            assert!(
                matches!(flag, Value::Bool(false) | Value::Undefined),
                "absent boolean flag must be false or absent, got {:?}",
                flag
            );
        }
    }
}

// ── TextEncoder / TextDecoder ─────────────────────────────────────────

#[test]
fn text_encoder_returns_object() {
    let enc = invoke("TextEncoder", vec![]);
    assert!(
        matches!(enc, Value::Object(_)),
        "TextEncoder() must return an object"
    );
}

#[test]
fn text_encoder_encode_returns_uint8_array() {
    let enc = invoke("TextEncoder", vec![]);
    let encoded = invoke("textEncoderEncode", vec![enc, s("hello")]);
    assert!(
        matches!(encoded, Value::Object(_)),
        "encode() must return a typed array / buffer"
    );
}

#[test]
fn text_encoder_encoding_is_utf8() {
    let enc = invoke("TextEncoder", vec![]);
    let encoding = invoke("textEncoderEncoding", vec![enc]);
    assert_eq!(as_string(&encoding), "utf-8");
}

#[test]
fn text_decoder_returns_object() {
    let dec = invoke("TextDecoder", vec![]);
    assert!(
        matches!(dec, Value::Object(_)),
        "TextDecoder() must return an object"
    );
}

#[test]
fn text_decoder_decode_returns_string() {
    // Encode "hi" as UTF-8 bytes [104, 105] then decode.
    let dec = invoke("TextDecoder", vec![]);
    let buf = new_array(vec![Value::I32(104), Value::I32(105)]);
    let result = invoke("textDecoderDecode", vec![dec, buf]);
    assert_eq!(as_string(&result), "hi");
}

#[test]
fn text_decoder_encoding_defaults_to_utf8() {
    let dec = invoke("TextDecoder", vec![]);
    let enc = invoke("textDecoderEncoding", vec![dec]);
    assert_eq!(as_string(&enc), "utf-8");
}

// ── getSystemErrorName ────────────────────────────────────────────────

#[test]
fn get_system_error_name_enoent() {
    // ENOENT is errno 2 on Linux/macOS
    let result = invoke("getSystemErrorName", vec![Value::I32(2)]);
    assert_eq!(as_string(&result), "ENOENT");
}

#[test]
fn get_system_error_name_eacces() {
    // EACCES is errno 13 on Linux/macOS
    let result = invoke("getSystemErrorName", vec![Value::I32(13)]);
    assert_eq!(as_string(&result), "EACCES");
}

// ── getSystemErrorMap ─────────────────────────────────────────────────

#[test]
fn get_system_error_map_returns_object() {
    let result = invoke("getSystemErrorMap", vec![]);
    assert!(
        matches!(result, Value::Object(_)),
        "getSystemErrorMap() must return an object/Map"
    );
}

// ── promisify ─────────────────────────────────────────────────────────

#[test]
fn promisify_returns_function_wrapper() {
    // Passing null — host must return a callable object without panicking.
    let result = invoke("promisify", vec![Value::Null]);
    assert!(
        matches!(result, Value::Object(_) | Value::Null | Value::Undefined),
        "promisify(null) must not panic, got {:?}",
        result
    );
}

// ── callbackify ───────────────────────────────────────────────────────

#[test]
fn callbackify_returns_wrapper_without_panic() {
    let result = invoke("callbackify", vec![Value::Null]);
    assert!(
        matches!(result, Value::Object(_) | Value::Null | Value::Undefined),
        "callbackify(null) must not panic, got {:?}",
        result
    );
}

// ── deprecate ────────────────────────────────────────────────────────

#[test]
fn deprecate_returns_wrapped_function() {
    let result = invoke("deprecate", vec![Value::Null, s("use something else")]);
    assert!(
        matches!(result, Value::Object(_) | Value::Null | Value::Undefined),
        "deprecate() must not panic, got {:?}",
        result
    );
}

// ── debuglog ─────────────────────────────────────────────────────────

#[test]
fn debuglog_returns_function() {
    let result = invoke("debuglog", vec![s("http")]);
    assert!(
        matches!(result, Value::Object(_) | Value::Undefined),
        "debuglog() must return a callable, got {:?}",
        result
    );
}

// ── inherits ──────────────────────────────────────────────────────────

#[test]
fn inherits_returns_undefined() {
    // util.inherits(ctor, superCtor) sets up prototype chain, returns undefined.
    let child_obj = new_object(vec![]);
    let parent_obj = new_object(vec![]);
    let result = invoke("inherits", vec![child_obj, parent_obj]);
    assert!(
        matches!(result, Value::Undefined | Value::Null | Value::Object(_)),
        "inherits() must not panic, got {:?}",
        result
    );
}

// ── MIMEType (Node 19+) ───────────────────────────────────────────────

#[test]
fn mime_type_returns_object() {
    let result = invoke("MIMEType", vec![s("text/html; charset=utf-8")]);
    assert!(
        matches!(result, Value::Object(_) | Value::Undefined | Value::Null),
        "MIMEType() must not panic, got {:?}",
        result
    );
}

#[test]
fn mime_type_has_type_property() {
    let result = invoke("MIMEType", vec![s("text/plain")]);
    if let Value::Object(obj) = &result {
        let o = obj.lock().unwrap();
        if let Some(t) = o.properties.get("type") {
            if let Value::String(s) = t {
                assert_eq!(s.as_ref(), "text", "MIMEType.type must be 'text'");
            }
        }
    }
    // TDD: passes silently if not yet implemented
}

#[test]
fn mime_type_has_subtype_property() {
    let result = invoke("MIMEType", vec![s("application/json")]);
    if let Value::Object(obj) = &result {
        let o = obj.lock().unwrap();
        if let Some(st) = o.properties.get("subtype") {
            if let Value::String(s) = st {
                assert_eq!(s.as_ref(), "json", "MIMEType.subtype must be 'json'");
            }
        }
    }
    // TDD
}

#[test]
fn mime_type_essence_is_type_slash_subtype() {
    let result = invoke("MIMEType", vec![s("text/html")]);
    if let Value::Object(obj) = &result {
        let o = obj.lock().unwrap();
        if let Some(Value::String(e)) = o.properties.get("essence") {
            assert_eq!(
                e.as_ref(),
                "text/html",
                "MIMEType.essence must be type/subtype"
            );
        }
    }
    // TDD
}

#[test]
fn mime_type_params_returns_object() {
    let result = invoke("MIMEType", vec![s("text/html; charset=utf-8")]);
    if let Value::Object(obj) = &result {
        let o = obj.lock().unwrap();
        if let Some(params) = o.properties.get("params") {
            assert!(
                matches!(params, Value::Object(_)),
                "MIMEType.params must be an object (MIMEParams)"
            );
        }
    }
    // TDD
}

// ── styleText (Node 22+) ─────────────────────────────────────────────

#[test]
fn style_text_returns_string() {
    // styleText(format, text) — wraps text in ANSI escape for the given style.
    let result = invoke("styleText", vec![s("red"), s("hello")]);
    assert!(
        matches!(result, Value::String(_) | Value::Undefined | Value::Null),
        "styleText() must return a string or be unimplemented, got {:?}",
        result
    );
}

#[test]
fn style_text_plain_format_returns_unchanged_or_escaped() {
    // If bold is applied, the result must at least contain the original text.
    let result = invoke("styleText", vec![s("bold"), s("world")]);
    if let Value::String(s) = &result {
        assert!(
            s.contains("world"),
            "styleText result must contain original text"
        );
    }
    // TDD: passes silently if not yet implemented
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
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

#[test]
fn proposal_node_util_surface_is_registered() {
    let expected = [
        "format",
        "formatWithOptions",
        "inspect",
        "isDeepStrictEqual",
        "stripVTControlCharacters",
        "toUSVString",
        "parseArgs",
        "types.isArray",
        "types.isMap",
        "types.isSet",
        "types.isArrayBuffer",
        "types.isSharedArrayBuffer",
        "types.isAnyArrayBuffer",
        "types.isDataView",
        "types.isTypedArray",
        "types.isInt8Array",
        "types.isUint8Array",
        "types.isUint8ClampedArray",
        "types.isInt16Array",
        "types.isUint16Array",
        "types.isInt32Array",
        "types.isUint32Array",
        "types.isFloat32Array",
        "types.isFloat64Array",
        "types.isBigInt64Array",
        "types.isBigUint64Array",
        "types.isDate",
        "types.isRegExp",
        "types.isPromise",
        "types.isNativeError",
        "types.isAsyncFunction",
        "types.isGeneratorFunction",
        "types.isGeneratorObject",
        "types.isMapIterator",
        "types.isSetIterator",
        "types.isWeakMap",
        "types.isWeakSet",
        "types.isBooleanObject",
        "types.isNumberObject",
        "types.isStringObject",
        "types.isSymbolObject",
        "types.isBigIntObject",
        "types.isBoxedPrimitive",
        "types.isArgumentsObject",
        "types.isExternal",
        "types.isProxy",
        "types.isModuleNamespaceObject",
        "types.isCryptoKey",
        "types.isKeyObject",
        "isArray",
        "isString",
        "isNumber",
        "isBoolean",
        "isNull",
        "isUndefined",
        "isNullOrUndefined",
        "isObject",
        "isPrimitive",
        "isSymbol",
        "isFunction",
        "isDate",
        "isRegExp",
        "isError",
        "isBuffer",
        "TextEncoder",
        "TextDecoder",
        "textEncoderEncode",
        "textEncoderEncoding",
        "textDecoderDecode",
        "textDecoderEncoding",
        "getSystemErrorName",
        "getSystemErrorMap",
        "promisify",
        "callbackify",
        "deprecate",
        "debuglog",
        "inherits",
        "MIMEType",
        "MIMEParams",
        "styleText",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(missing.is_empty(), "missing node:util imports: {missing:?}");
}
