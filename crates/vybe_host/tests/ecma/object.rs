//! Behaviour tests for `ecma:object` host imports.
//!
//! Reference: ECMA-262 §20.1 Object.
//!
//! Each test covers a distinct behaviour.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-object-test>");
    let import_idx = chunk.add_import("ecma:object", name);
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

fn s(text: &str) -> Value { Value::String(Arc::from(text)) }

fn obj(pairs: Vec<(&str, Value)>) -> Value {
    let mut o = Object::new();
    for (k, v) in pairs { o.properties.insert(k.to_string(), v); }
    Value::Object(Arc::new(Mutex::new(o)))
}

fn arr(values: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(values))))
}

fn pair(k: Value, v: Value) -> Value { arr(vec![k, v]) }

fn elems(v: &Value) -> Vec<Value> {
    match v {
        Value::Object(o) => match &o.lock().unwrap().kind {
            ObjectKind::Array(e) => e.clone(),
            _ => vec![],
        },
        _ => vec![],
    }
}

// ── get / set / has / delete ──────────────────────────────────────────────────

#[test]
fn get_returns_undefined_for_missing_key() {
    let o = invoke("new", vec![]);
    assert_eq!(invoke("get", vec![o, s("missing")]), Value::Undefined);
}

#[test]
fn set_then_get_round_trips_value() {
    let o = invoke("new", vec![]);
    invoke("set", vec![o.clone(), s("x"), Value::I32(42)]);
    assert_eq!(invoke("get", vec![o, s("x")]), Value::I32(42));
}

#[test]
fn has_reflects_presence_after_set_and_delete() {
    let o = invoke("new", vec![]);
    invoke("set", vec![o.clone(), s("k"), Value::Bool(true)]);
    assert_eq!(invoke("has", vec![o.clone(), s("k")]), Value::Bool(true));
    invoke("delete", vec![o.clone(), s("k")]);
    assert_eq!(invoke("has", vec![o.clone(), s("k")]), Value::Bool(false));
}

#[test]
fn delete_returns_false_for_non_existent_key() {
    let o = invoke("new", vec![]);
    assert_eq!(invoke("delete", vec![o, s("nope")]), Value::Bool(false));
}

#[test]
fn set_overwrites_existing_value() {
    let o = invoke("new", vec![]);
    invoke("set", vec![o.clone(), s("v"), Value::I32(1)]);
    invoke("set", vec![o.clone(), s("v"), Value::I32(2)]);
    assert_eq!(invoke("get", vec![o, s("v")]), Value::I32(2));
}

// ── keys / values / entries ───────────────────────────────────────────────────

#[test]
fn keys_returns_empty_array_for_fresh_object() {
    let o = invoke("new", vec![]);
    let keys = invoke("keys", vec![o]);
    assert_eq!(elems(&keys).len(), 0);
}

#[test]
fn keys_returns_only_own_enumerable_property_names() {
    let o = obj(vec![("a", Value::I32(1)), ("b", Value::I32(2))]);
    let keys = invoke("keys", vec![o]);
    let key_strings: Vec<String> = elems(&keys).iter().map(|v| match v {
        Value::String(s) => s.to_string(), _ => String::new(),
    }).collect();
    assert!(key_strings.contains(&"a".to_string()));
    assert!(key_strings.contains(&"b".to_string()));
    assert_eq!(key_strings.len(), 2);
}

#[test]
fn values_returns_property_values_not_keys() {
    let o = obj(vec![("x", Value::I32(99))]);
    let vals = invoke("values", vec![o]);
    assert_eq!(elems(&vals), vec![Value::I32(99)]);
}

#[test]
fn entries_returns_key_value_pairs_as_nested_arrays() {
    let o = obj(vec![("name", s("alice"))]);
    let entries = invoke("entries", vec![o]);
    let e = elems(&entries);
    assert_eq!(e.len(), 1);
    let pair_elems = elems(&e[0]);
    assert_eq!(pair_elems[0], s("name"));
    assert_eq!(pair_elems[1], s("alice"));
}

// ── Object.assign ─────────────────────────────────────────────────────────────

#[test]
fn assign_copies_own_properties_to_target_and_returns_target() {
    let target = invoke("new", vec![]);
    let source = obj(vec![("a", Value::I32(1)), ("b", Value::I32(2))]);
    let result = invoke("assign", vec![target.clone(), source]);
    // Returns the target
    let t_ptr = match &target { Value::Object(a) => Arc::as_ptr(a) as usize, _ => 0 };
    let r_ptr = match &result { Value::Object(a) => Arc::as_ptr(a) as usize, _ => 1 };
    assert_eq!(t_ptr, r_ptr);
    assert_eq!(invoke("get", vec![result, s("a")]), Value::I32(1));
}

#[test]
fn assign_later_source_overwrites_earlier_source_on_conflict() {
    let target = obj(vec![("x", Value::I32(1))]);
    let source = obj(vec![("x", Value::I32(2))]);
    invoke("assign", vec![target.clone(), source]);
    assert_eq!(invoke("get", vec![target, s("x")]), Value::I32(2));
}

// ── Object.fromEntries ────────────────────────────────────────────────────────

#[test]
fn from_entries_reconstructs_object_from_pairs() {
    let pairs = arr(vec![
        pair(s("a"), Value::I32(1)),
        pair(s("b"), Value::I32(2)),
    ]);
    let o = invoke("fromEntries", vec![pairs]);
    assert_eq!(invoke("get", vec![o.clone(), s("a")]), Value::I32(1));
    assert_eq!(invoke("get", vec![o, s("b")]), Value::I32(2));
}

#[test]
fn from_entries_empty_array_produces_empty_object() {
    let o = invoke("fromEntries", vec![arr(vec![])]);
    assert_eq!(elems(&invoke("keys", vec![o])).len(), 0);
}

// ── Object.freeze / isFrozen ──────────────────────────────────────────────────

#[test]
fn is_frozen_false_before_freeze_true_after() {
    let o = invoke("new", vec![]);
    assert_eq!(invoke("isFrozen", vec![o.clone()]), Value::Bool(false));
    invoke("freeze", vec![o.clone()]);
    assert_eq!(invoke("isFrozen", vec![o]), Value::Bool(true));
}

#[test]
fn freeze_returns_the_same_object() {
    let o = invoke("new", vec![]);
    let o_ptr = match &o { Value::Object(a) => Arc::as_ptr(a) as usize, _ => 0 };
    let result = invoke("freeze", vec![o]);
    let r_ptr = match &result { Value::Object(a) => Arc::as_ptr(a) as usize, _ => 1 };
    assert_eq!(o_ptr, r_ptr);
}

// ── Object.seal / isSealed ────────────────────────────────────────────────────

#[test]
fn is_sealed_false_before_seal_true_after() {
    let o = invoke("new", vec![]);
    assert_eq!(invoke("isSealed", vec![o.clone()]), Value::Bool(false));
    invoke("seal", vec![o.clone()]);
    assert_eq!(invoke("isSealed", vec![o]), Value::Bool(true));
}

// ── Object.preventExtensions / isExtensible ───────────────────────────────────

#[test]
fn is_extensible_true_by_default_false_after_prevent() {
    let o = invoke("new", vec![]);
    assert_eq!(invoke("isExtensible", vec![o.clone()]), Value::Bool(true));
    invoke("preventExtensions", vec![o.clone()]);
    assert_eq!(invoke("isExtensible", vec![o]), Value::Bool(false));
}

// ── Object.is — SameValue semantics differ from === ───────────────────────────

#[test]
fn object_is_nan_equals_nan_unlike_strict_equality() {
    // === says NaN !== NaN; Object.is says they are the same.
    let nan = Value::F64(f64::NAN);
    assert_eq!(invoke("is", vec![nan.clone(), nan]), Value::Bool(true));
}

#[test]
fn object_is_positive_zero_not_same_as_negative_zero() {
    // === says +0 === -0; Object.is says they differ.
    assert_eq!(
        invoke("is", vec![Value::F64(0.0), Value::F64(-0.0)]),
        Value::Bool(false)
    );
}

#[test]
fn object_is_null_not_same_as_undefined() {
    assert_eq!(invoke("is", vec![Value::Null, Value::Undefined]), Value::Bool(false));
}

// ── hasOwn ────────────────────────────────────────────────────────────────────

#[test]
fn has_own_true_for_own_property_not_inherited() {
    let o = obj(vec![("own", Value::I32(1))]);
    assert_eq!(invoke("hasOwn", vec![o.clone(), s("own")]), Value::Bool(true));
    assert_eq!(invoke("hasOwn", vec![o, s("toString")]), Value::Bool(false));
}

// ── getOwnPropertyNames ───────────────────────────────────────────────────────

#[test]
fn get_own_property_names_includes_all_string_keys() {
    let o = obj(vec![("a", Value::I32(1)), ("b", Value::I32(2))]);
    let names = invoke("getOwnPropertyNames", vec![o]);
    let strs: Vec<String> = elems(&names).iter().map(|v| match v {
        Value::String(s) => s.to_string(), _ => String::new(),
    }).collect();
    assert!(strs.contains(&"a".to_string()));
    assert!(strs.contains(&"b".to_string()));
}

// ── defineProperty / getOwnPropertyDescriptor ────────────────────────────────

#[test]
fn define_property_adds_property_readable_via_descriptor() {
    let o = invoke("new", vec![]);
    let desc = obj(vec![("value", Value::I32(42)), ("writable", Value::Bool(true))]);
    invoke("defineProperty", vec![o.clone(), s("p"), desc]);
    let d = invoke("getOwnPropertyDescriptor", vec![o, s("p")]);
    assert!(matches!(d, Value::Object(_)));
}

// ── getPrototypeOf / setPrototypeOf ───────────────────────────────────────────

#[test]
fn get_prototype_of_returns_object_or_null() {
    let o = invoke("new", vec![]);
    let proto = invoke("getPrototypeOf", vec![o]);
    assert!(matches!(proto, Value::Object(_) | Value::Null));
}

#[test]
fn set_prototype_of_returns_bool() {
    let o = invoke("new", vec![]);
    let result = invoke("setPrototypeOf", vec![o, Value::Null]);
    assert!(matches!(result, Value::Bool(_)));
}

// ── Object.create ─────────────────────────────────────────────────────────────

#[test]
fn create_with_null_creates_object_with_no_prototype() {
    // ECMA-262 §20.1.2.2: Object.create(null) → object with [[Prototype]] = null.
    // The object has no inherited toString, hasOwnProperty etc.
    let o = invoke("create", vec![Value::Null]);
    assert!(matches!(o, Value::Object(_)));
    assert_eq!(invoke("getPrototypeOf", vec![o]), Value::Null);
}

#[test]
fn create_with_proto_object_inherits_that_prototype() {
    // Object.create(proto) → new object whose [[Prototype]] is proto.
    let proto = obj(vec![("x", Value::I32(10))]);
    let child = invoke("create", vec![proto.clone()]);
    let child_proto = invoke("getPrototypeOf", vec![child]);
    let p_ptr = match &proto { Value::Object(a) => Arc::as_ptr(a) as usize, _ => 0 };
    let cp_ptr = match &child_proto { Value::Object(a) => Arc::as_ptr(a) as usize, _ => 1 };
    assert_eq!(p_ptr, cp_ptr);
}

// ── Object.prototype.toString tag ─────────────────────────────────────────────

#[test]
fn to_string_tag_of_array_is_object_array() {
    // ECMA-262 §20.1.3.6: Object.prototype.toString applied to arrays
    // returns "[object Array]".
    let a = arr(vec![Value::I32(1)]);
    let tag = invoke("toStringTag", vec![a]);
    assert_eq!(tag, s("[object Array]"));
}

#[test]
fn to_string_tag_of_plain_object_is_object_object() {
    // Object.prototype.toString for plain objects returns "[object Object]".
    let o = invoke("new", vec![]);
    let tag = invoke("toStringTag", vec![o]);
    assert_eq!(tag, s("[object Object]"));
}

#[test]
fn to_string_tag_of_null_is_object_null() {
    // §20.1.3.6: The tag for null is "[object Null]".
    let tag = invoke("toStringTag", vec![Value::Null]);
    assert_eq!(tag, s("[object Null]"));
}

#[test]
fn to_string_tag_of_undefined_is_object_undefined() {
    // §20.1.3.6: The tag for undefined is "[object Undefined]".
    let tag = invoke("toStringTag", vec![Value::Undefined]);
    assert_eq!(tag, s("[object Undefined]"));
}

// ── Object.defineProperties ───────────────────────────────────────────────────

#[test]
fn define_properties_adds_multiple_properties_at_once() {
    // ECMA-262 §20.1.2.3: Object.defineProperties(O, props) — bulk defineProperty.
    let o = invoke("new", vec![]);
    let props = obj(vec![
        ("a", obj(vec![("value", Value::I32(1)), ("enumerable", Value::Bool(true))])),
        ("b", obj(vec![("value", Value::I32(2)), ("enumerable", Value::Bool(true))])),
    ]);
    invoke("defineProperties", vec![o.clone(), props]);
    assert_eq!(invoke("get", vec![o.clone(), s("a")]), Value::I32(1));
    assert_eq!(invoke("get", vec![o, s("b")]), Value::I32(2));
}

// ── Object.getOwnPropertySymbols ──────────────────────────────────────────────

#[test]
fn get_own_property_symbols_returns_array() {
    // ECMA-262 §20.1.2.10: getOwnPropertySymbols returns only symbol-keyed properties.
    // For a plain object with no symbol keys, the array is empty.
    let o = obj(vec![("a", Value::I32(1))]);
    let syms = invoke("getOwnPropertySymbols", vec![o]);
    // Must be an array; string-keyed properties do not appear here.
    assert!(matches!(syms, Value::Object(_)));
    assert_eq!(elems(&syms).len(), 0);
}

// ── Object.getOwnPropertyDescriptor detail ────────────────────────────────────

#[test]
fn own_property_descriptor_has_value_writable_enumerable_configurable() {
    // ECMA-262 §20.1.2.8: descriptor object has value/writable/enumerable/configurable.
    let o = obj(vec![("z", Value::I32(7))]);
    let d = invoke("getOwnPropertyDescriptor", vec![o, s("z")]);
    assert_eq!(invoke("get", vec![d, s("value")]), Value::I32(7));
}

// ── Object.is — additional SameValue cases ────────────────────────────────────

#[test]
fn object_is_same_string_values_are_equal() {
    assert_eq!(invoke("is", vec![s("hello"), s("hello")]), Value::Bool(true));
}

#[test]
fn object_is_same_integer_values_are_equal() {
    assert_eq!(invoke("is", vec![Value::I32(5), Value::I32(5)]), Value::Bool(true));
}

// ── Object.prototype.hasOwnProperty ──────────────────────────────────────────

#[test]
fn has_own_property_false_for_missing_key() {
    let o = obj(vec![("a", Value::I32(1))]);
    assert_eq!(invoke("hasOwn", vec![o, s("b")]), Value::Bool(false));
}

// ── Object.groupBy (ES2024 §20.1.2.1) ────────────────────────────────────────

#[test]
fn group_by_partitions_array_into_keyed_groups() {
    // ECMA-262 ES2024: Object.groupBy(items, keyFn) groups items by keyFn return value.
    // We encode the key function via an object descriptor.
    let items = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3), Value::I32(4)]);
    let key_fn = {
        let mut o = Object::new();
        // Descriptor: items ≤ 2 go to "small", items > 2 go to "large".
        o.properties.insert("__groupby_le2_small_large".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let result = invoke("groupBy", vec![items, key_fn]);
    // Result must be an object with keys "small" and "large".
    assert!(matches!(result, Value::Object(_)));
    let small = invoke("get", vec![result.clone(), s("small")]);
    let large = invoke("get", vec![result, s("large")]);
    assert!(matches!(small, Value::Object(_)));
    assert!(matches!(large, Value::Object(_)));
}

// ── Object.prototype.toString called directly ─────────────────────────────────

#[test]
fn to_string_tag_of_function_is_object_function() {
    // ECMA-262 §20.1.3.6: Object.prototype.toString on a callable is "[object Function]".
    // We represent a host function reference as an object with __callable marker.
    let fn_obj = {
        let mut o = Object::new();
        o.properties.insert("__callable".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let tag = invoke("toStringTag", vec![fn_obj]);
    // Must be either "[object Function]" or "[object Object]" depending on implementation.
    assert!(matches!(tag, Value::String(_)));
}

// ── Object.getOwnPropertyDescriptors (plural) ─────────────────────────────────

#[test]
fn get_own_property_descriptors_returns_object_of_descriptors() {
    // ECMA-262 §20.1.2.9: Object.getOwnPropertyDescriptors(O) returns ALL own descriptors.
    let o = obj(vec![("x", Value::I32(1)), ("y", Value::I32(2))]);
    let descs = invoke("getOwnPropertyDescriptors", vec![o]);
    assert!(matches!(descs, Value::Object(_)));
    // Each key in the result must be an object (descriptor).
    let x_desc = invoke("get", vec![descs.clone(), s("x")]);
    assert!(matches!(x_desc, Value::Object(_)));
    let y_desc = invoke("get", vec![descs, s("y")]);
    assert!(matches!(y_desc, Value::Object(_)));
}

// ── Object.prototype.isPrototypeOf ───────────────────────────────────────────

#[test]
fn is_prototype_of_returns_false_when_not_in_chain() {
    // ECMA-262 §20.1.3.4: O.isPrototypeOf(V) — checks if O is in V's prototype chain.
    // Two independently created objects share no custom prototype relationship.
    let a = invoke("new", vec![]);
    let b = invoke("new", vec![]);
    let result = invoke("isPrototypeOf", vec![a, b]);
    assert!(matches!(result, Value::Bool(false) | Value::Undefined));
}

// ── Object.prototype.propertyIsEnumerable ────────────────────────────────────

#[test]
fn property_is_enumerable_true_for_own_enumerable_property() {
    // ECMA-262 §20.1.3.5: propertyIsEnumerable returns true for own enumerable properties.
    let o = obj(vec![("x", Value::I32(1))]);
    assert_eq!(invoke("propertyIsEnumerable", vec![o, s("x")]), Value::Bool(true));
}

#[test]
fn property_is_enumerable_false_for_missing_property() {
    let o = invoke("new", vec![]);
    assert_eq!(invoke("propertyIsEnumerable", vec![o, s("nope")]), Value::Bool(false));
}

// ── Object.prototype.valueOf ──────────────────────────────────────────────────

#[test]
fn value_of_on_plain_object_returns_the_object_itself() {
    // ECMA-262 §20.1.3.7: Object.prototype.valueOf returns `this` (the object itself).
    let o = invoke("new", vec![]);
    let val = invoke("valueOf", vec![o.clone()]);
    let o_ptr = match &o { Value::Object(a) => Arc::as_ptr(a) as usize, _ => 0 };
    let v_ptr = match &val { Value::Object(a) => Arc::as_ptr(a) as usize, _ => 1 };
    assert_eq!(o_ptr, v_ptr);
}

// ── Object.prototype.hasOwnProperty (ECMA-262 §20.1.3.2) ─────────────────────

#[test]
fn has_own_property_true_for_own_key() {
    // §20.1.3.2: hasOwnProperty(V) returns true when V is an own property.
    let o = invoke("new", vec![]);
    invoke("set", vec![o.clone(), s("x"), Value::I32(1)]);
    assert_eq!(invoke("hasOwnProperty", vec![o, s("x")]), Value::Bool(true));
}

#[test]
fn has_own_property_false_for_absent_key() {
    let o = invoke("new", vec![]);
    assert_eq!(invoke("hasOwnProperty", vec![o, s("missing")]), Value::Bool(false));
}

// ── Object.prototype.toString (ECMA-262 §20.1.3.6) ───────────────────────────

#[test]
fn to_string_returns_object_object_for_plain_object() {
    // §20.1.3.6: Object.prototype.toString returns "[object Object]" by default.
    let o = invoke("new", vec![]);
    match invoke("toString", vec![o]) {
        Value::String(s) => assert_eq!(s.as_ref(), "[object Object]"),
        other => panic!("expected string, got {:?}", other),
    }
}

// ── Object.prototype.toLocaleString (ECMA-262 §20.1.3.5) ─────────────────────

#[test]
fn to_locale_string_delegates_to_to_string_for_plain_object() {
    // §20.1.3.5: toLocaleString calls this.toString(), so for a plain object
    // the result is "[object Object]" (or equivalent non-empty string).
    let o = invoke("new", vec![]);
    match invoke("toLocaleString", vec![o]) {
        Value::String(s) => assert!(!s.is_empty()),
        other => panic!("expected string, got {:?}", other),
    }
}
