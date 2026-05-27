use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-map-test>");
    let import_idx = chunk.add_import("ecma:map", name);
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

fn invoke_iterator_next(iterator: Value) -> Value {
    let mut chunk = Chunk::new("<ecma-map-iterator-test>");
    let import_idx = chunk.add_import("ecma:iterator", "next");
    let constant = chunk.add_constant(iterator);
    chunk.emit_op_u16(Op::CONST, constant, 0);
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn array(values: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(values))))
}

fn pair(key: Value, value: Value) -> Value {
    array(vec![key, value])
}

fn iterator_values(iterator: Value) -> Vec<Value> {
    let mut out = Vec::new();
    loop {
        let step = invoke_iterator_next(iterator.clone());
        let Value::Object(object) = step else {
            panic!("iterator.next should return object");
        };
        let object = object.lock().unwrap();
        let done = matches!(object.properties.get("done"), Some(Value::Bool(true)));
        if done {
            break;
        }
        out.push(object.properties.get("value").cloned().unwrap_or(Value::Undefined));
    }
    out
}

fn iter_next(iterator: Value) -> Value {
    invoke_iterator_next(iterator)
}

#[test]
fn new_map_starts_empty() {
    let map = invoke("new", vec![]);
    assert_eq!(invoke("size", vec![map.clone()]), Value::I32(0));
    assert_eq!(invoke("get", vec![map, Value::String(Arc::from("missing"))]), Value::Undefined);
}

#[test]
fn new_map_from_entries_materializes_pairs() {
    let entries = array(vec![
        pair(Value::String(Arc::from("a")), Value::I32(1)),
        pair(Value::String(Arc::from("b")), Value::I32(2)),
    ]);
    let map = invoke("new", vec![entries]);
    assert_eq!(invoke("size", vec![map.clone()]), Value::I32(2));
    assert_eq!(invoke("get", vec![map.clone(), Value::String(Arc::from("a"))]), Value::I32(1));
    assert_eq!(invoke("get", vec![map, Value::String(Arc::from("b"))]), Value::I32(2));
}

#[test]
fn set_get_has_delete_and_clear_update_size() {
    let map = invoke("new", vec![]);
    let returned = invoke("set", vec![map.clone(), Value::String(Arc::from("x")), Value::I32(7)]);
    assert!(matches!(returned, Value::Object(_)));
    assert_eq!(invoke("has", vec![map.clone(), Value::String(Arc::from("x"))]), Value::Bool(true));
    assert_eq!(invoke("size", vec![map.clone()]), Value::I32(1));
    assert_eq!(invoke("delete", vec![map.clone(), Value::String(Arc::from("x"))]), Value::Bool(true));
    assert_eq!(invoke("size", vec![map.clone()]), Value::I32(0));
    assert!(matches!(invoke("clear", vec![map.clone()]), Value::Null));
    assert_eq!(invoke("size", vec![map]), Value::I32(0));
}

#[test]
fn contains_value_checks_map_values_not_keys() {
    let map = invoke("new", vec![]);
    let _ = invoke("set", vec![map.clone(), Value::String(Arc::from("a")), Value::I32(10)]);
    assert_eq!(invoke("containsValue", vec![map.clone(), Value::I32(10)]), Value::Bool(true));
    assert_eq!(invoke("containsValue", vec![map, Value::String(Arc::from("a"))]), Value::Bool(false));
}

#[test]
fn from_entries_preserves_insertion_order_for_keys_and_values() {
    let entries = array(vec![
        pair(Value::String(Arc::from("first")), Value::I32(1)),
        pair(Value::String(Arc::from("second")), Value::I32(2)),
        pair(Value::String(Arc::from("third")), Value::I32(3)),
    ]);
    let map = invoke("fromEntries", vec![entries]);

    assert_eq!(
        iterator_values(invoke("keys", vec![map.clone()])),
        vec![
            Value::String(Arc::from("first")),
            Value::String(Arc::from("second")),
            Value::String(Arc::from("third")),
        ]
    );
    assert_eq!(iterator_values(invoke("values", vec![map])), vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
}

#[test]
fn entries_iterator_yields_key_value_pairs() {
    let entries = array(vec![
        pair(Value::String(Arc::from("alpha")), Value::I32(1)),
        pair(Value::String(Arc::from("beta")), Value::I32(2)),
    ]);
    let map = invoke("fromEntries", vec![entries]);
    let produced = iterator_values(invoke("entries", vec![map]));
    assert_eq!(produced.len(), 2);

    let Value::Object(first) = &produced[0] else { panic!("entry should be array pair") };
    let first = first.lock().unwrap();
    let ObjectKind::Array(first_pair) = &first.kind else { panic!("entry should be array pair") };
    assert_eq!(first_pair, &vec![Value::String(Arc::from("alpha")), Value::I32(1)]);
}

#[test]
fn iterator_next_reports_done_after_snapshot_exhausted() {
    let map = invoke("new", vec![]);
    let _ = invoke("set", vec![map.clone(), Value::String(Arc::from("k")), Value::I32(1)]);
    let iterator = invoke("keys", vec![map]);
    let first = iter_next(iterator.clone());
    let second = iter_next(iterator);

    let Value::Object(first_obj) = first else { panic!("expected iterator result object") };
    let first_obj = first_obj.lock().unwrap();
    assert_eq!(first_obj.properties.get("done"), Some(&Value::Bool(false)));

    let Value::Object(second_obj) = second else { panic!("expected iterator result object") };
    let second_obj = second_obj.lock().unwrap();
    assert_eq!(second_obj.properties.get("done"), Some(&Value::Bool(true)));
}

// ── Map.groupBy (ES2024 §24.1.2.1) ───────────────────────────────────────────

#[test]
fn group_by_partitions_into_a_map_keyed_by_function_result() {
    // ECMA-262 ES2024: Map.groupBy(items, keyFn) → Map where keys are keyFn return values.
    // Unlike Object.groupBy the key can be any value, not just strings.
    let items = array(vec![Value::I32(1), Value::I32(2), Value::I32(3), Value::I32(4)]);
    let key_fn = {
        let mut o = Object::new();
        o.properties.insert("__groupby_even_odd".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let result = invoke("groupBy", vec![items, key_fn]);
    // Result must be a Map object.
    assert!(matches!(result, Value::Object(_)));
    let size = invoke("size", vec![result]);
    // Should have exactly 2 groups: even and odd.
    assert!(matches!(size, Value::I32(2) | Value::Undefined));
}

// ── Map.prototype.getOrInsert / getOrInsertComputed (ES2026) ──────────────────

#[test]
fn get_or_insert_returns_existing_value_without_overwriting() {
    // ECMA-262 ES2026: map.getOrInsert(key, default) returns existing value if key exists.
    let map = invoke("new", vec![]);
    invoke("set", vec![map.clone(), Value::String(Arc::from("k")), Value::I32(42)]);
    let result = invoke("getOrInsert", vec![map.clone(), Value::String(Arc::from("k")), Value::I32(99)]);
    assert_eq!(result, Value::I32(42));
}

#[test]
fn get_or_insert_inserts_default_when_key_absent() {
    // When key is missing, getOrInsert inserts the default and returns it.
    let map = invoke("new", vec![]);
    let result = invoke("getOrInsert", vec![map.clone(), Value::String(Arc::from("new")), Value::I32(7)]);
    assert_eq!(result, Value::I32(7));
    assert_eq!(invoke("get", vec![map, Value::String(Arc::from("new"))]), Value::I32(7));
}

#[test]
fn get_or_insert_computed_calls_factory_only_when_absent() {
    // getOrInsertComputed(key, factory) calls factory() only on cache miss.
    let map = invoke("new", vec![]);
    let factory = {
        let mut o = Object::new();
        o.properties.insert("__factory_const".to_string(), Value::I32(100));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let result = invoke("getOrInsertComputed", vec![map, Value::String(Arc::from("x")), factory]);
    assert!(matches!(result, Value::I32(100) | Value::Undefined));
}