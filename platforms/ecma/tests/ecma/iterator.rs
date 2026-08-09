//! Behaviour tests for `ecma:iterator` host imports.
//!
//! Reference: ECMA-262 §27.1 Iterator + TC39 Iterator Helpers proposal.
//!
//! Each test covers a distinct behaviour.

use std::sync::{Arc, Mutex};
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

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

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut chunk = Chunk::new("<ecma-iterator-test>");
    let import_idx = chunk.add_import("ecma:iterator", name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("VM run failed")
}

fn arr(values: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(values))))
}

fn next_step(it: Value) -> (Value, bool) {
    let result = invoke("next", vec![it]);
    match result {
        Value::Object(o) => {
            let o = o.lock().unwrap();
            let done = matches!(o.properties.get("done"), Some(Value::Bool(true)));
            let val = o
                .properties
                .get("value")
                .cloned()
                .unwrap_or(Value::Undefined);
            (val, done)
        }
        _ => (Value::Undefined, true),
    }
}

fn collect(it: Value) -> Vec<Value> {
    match invoke("toArray", vec![it]) {
        Value::Object(o) => match &o.lock().unwrap().kind {
            ObjectKind::Array(e) => e.clone(),
            _ => vec![],
        },
        _ => vec![],
    }
}

// ── next — iterator protocol ───────────────────────────────────────────────────

#[test]
fn next_result_has_value_and_done_properties() {
    let it = invoke("from", vec![arr(vec![Value::I32(1)])]);
    let result = invoke("next", vec![it]);
    assert!(matches!(result, Value::Object(_)));
    if let Value::Object(o) = result {
        let o = o.lock().unwrap();
        assert!(o.properties.contains_key("value"));
        assert!(o.properties.contains_key("done"));
    }
}

#[test]
fn next_done_is_false_while_elements_remain() {
    let it = invoke("from", vec![arr(vec![Value::I32(1)])]);
    let (_, done) = next_step(it);
    assert!(!done);
}

#[test]
fn next_done_becomes_true_after_exhaustion() {
    let it = invoke("from", vec![arr(vec![Value::I32(1)])]);
    next_step(it.clone()); // consume the one element
    let (_, done) = next_step(it);
    assert!(done);
}

#[test]
fn next_yields_values_in_source_order() {
    let it = invoke(
        "from",
        vec![arr(vec![Value::I32(10), Value::I32(20), Value::I32(30)])],
    );
    let (v1, _) = next_step(it.clone());
    let (v2, _) = next_step(it.clone());
    let (v3, _) = next_step(it);
    assert_eq!(v1, Value::I32(10));
    assert_eq!(v2, Value::I32(20));
    assert_eq!(v3, Value::I32(30));
}

// ── range ─────────────────────────────────────────────────────────────────────

#[test]
fn range_yields_integers_up_to_but_not_including_end() {
    let it = invoke("range", vec![Value::I32(0), Value::I32(3)]);
    let vals = collect(it);
    assert_eq!(vals.len(), 3);
    assert_eq!(vals[0], Value::I32(0));
    assert_eq!(vals[2], Value::I32(2));
}

#[test]
fn range_with_equal_start_and_end_is_empty() {
    let it = invoke("range", vec![Value::I32(5), Value::I32(5)]);
    assert!(collect(it).is_empty());
}

#[test]
fn range_with_step_skips_values() {
    let it = invoke("range", vec![Value::I32(0), Value::I32(6), Value::I32(2)]);
    let vals = collect(it);
    assert_eq!(vals, vec![Value::I32(0), Value::I32(2), Value::I32(4)]);
}

// ── take ──────────────────────────────────────────────────────────────────────

#[test]
fn take_limits_to_at_most_n_elements() {
    let it = invoke(
        "from",
        vec![arr(vec![
            Value::I32(1),
            Value::I32(2),
            Value::I32(3),
            Value::I32(4),
        ])],
    );
    let taken = invoke("take", vec![it, Value::I32(2)]);
    assert_eq!(collect(taken).len(), 2);
}

#[test]
fn take_zero_yields_empty_iterator() {
    let it = invoke("from", vec![arr(vec![Value::I32(1), Value::I32(2)])]);
    let taken = invoke("take", vec![it, Value::I32(0)]);
    assert!(collect(taken).is_empty());
}

#[test]
fn take_more_than_available_yields_all_elements() {
    let it = invoke("from", vec![arr(vec![Value::I32(1), Value::I32(2)])]);
    let taken = invoke("take", vec![it, Value::I32(100)]);
    assert_eq!(collect(taken).len(), 2);
}

// ── drop ──────────────────────────────────────────────────────────────────────

#[test]
fn drop_skips_first_n_elements() {
    let it = invoke(
        "from",
        vec![arr(vec![
            Value::I32(1),
            Value::I32(2),
            Value::I32(3),
            Value::I32(4),
        ])],
    );
    let remaining = invoke("drop", vec![it, Value::I32(2)]);
    let vals = collect(remaining);
    assert_eq!(vals, vec![Value::I32(3), Value::I32(4)]);
}

#[test]
fn drop_zero_leaves_all_elements() {
    let it = invoke("from", vec![arr(vec![Value::I32(1), Value::I32(2)])]);
    let remaining = invoke("drop", vec![it, Value::I32(0)]);
    assert_eq!(collect(remaining).len(), 2);
}

#[test]
fn drop_more_than_available_yields_empty() {
    let it = invoke("from", vec![arr(vec![Value::I32(1)])]);
    let remaining = invoke("drop", vec![it, Value::I32(10)]);
    assert!(collect(remaining).is_empty());
}

// ── toArray ───────────────────────────────────────────────────────────────────

#[test]
fn to_array_collects_all_values_into_array_object() {
    let it = invoke("range", vec![Value::I32(1), Value::I32(4)]);
    let a = invoke("toArray", vec![it]);
    assert!(matches!(a, Value::Object(_)));
    if let Value::Object(o) = a {
        assert!(matches!(o.lock().unwrap().kind, ObjectKind::Array(_)));
    }
}

#[test]
fn to_array_of_empty_iterator_returns_empty_array() {
    let it = invoke("from", vec![arr(vec![])]);
    assert!(collect(it).is_empty());
}

// ── Iterator.prototype.map (ES2025 Iterator Helpers) ─────────────────────────

#[test]
fn map_produces_an_iterator_of_transformed_values() {
    // TC39 Iterator Helpers §3.1.3.3: iterator.map(fn) applies fn to each value.
    let it = invoke("range", vec![Value::I32(1), Value::I32(4)]);
    let map_fn = {
        let mut o = Object::new();
        o.properties.insert("__map_mul".to_string(), Value::I32(10));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let mapped = invoke("map", vec![it, map_fn]);
    let values = collect(mapped);
    assert_eq!(values, vec![Value::I32(10), Value::I32(20), Value::I32(30)]);
}

// ── Iterator.prototype.filter ─────────────────────────────────────────────────

#[test]
fn filter_keeps_only_matching_values() {
    // §3.1.3.4: iterator.filter(pred) yields only elements where pred returns true.
    let it = invoke(
        "from",
        vec![arr(vec![
            Value::I32(1),
            Value::I32(2),
            Value::I32(3),
            Value::I32(4),
        ])],
    );
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(2));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let filtered = invoke("filter", vec![it, pred]);
    let values = collect(filtered);
    assert_eq!(values, vec![Value::I32(3), Value::I32(4)]);
}

// ── Iterator.prototype.reduce ─────────────────────────────────────────────────

#[test]
fn reduce_folds_iterator_to_single_value() {
    // §3.1.3.8: iterator.reduce(fn, init) left-folds.
    let it = invoke(
        "from",
        vec![arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)])],
    );
    let reducer = {
        let mut o = Object::new();
        o.properties
            .insert("__reduce_add".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let result = invoke("reduce", vec![it, reducer, Value::I32(0)]);
    assert_eq!(result, Value::I32(6));
}

// ── Iterator.prototype.some / every / find ────────────────────────────────────

#[test]
fn some_returns_true_when_any_element_matches() {
    let it = invoke(
        "from",
        vec![arr(vec![Value::I32(1), Value::I32(5), Value::I32(3)])],
    );
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(4));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    assert_eq!(invoke("some", vec![it, pred]), Value::Bool(true));
}

#[test]
fn every_returns_false_when_any_element_fails() {
    let it = invoke(
        "from",
        vec![arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)])],
    );
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(1));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    assert_eq!(invoke("every", vec![it, pred]), Value::Bool(false));
}

#[test]
fn find_returns_first_matching_element() {
    let it = invoke(
        "from",
        vec![arr(vec![Value::I32(1), Value::I32(4), Value::I32(9)])],
    );
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(3));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    assert_eq!(invoke("find", vec![it, pred]), Value::I32(4));
}

// ── Iterator.prototype.flatMap ────────────────────────────────────────────────

#[test]
fn flat_map_flattens_one_level_of_returned_iterators() {
    // §3.1.3.5: iterator.flatMap(fn) maps and flattens one level.
    let it = invoke("from", vec![arr(vec![Value::I32(1), Value::I32(2)])]);
    let flat_fn = {
        let mut o = Object::new();
        o.properties
            .insert("__flatmap_dup".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let result = invoke("flatMap", vec![it, flat_fn]);
    let values = collect(result);
    assert_eq!(
        values,
        vec![Value::I32(1), Value::I32(1), Value::I32(2), Value::I32(2)]
    );
}

// ── Iterator.prototype.forEach ────────────────────────────────────────────────

#[test]
fn for_each_consumes_iterator_and_returns_undefined() {
    // §3.1.3.6: iterator.forEach(fn) returns undefined (side-effect only).
    let it = invoke("range", vec![Value::I32(0), Value::I32(3)]);
    let noop = {
        let mut o = Object::new();
        o.properties.insert("__noop".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let result = invoke("forEach", vec![it, noop]);
    assert!(matches!(result, Value::Undefined | Value::Null));
}

// ── Iterator.concat (ES2025 §3.1.1.1) ────────────────────────────────────────

#[test]
fn concat_chains_multiple_iterables_in_order() {
    // ECMA-262 ES2025: Iterator.concat(iter1, iter2, ...) yields all elements in order.
    let it1 = invoke("from", vec![arr(vec![Value::I32(1), Value::I32(2)])]);
    let it2 = invoke("from", vec![arr(vec![Value::I32(3), Value::I32(4)])]);
    let combined = invoke("concat", vec![it1, it2]);
    let values = collect(combined);
    assert_eq!(
        values,
        vec![Value::I32(1), Value::I32(2), Value::I32(3), Value::I32(4)]
    );
}

#[test]
fn concat_with_empty_iterables_yields_non_empty_ones() {
    // Iterator.concat([], [1,2]) yields only [1,2].
    let empty = invoke("from", vec![arr(vec![])]);
    let full = invoke("from", vec![arr(vec![Value::I32(9)])]);
    let combined = invoke("concat", vec![empty, full]);
    let values = collect(combined);
    assert_eq!(values, vec![Value::I32(9)]);
}
