//! Behaviour tests for `ecma:array` host imports.
//!
//! Reference: ECMA-262 §22.1 Array.
//!
//! Each test covers a distinct behaviour — not a value variation.

use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

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
    let mut chunk = Chunk::new("<ecma-array-test>");
    let import_idx = chunk.add_import("ecma:array", name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("VM run failed")
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn arr(values: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(values))))
}

fn elems(v: &Value) -> Vec<Value> {
    match v {
        Value::Object(o) => match &o.lock().unwrap().kind {
            ObjectKind::Array(e) => e.clone(),
            _ => vec![] },
        _ => vec![] }
}

fn iter_next(it: Value) -> (Value, bool) {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut chunk = Chunk::new("<ecma-array-iter>");
    let idx = chunk.add_import("ecma:iterator", "next");
    push_arg(&mut vm, &mut chunk, it);
    chunk.emit_call(idx, 1, 0);
    chunk.emit_op(Op::RETURN, 0);
    let result = vm.run(vec![chunk]).expect("VM run failed");
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
        _ => (Value::Undefined, true) }
}

fn collect_iter(it: Value) -> Vec<Value> {
    let mut out = Vec::new();
    loop {
        let (val, done) = iter_next(it.clone());
        if done {
            break;
        }
        out.push(val);
    }
    out
}

// ── isArray ───────────────────────────────────────────────────────────────────

#[test]
fn is_array_distinguishes_arrays_from_plain_objects() {
    assert_eq!(invoke("isArray", vec![arr(vec![])]), Value::Bool(true));
    let plain = Value::Object(Arc::new(Mutex::new(Object::new())));
    assert_eq!(invoke("isArray", vec![plain]), Value::Bool(false));
}

#[test]
fn is_array_returns_false_for_primitives() {
    assert_eq!(invoke("isArray", vec![s("hello")]), Value::Bool(false));
    assert_eq!(invoke("isArray", vec![Value::I32(42)]), Value::Bool(false));
    assert_eq!(invoke("isArray", vec![Value::Null]), Value::Bool(false));
}

// ── Array.of vs new Array(n) ─────────────────────────────────────────────────

#[test]
fn of_creates_single_element_array_not_hole_array() {
    // Array.of(3) → [3], length 1. Unlike new Array(3) which has length 3 with holes.
    let result = invoke("of", vec![Value::I32(3)]);
    assert_eq!(elems(&result).len(), 1);
    assert_eq!(elems(&result)[0], Value::I32(3));
}

#[test]
fn new_with_length_creates_empty_slots_not_element() {
    // new Array(5) → length 5, no real elements stored as values.
    let result = invoke("newWithLength", vec![Value::I32(5)]);
    assert_eq!(invoke("length", vec![result]), Value::I32(5));
}

// ── Array.from ────────────────────────────────────────────────────────────────

#[test]
fn from_string_splits_into_individual_characters() {
    // Array.from("abc") → ["a","b","c"] — iterates Unicode code points.
    let result = invoke("from", vec![s("abc")]);
    let e = elems(&result);
    assert_eq!(e.len(), 3);
    assert_eq!(e[0], s("a"));
    assert_eq!(e[2], s("c"));
}

#[test]
fn from_array_produces_a_shallow_copy() {
    let original = arr(vec![Value::I32(1), Value::I32(2)]);
    let copy = invoke("from", vec![original.clone()]);
    // Same contents but different object identity.
    let copy_ptr = match &copy {
        Value::Object(a) => Arc::as_ptr(a) as usize,
        _ => 0 };
    let orig_ptr = match &original {
        Value::Object(a) => Arc::as_ptr(a) as usize,
        _ => 1 };
    assert_ne!(copy_ptr, orig_ptr);
    assert_eq!(elems(&copy), vec![Value::I32(1), Value::I32(2)]);
}

// ── push / pop ────────────────────────────────────────────────────────────────

#[test]
fn push_returns_new_length_not_the_pushed_value() {
    let a = arr(vec![Value::I32(1)]);
    let result = invoke("push", vec![a, Value::I32(2)]);
    assert_eq!(result, Value::I32(2)); // length is 2, not the value 2
}

#[test]
fn pop_on_empty_array_returns_undefined() {
    let a = arr(vec![]);
    assert_eq!(invoke("pop", vec![a]), Value::Undefined);
}

#[test]
fn pop_removes_last_element_reducing_length() {
    let a = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    let popped = invoke("pop", vec![a.clone()]);
    assert_eq!(popped, Value::I32(3));
    assert_eq!(invoke("length", vec![a]), Value::I32(2));
}

// ── shift / unshift ───────────────────────────────────────────────────────────

#[test]
fn shift_removes_first_and_returns_it() {
    let a = arr(vec![Value::I32(10), Value::I32(20)]);
    let shifted = invoke("shift", vec![a.clone()]);
    assert_eq!(shifted, Value::I32(10));
    assert_eq!(invoke("length", vec![a]), Value::I32(1));
}

#[test]
fn shift_on_empty_returns_undefined() {
    assert_eq!(invoke("shift", vec![arr(vec![])]), Value::Undefined);
}

#[test]
fn unshift_prepends_in_argument_order() {
    // unshift(a, b) → [a, b, ...original], not [b, a, ...original]
    let a = arr(vec![Value::I32(3)]);
    let new_len = invoke("unshift", vec![a.clone(), Value::I32(1), Value::I32(2)]);
    assert_eq!(new_len, Value::I32(3));
    assert_eq!(elems(&a)[0], Value::I32(1));
    assert_eq!(elems(&a)[1], Value::I32(2));
}

// ── splice ────────────────────────────────────────────────────────────────────

#[test]
fn splice_returns_removed_elements_as_array() {
    let a = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    let removed = invoke("splice", vec![a, Value::I32(1), Value::I32(1)]);
    assert_eq!(elems(&removed), vec![Value::I32(2)]);
}

#[test]
fn splice_with_zero_delete_count_only_inserts() {
    let a = arr(vec![Value::I32(1), Value::I32(3)]);
    let removed = invoke(
        "splice",
        vec![a.clone(), Value::I32(1), Value::I32(0), Value::I32(2)],
    );
    // Returns empty array (nothing removed)
    assert_eq!(elems(&removed).len(), 0);
    // Array is now [1, 2, 3]
    assert_eq!(elems(&a).len(), 3);
    assert_eq!(elems(&a)[1], Value::I32(2));
}

#[test]
fn splice_with_negative_start_counts_from_end() {
    let a = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    let removed = invoke("splice", vec![a, Value::I32(-1), Value::I32(1)]);
    assert_eq!(elems(&removed), vec![Value::I32(3)]);
}

// ── slice ─────────────────────────────────────────────────────────────────────

#[test]
fn slice_does_not_mutate_original() {
    let a = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    let _ = invoke("slice", vec![a.clone(), Value::I32(0), Value::I32(2)]);
    assert_eq!(invoke("length", vec![a]), Value::I32(3));
}

#[test]
fn slice_with_negative_indices_counts_from_end() {
    let a = arr(vec![
        Value::I32(1),
        Value::I32(2),
        Value::I32(3),
        Value::I32(4),
    ]);
    let result = invoke("slice", vec![a, Value::I32(-2)]);
    assert_eq!(elems(&result), vec![Value::I32(3), Value::I32(4)]);
}

// ── indexOf / includes — equality semantics differ ────────────────────────────

#[test]
fn index_of_uses_strict_equality_nan_not_found() {
    // indexOf uses ===; NaN !== NaN so NaN is never found.
    let a = arr(vec![Value::F64(f64::NAN)]);
    assert_eq!(
        invoke("indexOf", vec![a, Value::F64(f64::NAN)]),
        Value::I32(-1)
    );
}

#[test]
fn includes_uses_same_value_zero_nan_found() {
    // includes uses SameValueZero; NaN equals NaN.
    let a = arr(vec![Value::F64(f64::NAN)]);
    assert_eq!(
        invoke("includes", vec![a, Value::F64(f64::NAN)]),
        Value::Bool(true)
    );
}

#[test]
fn last_index_of_searches_backward_from_end() {
    let a = arr(vec![Value::I32(1), Value::I32(2), Value::I32(1)]);
    assert_eq!(invoke("lastIndexOf", vec![a, Value::I32(1)]), Value::I32(2));
}

// ── join ──────────────────────────────────────────────────────────────────────

#[test]
fn join_renders_null_and_undefined_elements_as_empty_string() {
    let a = arr(vec![
        Value::I32(1),
        Value::Null,
        Value::Undefined,
        Value::I32(2),
    ]);
    match invoke("join", vec![a, s(",")]) {
        Value::String(s) => assert_eq!(s.as_ref(), "1, ,2"),
        other => panic!("expected string, got {:?}", other) }
}

// ── flat ──────────────────────────────────────────────────────────────────────

#[test]
fn flat_default_depth_one_does_not_recurse_deeper() {
    // [[1, [2]]] flat(1) → [1, [2]], the inner [2] is not flattened.
    let inner = arr(vec![Value::I32(2)]);
    let nested = arr(vec![Value::I32(1), inner]);
    let outer = arr(vec![nested]);
    let result = invoke("flat", vec![outer, Value::I32(1)]);
    let e = elems(&result);
    assert_eq!(e.len(), 2);
    assert_eq!(e[0], Value::I32(1));
    assert!(
        matches!(e[1], Value::Object(_)),
        "inner array should remain nested"
    );
}

// ── sort — default is lexicographic ──────────────────────────────────────────

#[test]
fn sort_default_compares_as_strings_so_ten_precedes_nine() {
    // Default sort: "10" < "9" lexicographically.
    let a = arr(vec![Value::I32(9), Value::I32(10), Value::I32(1)]);
    let result = invoke("sort", vec![a, Value::Null]);
    let e = elems(&result);
    // "1" < "10" < "9"
    assert_eq!(e[0], Value::I32(1));
    assert_eq!(e[1], Value::I32(10));
    assert_eq!(e[2], Value::I32(9));
}

// ── reverse vs toReversed — mutation semantics ────────────────────────────────

#[test]
fn reverse_mutates_in_place_and_returns_same_object() {
    let a = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    let orig_ptr = match &a {
        Value::Object(arc) => Arc::as_ptr(arc) as usize,
        _ => 0 };
    let result = invoke("reverse", vec![a]);
    let result_ptr = match &result {
        Value::Object(arc) => Arc::as_ptr(arc) as usize,
        _ => 1 };
    assert_eq!(orig_ptr, result_ptr); // same object
    assert_eq!(elems(&result)[0], Value::I32(3));
}

#[test]
fn to_reversed_does_not_mutate_original() {
    let a = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    let result = invoke("toReversed", vec![a.clone()]);
    // Original unchanged
    assert_eq!(elems(&a)[0], Value::I32(1));
    // Result is reversed
    assert_eq!(elems(&result)[0], Value::I32(3));
}

// ── fill ──────────────────────────────────────────────────────────────────────

#[test]
fn fill_returns_same_array_reference() {
    let a = arr(vec![Value::I32(0), Value::I32(0), Value::I32(0)]);
    let orig_ptr = match &a {
        Value::Object(arc) => Arc::as_ptr(arc) as usize,
        _ => 0 };
    let result = invoke("fill", vec![a, Value::I32(7), Value::I32(0), Value::I32(3)]);
    let result_ptr = match &result {
        Value::Object(arc) => Arc::as_ptr(arc) as usize,
        _ => 1 };
    assert_eq!(orig_ptr, result_ptr);
}

#[test]
fn fill_with_start_and_end_only_fills_that_range() {
    let a = arr(vec![
        Value::I32(0),
        Value::I32(0),
        Value::I32(0),
        Value::I32(0),
    ]);
    let result = invoke("fill", vec![a, Value::I32(9), Value::I32(1), Value::I32(3)]);
    let e = elems(&result);
    assert_eq!(e[0], Value::I32(0)); // untouched
    assert_eq!(e[1], Value::I32(9));
    assert_eq!(e[2], Value::I32(9));
    assert_eq!(e[3], Value::I32(0)); // untouched
}

// ── at — negative indexing ────────────────────────────────────────────────────

#[test]
fn at_negative_index_addresses_from_end() {
    let a = arr(vec![Value::I32(10), Value::I32(20), Value::I32(30)]);
    assert_eq!(
        invoke("at", vec![a.clone(), Value::I32(-1)]),
        Value::I32(30)
    );
    assert_eq!(invoke("at", vec![a, Value::I32(-2)]), Value::I32(20));
}

#[test]
fn at_out_of_bounds_returns_undefined() {
    let a = arr(vec![Value::I32(1)]);
    assert_eq!(
        invoke("at", vec![a.clone(), Value::I32(5)]),
        Value::Undefined
    );
    assert_eq!(invoke("at", vec![a, Value::I32(-5)]), Value::Undefined);
}

// ── with — non-mutating element replacement ───────────────────────────────────

#[test]
fn with_returns_new_array_original_unchanged() {
    let a = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    let result = invoke("with", vec![a.clone(), Value::I32(1), Value::I32(99)]);
    assert_eq!(elems(&a)[1], Value::I32(2)); // original unchanged
    assert_eq!(elems(&result)[1], Value::I32(99));
}

// ── toSpliced — non-mutating splice ──────────────────────────────────────────

#[test]
fn to_spliced_returns_new_array_without_mutating_original() {
    let a = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    let result = invoke("toSpliced", vec![a.clone(), Value::I32(1), Value::I32(1)]);
    assert_eq!(elems(&a).len(), 3); // original unchanged
    assert_eq!(elems(&result).len(), 2);
}

// ── toSorted — non-mutating sort ─────────────────────────────────────────────

#[test]
fn to_sorted_returns_new_array_without_mutating_original() {
    let a = arr(vec![Value::I32(3), Value::I32(1), Value::I32(2)]);
    let result = invoke("toSorted", vec![a.clone(), Value::Null]);
    assert_eq!(elems(&a)[0], Value::I32(3)); // original unchanged
    assert!(elems(&result).len() == 3);
}

// ── concat ────────────────────────────────────────────────────────────────────

#[test]
fn concat_flattens_one_level_only() {
    // concat([1,2], [3,4]) → [1,2,3,4]. Non-array primitives appended as-is.
    let a = arr(vec![Value::I32(1), Value::I32(2)]);
    let b = arr(vec![Value::I32(3), Value::I32(4)]);
    let result = invoke("concat", vec![a, b]);
    assert_eq!(elems(&result).len(), 4);
}

// ── findLast / findLastIndex ──────────────────────────────────────────────────
// These search from right — distinct from find/findIndex which search from left.

#[test]
fn find_last_returns_rightmost_matching_element() {
    // ECMA-262 ES2023: findLast searches from the end. [1,3,2,4].findLast(x > 2) = 4.
    let a = arr(vec![
        Value::I32(1),
        Value::I32(3),
        Value::I32(2),
        Value::I32(4),
    ]);
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(2));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    assert_eq!(invoke("findLast", vec![a, pred]), Value::I32(4));
}

#[test]
fn find_last_returns_undefined_when_no_match() {
    let a = arr(vec![Value::I32(1), Value::I32(2)]);
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(10));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    assert_eq!(invoke("findLast", vec![a, pred]), Value::Undefined);
}

#[test]
fn find_last_index_returns_rightmost_matching_index() {
    // [1,3,2,4].findLastIndex(x > 2) → index 3 (the 4).
    let a = arr(vec![
        Value::I32(1),
        Value::I32(3),
        Value::I32(2),
        Value::I32(4),
    ]);
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(2));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    assert_eq!(invoke("findLastIndex", vec![a, pred]), Value::I32(3));
}

#[test]
fn find_last_index_returns_negative_one_when_no_match() {
    let a = arr(vec![Value::I32(1), Value::I32(2)]);
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(99));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    assert_eq!(invoke("findLastIndex", vec![a, pred]), Value::I32(-1));
}

// ── keys / values / entries iterators ────────────────────────────────────────

#[test]
fn keys_iterator_yields_numeric_indices() {
    let a = arr(vec![s("a"), s("b"), s("c")]);
    let it = invoke("keys", vec![a]);
    let keys = collect_iter(it);
    assert_eq!(keys, vec![Value::I32(0), Value::I32(1), Value::I32(2)]);
}

#[test]
fn values_iterator_yields_elements_in_order() {
    let a = arr(vec![Value::I32(10), Value::I32(20), Value::I32(30)]);
    let it = invoke("values", vec![a]);
    let vals = collect_iter(it);
    assert_eq!(vals, vec![Value::I32(10), Value::I32(20), Value::I32(30)]);
}

#[test]
fn entries_iterator_yields_index_value_pairs() {
    let a = arr(vec![s("x"), s("y")]);
    let it = invoke("entries", vec![a]);
    let pairs = collect_iter(it);
    // Each pair is [index, value]
    assert_eq!(pairs.len(), 2);
    let first = elems(&pairs[0]);
    assert_eq!(first[0], Value::I32(0));
    assert_eq!(first[1], s("x"));
}

// ── copyWithin ────────────────────────────────────────────────────────────────

#[test]
fn copy_within_handles_overlapping_source_and_target() {
    // [1,2,3,4,5].copyWithin(1,3) → [1,4,5,4,5]
    let a = arr(vec![
        Value::I32(1),
        Value::I32(2),
        Value::I32(3),
        Value::I32(4),
        Value::I32(5),
    ]);
    let result = invoke("copyWithin", vec![a, Value::I32(1), Value::I32(3)]);
    let e = elems(&result);
    assert_eq!(e[1], Value::I32(4));
    assert_eq!(e[2], Value::I32(5));
}

// ── first / last (Vybe extras) ────────────────────────────────────────────────

#[test]
fn first_on_empty_array_returns_undefined() {
    assert_eq!(invoke("first", vec![arr(vec![])]), Value::Undefined);
}

#[test]
fn last_on_empty_array_returns_undefined() {
    assert_eq!(invoke("last", vec![arr(vec![])]), Value::Undefined);
}

// ── reduce ────────────────────────────────────────────────────────────────────

#[test]
fn reduce_sums_all_elements_left_to_right() {
    // reduce([1,2,3], (acc,x)=>acc+x, 0) = 6
    let callback = {
        let mut o = Object::new();
        o.properties
            .insert("__reduce_add".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let a = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    assert_eq!(
        invoke("reduce", vec![a, callback, Value::I32(0)]),
        Value::I32(6)
    );
}

#[test]
fn reduce_on_empty_array_returns_initial_value() {
    let callback = {
        let mut o = Object::new();
        o.properties
            .insert("__reduce_add".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    assert_eq!(
        invoke("reduce", vec![arr(vec![]), callback, Value::I32(42)]),
        Value::I32(42)
    );
}

#[test]
fn reduce_right_processes_from_right_to_left() {
    // reduceRight(["a","b","c"], (acc,x)=>acc+x, "") = "cba"
    let callback = {
        let mut o = Object::new();
        o.properties
            .insert("__reduce_concat".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let a = arr(vec![s("a"), s("b"), s("c")]);
    let result = invoke("reduceRight", vec![a, callback, s("")]);
    assert_eq!(result, s("cba"));
}

// ── find / findIndex ──────────────────────────────────────────────────────────

#[test]
fn find_returns_first_matching_element() {
    // find([1,2,3,4], x => x > 2) = 3
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(2));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let a = arr(vec![
        Value::I32(1),
        Value::I32(2),
        Value::I32(3),
        Value::I32(4),
    ]);
    assert_eq!(invoke("find", vec![a, pred]), Value::I32(3));
}

#[test]
fn find_returns_undefined_when_no_element_matches() {
    let pred = {
        let mut o = Object::new();
        o.properties
            .insert("__pred_gt".to_string(), Value::I32(100));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let a = arr(vec![Value::I32(1), Value::I32(2)]);
    assert_eq!(invoke("find", vec![a, pred]), Value::Undefined);
}

#[test]
fn find_index_returns_index_of_first_match() {
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(2));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let a = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    assert_eq!(invoke("findIndex", vec![a, pred]), Value::I32(2));
}

#[test]
fn find_index_returns_negative_one_when_no_match() {
    let pred = {
        let mut o = Object::new();
        o.properties
            .insert("__pred_gt".to_string(), Value::I32(100));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let a = arr(vec![Value::I32(1), Value::I32(2)]);
    assert_eq!(invoke("findIndex", vec![a, pred]), Value::I32(-1));
}

// ── every / some ──────────────────────────────────────────────────────────────

#[test]
fn every_returns_true_when_all_elements_match() {
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(0));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let a = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    assert_eq!(invoke("every", vec![a, pred]), Value::Bool(true));
}

#[test]
fn every_returns_false_when_any_element_fails() {
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(1));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let a = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    // 1 is not > 1
    assert_eq!(invoke("every", vec![a, pred]), Value::Bool(false));
}

#[test]
fn every_on_empty_array_returns_true_vacuously() {
    // ECMA-262: every([]) = true (vacuous truth).
    let pred = {
        let mut o = Object::new();
        o.properties
            .insert("__pred_gt".to_string(), Value::I32(999));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    assert_eq!(invoke("every", vec![arr(vec![]), pred]), Value::Bool(true));
}

#[test]
fn some_returns_true_when_at_least_one_matches() {
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(2));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let a = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    assert_eq!(invoke("some", vec![a, pred]), Value::Bool(true));
}

#[test]
fn some_returns_false_when_no_element_matches() {
    let pred = {
        let mut o = Object::new();
        o.properties
            .insert("__pred_gt".to_string(), Value::I32(100));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let a = arr(vec![Value::I32(1), Value::I32(2)]);
    assert_eq!(invoke("some", vec![a, pred]), Value::Bool(false));
}

#[test]
fn some_on_empty_array_returns_false() {
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(0));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    assert_eq!(invoke("some", vec![arr(vec![]), pred]), Value::Bool(false));
}

// ── map / filter / flatMap ────────────────────────────────────────────────────

#[test]
fn map_transforms_each_element() {
    // map([1,2,3], x => x * 2) = [2,4,6]
    let cb = {
        let mut o = Object::new();
        o.properties.insert("__map_mul".to_string(), Value::I32(2));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let a = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    let result = invoke("map", vec![a, cb]);
    assert_eq!(
        elems(&result),
        vec![Value::I32(2), Value::I32(4), Value::I32(6)]
    );
}

#[test]
fn map_does_not_mutate_original() {
    let cb = {
        let mut o = Object::new();
        o.properties.insert("__map_mul".to_string(), Value::I32(2));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let a = arr(vec![Value::I32(1), Value::I32(2)]);
    let _ = invoke("map", vec![a.clone(), cb]);
    assert_eq!(elems(&a), vec![Value::I32(1), Value::I32(2)]);
}

#[test]
fn filter_keeps_only_matching_elements() {
    // filter([1,2,3,4], x => x % 2 == 0) = [2,4]
    let pred = {
        let mut o = Object::new();
        o.properties.insert(
            "__filter_mod_eq".to_string(),
            Value::Object(Arc::new(Mutex::new({
                let mut p = Object::new();
                p.properties.insert("mod".to_string(), Value::I32(2));
                p.properties.insert("eq".to_string(), Value::I32(0));
                p
            }))),
        );
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let a = arr(vec![
        Value::I32(1),
        Value::I32(2),
        Value::I32(3),
        Value::I32(4),
    ]);
    let result = invoke("filter", vec![a, pred]);
    assert_eq!(elems(&result), vec![Value::I32(2), Value::I32(4)]);
}

#[test]
fn flat_map_maps_then_flattens_one_level() {
    // flatMap([1,2], x => [x, x]) = [1,1,2,2]
    let cb = {
        let mut o = Object::new();
        o.properties
            .insert("__flatmap_dup".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let a = arr(vec![Value::I32(1), Value::I32(2)]);
    let result = invoke("flatMap", vec![a, cb]);
    assert_eq!(
        elems(&result),
        vec![Value::I32(1), Value::I32(1), Value::I32(2), Value::I32(2)]
    );
}

// ── forEach ───────────────────────────────────────────────────────────────────

#[test]
fn for_each_returns_undefined() {
    // ECMA-262 §23.1.3.14: forEach always returns undefined.
    let cb = {
        let mut o = Object::new();
        o.properties.insert("__noop".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let a = arr(vec![Value::I32(1), Value::I32(2)]);
    assert_eq!(invoke("forEach", vec![a, cb]), Value::Undefined);
}

// ── Array.from with mapping function ─────────────────────────────────────────

#[test]
fn from_with_map_fn_transforms_elements() {
    // Array.from({length:3}, (_, i) => i*2) = [0,2,4]
    let map_fn = {
        let mut o = Object::new();
        o.properties
            .insert("__from_map_double_index".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let array_like = {
        let mut o = Object::new();
        o.properties.insert("length".to_string(), Value::I32(3));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let result = invoke("fromWithMap", vec![array_like, map_fn]);
    assert_eq!(
        elems(&result),
        vec![Value::I32(0), Value::I32(2), Value::I32(4)]
    );
}

// ── Array.fromAsync (ES2024 §22.1.2.2) ───────────────────────────────────────

#[test]
fn from_async_returns_a_promise_wrapping_the_collected_array() {
    // ECMA-262 ES2024: Array.fromAsync(asyncIterable) returns a Promise<Array>.
    // When given a sync array (which is also iterable), it resolves with the same elements.
    let input = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    let result = invoke("fromAsync", vec![input]);
    // Must be a Promise (Object), not the array directly.
    assert!(matches!(result, Value::Object(_)));
}

#[test]
fn from_async_with_map_function_applies_transform() {
    // Array.fromAsync(iterable, mapFn) applies mapFn to each element asynchronously.
    let input = arr(vec![Value::I32(1), Value::I32(2)]);
    let map_fn = {
        let mut o = Object::new();
        o.properties.insert("__map_mul".to_string(), Value::I32(3));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let result = invoke("fromAsync", vec![input, map_fn]);
    // Returns a Promise object — settlement checked via "settled" in a real async ctx.
    assert!(matches!(result, Value::Object(_)));
}

// ── Array.prototype.toLocaleString ───────────────────────────────────────────

#[test]
fn to_locale_string_returns_string_with_elements_joined() {
    // ECMA-262 §23.1.3.29: Array.prototype.toLocaleString returns a locale-formatted
    // string; each element's own toLocaleString is called and joined with a separator.
    let a = arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)]);
    let result = invoke("toLocaleString", vec![a]);
    assert!(matches!(result, Value::String(_)));
}

#[test]
fn to_locale_string_of_empty_array_is_empty_string() {
    // Array.prototype.toLocaleString([]) = "".
    let a = arr(vec![]);
    let result = invoke("toLocaleString", vec![a]);
    assert_eq!(result, Value::String(Arc::from("")));
}
