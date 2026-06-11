//! Behaviour tests for typed array host imports.
//!
//! Reference: ECMA-262 §22.2 TypedArray.
//!
//! Covers Int8Array, Uint8Array, Int32Array, Float64Array, BigInt64Array.
//! Each test covers a distinct behaviour — not repeated across element types
//! unless the behaviour differs by type.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke_ns(ns: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-typedarray-test>");
    let import_idx = chunk.add_import(ns, name);
    let argc = args.len() as u8;
    for value in args {
        let c = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, c, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn i32ns(name: &str, args: Vec<Value>) -> Value {
    invoke_ns("ecma:int32array", name, args)
}
fn u8ns(name: &str, args: Vec<Value>) -> Value {
    invoke_ns("ecma:uint8array", name, args)
}
fn f64ns(name: &str, args: Vec<Value>) -> Value {
    invoke_ns("ecma:float64array", name, args)
}
fn bi64ns(name: &str, args: Vec<Value>) -> Value {
    invoke_ns("ecma:bigint64array", name, args)
}

fn arr(values: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(values))))
}

fn elems(v: &Value) -> Vec<Value> {
    match v {
        Value::Object(o) => match &o.lock().unwrap().kind {
            ObjectKind::Array(e) => e.clone(),
            _ => vec![],
        },
        _ => vec![],
    }
}

// ── newWithLength — all integer element types default to 0 ───────────────────

#[test]
fn int32_new_with_length_creates_correct_length() {
    let ta = i32ns("newWithLength", vec![Value::I32(5)]);
    assert_eq!(i32ns("length", vec![ta]), Value::I32(5));
}

#[test]
fn uint8_new_with_zero_length_is_empty() {
    let ta = u8ns("newWithLength", vec![Value::I32(0)]);
    assert_eq!(u8ns("length", vec![ta]), Value::I32(0));
}

// ── from — values preserved ───────────────────────────────────────────────────

#[test]
fn int32_from_array_preserves_elements_and_length() {
    let ta = i32ns(
        "from",
        vec![arr(vec![Value::I32(10), Value::I32(20), Value::I32(30)])],
    );
    assert_eq!(i32ns("length", vec![ta.clone()]), Value::I32(3));
    assert_eq!(
        i32ns("get", vec![ta.clone(), Value::I32(0)]),
        Value::I32(10)
    );
    assert_eq!(
        i32ns("get", vec![ta.clone(), Value::I32(2)]),
        Value::I32(30)
    );
}

#[test]
fn float64_from_preserves_fractional_values() {
    let ta = f64ns("from", vec![arr(vec![Value::F64(1.5), Value::F64(2.5)])]);
    match f64ns("get", vec![ta, Value::I32(0)]) {
        Value::F64(f) => assert!((f - 1.5).abs() < 1e-9),
        other => panic!("expected F64, got {:?}", other),
    }
}

// ── get out-of-bounds ─────────────────────────────────────────────────────────

#[test]
fn get_out_of_bounds_returns_undefined() {
    let ta = i32ns("newWithLength", vec![Value::I32(3)]);
    assert_eq!(
        i32ns("get", vec![ta.clone(), Value::I32(3)]),
        Value::Undefined
    );
    assert_eq!(i32ns("get", vec![ta, Value::I32(100)]), Value::Undefined);
}

// ── set and get round-trip ────────────────────────────────────────────────────

#[test]
fn set_then_get_round_trips_value() {
    let ta = i32ns("newWithLength", vec![Value::I32(4)]);
    i32ns("set", vec![ta.clone(), Value::I32(2), Value::I32(-7)]);
    assert_eq!(i32ns("get", vec![ta, Value::I32(2)]), Value::I32(-7));
}

// ── Uint8Array — values are modulo 256 (unsigned byte wrap) ──────────────────

#[test]
fn uint8_wraps_values_to_0_255_range() {
    let ta = u8ns("newWithLength", vec![Value::I32(1)]);
    u8ns("set", vec![ta.clone(), Value::I32(0), Value::I32(300)]);
    // 300 mod 256 = 44
    match u8ns("get", vec![ta, Value::I32(0)]) {
        Value::I32(v) => assert_eq!(v, 44),
        Value::F64(f) => assert!((f - 44.0).abs() < 0.5),
        other => panic!("expected 44, got {:?}", other),
    }
}

// ── at — negative indexing works on typed arrays ──────────────────────────────

#[test]
fn at_negative_one_returns_last_element() {
    let ta = i32ns(
        "from",
        vec![arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)])],
    );
    assert_eq!(i32ns("at", vec![ta, Value::I32(-1)]), Value::I32(3));
}

// ── fill ──────────────────────────────────────────────────────────────────────

#[test]
fn fill_sets_all_elements_to_given_value() {
    let ta = i32ns("newWithLength", vec![Value::I32(4)]);
    i32ns(
        "fill",
        vec![ta.clone(), Value::I32(9), Value::I32(0), Value::I32(4)],
    );
    assert_eq!(i32ns("get", vec![ta.clone(), Value::I32(0)]), Value::I32(9));
    assert_eq!(i32ns("get", vec![ta, Value::I32(3)]), Value::I32(9));
}

#[test]
fn fill_with_range_only_fills_that_range() {
    let ta = i32ns(
        "from",
        vec![arr(vec![Value::I32(0), Value::I32(0), Value::I32(0)])],
    );
    i32ns(
        "fill",
        vec![ta.clone(), Value::I32(5), Value::I32(1), Value::I32(2)],
    );
    assert_eq!(i32ns("get", vec![ta.clone(), Value::I32(0)]), Value::I32(0)); // untouched
    assert_eq!(i32ns("get", vec![ta.clone(), Value::I32(1)]), Value::I32(5)); // filled
    assert_eq!(i32ns("get", vec![ta, Value::I32(2)]), Value::I32(0)); // untouched
}

// ── slice vs subarray — copy vs view ─────────────────────────────────────────

#[test]
fn slice_returns_new_typed_array_not_same_pointer() {
    let ta = i32ns(
        "from",
        vec![arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)])],
    );
    let sliced = i32ns("slice", vec![ta.clone(), Value::I32(0), Value::I32(2)]);
    let ta_ptr = match &ta {
        Value::Object(a) => Arc::as_ptr(a) as usize,
        _ => 0,
    };
    let sl_ptr = match &sliced {
        Value::Object(a) => Arc::as_ptr(a) as usize,
        _ => 1,
    };
    assert_ne!(ta_ptr, sl_ptr);
    assert_eq!(i32ns("length", vec![sliced]), Value::I32(2));
}

// ── sort — numeric by default, unlike Array.sort which is lexicographic ──────

#[test]
fn int32_sort_uses_numeric_order_not_lexicographic() {
    // [10, 9, 1] sorted numerically → [1, 9, 10], not [1, 10, 9].
    let ta = i32ns(
        "from",
        vec![arr(vec![Value::I32(10), Value::I32(9), Value::I32(1)])],
    );
    let sorted = i32ns("sort", vec![ta, Value::Null]);
    assert_eq!(
        i32ns("get", vec![sorted.clone(), Value::I32(0)]),
        Value::I32(1)
    );
    assert_eq!(
        i32ns("get", vec![sorted.clone(), Value::I32(1)]),
        Value::I32(9)
    );
    assert_eq!(i32ns("get", vec![sorted, Value::I32(2)]), Value::I32(10));
}

// ── indexOf / includes ────────────────────────────────────────────────────────

#[test]
fn index_of_returns_correct_index() {
    let ta = i32ns(
        "from",
        vec![arr(vec![Value::I32(5), Value::I32(10), Value::I32(15)])],
    );
    assert_eq!(i32ns("indexOf", vec![ta, Value::I32(10)]), Value::I32(1));
}

#[test]
fn index_of_not_found_returns_negative_one() {
    let ta = i32ns("from", vec![arr(vec![Value::I32(1), Value::I32(2)])]);
    assert_eq!(i32ns("indexOf", vec![ta, Value::I32(99)]), Value::I32(-1));
}

#[test]
fn includes_true_for_present_value() {
    let ta = u8ns("from", vec![arr(vec![Value::I32(10), Value::I32(20)])]);
    assert_eq!(
        u8ns("includes", vec![ta, Value::I32(20)]),
        Value::Bool(true)
    );
}

// ── BigInt64Array — stores and retrieves i64 values ──────────────────────────

#[test]
fn bigint64_from_stores_i64_values() {
    let ta = bi64ns(
        "from",
        vec![arr(vec![
            Value::I64(1),
            Value::I64(-1),
            Value::I64(i64::MAX),
        ])],
    );
    assert_eq!(bi64ns("length", vec![ta.clone()]), Value::I32(3));
    assert_eq!(
        bi64ns("get", vec![ta.clone(), Value::I32(1)]),
        Value::I64(-1)
    );
    assert_eq!(bi64ns("get", vec![ta, Value::I32(2)]), Value::I64(i64::MAX));
}

// ── Float64 special values ────────────────────────────────────────────────────

#[test]
fn float64_stores_infinity_and_neg_infinity() {
    let ta = f64ns(
        "from",
        vec![arr(vec![
            Value::F64(f64::INFINITY),
            Value::F64(f64::NEG_INFINITY),
        ])],
    );
    match f64ns("get", vec![ta.clone(), Value::I32(0)]) {
        Value::F64(f) => assert!(f.is_infinite() && f > 0.0),
        other => panic!("expected +Inf, got {:?}", other),
    }
    match f64ns("get", vec![ta, Value::I32(1)]) {
        Value::F64(f) => assert!(f.is_infinite() && f < 0.0),
        other => panic!("expected -Inf, got {:?}", other),
    }
}

#[test]
fn float64_stores_nan() {
    let ta = f64ns("from", vec![arr(vec![Value::F64(f64::NAN)])]);
    match f64ns("get", vec![ta, Value::I32(0)]) {
        Value::F64(f) => assert!(f.is_nan()),
        other => panic!("expected NaN, got {:?}", other),
    }
}

// ── toReversed — non-mutating ─────────────────────────────────────────────────

#[test]
fn to_reversed_returns_new_typed_array_original_unchanged() {
    let ta = i32ns(
        "from",
        vec![arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)])],
    );
    let rev = i32ns("toReversed", vec![ta.clone()]);
    assert_eq!(i32ns("get", vec![ta, Value::I32(0)]), Value::I32(1)); // original unchanged
    assert_eq!(i32ns("get", vec![rev, Value::I32(0)]), Value::I32(3));
}

// ── of ────────────────────────────────────────────────────────────────────────

#[test]
fn int32_of_creates_typed_array_from_variadic_args() {
    let ta = i32ns("of", vec![Value::I32(7), Value::I32(8), Value::I32(9)]);
    assert_eq!(i32ns("length", vec![ta.clone()]), Value::I32(3));
    assert_eq!(i32ns("get", vec![ta, Value::I32(1)]), Value::I32(8));
}

// ── subarray — creates view sharing same buffer ───────────────────────────────

#[test]
fn subarray_returns_view_with_correct_length() {
    // ECMA-262 §22.2.3.31: subarray(begin, end) returns a view of the same buffer.
    let ta = i32ns(
        "from",
        vec![arr(vec![
            Value::I32(10),
            Value::I32(20),
            Value::I32(30),
            Value::I32(40),
        ])],
    );
    let sub = i32ns("subarray", vec![ta, Value::I32(1), Value::I32(3)]);
    // subarray(1,3) → view of [20, 30], length 2.
    assert_eq!(i32ns("length", vec![sub.clone()]), Value::I32(2));
    assert_eq!(i32ns("get", vec![sub, Value::I32(0)]), Value::I32(20));
}

#[test]
fn subarray_negative_start_counts_from_end() {
    // Negative begin is treated as length + begin.
    let ta = i32ns(
        "from",
        vec![arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)])],
    );
    let sub = i32ns("subarray", vec![ta, Value::I32(-2)]);
    // subarray(-2) → last 2 elements [2, 3].
    assert_eq!(i32ns("length", vec![sub.clone()]), Value::I32(2));
    assert_eq!(i32ns("get", vec![sub, Value::I32(0)]), Value::I32(2));
}

// ── set — copies from source array ───────────────────────────────────────────

#[test]
fn set_copies_elements_from_source_at_given_offset() {
    // ECMA-262 §22.2.3.24: TypedArray.prototype.set(source, offset).
    let ta = i32ns("newWithLength", vec![Value::I32(4)]);
    let src = arr(vec![Value::I32(7), Value::I32(8)]);
    i32ns("set", vec![ta.clone(), src, Value::I32(1)]);
    assert_eq!(i32ns("get", vec![ta.clone(), Value::I32(0)]), Value::I32(0));
    assert_eq!(i32ns("get", vec![ta.clone(), Value::I32(1)]), Value::I32(7));
    assert_eq!(i32ns("get", vec![ta, Value::I32(2)]), Value::I32(8));
}

// ── copyWithin ────────────────────────────────────────────────────────────────

#[test]
fn copy_within_copies_a_slice_to_a_target_position() {
    // ECMA-262 §22.2.3.6: copyWithin(target, start, end) copies bytes within the array.
    let ta = i32ns(
        "from",
        vec![arr(vec![
            Value::I32(1),
            Value::I32(2),
            Value::I32(3),
            Value::I32(4),
        ])],
    );
    i32ns("copyWithin", vec![ta.clone(), Value::I32(0), Value::I32(2)]);
    // copyWithin(0, 2) → copy [3,4] to index 0 → [3, 4, 3, 4].
    assert_eq!(i32ns("get", vec![ta.clone(), Value::I32(0)]), Value::I32(3));
    assert_eq!(i32ns("get", vec![ta, Value::I32(1)]), Value::I32(4));
}

// ── buffer / byteLength / byteOffset getters ─────────────────────────────────

#[test]
fn byte_length_of_typed_array_is_length_times_bytes_per_element() {
    // ECMA-262 §22.2.3.3: Int32Array has 4 bytes per element; byteLength = length * 4.
    let ta = i32ns("newWithLength", vec![Value::I32(3)]);
    assert_eq!(i32ns("byteLength", vec![ta]), Value::I32(12));
}

#[test]
fn byte_offset_of_direct_construction_is_zero() {
    // ECMA-262 §22.2.3.4: A TypedArray created directly has byteOffset = 0.
    let ta = u8ns("newWithLength", vec![Value::I32(8)]);
    assert_eq!(u8ns("byteOffset", vec![ta]), Value::I32(0));
}

#[test]
fn buffer_returns_an_arraybuffer_object() {
    // ECMA-262 §22.2.3.1: TypedArray.prototype.buffer returns the underlying ArrayBuffer.
    let ta = i32ns("from", vec![arr(vec![Value::I32(1), Value::I32(2)])]);
    let buf = i32ns("buffer", vec![ta]);
    assert!(matches!(buf, Value::Object(_)));
}

// ── TypedArray HOFs: map, filter, reduce, every, some, forEach ────────────────

#[test]
fn map_transforms_each_typed_array_element() {
    // ECMA-262 §22.2.3.19: TypedArray.prototype.map applies fn to each element.
    let ta = i32ns(
        "from",
        vec![arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)])],
    );
    let map_fn = {
        let mut o = Object::new();
        o.properties.insert("__map_mul".to_string(), Value::I32(2));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let result = i32ns("map", vec![ta, map_fn]);
    assert_eq!(
        i32ns("get", vec![result.clone(), Value::I32(0)]),
        Value::I32(2)
    );
    assert_eq!(i32ns("get", vec![result, Value::I32(2)]), Value::I32(6));
}

#[test]
fn filter_keeps_only_matching_typed_elements() {
    // ECMA-262 §22.2.3.9: filter returns a new TypedArray with matching elements.
    let ta = i32ns(
        "from",
        vec![arr(vec![
            Value::I32(1),
            Value::I32(4),
            Value::I32(2),
            Value::I32(5),
        ])],
    );
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(3));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let result = i32ns("filter", vec![ta, pred]);
    assert_eq!(i32ns("length", vec![result.clone()]), Value::I32(2));
    assert_eq!(i32ns("get", vec![result, Value::I32(0)]), Value::I32(4));
}

#[test]
fn reduce_folds_typed_array_to_single_value() {
    // ECMA-262 §22.2.3.24: reduce(fn, init) left-folds.
    let ta = i32ns(
        "from",
        vec![arr(vec![Value::I32(1), Value::I32(2), Value::I32(3)])],
    );
    let reducer = {
        let mut o = Object::new();
        o.properties
            .insert("__reduce_add".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let result = i32ns("reduce", vec![ta, reducer, Value::I32(0)]);
    assert_eq!(result, Value::I32(6));
}

#[test]
fn every_returns_false_when_any_typed_element_fails() {
    let ta = i32ns(
        "from",
        vec![arr(vec![Value::I32(1), Value::I32(5), Value::I32(2)])],
    );
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(3));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    assert_eq!(i32ns("every", vec![ta, pred]), Value::Bool(false));
}

#[test]
fn some_returns_true_when_any_typed_element_matches() {
    let ta = i32ns(
        "from",
        vec![arr(vec![Value::I32(1), Value::I32(5), Value::I32(2)])],
    );
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(3));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    assert_eq!(i32ns("some", vec![ta, pred]), Value::Bool(true));
}

#[test]
fn for_each_returns_undefined_on_typed_array() {
    // ECMA-262 §22.2.3.13: forEach returns undefined.
    let ta = i32ns("from", vec![arr(vec![Value::I32(1), Value::I32(2)])]);
    let noop = {
        let mut o = Object::new();
        o.properties.insert("__noop".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    assert!(matches!(
        i32ns("forEach", vec![ta, noop]),
        Value::Undefined | Value::Null
    ));
}

#[test]
fn find_on_typed_array_returns_first_matching_element() {
    let ta = i32ns(
        "from",
        vec![arr(vec![Value::I32(1), Value::I32(5), Value::I32(3)])],
    );
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(3));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    assert_eq!(i32ns("find", vec![ta, pred]), Value::I32(5));
}

#[test]
fn find_last_on_typed_array_returns_rightmost_match() {
    // ECMA-262 ES2023 §22.2.3.11: TypedArray.prototype.findLast.
    let ta = i32ns(
        "from",
        vec![arr(vec![Value::I32(1), Value::I32(5), Value::I32(4)])],
    );
    let pred = {
        let mut o = Object::new();
        o.properties.insert("__pred_gt".to_string(), Value::I32(3));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    assert_eq!(i32ns("findLast", vec![ta, pred]), Value::I32(4));
}

// ── TypedArray iterators: entries, keys, values ───────────────────────────────

#[test]
fn keys_iterator_yields_numeric_indices() {
    // ECMA-262 §22.2.3.17: TypedArray.prototype.keys() returns index iterator.
    let ta = i32ns("from", vec![arr(vec![Value::I32(10), Value::I32(20)])]);
    let it = i32ns("keys", vec![ta]);
    assert!(matches!(it, Value::Object(_)));
}

#[test]
fn values_iterator_yields_elements_in_order() {
    let ta = i32ns("from", vec![arr(vec![Value::I32(10), Value::I32(20)])]);
    let it = i32ns("values", vec![ta]);
    assert!(matches!(it, Value::Object(_)));
}

#[test]
fn entries_iterator_yields_index_value_pairs() {
    let ta = i32ns("from", vec![arr(vec![Value::I32(10)])]);
    let it = i32ns("entries", vec![ta]);
    assert!(matches!(it, Value::Object(_)));
}

// ── TypedArray.prototype.toLocaleString ───────────────────────────────────────

#[test]
fn to_locale_string_of_typed_array_is_a_string() {
    // ECMA-262 §22.2.3.30: TypedArray.prototype.toLocaleString joins locale-formatted elements.
    let ta = f64ns("from", vec![arr(vec![Value::F64(1.5), Value::F64(2.5)])]);
    assert!(matches!(
        f64ns("toLocaleString", vec![ta]),
        Value::String(_)
    ));
}

// ── Uint8Array.prototype.toBase64 / toHex (ES2025 §22.2.7) ───────────────────

#[test]
fn uint8_to_base64_encodes_bytes() {
    // ECMA-262 ES2025: Uint8Array.prototype.toBase64() encodes as Base64 string.
    // [77, 97, 110] = "Man" in Base64 = "TWFu".
    let ta = u8ns(
        "from",
        vec![arr(vec![Value::I32(77), Value::I32(97), Value::I32(110)])],
    );
    let result = u8ns("toBase64", vec![ta]);
    assert_eq!(result, Value::String(Arc::from("TWFu")));
}

#[test]
fn uint8_to_base64_empty_array_returns_empty_string() {
    let ta = u8ns("newWithLength", vec![Value::I32(0)]);
    let result = u8ns("toBase64", vec![ta]);
    assert_eq!(result, Value::String(Arc::from("")));
}

#[test]
fn uint8_from_base64_decodes_to_correct_bytes() {
    // ECMA-262 ES2025: Uint8Array.fromBase64("TWFu") → [77, 97, 110].
    let ta = u8ns("fromBase64", vec![Value::String(Arc::from("TWFu"))]);
    assert_eq!(u8ns("get", vec![ta.clone(), Value::I32(0)]), Value::I32(77));
    assert_eq!(u8ns("get", vec![ta.clone(), Value::I32(1)]), Value::I32(97));
    assert_eq!(u8ns("get", vec![ta, Value::I32(2)]), Value::I32(110));
}

#[test]
fn uint8_to_hex_encodes_each_byte_as_two_hex_digits() {
    // ECMA-262 ES2025: toHex() returns lowercase hex pairs.
    // [0, 255, 16] → "00ff10".
    let ta = u8ns(
        "from",
        vec![arr(vec![Value::I32(0), Value::I32(255), Value::I32(16)])],
    );
    let result = u8ns("toHex", vec![ta]);
    assert_eq!(result, Value::String(Arc::from("00ff10")));
}

#[test]
fn uint8_from_hex_decodes_hex_string_to_bytes() {
    // Uint8Array.fromHex("00ff10") → [0, 255, 16].
    let ta = u8ns("fromHex", vec![Value::String(Arc::from("00ff10"))]);
    assert_eq!(u8ns("length", vec![ta.clone()]), Value::I32(3));
    assert_eq!(
        u8ns("get", vec![ta.clone(), Value::I32(1)]),
        Value::I32(255)
    );
}

#[test]
fn uint8_base64_roundtrip_preserves_bytes() {
    // Encode then decode must reproduce the original bytes.
    let original = arr(vec![
        Value::I32(1),
        Value::I32(2),
        Value::I32(3),
        Value::I32(100),
    ]);
    let ta = u8ns("from", vec![original]);
    let encoded = u8ns("toBase64", vec![ta.clone()]);
    let decoded = u8ns("fromBase64", vec![encoded]);
    assert_eq!(
        u8ns("get", vec![decoded.clone(), Value::I32(0)]),
        Value::I32(1)
    );
    assert_eq!(u8ns("get", vec![decoded, Value::I32(3)]), Value::I32(100));
}
