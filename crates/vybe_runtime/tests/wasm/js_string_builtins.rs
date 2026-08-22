//! Tests for the js-string-builtins WASM proposal (merged; V8 native).
//! Spec: `proposals/js-string-builtins/proposals/js-string-builtins/Overview.md`
//!
//! Also covers the wasm:js-string extensions from the Stage-1
//! js-primitive-builtins proposal: fromI32, fromU32, fromI64, fromU64, fromF64.

use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};

/// Unique names for test-argument globals, so reused VMs never collide.
static TEST_GLOBAL_SEQ: AtomicUsize = AtomicUsize::new(0);

/// Emit a spec const for `v` when a const emitter exists; otherwise route the
/// value through a VM global (`global.get` of a unique imported name) and
/// return the pending `(name, value)` binding for the VM that runs the chunk.
fn push_arg(chunk: &mut Chunk, v: Value) -> Option<(String, Value)> {
    match v {
        Value::I32(n) => chunk.emit_i32_const(n, 0),
        Value::I64(n) => chunk.emit_i64_const(n, 0),
        Value::F32(f) => chunk.emit_f32_const(f, 0),
        Value::F64(f) => chunk.emit_f64_const(f, 0),
        Value::Bool(b) => chunk.emit_bool_const(b, 0),
        Value::String(s) => chunk.emit_string_const(&s, 0),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
        other => {
            let name = format!(
                "__test_arg_{}",
                TEST_GLOBAL_SEQ.fetch_add(1, Ordering::Relaxed)
            );
            let ci = chunk.intern_string_constant(&name);
            chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
            return Some((name, other));
        }
    }
    None
}

fn call(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    let mut pending = Vec::new();
    for v in args {
        pending.extend(push_arg(&mut chunk, v));
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    vybe_runtime::js_builtins::register(&mut vm);
    for (name, value) in pending {
        vm.set_global_owned(name, value);
    }
    vm.run(vec![chunk]).expect("VM run failed")
}

fn expect_trap(module: &str, name: &str, args: Vec<Value>) {
    let mut chunk = Chunk::new("<test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    let mut pending = Vec::new();
    for v in args {
        pending.extend(push_arg(&mut chunk, v));
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    vybe_runtime::js_builtins::register(&mut vm);
    for (name, value) in pending {
        vm.set_global_owned(name, value);
    }
    assert!(vm.run(vec![chunk]).is_err(), "{module}.{name} should trap");
}

fn str(s: &str) -> Value {
    Value::String(Arc::from(s))
}

fn new_i16_array(values: &[i16]) -> Value {
    let elems: Vec<Value> = values.iter().map(|&v| Value::I32(v as i32)).collect();
    Value::Object(Arc::new(Mutex::new(Object::new_array(elems))))
}

fn read_i16_array(v: &Value) -> Vec<i16> {
    match v {
        Value::Object(o) => {
            let obj = o.lock().unwrap();
            match &obj.kind {
                ObjectKind::Array(elems) => elems.iter().map(|e| e.as_i32() as i16).collect(),
                _ => vec![],
            }
        }
        _ => vec![],
    }
}

const MOD: &str = "wasm:js-string";

// ── test ─────────────────────────────────────────────────────────────

#[test]
fn test_returns_1_for_string() {
    assert_eq!(call(MOD, "test", vec![str("hello")]).as_i32(), 1);
    assert_eq!(call(MOD, "test", vec![str("")]).as_i32(), 1);
}

#[test]
fn test_returns_0_for_non_string() {
    assert_eq!(call(MOD, "test", vec![Value::Null]).as_i32(), 0);
    assert_eq!(call(MOD, "test", vec![Value::I32(0)]).as_i32(), 0);
    assert_eq!(call(MOD, "test", vec![Value::Bool(true)]).as_i32(), 0);
}

// ── cast ─────────────────────────────────────────────────────────────

#[test]
fn cast_returns_string_unchanged() {
    let v = call(MOD, "cast", vec![str("hi")]);
    assert_eq!(format!("{}", v), "hi");
}

#[test]
fn cast_traps_on_null() {
    expect_trap(MOD, "cast", vec![Value::Null]);
}

#[test]
fn cast_traps_on_non_string() {
    expect_trap(MOD, "cast", vec![Value::I32(1)]);
}

// ── length ───────────────────────────────────────────────────────────

#[test]
fn length_returns_utf16_code_unit_count() {
    assert_eq!(call(MOD, "length", vec![str("hello")]).as_i32(), 5);
    assert_eq!(call(MOD, "length", vec![str("")]).as_i32(), 0);
    // "é" is U+00E9, 1 UTF-16 unit
    assert_eq!(call(MOD, "length", vec![str("é")]).as_i32(), 1);
    // U+1F600 (emoji) is a surrogate pair — 2 UTF-16 units
    assert_eq!(call(MOD, "length", vec![str("\u{1F600}")]).as_i32(), 2);
}

#[test]
fn length_traps_on_null() {
    expect_trap(MOD, "length", vec![Value::Null]);
}

// ── concat ───────────────────────────────────────────────────────────

#[test]
fn concat_joins_strings() {
    assert_eq!(
        format!("{}", call(MOD, "concat", vec![str("foo"), str("bar")])),
        "foobar"
    );
    assert_eq!(
        format!("{}", call(MOD, "concat", vec![str(""), str("x")])),
        "x"
    );
}

#[test]
fn concat_traps_on_non_string() {
    expect_trap(MOD, "concat", vec![Value::Null, str("x")]);
    expect_trap(MOD, "concat", vec![str("x"), Value::Null]);
}

// ── substring ────────────────────────────────────────────────────────

#[test]
fn substring_basic_slice() {
    assert_eq!(
        format!(
            "{}",
            call(
                MOD,
                "substring",
                vec![str("hello"), Value::I32(1), Value::I32(4)]
            )
        ),
        "ell"
    );
}

#[test]
fn substring_clamps_to_length() {
    assert_eq!(
        format!(
            "{}",
            call(
                MOD,
                "substring",
                vec![str("hi"), Value::I32(0), Value::I32(100)]
            )
        ),
        "hi"
    );
}

#[test]
fn substring_swaps_if_start_greater_than_end() {
    // JS substring swaps start/end if start > end
    assert_eq!(
        format!(
            "{}",
            call(
                MOD,
                "substring",
                vec![str("hello"), Value::I32(4), Value::I32(1)]
            )
        ),
        "ell"
    );
}

#[test]
fn substring_treats_indices_as_u32() {
    // Negative i32 treated as u32 (large number) → clamped to length
    assert_eq!(
        format!(
            "{}",
            call(
                MOD,
                "substring",
                vec![str("hi"), Value::I32(0), Value::I32(-1_i32)]
            )
        ),
        "hi"
    );
}

#[test]
fn substring_works_on_utf16_units() {
    // U+1F600 is 2 UTF-16 units; substring(0, 2) should return the full emoji
    let emoji = "\u{1F600}";
    let result = call(
        MOD,
        "substring",
        vec![str(emoji), Value::I32(0), Value::I32(2)],
    );
    assert_eq!(format!("{}", result), emoji);
}

#[test]
fn substring_traps_on_non_string() {
    expect_trap(
        MOD,
        "substring",
        vec![Value::Null, Value::I32(0), Value::I32(1)],
    );
}

// ── equals ───────────────────────────────────────────────────────────

#[test]
fn equals_same_string_is_1() {
    assert_eq!(
        call(MOD, "equals", vec![str("hello"), str("hello")]).as_i32(),
        1
    );
}

#[test]
fn equals_different_strings_is_0() {
    assert_eq!(call(MOD, "equals", vec![str("a"), str("b")]).as_i32(), 0);
}

#[test]
fn equals_null_null_is_1() {
    assert_eq!(
        call(MOD, "equals", vec![Value::Null, Value::Null]).as_i32(),
        1
    );
}

#[test]
fn equals_string_and_null_is_0() {
    assert_eq!(call(MOD, "equals", vec![str("x"), Value::Null]).as_i32(), 0);
}

#[test]
fn equals_traps_on_non_string_non_null() {
    expect_trap(MOD, "equals", vec![Value::I32(1), str("x")]);
}

// ── compare ──────────────────────────────────────────────────────────

#[test]
fn compare_returns_0_for_equal() {
    assert_eq!(
        call(MOD, "compare", vec![str("abc"), str("abc")]).as_i32(),
        0
    );
}

#[test]
fn compare_returns_negative_for_less() {
    assert!(call(MOD, "compare", vec![str("a"), str("b")]).as_i32() < 0);
}

#[test]
fn compare_returns_positive_for_greater() {
    assert!(call(MOD, "compare", vec![str("b"), str("a")]).as_i32() > 0);
}

#[test]
fn compare_traps_on_null() {
    expect_trap(MOD, "compare", vec![Value::Null, str("x")]);
    expect_trap(MOD, "compare", vec![str("x"), Value::Null]);
}

// ── charCodeAt ───────────────────────────────────────────────────────

#[test]
fn char_code_at_ascii() {
    assert_eq!(
        call(MOD, "charCodeAt", vec![str("ABC"), Value::I32(0)]).as_i32(),
        65
    ); // 'A'
    assert_eq!(
        call(MOD, "charCodeAt", vec![str("ABC"), Value::I32(2)]).as_i32(),
        67
    ); // 'C'
}

#[test]
fn char_code_at_returns_utf16_unit() {
    // U+1F600 encodes as surrogate pair: 0xD83D, 0xDE00
    let emoji = "\u{1F600}";
    assert_eq!(
        call(MOD, "charCodeAt", vec![str(emoji), Value::I32(0)]).as_i32(),
        0xD83D
    );
    assert_eq!(
        call(MOD, "charCodeAt", vec![str(emoji), Value::I32(1)]).as_i32(),
        0xDE00
    );
}

#[test]
fn char_code_at_traps_out_of_bounds() {
    expect_trap(MOD, "charCodeAt", vec![str("hi"), Value::I32(5)]);
}

#[test]
fn char_code_at_traps_on_non_string() {
    expect_trap(MOD, "charCodeAt", vec![Value::Null, Value::I32(0)]);
}

// ── codePointAt ──────────────────────────────────────────────────────

#[test]
fn code_point_at_ascii() {
    assert_eq!(
        call(MOD, "codePointAt", vec![str("A"), Value::I32(0)]).as_i32(),
        65
    );
}

#[test]
fn code_point_at_returns_full_supplementary_codepoint() {
    // U+1F600 at index 0 — should return 0x1F600, not the surrogate
    let emoji = "\u{1F600}";
    assert_eq!(
        call(MOD, "codePointAt", vec![str(emoji), Value::I32(0)]).as_i32(),
        0x1F600
    );
}

#[test]
fn code_point_at_lone_surrogate_at_index_1() {
    // Index 1 is the low surrogate — returned as-is (no pair to complete)
    let emoji = "\u{1F600}";
    assert_eq!(
        call(MOD, "codePointAt", vec![str(emoji), Value::I32(1)]).as_i32(),
        0xDE00
    );
}

#[test]
fn code_point_at_traps_out_of_bounds() {
    expect_trap(MOD, "codePointAt", vec![str("hi"), Value::I32(99)]);
}

// ── fromCharCode ─────────────────────────────────────────────────────

#[test]
fn from_char_code_ascii() {
    assert_eq!(
        format!("{}", call(MOD, "fromCharCode", vec![Value::I32(65)])),
        "A"
    );
}

#[test]
fn from_char_code_treats_as_u16() {
    // 0xE9 = 233 = 'é'
    assert_eq!(
        format!("{}", call(MOD, "fromCharCode", vec![Value::I32(0xE9)])),
        "é"
    );
}

#[test]
fn from_char_code_wraps_negative_to_u16() {
    // -1 as u16 = 0xFFFF
    let result = call(MOD, "fromCharCode", vec![Value::I32(-1)]);
    let s = format!("{}", result);
    let units: Vec<u16> = s.encode_utf16().collect();
    assert_eq!(units, &[0xFFFF]);
}

// ── fromCodePoint ────────────────────────────────────────────────────

#[test]
fn from_code_point_ascii() {
    assert_eq!(
        format!("{}", call(MOD, "fromCodePoint", vec![Value::I32(65)])),
        "A"
    );
}

#[test]
fn from_code_point_supplementary() {
    assert_eq!(
        format!("{}", call(MOD, "fromCodePoint", vec![Value::I32(0x1F600)])),
        "\u{1F600}"
    );
}

#[test]
fn from_code_point_traps_above_max() {
    expect_trap(MOD, "fromCodePoint", vec![Value::I32(0x110000)]);
}

// ── fromCharCodeArray ────────────────────────────────────────────────

#[test]
fn from_char_code_array_basic() {
    let arr = new_i16_array(&[72, 101, 108, 108, 111]); // "Hello"
    assert_eq!(
        format!(
            "{}",
            call(
                MOD,
                "fromCharCodeArray",
                vec![arr, Value::I32(0), Value::I32(5)]
            )
        ),
        "Hello"
    );
}

#[test]
fn from_char_code_array_sub_range() {
    let arr = new_i16_array(&[65, 66, 67, 68]); // "ABCD"
    assert_eq!(
        format!(
            "{}",
            call(
                MOD,
                "fromCharCodeArray",
                vec![arr, Value::I32(1), Value::I32(3)]
            )
        ),
        "BC"
    );
}

#[test]
fn from_char_code_array_traps_on_invalid_range() {
    let arr = new_i16_array(&[65, 66]);
    expect_trap(
        MOD,
        "fromCharCodeArray",
        vec![arr.clone(), Value::I32(0), Value::I32(10)],
    );
}

// ── intoCharCodeArray ────────────────────────────────────────────────

#[test]
fn into_char_code_array_writes_and_returns_count() {
    let arr = new_i16_array(&[0, 0, 0, 0, 0]);
    let count = call(
        MOD,
        "intoCharCodeArray",
        vec![str("Hello"), arr.clone(), Value::I32(0)],
    )
    .as_i32();
    assert_eq!(count, 5);
    assert_eq!(read_i16_array(&arr), vec![72, 101, 108, 108, 111]);
}

#[test]
fn into_char_code_array_writes_at_offset() {
    let arr = new_i16_array(&[0, 0, 0, 0, 0]);
    call(
        MOD,
        "intoCharCodeArray",
        vec![str("Hi"), arr.clone(), Value::I32(2)],
    );
    assert_eq!(read_i16_array(&arr)[2..4], [72i16, 105]);
}

#[test]
fn into_char_code_array_traps_if_no_fit() {
    let arr = new_i16_array(&[0, 0]);
    expect_trap(
        MOD,
        "intoCharCodeArray",
        vec![str("Hello"), arr, Value::I32(0)],
    );
}

// ── fromI32 / fromU32 / fromI64 / fromU64 / fromF64 (js-primitive-builtins extensions) ──

#[test]
fn from_i32_signed_decimal() {
    assert_eq!(
        format!("{}", call(MOD, "fromI32", vec![Value::I32(-7)])),
        "-7"
    );
    assert_eq!(
        format!("{}", call(MOD, "fromI32", vec![Value::I32(0)])),
        "0"
    );
}

#[test]
fn from_u32_unsigned_decimal() {
    assert_eq!(
        format!("{}", call(MOD, "fromU32", vec![Value::I32(-1)])),
        "4294967295"
    );
    assert_eq!(
        format!("{}", call(MOD, "fromU32", vec![Value::I32(0)])),
        "0"
    );
}

#[test]
fn from_i64_signed_decimal() {
    assert_eq!(
        format!("{}", call(MOD, "fromI64", vec![Value::I64(-1)])),
        "-1"
    );
    assert_eq!(
        format!("{}", call(MOD, "fromI64", vec![Value::I64(i64::MAX)])),
        "9223372036854775807"
    );
}

#[test]
fn from_u64_unsigned_decimal() {
    assert_eq!(
        format!("{}", call(MOD, "fromU64", vec![Value::I64(-1_i64)])),
        "18446744073709551615"
    );
}

#[test]
fn from_f64_finite() {
    assert_eq!(
        format!("{}", call(MOD, "fromF64", vec![Value::F64(3.14)])),
        "3.14"
    );
    assert_eq!(
        format!("{}", call(MOD, "fromF64", vec![Value::F64(0.0)])),
        "0"
    );
}

#[test]
fn from_f64_special_values() {
    assert_eq!(
        format!("{}", call(MOD, "fromF64", vec![Value::F64(f64::NAN)])),
        "NaN"
    );
    assert_eq!(
        format!("{}", call(MOD, "fromF64", vec![Value::F64(f64::INFINITY)])),
        "Infinity"
    );
    assert_eq!(
        format!(
            "{}",
            call(MOD, "fromF64", vec![Value::F64(f64::NEG_INFINITY)])
        ),
        "-Infinity"
    );
}
