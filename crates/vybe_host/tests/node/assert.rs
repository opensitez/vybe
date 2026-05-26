//! Behaviour tests for `node:assert` host imports.
//!
//! Reference: <https://nodejs.org/api/assert.html>.
//!
//! Coverage:
//!   - `ok(value[, message])` — truthy check
//!   - `equal(actual, expected[, message])` — loose ==
//!   - `notEqual(actual, expected[, message])` — loose !=
//!   - `strictEqual(actual, expected[, message])` — strict ===
//!   - `notStrictEqual(actual, expected[, message])` — strict !==
//!   - `deepEqual(actual, expected[, message])` — deep loose
//!   - `notDeepEqual(actual, expected[, message])` — deep loose negative
//!   - `deepStrictEqual(actual, expected[, message])` — deep strict
//!   - `notDeepStrictEqual(actual, expected[, message])` — deep strict negative
//!   - `throws(fn[, error][, message])` — fn must throw
//!   - `doesNotThrow(fn[, error][, message])` — fn must not throw
//!   - `ifError(value)` — throws if value is truthy
//!   - `fail([message])` — always throws
//!   - `match(string, regexp[, message])` — string must match regex
//!   - `doesNotMatch(string, regexp[, message])` — string must not match
//!
//! Deferred:
//!   - `rejects`, `doesNotReject` — require promise infrastructure
//!   - `AssertionError` constructor — needs class infrastructure

use std::sync::Arc;
use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn call_assert(name: &str, args: Vec<Value>) -> Result<Value, String> {
    let mut chunk = Chunk::new("<node-assert-test>");
    let import_idx = chunk.add_import("node:assert", name);
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
    vm.run(vec![chunk]).map_err(|e| format!("{e}"))
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:assert"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

// ── ok ────────────────────────────────────────────────────────────────────────

#[test]
fn ok_truthy_value_passes() {
    assert!(call_assert("ok", vec![Value::Bool(true)]).is_ok());
}

#[test]
fn ok_non_zero_integer_passes() {
    assert!(call_assert("ok", vec![Value::I32(1)]).is_ok());
}

#[test]
fn ok_non_empty_string_passes() {
    assert!(call_assert("ok", vec![s("hello")]).is_ok());
}

#[test]
fn ok_false_throws_assertion_error() {
    assert!(call_assert("ok", vec![Value::Bool(false)]).is_err());
}

#[test]
fn ok_zero_throws() {
    assert!(call_assert("ok", vec![Value::I32(0)]).is_err());
}

#[test]
fn ok_null_throws() {
    assert!(call_assert("ok", vec![Value::Null]).is_err());
}

#[test]
fn ok_undefined_throws() {
    assert!(call_assert("ok", vec![Value::Undefined]).is_err());
}

#[test]
fn ok_empty_string_throws() {
    assert!(call_assert("ok", vec![s("")]).is_err());
}

// ── equal (loose) ─────────────────────────────────────────────────────────────

#[test]
fn equal_same_integers_passes() {
    assert!(call_assert("equal", vec![Value::I32(42), Value::I32(42)]).is_ok());
}

#[test]
fn equal_int_and_float_same_value_passes() {
    assert!(call_assert("equal", vec![Value::I32(1), Value::F64(1.0)]).is_ok());
}

#[test]
fn equal_number_and_string_same_digits_passes() {
    // Loose equality: 1 == "1"
    assert!(call_assert("equal", vec![Value::I32(1), s("1")]).is_ok());
}

#[test]
fn equal_different_values_throws() {
    assert!(call_assert("equal", vec![Value::I32(1), Value::I32(2)]).is_err());
}

// ── notEqual ──────────────────────────────────────────────────────────────────

#[test]
fn not_equal_different_values_passes() {
    assert!(call_assert("notEqual", vec![Value::I32(1), Value::I32(2)]).is_ok());
}

#[test]
fn not_equal_same_value_throws() {
    assert!(call_assert("notEqual", vec![Value::I32(1), Value::I32(1)]).is_err());
}

// ── strictEqual ───────────────────────────────────────────────────────────────

#[test]
fn strict_equal_same_type_and_value_passes() {
    assert!(call_assert("strictEqual", vec![Value::I32(5), Value::I32(5)]).is_ok());
}

#[test]
fn strict_equal_same_string_passes() {
    assert!(call_assert("strictEqual", vec![s("abc"), s("abc")]).is_ok());
}

#[test]
fn strict_equal_bool_true_true_passes() {
    assert!(call_assert("strictEqual", vec![Value::Bool(true), Value::Bool(true)]).is_ok());
}

#[test]
fn strict_equal_int_and_string_same_digits_throws() {
    // Strict: 1 !== "1"
    assert!(call_assert("strictEqual", vec![Value::I32(1), s("1")]).is_err());
}

#[test]
fn strict_equal_null_and_undefined_throws() {
    assert!(call_assert("strictEqual", vec![Value::Null, Value::Undefined]).is_err());
}

// ── notStrictEqual ────────────────────────────────────────────────────────────

#[test]
fn not_strict_equal_int_and_string_passes() {
    assert!(call_assert("notStrictEqual", vec![Value::I32(1), s("1")]).is_ok());
}

#[test]
fn not_strict_equal_same_value_same_type_throws() {
    assert!(call_assert("notStrictEqual", vec![Value::I32(1), Value::I32(1)]).is_err());
}

// ── deepEqual ─────────────────────────────────────────────────────────────────

#[test]
fn deep_equal_identical_primitives_passes() {
    assert!(call_assert("deepEqual", vec![Value::I32(42), Value::I32(42)]).is_ok());
}

#[test]
fn deep_equal_different_primitives_throws() {
    assert!(call_assert("deepEqual", vec![Value::I32(1), Value::I32(2)]).is_err());
}

// ── deepStrictEqual ───────────────────────────────────────────────────────────

#[test]
fn deep_strict_equal_same_int_passes() {
    assert!(call_assert("deepStrictEqual", vec![Value::I32(7), Value::I32(7)]).is_ok());
}

#[test]
fn deep_strict_equal_int_vs_float_different_type_throws() {
    // deepStrictEqual uses Object.is semantics for primitives — 1 !== 1.0 in strict mode
    assert!(call_assert("deepStrictEqual", vec![Value::I32(1), Value::F64(1.0)]).is_err());
}

// ── ifError ───────────────────────────────────────────────────────────────────

#[test]
fn if_error_null_does_not_throw() {
    assert!(call_assert("ifError", vec![Value::Null]).is_ok());
}

#[test]
fn if_error_undefined_does_not_throw() {
    assert!(call_assert("ifError", vec![Value::Undefined]).is_ok());
}

#[test]
fn if_error_truthy_value_throws() {
    assert!(call_assert("ifError", vec![s("some error")]).is_err());
}

// ── fail ──────────────────────────────────────────────────────────────────────

#[test]
fn fail_always_throws() {
    assert!(call_assert("fail", vec![]).is_err());
}

#[test]
fn fail_with_message_throws_with_that_message() {
    let result = call_assert("fail", vec![s("my failure")]);
    assert!(result.is_err());
    assert!(result.unwrap_err().contains("my failure"));
}

// ── match / doesNotMatch ──────────────────────────────────────────────────────

#[test]
fn match_passing_string_against_matching_pattern() {
    // "hello world" matches /hello/
    let result = call_assert("match", vec![s("hello world"), s("/hello/")]);
    assert!(result.is_ok());
}

#[test]
fn match_string_not_matching_pattern_throws() {
    let result = call_assert("match", vec![s("hello"), s("/^\\d+$/")]);
    assert!(result.is_err());
}

#[test]
fn does_not_match_non_matching_string_passes() {
    let result = call_assert("doesNotMatch", vec![s("hello"), s("/^\\d+$/")]);
    assert!(result.is_ok());
}

#[test]
fn does_not_match_matching_string_throws() {
    let result = call_assert("doesNotMatch", vec![s("hello"), s("/hello/")]);
    assert!(result.is_err());
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_node_assert_surface_is_registered() {
    let expected = [
        "ok",
        "equal",
        "notEqual",
        "strictEqual",
        "notStrictEqual",
        "deepEqual",
        "notDeepEqual",
        "deepStrictEqual",
        "notDeepStrictEqual",
        "ifError",
        "fail",
        "throws",
        "doesNotThrow",
        "match",
        "doesNotMatch",
        "rejects",
        "doesNotReject",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(missing.is_empty(), "missing node:assert imports: {missing:?}");
}
