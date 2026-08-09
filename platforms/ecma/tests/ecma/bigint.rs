//! Behaviour tests for `ecma:bigint` host imports.
//!
//! Reference: ECMA-262 §21.2 BigInt.
//!
//! Each test covers a distinct behaviour — particularly edge cases in
//! truncating division, two's-complement bitwise ops, and width clamping.

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::Value;
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
    let mut chunk = Chunk::new("<ecma-bigint-test>");
    let import_idx = chunk.add_import("ecma:bigint", name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("VM run failed")
}

fn bi(n: i64) -> Value {
    Value::bigint_i64(n)
}

// ── BigInt() coercion ─────────────────────────────────────────────────────────

#[test]
fn bigint_coerces_i32_to_i64() {
    // BigInt() constructor returns Value::BigInt (not I64) per §21.2.1.1
    assert_eq!(invoke("BigInt", vec![Value::I32(7)]), Value::bigint_i64(7));
}

#[test]
fn bigint_of_zero_stays_zero() {
    assert_eq!(invoke("BigInt", vec![Value::I32(0)]), Value::bigint_i64(0));
}

// ── Division truncates toward zero, not floor ────────────────────────────────

#[test]
fn div_negative_dividend_truncates_toward_zero_not_floor() {
    // -7n / 2n = -3n (truncate), not -4n (floor).
    assert_eq!(invoke("div", vec![bi(-7), bi(2)]), bi(-3));
}

#[test]
fn rem_sign_follows_dividend() {
    // -7n % 2n = -1n (sign of dividend), not 1n.
    assert_eq!(invoke("rem", vec![bi(-7), bi(2)]), bi(-1));
}

// ── Arithmetic identity laws ──────────────────────────────────────────────────

#[test]
fn add_zero_is_identity() {
    assert_eq!(invoke("add", vec![bi(42), bi(0)]), bi(42));
}

#[test]
fn mul_by_zero_yields_zero() {
    assert_eq!(invoke("mul", vec![bi(999), bi(0)]), bi(0));
}

#[test]
fn sub_self_is_zero() {
    assert_eq!(invoke("sub", vec![bi(12345), bi(12345)]), bi(0));
}

#[test]
fn neg_double_negation_is_identity() {
    assert_eq!(invoke("neg", vec![invoke("neg", vec![bi(42)])]), bi(42));
}

// ── Bitwise — two's complement semantics ─────────────────────────────────────

#[test]
fn not_zero_equals_negative_one() {
    // ~0n = -1n in two's complement.
    assert_eq!(invoke("not", vec![bi(0)]), bi(-1));
}

#[test]
fn not_negative_one_equals_zero() {
    assert_eq!(invoke("not", vec![bi(-1)]), bi(0));
}

#[test]
fn xor_with_self_always_zero() {
    assert_eq!(invoke("xor", vec![bi(0xDEAD), bi(0xDEAD)]), bi(0));
}

#[test]
fn shr_negative_is_arithmetic_sign_extending() {
    // -8n >> 1 = -4n (arithmetic shift, not logical).
    assert_eq!(invoke("shr", vec![bi(-8), bi(1)]), bi(-4));
}

#[test]
fn shl_one_left_four_equals_sixteen() {
    assert_eq!(invoke("shl", vec![bi(1), bi(4)]), bi(16));
}

// ── asIntN — signed width clamping ────────────────────────────────────────────

#[test]
fn as_int_n_8_wraps_128_to_negative_128() {
    // 128 is just outside the signed 8-bit range; wraps to -128.
    assert_eq!(invoke("asIntN", vec![Value::I32(8), bi(128)]), bi(-128));
}

#[test]
fn as_int_n_8_keeps_127_unchanged() {
    assert_eq!(invoke("asIntN", vec![Value::I32(8), bi(127)]), bi(127));
}

#[test]
fn as_int_n_1_of_one_is_negative_one() {
    // 1-bit signed: only values are 0 and -1. 1n → -1n.
    assert_eq!(invoke("asIntN", vec![Value::I32(1), bi(1)]), bi(-1));
}

// ── asUintN — unsigned width clamping ─────────────────────────────────────────

#[test]
fn as_uint_n_8_wraps_256_to_zero() {
    assert_eq!(invoke("asUintN", vec![Value::I32(8), bi(256)]), bi(0));
}

#[test]
fn as_uint_n_8_of_negative_one_is_255() {
    // -1n in unsigned 8-bit = 255 (all bits set).
    assert_eq!(invoke("asUintN", vec![Value::I32(8), bi(-1)]), bi(255));
}

#[test]
fn as_uint_n_1_of_one_is_one() {
    assert_eq!(invoke("asUintN", vec![Value::I32(1), bi(1)]), bi(1));
}

// ── Comparisons ───────────────────────────────────────────────────────────────

#[test]
fn eq_only_true_for_same_value() {
    assert_eq!(invoke("eq", vec![bi(3), bi(3)]), Value::Bool(true));
    assert_eq!(invoke("eq", vec![bi(3), bi(4)]), Value::Bool(false));
}

#[test]
fn lt_respects_sign() {
    assert_eq!(invoke("lt", vec![bi(-1), bi(0)]), Value::Bool(true));
    assert_eq!(invoke("lt", vec![bi(0), bi(-1)]), Value::Bool(false));
}

#[test]
fn ge_is_true_for_equal_values() {
    assert_eq!(invoke("ge", vec![bi(5), bi(5)]), Value::Bool(true));
}

// ── toString ──────────────────────────────────────────────────────────────────

#[test]
fn to_string_of_negative_bigint_includes_minus_sign() {
    match invoke("toString", vec![bi(-42)]) {
        Value::String(s) => assert_eq!(s.as_ref(), "-42"),
        other => panic!("expected string, got {:?}", other),
    }
}

#[test]
fn to_string_of_zero_is_single_zero_digit() {
    match invoke("toString", vec![bi(0)]) {
        Value::String(s) => assert_eq!(s.as_ref(), "0"),
        other => panic!("expected string, got {:?}", other),
    }
}

// ── toString(radix) ───────────────────────────────────────────────────────────

#[test]
fn to_string_with_radix_16_produces_hex() {
    // ECMA-262 §21.2.3.4: BigInt.prototype.toString(radix) converts to given base.
    // 255n.toString(16) = "ff"
    match invoke("toStringRadix", vec![bi(255), Value::I32(16)]) {
        Value::String(s) => assert_eq!(s.as_ref(), "ff"),
        other => panic!("expected string, got {:?}", other),
    }
}

#[test]
fn to_string_with_radix_2_produces_binary() {
    // 8n.toString(2) = "1000"
    match invoke("toStringRadix", vec![bi(8), Value::I32(2)]) {
        Value::String(s) => assert_eq!(s.as_ref(), "1000"),
        other => panic!("expected string, got {:?}", other),
    }
}

// ── pow (exponentiation ** operator) ─────────────────────────────────────────

#[test]
fn pow_two_to_ten_is_1024() {
    // ECMA-262: BigInt ** BigInt; 2n ** 10n = 1024n.
    assert_eq!(invoke("pow", vec![bi(2), bi(10)]), bi(1024));
}

#[test]
fn pow_any_bigint_to_zero_is_one() {
    // x ** 0n = 1n for any x.
    assert_eq!(invoke("pow", vec![bi(9999), bi(0)]), bi(1));
}

// ── BigInt.prototype.toLocaleString ──────────────────────────────────────────

#[test]
fn to_locale_string_returns_non_empty_string() {
    // ECMA-262 §21.2.3.3: BigInt.prototype.toLocaleString returns locale-formatted string.
    let result = invoke("toLocaleString", vec![bi(1234567)]);
    assert!(matches!(result, Value::String(ref s) if !s.is_empty()));
}

// ── BigInt.prototype.valueOf ──────────────────────────────────────────────────

#[test]
fn value_of_returns_the_bigint_primitive() {
    // §21.2.3.4: valueOf returns the BigInt primitive (same value).
    assert_eq!(invoke("valueOf", vec![bi(42)]), bi(42));
}
