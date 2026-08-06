//! Tests for the nontrapping-float-to-int-conversions WASM proposal.
//! Spec: `proposals/nontrapping-float-to-int-conversions/`, opcodes 0xFC 0x00–0x07.

use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

fn run(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut chunk = Chunk::new("<script>");
    emit(&mut chunk);
    chunk.emit_op(Op::RETURN, 0);
    VM::new().run(vec![chunk]).expect("VM execution failed")
}

fn push(c: &mut Chunk, v: f64) {
    c.emit_f64_const(v, 0);
}

// ── i32.trunc_sat_f32_s (0xFC 0x00) ──────────────────────────────────

#[test]
fn i32_trunc_sat_f32_s_zero_and_subnormal() {
    for v in [
        0.0_f64,
        -0.0,
        f32::from_bits(1) as f64,
        -(f32::from_bits(1) as f64),
    ] {
        assert_eq!(
            run(|c| {
                push(c, v);
                c.emit_op(Op::I32_TRUNC_SAT_F32_S, 0);
            })
            .as_i32(),
            0
        );
    }
}

#[test]
fn i32_trunc_sat_f32_s_truncates_toward_zero() {
    assert_eq!(
        run(|c| {
            push(c, 1.5);
            c.emit_op(Op::I32_TRUNC_SAT_F32_S, 0);
        })
        .as_i32(),
        1
    );
    assert_eq!(
        run(|c| {
            push(c, -1.9);
            c.emit_op(Op::I32_TRUNC_SAT_F32_S, 0);
        })
        .as_i32(),
        -1
    );
    assert_eq!(
        run(|c| {
            push(c, -2.0);
            c.emit_op(Op::I32_TRUNC_SAT_F32_S, 0);
        })
        .as_i32(),
        -2
    );
}

#[test]
fn i32_trunc_sat_f32_s_in_range_boundary() {
    // f32 cannot represent 2147483647 exactly; the largest f32 ≤ i32::MAX is 2147483520
    assert_eq!(
        run(|c| {
            push(c, 2147483520.0);
            c.emit_op(Op::I32_TRUNC_SAT_F32_S, 0);
        })
        .as_i32(),
        2147483520
    );
    assert_eq!(
        run(|c| {
            push(c, -2147483648.0);
            c.emit_op(Op::I32_TRUNC_SAT_F32_S, 0);
        })
        .as_i32(),
        -2147483648
    );
}

#[test]
fn i32_trunc_sat_f32_s_positive_overflow_saturates() {
    assert_eq!(
        run(|c| {
            push(c, 2147483648.0);
            c.emit_op(Op::I32_TRUNC_SAT_F32_S, 0);
        })
        .as_i32(),
        i32::MAX
    );
}

#[test]
fn i32_trunc_sat_f32_s_negative_overflow_saturates() {
    assert_eq!(
        run(|c| {
            push(c, -2147483904.0);
            c.emit_op(Op::I32_TRUNC_SAT_F32_S, 0);
        })
        .as_i32(),
        i32::MIN
    );
}

#[test]
fn i32_trunc_sat_f32_s_inf_saturates() {
    assert_eq!(
        run(|c| {
            push(c, f64::INFINITY);
            c.emit_op(Op::I32_TRUNC_SAT_F32_S, 0);
        })
        .as_i32(),
        i32::MAX
    );
    assert_eq!(
        run(|c| {
            push(c, f64::NEG_INFINITY);
            c.emit_op(Op::I32_TRUNC_SAT_F32_S, 0);
        })
        .as_i32(),
        i32::MIN
    );
}

#[test]
fn i32_trunc_sat_f32_s_nan_returns_zero() {
    assert_eq!(
        run(|c| {
            push(c, f64::NAN);
            c.emit_op(Op::I32_TRUNC_SAT_F32_S, 0);
        })
        .as_i32(),
        0
    );
    assert_eq!(
        run(|c| {
            push(c, -f64::NAN);
            c.emit_op(Op::I32_TRUNC_SAT_F32_S, 0);
        })
        .as_i32(),
        0
    );
}

// ── i32.trunc_sat_f32_u (0xFC 0x01) ──────────────────────────────────

#[test]
fn i32_trunc_sat_f32_u_zero_and_subnormal() {
    for v in [
        0.0_f64,
        -0.0,
        f32::from_bits(1) as f64,
        -(f32::from_bits(1) as f64),
    ] {
        assert_eq!(
            run(|c| {
                push(c, v);
                c.emit_op(Op::I32_TRUNC_SAT_F32_U, 0);
            })
            .as_i32(),
            0
        );
    }
}

#[test]
fn i32_trunc_sat_f32_u_truncates_toward_zero() {
    assert_eq!(
        run(|c| {
            push(c, 1.9);
            c.emit_op(Op::I32_TRUNC_SAT_F32_U, 0);
        })
        .as_i32(),
        1
    );
    assert_eq!(
        run(|c| {
            push(c, 2.0);
            c.emit_op(Op::I32_TRUNC_SAT_F32_U, 0);
        })
        .as_i32(),
        2
    );
}

#[test]
fn i32_trunc_sat_f32_u_mid_range_bit_pattern() {
    // 2^31 stored as 0x80000000 = -2147483648 as i32
    assert_eq!(
        run(|c| {
            push(c, 2147483648.0);
            c.emit_op(Op::I32_TRUNC_SAT_F32_U, 0);
        })
        .as_i32(),
        -2147483648_i32
    );
    // 0xFFFFFF00 = -256 as i32
    assert_eq!(
        run(|c| {
            push(c, 4294967040.0);
            c.emit_op(Op::I32_TRUNC_SAT_F32_U, 0);
        })
        .as_i32(),
        -256_i32
    );
}

#[test]
fn i32_trunc_sat_f32_u_negative_saturates_to_zero() {
    assert_eq!(
        run(|c| {
            push(c, -1.0);
            c.emit_op(Op::I32_TRUNC_SAT_F32_U, 0);
        })
        .as_i32(),
        0
    );
}

#[test]
fn i32_trunc_sat_f32_u_positive_overflow_saturates() {
    // 2^32 > u32::MAX → 0xFFFFFFFF = -1 as i32
    assert_eq!(
        run(|c| {
            push(c, 4294967296.0);
            c.emit_op(Op::I32_TRUNC_SAT_F32_U, 0);
        })
        .as_i32(),
        -1_i32
    );
}

#[test]
fn i32_trunc_sat_f32_u_inf_saturates() {
    assert_eq!(
        run(|c| {
            push(c, f64::INFINITY);
            c.emit_op(Op::I32_TRUNC_SAT_F32_U, 0);
        })
        .as_i32(),
        -1_i32
    );
    assert_eq!(
        run(|c| {
            push(c, f64::NEG_INFINITY);
            c.emit_op(Op::I32_TRUNC_SAT_F32_U, 0);
        })
        .as_i32(),
        0
    );
}

#[test]
fn i32_trunc_sat_f32_u_nan_returns_zero() {
    assert_eq!(
        run(|c| {
            push(c, f64::NAN);
            c.emit_op(Op::I32_TRUNC_SAT_F32_U, 0);
        })
        .as_i32(),
        0
    );
    assert_eq!(
        run(|c| {
            push(c, -f64::NAN);
            c.emit_op(Op::I32_TRUNC_SAT_F32_U, 0);
        })
        .as_i32(),
        0
    );
}

// ── i32.trunc_sat_f64_s (0xFC 0x02) ──────────────────────────────────

#[test]
fn i32_trunc_sat_f64_s_zero_and_subnormal() {
    for v in [0.0_f64, -0.0, f64::from_bits(1), -f64::from_bits(1)] {
        assert_eq!(
            run(|c| {
                push(c, v);
                c.emit_op(Op::I32_TRUNC_SAT_F64_S, 0);
            })
            .as_i32(),
            0
        );
    }
}

#[test]
fn i32_trunc_sat_f64_s_truncates_toward_zero() {
    assert_eq!(
        run(|c| {
            push(c, 1.5);
            c.emit_op(Op::I32_TRUNC_SAT_F64_S, 0);
        })
        .as_i32(),
        1
    );
    assert_eq!(
        run(|c| {
            push(c, -1.9);
            c.emit_op(Op::I32_TRUNC_SAT_F64_S, 0);
        })
        .as_i32(),
        -1
    );
}

#[test]
fn i32_trunc_sat_f64_s_in_range_exact() {
    // f64 can represent i32::MAX exactly unlike f32
    assert_eq!(
        run(|c| {
            push(c, 2147483647.0);
            c.emit_op(Op::I32_TRUNC_SAT_F64_S, 0);
        })
        .as_i32(),
        2147483647
    );
    assert_eq!(
        run(|c| {
            push(c, -2147483648.0);
            c.emit_op(Op::I32_TRUNC_SAT_F64_S, 0);
        })
        .as_i32(),
        -2147483648
    );
}

#[test]
fn i32_trunc_sat_f64_s_positive_overflow_saturates() {
    assert_eq!(
        run(|c| {
            push(c, 2147483648.0);
            c.emit_op(Op::I32_TRUNC_SAT_F64_S, 0);
        })
        .as_i32(),
        i32::MAX
    );
}

#[test]
fn i32_trunc_sat_f64_s_negative_overflow_saturates() {
    assert_eq!(
        run(|c| {
            push(c, -2147483649.0);
            c.emit_op(Op::I32_TRUNC_SAT_F64_S, 0);
        })
        .as_i32(),
        i32::MIN
    );
}

#[test]
fn i32_trunc_sat_f64_s_inf_saturates() {
    assert_eq!(
        run(|c| {
            push(c, f64::INFINITY);
            c.emit_op(Op::I32_TRUNC_SAT_F64_S, 0);
        })
        .as_i32(),
        i32::MAX
    );
    assert_eq!(
        run(|c| {
            push(c, f64::NEG_INFINITY);
            c.emit_op(Op::I32_TRUNC_SAT_F64_S, 0);
        })
        .as_i32(),
        i32::MIN
    );
}

#[test]
fn i32_trunc_sat_f64_s_nan_returns_zero() {
    assert_eq!(
        run(|c| {
            push(c, f64::NAN);
            c.emit_op(Op::I32_TRUNC_SAT_F64_S, 0);
        })
        .as_i32(),
        0
    );
    assert_eq!(
        run(|c| {
            push(c, -f64::NAN);
            c.emit_op(Op::I32_TRUNC_SAT_F64_S, 0);
        })
        .as_i32(),
        0
    );
}

// ── i32.trunc_sat_f64_u (0xFC 0x03) ──────────────────────────────────

#[test]
fn i32_trunc_sat_f64_u_zero_and_subnormal() {
    for v in [0.0_f64, -0.0, f64::from_bits(1), -f64::from_bits(1)] {
        assert_eq!(
            run(|c| {
                push(c, v);
                c.emit_op(Op::I32_TRUNC_SAT_F64_U, 0);
            })
            .as_i32(),
            0
        );
    }
}

#[test]
fn i32_trunc_sat_f64_u_truncates_toward_zero() {
    assert_eq!(
        run(|c| {
            push(c, 1.9);
            c.emit_op(Op::I32_TRUNC_SAT_F64_U, 0);
        })
        .as_i32(),
        1
    );
    assert_eq!(
        run(|c| {
            push(c, 1e8);
            c.emit_op(Op::I32_TRUNC_SAT_F64_U, 0);
        })
        .as_i32(),
        100000000
    );
}

#[test]
fn i32_trunc_sat_f64_u_in_range_bit_patterns() {
    // f64 can represent u32::MAX exactly
    assert_eq!(
        run(|c| {
            push(c, 4294967295.0);
            c.emit_op(Op::I32_TRUNC_SAT_F64_U, 0);
        })
        .as_i32(),
        -1_i32
    );
    // 2^31 = 0x80000000 = -2147483648 as i32
    assert_eq!(
        run(|c| {
            push(c, 2147483648.0);
            c.emit_op(Op::I32_TRUNC_SAT_F64_U, 0);
        })
        .as_i32(),
        -2147483648_i32
    );
}

#[test]
fn i32_trunc_sat_f64_u_negative_saturates_to_zero() {
    assert_eq!(
        run(|c| {
            push(c, -1.0);
            c.emit_op(Op::I32_TRUNC_SAT_F64_U, 0);
        })
        .as_i32(),
        0
    );
}

#[test]
fn i32_trunc_sat_f64_u_positive_overflow_saturates() {
    for v in [4294967296.0, 1e16, 1e30] {
        assert_eq!(
            run(|c| {
                push(c, v);
                c.emit_op(Op::I32_TRUNC_SAT_F64_U, 0);
            })
            .as_i32(),
            -1_i32
        );
    }
}

#[test]
fn i32_trunc_sat_f64_u_inf_saturates() {
    assert_eq!(
        run(|c| {
            push(c, f64::INFINITY);
            c.emit_op(Op::I32_TRUNC_SAT_F64_U, 0);
        })
        .as_i32(),
        -1_i32
    );
    assert_eq!(
        run(|c| {
            push(c, f64::NEG_INFINITY);
            c.emit_op(Op::I32_TRUNC_SAT_F64_U, 0);
        })
        .as_i32(),
        0
    );
}

#[test]
fn i32_trunc_sat_f64_u_nan_returns_zero() {
    assert_eq!(
        run(|c| {
            push(c, f64::NAN);
            c.emit_op(Op::I32_TRUNC_SAT_F64_U, 0);
        })
        .as_i32(),
        0
    );
    assert_eq!(
        run(|c| {
            push(c, -f64::NAN);
            c.emit_op(Op::I32_TRUNC_SAT_F64_U, 0);
        })
        .as_i32(),
        0
    );
}

// ── i64.trunc_sat_f32_s (0xFC 0x04) ──────────────────────────────────

#[test]
fn i64_trunc_sat_f32_s_zero_and_subnormal() {
    for v in [
        0.0_f64,
        -0.0,
        f32::from_bits(1) as f64,
        -(f32::from_bits(1) as f64),
    ] {
        assert_eq!(
            run(|c| {
                push(c, v);
                c.emit_op(Op::I64_TRUNC_SAT_F32_S, 0);
            })
            .as_i64(),
            0
        );
    }
}

#[test]
fn i64_trunc_sat_f32_s_truncates_toward_zero() {
    assert_eq!(
        run(|c| {
            push(c, 1.5);
            c.emit_op(Op::I64_TRUNC_SAT_F32_S, 0);
        })
        .as_i64(),
        1
    );
    assert_eq!(
        run(|c| {
            push(c, -1.9);
            c.emit_op(Op::I64_TRUNC_SAT_F32_S, 0);
        })
        .as_i64(),
        -1
    );
    assert_eq!(
        run(|c| {
            push(c, -2.0);
            c.emit_op(Op::I64_TRUNC_SAT_F32_S, 0);
        })
        .as_i64(),
        -2
    );
}

#[test]
fn i64_trunc_sat_f32_s_in_range() {
    // 2^32 is exactly representable in f32
    assert_eq!(
        run(|c| {
            push(c, 4294967296.0);
            c.emit_op(Op::I64_TRUNC_SAT_F32_S, 0);
        })
        .as_i64(),
        4294967296
    );
    assert_eq!(
        run(|c| {
            push(c, -4294967296.0);
            c.emit_op(Op::I64_TRUNC_SAT_F32_S, 0);
        })
        .as_i64(),
        -4294967296
    );
}

#[test]
fn i64_trunc_sat_f32_s_positive_overflow_saturates() {
    assert_eq!(
        run(|c| {
            push(c, 9223372036854775808.0);
            c.emit_op(Op::I64_TRUNC_SAT_F32_S, 0);
        })
        .as_i64(),
        i64::MAX
    );
}

#[test]
fn i64_trunc_sat_f32_s_negative_overflow_saturates() {
    assert_eq!(
        run(|c| {
            push(c, -9223373136366403584.0);
            c.emit_op(Op::I64_TRUNC_SAT_F32_S, 0);
        })
        .as_i64(),
        i64::MIN
    );
}

#[test]
fn i64_trunc_sat_f32_s_inf_saturates() {
    assert_eq!(
        run(|c| {
            push(c, f64::INFINITY);
            c.emit_op(Op::I64_TRUNC_SAT_F32_S, 0);
        })
        .as_i64(),
        i64::MAX
    );
    assert_eq!(
        run(|c| {
            push(c, f64::NEG_INFINITY);
            c.emit_op(Op::I64_TRUNC_SAT_F32_S, 0);
        })
        .as_i64(),
        i64::MIN
    );
}

#[test]
fn i64_trunc_sat_f32_s_nan_returns_zero() {
    assert_eq!(
        run(|c| {
            push(c, f64::NAN);
            c.emit_op(Op::I64_TRUNC_SAT_F32_S, 0);
        })
        .as_i64(),
        0
    );
    assert_eq!(
        run(|c| {
            push(c, -f64::NAN);
            c.emit_op(Op::I64_TRUNC_SAT_F32_S, 0);
        })
        .as_i64(),
        0
    );
}

// ── i64.trunc_sat_f32_u (0xFC 0x05) ──────────────────────────────────

#[test]
fn i64_trunc_sat_f32_u_zero_and_subnormal() {
    for v in [
        0.0_f64,
        -0.0,
        f32::from_bits(1) as f64,
        -(f32::from_bits(1) as f64),
    ] {
        assert_eq!(
            run(|c| {
                push(c, v);
                c.emit_op(Op::I64_TRUNC_SAT_F32_U, 0);
            })
            .as_i64(),
            0
        );
    }
}

#[test]
fn i64_trunc_sat_f32_u_truncates_toward_zero() {
    assert_eq!(
        run(|c| {
            push(c, 1.5);
            c.emit_op(Op::I64_TRUNC_SAT_F32_U, 0);
        })
        .as_i64(),
        1
    );
    assert_eq!(
        run(|c| {
            push(c, 4294967296.0);
            c.emit_op(Op::I64_TRUNC_SAT_F32_U, 0);
        })
        .as_i64(),
        4294967296
    );
}

#[test]
fn i64_trunc_sat_f32_u_negative_saturates_to_zero() {
    assert_eq!(
        run(|c| {
            push(c, -1.0);
            c.emit_op(Op::I64_TRUNC_SAT_F32_U, 0);
        })
        .as_i64(),
        0
    );
}

#[test]
fn i64_trunc_sat_f32_u_positive_overflow_saturates() {
    // 2^64 > u64::MAX → 0xFFFFFFFFFFFFFFFF = -1 as i64
    assert_eq!(
        run(|c| {
            push(c, 18446744073709551616.0);
            c.emit_op(Op::I64_TRUNC_SAT_F32_U, 0);
        })
        .as_i64(),
        -1_i64
    );
}

#[test]
fn i64_trunc_sat_f32_u_inf_saturates() {
    assert_eq!(
        run(|c| {
            push(c, f64::INFINITY);
            c.emit_op(Op::I64_TRUNC_SAT_F32_U, 0);
        })
        .as_i64(),
        -1_i64
    );
    assert_eq!(
        run(|c| {
            push(c, f64::NEG_INFINITY);
            c.emit_op(Op::I64_TRUNC_SAT_F32_U, 0);
        })
        .as_i64(),
        0
    );
}

#[test]
fn i64_trunc_sat_f32_u_nan_returns_zero() {
    assert_eq!(
        run(|c| {
            push(c, f64::NAN);
            c.emit_op(Op::I64_TRUNC_SAT_F32_U, 0);
        })
        .as_i64(),
        0
    );
    assert_eq!(
        run(|c| {
            push(c, -f64::NAN);
            c.emit_op(Op::I64_TRUNC_SAT_F32_U, 0);
        })
        .as_i64(),
        0
    );
}

// ── i64.trunc_sat_f64_s (0xFC 0x06) ──────────────────────────────────

#[test]
fn i64_trunc_sat_f64_s_zero_and_subnormal() {
    for v in [0.0_f64, -0.0, f64::from_bits(1), -f64::from_bits(1)] {
        assert_eq!(
            run(|c| {
                push(c, v);
                c.emit_op(Op::I64_TRUNC_SAT_F64_S, 0);
            })
            .as_i64(),
            0
        );
    }
}

#[test]
fn i64_trunc_sat_f64_s_truncates_toward_zero() {
    assert_eq!(
        run(|c| {
            push(c, 1.5);
            c.emit_op(Op::I64_TRUNC_SAT_F64_S, 0);
        })
        .as_i64(),
        1
    );
    assert_eq!(
        run(|c| {
            push(c, -1.9);
            c.emit_op(Op::I64_TRUNC_SAT_F64_S, 0);
        })
        .as_i64(),
        -1
    );
}

#[test]
fn i64_trunc_sat_f64_s_in_range() {
    assert_eq!(
        run(|c| {
            push(c, 4294967296.0);
            c.emit_op(Op::I64_TRUNC_SAT_F64_S, 0);
        })
        .as_i64(),
        4294967296
    );
    assert_eq!(
        run(|c| {
            push(c, -4294967296.0);
            c.emit_op(Op::I64_TRUNC_SAT_F64_S, 0);
        })
        .as_i64(),
        -4294967296
    );
    // Largest f64 < 2^63
    assert_eq!(
        run(|c| {
            push(c, 9223372036854774784.0);
            c.emit_op(Op::I64_TRUNC_SAT_F64_S, 0);
        })
        .as_i64(),
        9223372036854774784
    );
}

#[test]
fn i64_trunc_sat_f64_s_positive_overflow_saturates() {
    assert_eq!(
        run(|c| {
            push(c, 9223372036854775808.0);
            c.emit_op(Op::I64_TRUNC_SAT_F64_S, 0);
        })
        .as_i64(),
        i64::MAX
    );
}

#[test]
fn i64_trunc_sat_f64_s_negative_overflow_saturates() {
    assert_eq!(
        run(|c| {
            push(c, -9223372036854777856.0);
            c.emit_op(Op::I64_TRUNC_SAT_F64_S, 0);
        })
        .as_i64(),
        i64::MIN
    );
}

#[test]
fn i64_trunc_sat_f64_s_inf_saturates() {
    assert_eq!(
        run(|c| {
            push(c, f64::INFINITY);
            c.emit_op(Op::I64_TRUNC_SAT_F64_S, 0);
        })
        .as_i64(),
        i64::MAX
    );
    assert_eq!(
        run(|c| {
            push(c, f64::NEG_INFINITY);
            c.emit_op(Op::I64_TRUNC_SAT_F64_S, 0);
        })
        .as_i64(),
        i64::MIN
    );
}

#[test]
fn i64_trunc_sat_f64_s_nan_returns_zero() {
    assert_eq!(
        run(|c| {
            push(c, f64::NAN);
            c.emit_op(Op::I64_TRUNC_SAT_F64_S, 0);
        })
        .as_i64(),
        0
    );
    assert_eq!(
        run(|c| {
            push(c, -f64::NAN);
            c.emit_op(Op::I64_TRUNC_SAT_F64_S, 0);
        })
        .as_i64(),
        0
    );
}

// ── i64.trunc_sat_f64_u (0xFC 0x07) ──────────────────────────────────

#[test]
fn i64_trunc_sat_f64_u_zero_and_subnormal() {
    for v in [0.0_f64, -0.0, f64::from_bits(1), -f64::from_bits(1)] {
        assert_eq!(
            run(|c| {
                push(c, v);
                c.emit_op(Op::I64_TRUNC_SAT_F64_U, 0);
            })
            .as_i64(),
            0
        );
    }
}

#[test]
fn i64_trunc_sat_f64_u_truncates_toward_zero() {
    assert_eq!(
        run(|c| {
            push(c, 1.5);
            c.emit_op(Op::I64_TRUNC_SAT_F64_U, 0);
        })
        .as_i64(),
        1
    );
    assert_eq!(
        run(|c| {
            push(c, 1e8);
            c.emit_op(Op::I64_TRUNC_SAT_F64_U, 0);
        })
        .as_i64(),
        100000000
    );
    assert_eq!(
        run(|c| {
            push(c, 1e16);
            c.emit_op(Op::I64_TRUNC_SAT_F64_U, 0);
        })
        .as_i64(),
        10000000000000000
    );
}

#[test]
fn i64_trunc_sat_f64_u_in_range_bit_patterns() {
    assert_eq!(
        run(|c| {
            push(c, 4294967295.0);
            c.emit_op(Op::I64_TRUNC_SAT_F64_U, 0);
        })
        .as_i64(),
        0xffffffff
    );
    assert_eq!(
        run(|c| {
            push(c, 4294967296.0);
            c.emit_op(Op::I64_TRUNC_SAT_F64_U, 0);
        })
        .as_i64(),
        0x100000000
    );
    // 2^63 fits in u64 but not i64 → stored as i64::MIN bit pattern
    assert_eq!(
        run(|c| {
            push(c, 9223372036854775808.0);
            c.emit_op(Op::I64_TRUNC_SAT_F64_U, 0);
        })
        .as_i64(),
        i64::MIN
    );
}

#[test]
fn i64_trunc_sat_f64_u_negative_saturates_to_zero() {
    assert_eq!(
        run(|c| {
            push(c, -1.0);
            c.emit_op(Op::I64_TRUNC_SAT_F64_U, 0);
        })
        .as_i64(),
        0
    );
}

#[test]
fn i64_trunc_sat_f64_u_positive_overflow_saturates() {
    // Above u64::MAX → all ones = -1 as i64
    assert_eq!(
        run(|c| {
            push(c, 18446744073709551616.0);
            c.emit_op(Op::I64_TRUNC_SAT_F64_U, 0);
        })
        .as_i64(),
        -1_i64
    );
}

#[test]
fn i64_trunc_sat_f64_u_inf_saturates() {
    assert_eq!(
        run(|c| {
            push(c, f64::INFINITY);
            c.emit_op(Op::I64_TRUNC_SAT_F64_U, 0);
        })
        .as_i64(),
        -1_i64
    );
    assert_eq!(
        run(|c| {
            push(c, f64::NEG_INFINITY);
            c.emit_op(Op::I64_TRUNC_SAT_F64_U, 0);
        })
        .as_i64(),
        0
    );
}

#[test]
fn i64_trunc_sat_f64_u_nan_returns_zero() {
    assert_eq!(
        run(|c| {
            push(c, f64::NAN);
            c.emit_op(Op::I64_TRUNC_SAT_F64_U, 0);
        })
        .as_i64(),
        0
    );
    assert_eq!(
        run(|c| {
            push(c, -f64::NAN);
            c.emit_op(Op::I64_TRUNC_SAT_F64_U, 0);
        })
        .as_i64(),
        0
    );
}
