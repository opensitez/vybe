//! Tests for all f32 instructions from the WASM spec (§5.3 numeric).
//! Covers: const, comparisons (eq/ne/lt/gt/le/ge),
//!         unary (abs/neg/ceil/floor/trunc/nearest/sqrt),
//!         binary (add/sub/mul/div/min/max/copysign).
//! The VM holds f32 values as F64; arithmetic is done at f32 precision.

use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};

fn run(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    VM::new().run(vec![c]).expect("run failed")
}

fn push(c: &mut Chunk, v: f32) {
    let k = c.add_constant(Value::F64(v as f64));
    c.emit_op_u16(Op::CONST, k, 0);
}

fn f32_result(v: Value) -> f32 {
    v.as_f64() as f32
}
fn bool_result(v: Value) -> i32 {
    v.as_i32()
}

// ── f32.const ────────────────────────────────────────────────────────────

#[test]
fn f32_const_zero() {
    assert_eq!(f32_result(run(|c| push(c, 0.0))), 0.0);
}
#[test]
fn f32_const_one() {
    assert_eq!(f32_result(run(|c| push(c, 1.0))), 1.0);
}
#[test]
fn f32_const_negative() {
    assert_eq!(f32_result(run(|c| push(c, -1.5))), -1.5);
}

// ── f32 comparisons ───────────────────────────────────────────────────────

#[test]
fn f32_eq_true() {
    assert_eq!(
        bool_result(run(|c| {
            push(c, 1.0);
            push(c, 1.0);
            c.emit_op(Op::F32_EQ, 0);
        })),
        1
    );
}
#[test]
fn f32_eq_false() {
    assert_eq!(
        bool_result(run(|c| {
            push(c, 1.0);
            push(c, 2.0);
            c.emit_op(Op::F32_EQ, 0);
        })),
        0
    );
}
#[test]
fn f32_ne_true() {
    assert_eq!(
        bool_result(run(|c| {
            push(c, 1.0);
            push(c, 2.0);
            c.emit_op(Op::F32_NE, 0);
        })),
        1
    );
}
#[test]
fn f32_ne_false() {
    assert_eq!(
        bool_result(run(|c| {
            push(c, 1.0);
            push(c, 1.0);
            c.emit_op(Op::F32_NE, 0);
        })),
        0
    );
}
#[test]
fn f32_lt_true() {
    assert_eq!(
        bool_result(run(|c| {
            push(c, 1.0);
            push(c, 2.0);
            c.emit_op(Op::F32_LT, 0);
        })),
        1
    );
}
#[test]
fn f32_lt_false() {
    assert_eq!(
        bool_result(run(|c| {
            push(c, 2.0);
            push(c, 1.0);
            c.emit_op(Op::F32_LT, 0);
        })),
        0
    );
}
#[test]
fn f32_gt_true() {
    assert_eq!(
        bool_result(run(|c| {
            push(c, 3.0);
            push(c, 1.0);
            c.emit_op(Op::F32_GT, 0);
        })),
        1
    );
}
#[test]
fn f32_gt_false() {
    assert_eq!(
        bool_result(run(|c| {
            push(c, 1.0);
            push(c, 3.0);
            c.emit_op(Op::F32_GT, 0);
        })),
        0
    );
}
#[test]
fn f32_le_equal() {
    assert_eq!(
        bool_result(run(|c| {
            push(c, 2.0);
            push(c, 2.0);
            c.emit_op(Op::F32_LE, 0);
        })),
        1
    );
}
#[test]
fn f32_le_less() {
    assert_eq!(
        bool_result(run(|c| {
            push(c, 1.0);
            push(c, 2.0);
            c.emit_op(Op::F32_LE, 0);
        })),
        1
    );
}
#[test]
fn f32_ge_equal() {
    assert_eq!(
        bool_result(run(|c| {
            push(c, 2.0);
            push(c, 2.0);
            c.emit_op(Op::F32_GE, 0);
        })),
        1
    );
}
#[test]
fn f32_ge_greater() {
    assert_eq!(
        bool_result(run(|c| {
            push(c, 3.0);
            push(c, 2.0);
            c.emit_op(Op::F32_GE, 0);
        })),
        1
    );
}
#[test]
fn f32_eq_nan_false() {
    assert_eq!(
        bool_result(run(|c| {
            push(c, f32::NAN);
            push(c, f32::NAN);
            c.emit_op(Op::F32_EQ, 0);
        })),
        0
    );
}
#[test]
fn f32_ne_nan_true() {
    assert_eq!(
        bool_result(run(|c| {
            push(c, f32::NAN);
            push(c, 1.0);
            c.emit_op(Op::F32_NE, 0);
        })),
        1
    );
}
#[test]
fn f32_lt_nan_false() {
    assert_eq!(
        bool_result(run(|c| {
            push(c, f32::NAN);
            push(c, 1.0);
            c.emit_op(Op::F32_LT, 0);
        })),
        0
    );
}
#[test]
fn f32_ordered_comparisons_with_nan_are_false() {
    for op in [Op::F32_GT, Op::F32_LE, Op::F32_GE] {
        assert_eq!(
            bool_result(run(|c| {
                push(c, f32::NAN);
                push(c, 1.0);
                c.emit_op(op, 0);
            })),
            0
        );
        assert_eq!(
            bool_result(run(|c| {
                push(c, 1.0);
                push(c, f32::NAN);
                c.emit_op(op, 0);
            })),
            0
        );
    }
}

// ── f32 unary ─────────────────────────────────────────────────────────────

#[test]
fn f32_abs_positive() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, 3.0);
            c.emit_op(Op::F32_ABS, 0);
        })),
        3.0
    );
}
#[test]
fn f32_abs_negative() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, -3.0);
            c.emit_op(Op::F32_ABS, 0);
        })),
        3.0
    );
}
#[test]
fn f32_neg() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, 3.0);
            c.emit_op(Op::F32_NEG, 0);
        })),
        -3.0
    );
}
#[test]
fn f32_neg_positive_zero_produces_negative_zero() {
    assert!(
        f32_result(run(|c| {
            push(c, 0.0);
            c.emit_op(Op::F32_NEG, 0);
        }))
        .is_sign_negative()
    );
}
#[test]
fn f32_abs_negative_zero_produces_positive_zero() {
    assert!(
        f32_result(run(|c| {
            push(c, -0.0);
            c.emit_op(Op::F32_ABS, 0);
        }))
        .is_sign_positive()
    );
}
#[test]
fn f32_ceil() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, 1.2);
            c.emit_op(Op::F32_CEIL, 0);
        })),
        2.0
    );
}
#[test]
fn f32_floor() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, 1.8);
            c.emit_op(Op::F32_FLOOR, 0);
        })),
        1.0
    );
}
#[test]
fn f32_trunc() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, -1.9);
            c.emit_op(Op::F32_TRUNC, 0);
        })),
        -1.0
    );
}
#[test]
fn f32_nearest_even() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, 0.5);
            c.emit_op(Op::F32_NEAREST, 0);
        })),
        0.0
    );
}
#[test]
fn f32_sqrt() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, 4.0);
            c.emit_op(Op::F32_SQRT, 0);
        })),
        2.0
    );
}
#[test]
fn f32_sqrt_negative_is_nan() {
    assert!(
        f32_result(run(|c| {
            push(c, -1.0);
            c.emit_op(Op::F32_SQRT, 0);
        }))
        .is_nan()
    );
}

// ── f32 binary ────────────────────────────────────────────────────────────

#[test]
fn f32_add() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, 1.5);
            push(c, 2.5);
            c.emit_op(Op::F32_ADD, 0);
        })),
        4.0
    );
}
#[test]
fn f32_sub() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, 5.0);
            push(c, 3.0);
            c.emit_op(Op::F32_SUB, 0);
        })),
        2.0
    );
}
#[test]
fn f32_mul() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, 3.0);
            push(c, 4.0);
            c.emit_op(Op::F32_MUL, 0);
        })),
        12.0
    );
}
#[test]
fn f32_div() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, 7.0);
            push(c, 2.0);
            c.emit_op(Op::F32_DIV, 0);
        })),
        3.5
    );
}
#[test]
fn f32_min() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, 1.0);
            push(c, 2.0);
            c.emit_op(Op::F32_MIN, 0);
        })),
        1.0
    );
}
#[test]
fn f32_max() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, 1.0);
            push(c, 2.0);
            c.emit_op(Op::F32_MAX, 0);
        })),
        2.0
    );
}
#[test]
fn f32_copysign_positive() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, -3.0);
            push(c, 1.0);
            c.emit_op(Op::F32_COPYSIGN, 0);
        })),
        3.0
    );
}
#[test]
fn f32_copysign_negative() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, 3.0);
            push(c, -1.0);
            c.emit_op(Op::F32_COPYSIGN, 0);
        })),
        -3.0
    );
}
#[test]
fn f32_copysign_uses_negative_zero_sign() {
    assert!(
        f32_result(run(|c| {
            push(c, 3.0);
            push(c, -0.0);
            c.emit_op(Op::F32_COPYSIGN, 0);
        }))
        .is_sign_negative()
    );
}
#[test]
fn f32_div_by_zero_infinity() {
    assert!(
        f32_result(run(|c| {
            push(c, 1.0);
            push(c, 0.0);
            c.emit_op(Op::F32_DIV, 0);
        }))
        .is_infinite()
    );
}

// ── Spec-required edge cases ──────────────────────────────────────────────

#[test]
fn f32_nearest_one_point_five() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, 1.5);
            c.emit_op(Op::F32_NEAREST, 0);
        })),
        2.0
    );
}
#[test]
fn f32_nearest_two_point_five() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, 2.5);
            c.emit_op(Op::F32_NEAREST, 0);
        })),
        2.0
    );
}
#[test]
fn f32_nearest_neg_half() {
    assert_eq!(
        f32_result(run(|c| {
            push(c, -0.5);
            c.emit_op(Op::F32_NEAREST, 0);
        })),
        0.0
    );
}
#[test]
fn f32_min_nan_propagates() {
    assert!(
        f32_result(run(|c| {
            push(c, f32::NAN);
            push(c, 5.0);
            c.emit_op(Op::F32_MIN, 0);
        }))
        .is_nan()
    );
}
#[test]
fn f32_max_nan_propagates() {
    assert!(
        f32_result(run(|c| {
            push(c, f32::NAN);
            push(c, 5.0);
            c.emit_op(Op::F32_MAX, 0);
        }))
        .is_nan()
    );
}
#[test]
fn f32_min_nan_second() {
    assert!(
        f32_result(run(|c| {
            push(c, 5.0);
            push(c, f32::NAN);
            c.emit_op(Op::F32_MIN, 0);
        }))
        .is_nan()
    );
}

#[test]
fn f32_min_neg_zero() {
    assert!(
        f32_result(run(|c| {
            push(c, -0.0);
            push(c, 0.0);
            c.emit_op(Op::F32_MIN, 0);
        }))
        .is_sign_negative()
    );
}

#[test]
fn f32_max_pos_zero() {
    assert!(
        f32_result(run(|c| {
            push(c, -0.0);
            push(c, 0.0);
            c.emit_op(Op::F32_MAX, 0);
        }))
        .is_sign_positive()
    );
}
