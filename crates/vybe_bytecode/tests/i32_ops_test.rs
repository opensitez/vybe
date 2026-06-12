//! Tests for all i32 instructions from the WASM spec (§5.3 numeric).
//! Covers: const, eqz, comparisons (eq/ne/lt_s/lt_u/gt_s/gt_u/le_s/le_u/ge_s/ge_u),
//!         clz/ctz/popcnt, add/sub/mul/div_s/div_u/rem_s/rem_u,
//!         and/or/xor/shl/shr_s/shr_u/rotl/rotr.

use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};

fn run(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    VM::new().run(vec![c]).expect("run failed")
}

fn i32(c: &mut Chunk, v: i32) -> u16 {
    let k = c.add_constant(Value::I32(v));
    k
}

fn push(c: &mut Chunk, v: i32) {
    let k = i32(c, v);
    c.emit_op_u16(Op::CONST, k, 0);
}

// ── i32.const ────────────────────────────────────────────────────────────

#[test]
fn i32_const_zero() {
    assert_eq!(run(|c| push(c, 0)).as_i32(), 0);
}
#[test]
fn i32_const_positive() {
    assert_eq!(run(|c| push(c, 42)).as_i32(), 42);
}
#[test]
fn i32_const_negative() {
    assert_eq!(run(|c| push(c, -1)).as_i32(), -1);
}
#[test]
fn i32_const_min() {
    assert_eq!(run(|c| push(c, i32::MIN)).as_i32(), i32::MIN);
}
#[test]
fn i32_const_max() {
    assert_eq!(run(|c| push(c, i32::MAX)).as_i32(), i32::MAX);
}

// ── i32.eqz ──────────────────────────────────────────────────────────────

#[test]
fn i32_eqz_zero() {
    assert_eq!(
        run(|c| {
            push(c, 0);
            c.emit_op(Op::I32_EQZ, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i32_eqz_nonzero() {
    assert_eq!(
        run(|c| {
            push(c, 5);
            c.emit_op(Op::I32_EQZ, 0);
        })
        .as_i32(),
        0
    );
}

// ── i32 comparisons (signed) ──────────────────────────────────────────────

#[test]
fn i32_eq_true() {
    assert_eq!(
        run(|c| {
            push(c, 3);
            push(c, 3);
            c.emit_op(Op::I32_EQ, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i32_eq_false() {
    assert_eq!(
        run(|c| {
            push(c, 1);
            push(c, 2);
            c.emit_op(Op::I32_EQ, 0);
        })
        .as_i32(),
        0
    );
}
#[test]
fn i32_ne_true() {
    assert_eq!(
        run(|c| {
            push(c, 1);
            push(c, 2);
            c.emit_op(Op::I32_NE, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i32_ne_false() {
    assert_eq!(
        run(|c| {
            push(c, 3);
            push(c, 3);
            c.emit_op(Op::I32_NE, 0);
        })
        .as_i32(),
        0
    );
}
#[test]
fn i32_lt_s_true() {
    assert_eq!(
        run(|c| {
            push(c, -1);
            push(c, 0);
            c.emit_op(Op::I32_LT_S, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i32_lt_s_false() {
    assert_eq!(
        run(|c| {
            push(c, 1);
            push(c, 0);
            c.emit_op(Op::I32_LT_S, 0);
        })
        .as_i32(),
        0
    );
}
#[test]
fn i32_gt_s_true() {
    assert_eq!(
        run(|c| {
            push(c, 5);
            push(c, 3);
            c.emit_op(Op::I32_GT_S, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i32_gt_s_false() {
    assert_eq!(
        run(|c| {
            push(c, 3);
            push(c, 5);
            c.emit_op(Op::I32_GT_S, 0);
        })
        .as_i32(),
        0
    );
}
#[test]
fn i32_le_s_equal() {
    assert_eq!(
        run(|c| {
            push(c, 3);
            push(c, 3);
            c.emit_op(Op::I32_LE_S, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i32_le_s_less() {
    assert_eq!(
        run(|c| {
            push(c, 2);
            push(c, 3);
            c.emit_op(Op::I32_LE_S, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i32_ge_s_equal() {
    assert_eq!(
        run(|c| {
            push(c, 3);
            push(c, 3);
            c.emit_op(Op::I32_GE_S, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i32_ge_s_greater() {
    assert_eq!(
        run(|c| {
            push(c, 5);
            push(c, 3);
            c.emit_op(Op::I32_GE_S, 0);
        })
        .as_i32(),
        1
    );
}

// ── i32 comparisons (unsigned) ────────────────────────────────────────────

#[test]
fn i32_lt_u_unsigned() {
    // -1 as u32 = 4294967295, which is > 1 unsigned
    assert_eq!(
        run(|c| {
            push(c, 1);
            push(c, -1);
            c.emit_op(Op::I32_LT_U, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i32_gt_u_unsigned() {
    assert_eq!(
        run(|c| {
            push(c, -1);
            push(c, 1);
            c.emit_op(Op::I32_GT_U, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i32_le_u_equal() {
    assert_eq!(
        run(|c| {
            push(c, 5);
            push(c, 5);
            c.emit_op(Op::I32_LE_U, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i32_ge_u_unsigned() {
    assert_eq!(
        run(|c| {
            push(c, -1);
            push(c, 0);
            c.emit_op(Op::I32_GE_U, 0);
        })
        .as_i32(),
        1
    );
}

// ── i32 bit counting ─────────────────────────────────────────────────────

#[test]
fn i32_clz_zero() {
    assert_eq!(
        run(|c| {
            push(c, 0);
            c.emit_op(Op::I32_CLZ, 0);
        })
        .as_i32(),
        32
    );
}
#[test]
fn i32_clz_one() {
    assert_eq!(
        run(|c| {
            push(c, 1);
            c.emit_op(Op::I32_CLZ, 0);
        })
        .as_i32(),
        31
    );
}
#[test]
fn i32_clz_min() {
    assert_eq!(
        run(|c| {
            push(c, i32::MIN);
            c.emit_op(Op::I32_CLZ, 0);
        })
        .as_i32(),
        0
    );
}
#[test]
fn i32_ctz_zero() {
    assert_eq!(
        run(|c| {
            push(c, 0);
            c.emit_op(Op::I32_CTZ, 0);
        })
        .as_i32(),
        32
    );
}
#[test]
fn i32_ctz_one() {
    assert_eq!(
        run(|c| {
            push(c, 1);
            c.emit_op(Op::I32_CTZ, 0);
        })
        .as_i32(),
        0
    );
}
#[test]
fn i32_ctz_four() {
    assert_eq!(
        run(|c| {
            push(c, 4);
            c.emit_op(Op::I32_CTZ, 0);
        })
        .as_i32(),
        2
    );
}
#[test]
fn i32_popcnt_zero() {
    assert_eq!(
        run(|c| {
            push(c, 0);
            c.emit_op(Op::I32_POPCNT, 0);
        })
        .as_i32(),
        0
    );
}
#[test]
fn i32_popcnt_ones() {
    assert_eq!(
        run(|c| {
            push(c, -1);
            c.emit_op(Op::I32_POPCNT, 0);
        })
        .as_i32(),
        32
    );
}
#[test]
fn i32_popcnt_five() {
    assert_eq!(
        run(|c| {
            push(c, 0b10101);
            c.emit_op(Op::I32_POPCNT, 0);
        })
        .as_i32(),
        3
    );
}

// ── i32 arithmetic ────────────────────────────────────────────────────────

#[test]
fn i32_add() {
    assert_eq!(
        run(|c| {
            push(c, 20);
            push(c, 22);
            c.emit_op(Op::I32_ADD, 0);
        })
        .as_i32(),
        42
    );
}
#[test]
fn i32_add_overflow() {
    assert_eq!(
        run(|c| {
            push(c, i32::MAX);
            push(c, 1);
            c.emit_op(Op::I32_ADD, 0);
        })
        .as_i32(),
        i32::MIN
    );
}
#[test]
fn i32_sub() {
    assert_eq!(
        run(|c| {
            push(c, 50);
            push(c, 8);
            c.emit_op(Op::I32_SUB, 0);
        })
        .as_i32(),
        42
    );
}
#[test]
fn i32_sub_negative() {
    assert_eq!(
        run(|c| {
            push(c, 0);
            push(c, 1);
            c.emit_op(Op::I32_SUB, 0);
        })
        .as_i32(),
        -1
    );
}
#[test]
fn i32_mul() {
    assert_eq!(
        run(|c| {
            push(c, 6);
            push(c, 7);
            c.emit_op(Op::I32_MUL, 0);
        })
        .as_i32(),
        42
    );
}
#[test]
fn i32_div_s() {
    assert_eq!(
        run(|c| {
            push(c, 84);
            push(c, 2);
            c.emit_op(Op::I32_DIV_S, 0);
        })
        .as_i32(),
        42
    );
}
#[test]
fn i32_div_s_neg() {
    assert_eq!(
        run(|c| {
            push(c, -84);
            push(c, 2);
            c.emit_op(Op::I32_DIV_S, 0);
        })
        .as_i32(),
        -42
    );
}
#[test]
fn i32_div_u() {
    assert_eq!(
        run(|c| {
            push(c, 84);
            push(c, 2);
            c.emit_op(Op::I32_DIV_U, 0);
        })
        .as_i32(),
        42
    );
}
#[test]
fn i32_rem_s() {
    assert_eq!(
        run(|c| {
            push(c, 85);
            push(c, 2);
            c.emit_op(Op::I32_REM_S, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i32_rem_s_neg() {
    assert_eq!(
        run(|c| {
            push(c, -7);
            push(c, 3);
            c.emit_op(Op::I32_REM_S, 0);
        })
        .as_i32(),
        -1
    );
}
#[test]
fn i32_rem_u() {
    assert_eq!(
        run(|c| {
            push(c, 7);
            push(c, 3);
            c.emit_op(Op::I32_REM_U, 0);
        })
        .as_i32(),
        1
    );
}

// ── i32 bitwise ───────────────────────────────────────────────────────────

#[test]
fn i32_and() {
    assert_eq!(
        run(|c| {
            push(c, 0b1100);
            push(c, 0b1010);
            c.emit_op(Op::I32_AND, 0);
        })
        .as_i32(),
        0b1000
    );
}
#[test]
fn i32_or() {
    assert_eq!(
        run(|c| {
            push(c, 0b1100);
            push(c, 0b1010);
            c.emit_op(Op::I32_OR, 0);
        })
        .as_i32(),
        0b1110
    );
}
#[test]
fn i32_xor() {
    assert_eq!(
        run(|c| {
            push(c, 0b1100);
            push(c, 0b1010);
            c.emit_op(Op::I32_XOR, 0);
        })
        .as_i32(),
        0b0110
    );
}
#[test]
fn i32_shl() {
    assert_eq!(
        run(|c| {
            push(c, 1);
            push(c, 3);
            c.emit_op(Op::I32_SHL, 0);
        })
        .as_i32(),
        8
    );
}
#[test]
fn i32_shr_s() {
    assert_eq!(
        run(|c| {
            push(c, -8);
            push(c, 1);
            c.emit_op(Op::I32_SHR_S, 0);
        })
        .as_i32(),
        -4
    );
}
#[test]
fn i32_shr_u() {
    assert_eq!(
        run(|c| {
            push(c, -1);
            push(c, 1);
            c.emit_op(Op::I32_SHR_U, 0);
        })
        .as_i32(),
        i32::MAX
    );
}
#[test]
fn i32_rotl() {
    assert_eq!(
        run(|c| {
            push(c, 0x80000000u32 as i32);
            push(c, 1);
            c.emit_op(Op::I32_ROTL, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i32_rotr() {
    assert_eq!(
        run(|c| {
            push(c, 1);
            push(c, 1);
            c.emit_op(Op::I32_ROTR, 0);
        })
        .as_i32(),
        0x80000000u32 as i32
    );
}

// ── Spec-required trap edge cases ─────────────────────────────────────────

fn run_err(emit: impl FnOnce(&mut Chunk)) -> String {
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    VM::new().run(vec![c]).unwrap_err().to_string()
}

#[test]
fn i32_div_s_by_zero_traps() {
    assert!(
        run_err(|c| {
            push(c, 1);
            push(c, 0);
            c.emit_op(Op::I32_DIV_S, 0);
        })
        .contains("trap")
    );
}
#[test]
fn i32_div_u_by_zero_traps() {
    assert!(
        run_err(|c| {
            push(c, 1);
            push(c, 0);
            c.emit_op(Op::I32_DIV_U, 0);
        })
        .contains("trap")
    );
}
#[test]
fn i32_rem_s_by_zero_traps() {
    assert!(
        run_err(|c| {
            push(c, 1);
            push(c, 0);
            c.emit_op(Op::I32_REM_S, 0);
        })
        .contains("trap")
    );
}
#[test]
fn i32_rem_u_by_zero_traps() {
    assert!(
        run_err(|c| {
            push(c, 1u32 as i32);
            push(c, 0);
            c.emit_op(Op::I32_REM_U, 0);
        })
        .contains("trap")
    );
}
#[test]
fn i32_div_s_min_neg1_traps() {
    assert!(
        run_err(|c| {
            push(c, i32::MIN);
            push(c, -1);
            c.emit_op(Op::I32_DIV_S, 0);
        })
        .contains("trap")
    );
}
#[test]
fn i32_rem_s_min_neg1_is_zero() {
    assert_eq!(
        run(|c| {
            push(c, i32::MIN);
            push(c, -1);
            c.emit_op(Op::I32_REM_S, 0);
        })
        .as_i32(),
        0
    );
}
#[test]
fn i32_shl_by_32_same_as_0() {
    assert_eq!(
        run(|c| {
            push(c, 1);
            push(c, 32);
            c.emit_op(Op::I32_SHL, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i32_shr_s_by_32_same_as_0() {
    assert_eq!(
        run(|c| {
            push(c, -1);
            push(c, 32);
            c.emit_op(Op::I32_SHR_S, 0);
        })
        .as_i32(),
        -1
    );
}

#[test]
fn i32_shr_u_by_32_same_as_0() {
    assert_eq!(
        run(|c| {
            push(c, -1);
            push(c, 32);
            c.emit_op(Op::I32_SHR_U, 0);
        })
        .as_i32(),
        -1
    );
}

#[test]
fn i32_rotl_by_32_same_as_0() {
    assert_eq!(
        run(|c| {
            push(c, 0x1234_5678);
            push(c, 32);
            c.emit_op(Op::I32_ROTL, 0);
        })
        .as_i32(),
        0x1234_5678
    );
}

#[test]
fn i32_rotr_by_33_same_as_1() {
    assert_eq!(
        run(|c| {
            push(c, 2);
            push(c, 33);
            c.emit_op(Op::I32_ROTR, 0);
        })
        .as_i32(),
        1
    );
}
