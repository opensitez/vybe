//! Tests for all i64 instructions from the WASM spec (§5.3 numeric).
//! Covers: const, eqz, comparisons, clz/ctz/popcnt, arithmetic, bitwise.

use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

fn run(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    VM::new().run(vec![c]).expect("run failed")
}

fn push(c: &mut Chunk, v: i64) {
    let k = c.add_constant(Value::I64(v));
    c.emit_op_u16(Op::CONST, k, 0);
}

// ── i64.const ────────────────────────────────────────────────────────────

#[test]
fn i64_const_zero() {
    assert_eq!(run(|c| push(c, 0)).as_i64(), 0);
}
#[test]
fn i64_const_large() {
    assert_eq!(run(|c| push(c, i64::MAX)).as_i64(), i64::MAX);
}
#[test]
fn i64_const_negative() {
    assert_eq!(run(|c| push(c, -1)).as_i64(), -1);
}

// ── i64.eqz ──────────────────────────────────────────────────────────────

#[test]
fn i64_eqz_zero() {
    assert_eq!(
        run(|c| {
            push(c, 0);
            c.emit_op(Op::I64_EQZ, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i64_eqz_nonzero() {
    assert_eq!(
        run(|c| {
            push(c, 1);
            c.emit_op(Op::I64_EQZ, 0);
        })
        .as_i32(),
        0
    );
}

// ── i64 comparisons (signed) ──────────────────────────────────────────────

#[test]
fn i64_eq_true() {
    assert_eq!(
        run(|c| {
            push(c, 7);
            push(c, 7);
            c.emit_op(Op::I64_EQ, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i64_eq_false() {
    assert_eq!(
        run(|c| {
            push(c, 1);
            push(c, 2);
            c.emit_op(Op::I64_EQ, 0);
        })
        .as_i32(),
        0
    );
}
#[test]
fn i64_ne_true() {
    assert_eq!(
        run(|c| {
            push(c, 1);
            push(c, 2);
            c.emit_op(Op::I64_NE, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i64_ne_false() {
    assert_eq!(
        run(|c| {
            push(c, 7);
            push(c, 7);
            c.emit_op(Op::I64_NE, 0);
        })
        .as_i32(),
        0
    );
}
#[test]
fn i64_lt_s_true() {
    assert_eq!(
        run(|c| {
            push(c, -1);
            push(c, 0);
            c.emit_op(Op::I64_LT_S, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i64_lt_s_false() {
    assert_eq!(
        run(|c| {
            push(c, 1);
            push(c, 0);
            c.emit_op(Op::I64_LT_S, 0);
        })
        .as_i32(),
        0
    );
}
#[test]
fn i64_gt_s_true() {
    assert_eq!(
        run(|c| {
            push(c, 5);
            push(c, 3);
            c.emit_op(Op::I64_GT_S, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i64_gt_s_false() {
    assert_eq!(
        run(|c| {
            push(c, 3);
            push(c, 5);
            c.emit_op(Op::I64_GT_S, 0);
        })
        .as_i32(),
        0
    );
}
#[test]
fn i64_le_s_equal() {
    assert_eq!(
        run(|c| {
            push(c, 3);
            push(c, 3);
            c.emit_op(Op::I64_LE_S, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i64_ge_s_equal() {
    assert_eq!(
        run(|c| {
            push(c, 3);
            push(c, 3);
            c.emit_op(Op::I64_GE_S, 0);
        })
        .as_i32(),
        1
    );
}

// ── i64 comparisons (unsigned) ────────────────────────────────────────────

#[test]
fn i64_lt_u_unsigned() {
    assert_eq!(
        run(|c| {
            push(c, 1);
            push(c, -1);
            c.emit_op(Op::I64_LT_U, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i64_gt_u_unsigned() {
    assert_eq!(
        run(|c| {
            push(c, -1);
            push(c, 1);
            c.emit_op(Op::I64_GT_U, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i64_le_u_equal() {
    assert_eq!(
        run(|c| {
            push(c, 5);
            push(c, 5);
            c.emit_op(Op::I64_LE_U, 0);
        })
        .as_i32(),
        1
    );
}
#[test]
fn i64_ge_u_unsigned() {
    assert_eq!(
        run(|c| {
            push(c, -1);
            push(c, 0);
            c.emit_op(Op::I64_GE_U, 0);
        })
        .as_i32(),
        1
    );
}

// ── i64 bit counting ─────────────────────────────────────────────────────

#[test]
fn i64_clz_zero() {
    assert_eq!(
        run(|c| {
            push(c, 0);
            c.emit_op(Op::I64_CLZ, 0);
        })
        .as_i64(),
        64
    );
}
#[test]
fn i64_clz_one() {
    assert_eq!(
        run(|c| {
            push(c, 1);
            c.emit_op(Op::I64_CLZ, 0);
        })
        .as_i64(),
        63
    );
}
#[test]
fn i64_ctz_zero() {
    assert_eq!(
        run(|c| {
            push(c, 0);
            c.emit_op(Op::I64_CTZ, 0);
        })
        .as_i64(),
        64
    );
}
#[test]
fn i64_ctz_four() {
    assert_eq!(
        run(|c| {
            push(c, 4);
            c.emit_op(Op::I64_CTZ, 0);
        })
        .as_i64(),
        2
    );
}
#[test]
fn i64_popcnt_zero() {
    assert_eq!(
        run(|c| {
            push(c, 0);
            c.emit_op(Op::I64_POPCNT, 0);
        })
        .as_i64(),
        0
    );
}
#[test]
fn i64_popcnt_ones() {
    assert_eq!(
        run(|c| {
            push(c, -1);
            c.emit_op(Op::I64_POPCNT, 0);
        })
        .as_i64(),
        64
    );
}

// ── i64 arithmetic ────────────────────────────────────────────────────────

#[test]
fn i64_add() {
    assert_eq!(
        run(|c| {
            push(c, 20);
            push(c, 22);
            c.emit_op(Op::I64_ADD, 0);
        })
        .as_i64(),
        42
    );
}
#[test]
fn i64_add_overflow() {
    assert_eq!(
        run(|c| {
            push(c, i64::MAX);
            push(c, 1);
            c.emit_op(Op::I64_ADD, 0);
        })
        .as_i64(),
        i64::MIN
    );
}
#[test]
fn i64_sub() {
    assert_eq!(
        run(|c| {
            push(c, 50);
            push(c, 8);
            c.emit_op(Op::I64_SUB, 0);
        })
        .as_i64(),
        42
    );
}
#[test]
fn i64_mul() {
    assert_eq!(
        run(|c| {
            push(c, 6);
            push(c, 7);
            c.emit_op(Op::I64_MUL, 0);
        })
        .as_i64(),
        42
    );
}
#[test]
fn i64_div_s() {
    assert_eq!(
        run(|c| {
            push(c, 84);
            push(c, 2);
            c.emit_op(Op::I64_DIV_S, 0);
        })
        .as_i64(),
        42
    );
}
#[test]
fn i64_div_u() {
    assert_eq!(
        run(|c| {
            push(c, 84);
            push(c, 2);
            c.emit_op(Op::I64_DIV_U, 0);
        })
        .as_i64(),
        42
    );
}
#[test]
fn i64_rem_s() {
    assert_eq!(
        run(|c| {
            push(c, 85);
            push(c, 2);
            c.emit_op(Op::I64_REM_S, 0);
        })
        .as_i64(),
        1
    );
}
#[test]
fn i64_rem_u() {
    assert_eq!(
        run(|c| {
            push(c, 7);
            push(c, 3);
            c.emit_op(Op::I64_REM_U, 0);
        })
        .as_i64(),
        1
    );
}

// ── i64 bitwise ───────────────────────────────────────────────────────────

#[test]
fn i64_and() {
    assert_eq!(
        run(|c| {
            push(c, 0b1100);
            push(c, 0b1010);
            c.emit_op(Op::I64_AND, 0);
        })
        .as_i64(),
        0b1000
    );
}
#[test]
fn i64_or() {
    assert_eq!(
        run(|c| {
            push(c, 0b1100);
            push(c, 0b1010);
            c.emit_op(Op::I64_OR, 0);
        })
        .as_i64(),
        0b1110
    );
}
#[test]
fn i64_xor() {
    assert_eq!(
        run(|c| {
            push(c, 0b1100);
            push(c, 0b1010);
            c.emit_op(Op::I64_XOR, 0);
        })
        .as_i64(),
        0b0110
    );
}
#[test]
fn i64_shl() {
    assert_eq!(
        run(|c| {
            push(c, 1);
            push(c, 3);
            c.emit_op(Op::I64_SHL, 0);
        })
        .as_i64(),
        8
    );
}
#[test]
fn i64_shr_s() {
    assert_eq!(
        run(|c| {
            push(c, -8);
            push(c, 1);
            c.emit_op(Op::I64_SHR_S, 0);
        })
        .as_i64(),
        -4
    );
}
#[test]
fn i64_shr_u() {
    assert_eq!(
        run(|c| {
            push(c, -1);
            push(c, 1);
            c.emit_op(Op::I64_SHR_U, 0);
        })
        .as_i64(),
        i64::MAX
    );
}
#[test]
fn i64_rotl() {
    let r = run(|c| {
        push(c, 0x8000000000000000u64 as i64);
        push(c, 1);
        c.emit_op(Op::I64_ROTL, 0);
    });
    assert_eq!(r.as_i64(), 1);
}
#[test]
fn i64_rotr() {
    let r = run(|c| {
        push(c, 1);
        push(c, 1);
        c.emit_op(Op::I64_ROTR, 0);
    });
    assert_eq!(r.as_i64(), 0x8000000000000000u64 as i64);
}

// ── Spec-required trap edge cases ─────────────────────────────────────────

fn run_err(emit: impl FnOnce(&mut Chunk)) -> String {
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    VM::new().run(vec![c]).unwrap_err().to_string()
}

fn push_i64(c: &mut Chunk, v: i64) {
    let k = c.add_constant(Value::I64(v));
    c.emit_op_u16(Op::CONST, k, 0);
}

#[test]
fn i64_div_s_by_zero_traps() {
    assert!(
        run_err(|c| {
            push_i64(c, 1);
            push_i64(c, 0);
            c.emit_op(Op::I64_DIV_S, 0);
        })
        .contains("trap")
    );
}
#[test]
fn i64_div_u_by_zero_traps() {
    assert!(
        run_err(|c| {
            push_i64(c, 1);
            push_i64(c, 0);
            c.emit_op(Op::I64_DIV_U, 0);
        })
        .contains("trap")
    );
}
#[test]
fn i64_rem_s_by_zero_traps() {
    assert!(
        run_err(|c| {
            push_i64(c, 1);
            push_i64(c, 0);
            c.emit_op(Op::I64_REM_S, 0);
        })
        .contains("trap")
    );
}
#[test]
fn i64_rem_u_by_zero_traps() {
    assert!(
        run_err(|c| {
            push_i64(c, 1);
            push_i64(c, 0);
            c.emit_op(Op::I64_REM_U, 0);
        })
        .contains("trap")
    );
}
#[test]
fn i64_div_s_min_neg1_traps() {
    assert!(
        run_err(|c| {
            push_i64(c, i64::MIN);
            push_i64(c, -1);
            c.emit_op(Op::I64_DIV_S, 0);
        })
        .contains("trap")
    );
}
#[test]
fn i64_rem_s_min_neg1_is_zero() {
    assert_eq!(
        run(|c| {
            push_i64(c, i64::MIN);
            push_i64(c, -1);
            c.emit_op(Op::I64_REM_S, 0);
        })
        .as_i64(),
        0
    );
}
#[test]
fn i64_shl_by_64_same_as_0() {
    assert_eq!(
        run(|c| {
            push_i64(c, 1);
            push_i64(c, 64);
            c.emit_op(Op::I64_SHL, 0);
        })
        .as_i64(),
        1
    );
}

#[test]
fn i64_shr_s_by_64_same_as_0() {
    assert_eq!(
        run(|c| {
            push_i64(c, -1);
            push_i64(c, 64);
            c.emit_op(Op::I64_SHR_S, 0);
        })
        .as_i64(),
        -1
    );
}

#[test]
fn i64_shr_u_by_64_same_as_0() {
    assert_eq!(
        run(|c| {
            push_i64(c, -1);
            push_i64(c, 64);
            c.emit_op(Op::I64_SHR_U, 0);
        })
        .as_i64(),
        -1
    );
}

#[test]
fn i64_rotl_by_64_same_as_0() {
    assert_eq!(
        run(|c| {
            push_i64(c, 0x1234_5678_9abc_def0_u64 as i64);
            push_i64(c, 64);
            c.emit_op(Op::I64_ROTL, 0);
        })
        .as_i64(),
        0x1234_5678_9abc_def0_u64 as i64
    );
}

#[test]
fn i64_rotr_by_65_same_as_1() {
    assert_eq!(
        run(|c| {
            push_i64(c, 2);
            push_i64(c, 65);
            c.emit_op(Op::I64_ROTR, 0);
        })
        .as_i64(),
        1
    );
}
