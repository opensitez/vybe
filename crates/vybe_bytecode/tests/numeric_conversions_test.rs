//! Tests for all numeric conversion instructions from the WASM spec (§5.3).
//! Covers: wrap, extend, trunc (trapping), convert, promote, demote, reinterpret.

use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};

fn run(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    VM::new().run(vec![c]).expect("run failed")
}

fn run_err(emit: impl FnOnce(&mut Chunk)) -> String {
    let mut c = Chunk::new("<script>");
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    VM::new().run(vec![c]).unwrap_err().to_string()
}

fn push_i32(c: &mut Chunk, v: i32) {
    let k = c.add_constant(Value::I32(v));
    c.emit_op_u16(Op::CONST, k, 0);
}
fn push_i64(c: &mut Chunk, v: i64) {
    let k = c.add_constant(Value::I64(v));
    c.emit_op_u16(Op::CONST, k, 0);
}
fn push_f64(c: &mut Chunk, v: f64) {
    let k = c.add_constant(Value::F64(v));
    c.emit_op_u16(Op::CONST, k, 0);
}

// ── i32.wrap_i64 (0xA7) ───────────────────────────────────────────────────

#[test]
fn i32_wrap_i64_basic() {
    assert_eq!(
        run(|c| {
            push_i64(c, 42);
            c.emit_op(Op::I32_WRAP_I64, 0);
        })
        .as_i32(),
        42
    );
}
#[test]
fn i32_wrap_i64_truncates() {
    assert_eq!(
        run(|c| {
            push_i64(c, 0x1_0000_0001);
            c.emit_op(Op::I32_WRAP_I64, 0);
        })
        .as_i32(),
        1
    );
}

// ── i32.trunc_f32_s (0xA8) / i32.trunc_f32_u (0xA9) ──────────────────────

#[test]
fn i32_trunc_f32_s_positive() {
    assert_eq!(
        run(|c| {
            push_f64(c, 3.7);
            c.emit_op(Op::I32_TRUNC_F32_S, 0);
        })
        .as_i32(),
        3
    );
}
#[test]
fn i32_trunc_f32_s_negative() {
    assert_eq!(
        run(|c| {
            push_f64(c, -2.9);
            c.emit_op(Op::I32_TRUNC_F32_S, 0);
        })
        .as_i32(),
        -2
    );
}
#[test]
fn i32_trunc_f32_s_overflow() {
    assert!(
        run_err(|c| {
            push_f64(c, 2_147_483_648.0);
            c.emit_op(Op::I32_TRUNC_F32_S, 0);
        })
        .contains("trap")
    );
}
#[test]
fn i32_trunc_f32_s_nan_traps() {
    assert!(
        run_err(|c| {
            push_f64(c, f64::NAN);
            c.emit_op(Op::I32_TRUNC_F32_S, 0);
        })
        .contains("trap")
    );
}
#[test]
fn i32_trunc_f32_u_positive() {
    assert_eq!(
        run(|c| {
            push_f64(c, 300.9);
            c.emit_op(Op::I32_TRUNC_F32_U, 0);
        })
        .as_i32() as u32,
        300
    );
}
#[test]
fn i32_trunc_f32_u_neg_traps() {
    assert!(
        run_err(|c| {
            push_f64(c, -1.0);
            c.emit_op(Op::I32_TRUNC_F32_U, 0);
        })
        .contains("trap")
    );
}

// ── i32.trunc_f64_s (0xAA) / i32.trunc_f64_u (0xAB) ──────────────────────

#[test]
fn i32_trunc_f64_s_positive() {
    assert_eq!(
        run(|c| {
            push_f64(c, 3.9);
            c.emit_op(Op::I32_FROM_F64, 0);
        })
        .as_i32(),
        3
    );
}
#[test]
fn i32_trunc_f64_s_negative() {
    assert_eq!(
        run(|c| {
            push_f64(c, -2.1);
            c.emit_op(Op::I32_FROM_F64, 0);
        })
        .as_i32(),
        -2
    );
}
#[test]
fn i32_trunc_f64_u_positive() {
    assert_eq!(
        run(|c| {
            push_f64(c, 65535.9);
            c.emit_op(Op::I32_TRUNC_F64_U, 0);
        })
        .as_i32() as u32,
        65535
    );
}
#[test]
fn i32_trunc_f64_u_overflow() {
    assert!(
        run_err(|c| {
            push_f64(c, 4_294_967_296.0);
            c.emit_op(Op::I32_TRUNC_F64_U, 0);
        })
        .contains("trap")
    );
}
#[test]
fn i32_trunc_f64_u_neg_traps() {
    assert!(
        run_err(|c| {
            push_f64(c, -0.5);
            c.emit_op(Op::I32_TRUNC_F64_U, 0);
        })
        .contains("trap")
    );
}

// ── i64.extend_i32_s (0xAC) / i64.extend_i32_u (0xAD) ────────────────────

#[test]
fn i64_extend_i32_s_positive() {
    assert_eq!(
        run(|c| {
            push_i32(c, 42);
            c.emit_op(Op::I64_EXTEND_I32_S, 0);
        })
        .as_i64(),
        42
    );
}
#[test]
fn i64_extend_i32_s_negative() {
    assert_eq!(
        run(|c| {
            push_i32(c, -1);
            c.emit_op(Op::I64_EXTEND_I32_S, 0);
        })
        .as_i64(),
        -1
    );
}
#[test]
fn i64_extend_i32_u_wraps() {
    assert_eq!(
        run(|c| {
            push_i32(c, -1);
            c.emit_op(Op::I64_EXTEND_I32_U, 0);
        })
        .as_i64(),
        4_294_967_295
    );
}

// ── i64.trunc_f32_s (0xAE) / i64.trunc_f32_u (0xAF) ──────────────────────

#[test]
fn i64_trunc_f32_s_positive() {
    assert_eq!(
        run(|c| {
            push_f64(c, 1_000_000.0);
            c.emit_op(Op::I64_TRUNC_F32_S, 0);
        })
        .as_i64(),
        1_000_000
    );
}
#[test]
fn i64_trunc_f32_s_negative() {
    assert_eq!(
        run(|c| {
            push_f64(c, -42.9);
            c.emit_op(Op::I64_TRUNC_F32_S, 0);
        })
        .as_i64(),
        -42
    );
}
#[test]
fn i64_trunc_f32_s_nan_traps() {
    assert!(
        run_err(|c| {
            push_f64(c, f64::NAN);
            c.emit_op(Op::I64_TRUNC_F32_S, 0);
        })
        .contains("trap")
    );
}
#[test]
fn i64_trunc_f32_u_positive() {
    assert!(
        run(|c| {
            push_f64(c, 4_000_000_000.0);
            c.emit_op(Op::I64_TRUNC_F32_U, 0);
        })
        .as_i64() as u64
            > 3_000_000_000
    );
}
#[test]
fn i64_trunc_f32_u_neg_traps() {
    assert!(
        run_err(|c| {
            push_f64(c, -1.0);
            c.emit_op(Op::I64_TRUNC_F32_U, 0);
        })
        .contains("trap")
    );
}

// ── i64.trunc_f64_s (0xB0) / i64.trunc_f64_u (0xB1) ──────────────────────

#[test]
fn i64_trunc_f64_s() {
    assert_eq!(
        run(|c| {
            push_f64(c, 9.9);
            c.emit_op(Op::I64_TRUNC_F64_S, 0);
        })
        .as_i64(),
        9
    );
}
#[test]
fn i64_trunc_f64_u() {
    assert_eq!(
        run(|c| {
            push_f64(c, 4_000_000_000.0);
            c.emit_op(Op::I64_TRUNC_F64_U, 0);
        })
        .as_i64() as u64,
        4_000_000_000
    );
}

// ── f32.convert_i32_s/u (0xB2-0xB3) ──────────────────────────────────────

#[test]
fn f32_convert_i32_s_pos() {
    assert_eq!(
        run(|c| {
            push_i32(c, 100);
            c.emit_op(Op::F32_CONVERT_I32_S, 0);
        })
        .as_f64() as f32,
        100.0
    );
}
#[test]
fn f32_convert_i32_s_neg() {
    assert_eq!(
        run(|c| {
            push_i32(c, -50);
            c.emit_op(Op::F32_CONVERT_I32_S, 0);
        })
        .as_f64() as f32,
        -50.0
    );
}
#[test]
fn f32_convert_i32_u() {
    assert_eq!(
        run(|c| {
            push_i32(c, i32::MIN);
            c.emit_op(Op::F32_CONVERT_I32_U, 0);
        })
        .as_f64() as f32,
        2_147_483_648.0
    );
}

// ── f32.convert_i64_s/u (0xB4-0xB5) ──────────────────────────────────────

#[test]
fn f32_convert_i64_s() {
    assert_eq!(
        run(|c| {
            push_i64(c, -1_000_000);
            c.emit_op(Op::F32_CONVERT_I64_S, 0);
        })
        .as_f64() as f32,
        -1_000_000.0
    );
}
#[test]
fn f32_convert_i64_u() {
    assert_eq!(
        run(|c| {
            push_i64(c, 4_000_000_000);
            c.emit_op(Op::F32_CONVERT_I64_U, 0);
        })
        .as_f64() as f32,
        4_000_000_000.0
    );
}

// ── f32.demote_f64 (0xB6) ─────────────────────────────────────────────────

#[test]
fn f32_demote_f64() {
    assert_eq!(
        run(|c| {
            push_f64(c, 3.14);
            c.emit_op(Op::F32_DEMOTE_F64, 0);
        })
        .as_f64() as f32,
        3.14f32
    );
}

// ── f64.convert_i32_s (0xB7) / f64.convert_i32_u (0xB8) ──────────────────

#[test]
fn f64_convert_i32_s_neg() {
    assert_eq!(
        run(|c| {
            push_i32(c, -42);
            c.emit_op(Op::F64_FROM_I32, 0);
        })
        .as_f64(),
        -42.0
    );
}
#[test]
fn f64_convert_i32_u() {
    assert_eq!(
        run(|c| {
            push_i32(c, -1);
            c.emit_op(Op::F64_CONVERT_I32_U, 0);
        })
        .as_f64(),
        4_294_967_295.0
    );
}

// ── f64.convert_i64_s/u (0xB9-0xBA) ──────────────────────────────────────

#[test]
fn f64_convert_i64_s() {
    assert_eq!(
        run(|c| {
            push_i64(c, -1);
            c.emit_op(Op::F64_CONVERT_I64_S, 0);
        })
        .as_f64(),
        -1.0
    );
}
#[test]
fn f64_convert_i64_u() {
    assert_eq!(
        run(|c| {
            push_i64(c, i64::MAX);
            c.emit_op(Op::F64_CONVERT_I64_U, 0);
        })
        .as_f64(),
        i64::MAX as u64 as f64
    );
}

// ── f64.promote_f32 (0xBB) ────────────────────────────────────────────────

#[test]
fn f64_promote_f32() {
    assert_eq!(
        run(|c| {
            push_f64(c, 1.0f32 as f64);
            c.emit_op(Op::F64_PROMOTE_F32, 0);
        })
        .as_f64(),
        1.0
    );
}

// ── reinterpret (0xBC-0xBF) ───────────────────────────────────────────────

#[test]
fn i32_reinterpret_f32() {
    // 1.0f32 bits = 0x3F800000
    let r = run(|c| {
        push_f64(c, 1.0f32 as f64);
        c.emit_op(Op::I32_REINTERPRET_F32, 0);
    });
    assert_eq!(r.as_i32(), 0x3F800000u32 as i32);
}
#[test]
fn i64_reinterpret_f64() {
    // 1.0f64 bits = 0x3FF0000000000000
    let r = run(|c| {
        push_f64(c, 1.0);
        c.emit_op(Op::I64_REINTERPRET_F64, 0);
    });
    assert_eq!(r.as_i64(), 0x3FF0000000000000u64 as i64);
}
#[test]
fn f32_reinterpret_i32() {
    // 0x3F800000 → 1.0f32
    let r = run(|c| {
        push_i32(c, 0x3F800000u32 as i32);
        c.emit_op(Op::F32_REINTERPRET_I32, 0);
    });
    assert_eq!(r.as_f64() as f32, 1.0f32);
}
#[test]
fn f64_reinterpret_i64() {
    let r = run(|c| {
        push_i64(c, 0x3FF0000000000000u64 as i64);
        c.emit_op(Op::F64_REINTERPRET_I64, 0);
    });
    assert_eq!(r.as_f64(), 1.0f64);
}
