//! Tests for the sign-extension-ops WASM proposal.
//! Spec: `proposals/sign-extension-ops/`, opcodes 0xC0–0xC4.

use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

fn run(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut chunk = Chunk::new("<script>");
    emit(&mut chunk);
    chunk.emit_op(Op::RETURN, 0);
    VM::new().run(vec![chunk]).expect("VM execution failed")
}

fn push_i32(c: &mut Chunk, v: i32) {
    c.emit_i32_const(v, 0);
}

fn push_i64(c: &mut Chunk, v: i64) {
    c.emit_i64_const(v, 0);
}

// ── i32.extend8_s (0xC0) ─────────────────────────────────────────────

#[test]
fn i32_extend8_s_zero() {
    assert_eq!(
        run(|c| {
            push_i32(c, 0);
            c.emit_op(Op::I32_EXTEND8_S, 0);
        })
        .as_i32(),
        0
    );
}

#[test]
fn i32_extend8_s_positive_max() {
    // 0x7f = 127 — largest positive i8; high bit clear, no sign extension
    assert_eq!(
        run(|c| {
            push_i32(c, 0x7f);
            c.emit_op(Op::I32_EXTEND8_S, 0);
        })
        .as_i32(),
        127
    );
}

#[test]
fn i32_extend8_s_sign_bit_set() {
    // 0x80 = bit 7 set → sign-extends to -128
    assert_eq!(
        run(|c| {
            push_i32(c, 0x80);
            c.emit_op(Op::I32_EXTEND8_S, 0);
        })
        .as_i32(),
        -128
    );
}

#[test]
fn i32_extend8_s_all_ones() {
    // 0xff as i8 = -1 → sign-extends to -1
    assert_eq!(
        run(|c| {
            push_i32(c, 0xff);
            c.emit_op(Op::I32_EXTEND8_S, 0);
        })
        .as_i32(),
        -1
    );
}

#[test]
fn i32_extend8_s_high_bits_discarded() {
    // Upper 24 bits are ignored; only bits 0–7 matter
    assert_eq!(
        run(|c| {
            push_i32(c, 0x0123_4500);
            c.emit_op(Op::I32_EXTEND8_S, 0);
        })
        .as_i32(),
        0
    );
    assert_eq!(
        run(|c| {
            push_i32(c, 0xfedc_ba80u32 as i32);
            c.emit_op(Op::I32_EXTEND8_S, 0);
        })
        .as_i32(),
        -0x80
    );
}

#[test]
fn i32_extend8_s_minus_one_input() {
    assert_eq!(
        run(|c| {
            push_i32(c, -1);
            c.emit_op(Op::I32_EXTEND8_S, 0);
        })
        .as_i32(),
        -1
    );
}

// ── i32.extend16_s (0xC1) ────────────────────────────────────────────

#[test]
fn i32_extend16_s_zero() {
    assert_eq!(
        run(|c| {
            push_i32(c, 0);
            c.emit_op(Op::I32_EXTEND16_S, 0);
        })
        .as_i32(),
        0
    );
}

#[test]
fn i32_extend16_s_positive_max() {
    // 0x7fff = 32767 — largest positive i16
    assert_eq!(
        run(|c| {
            push_i32(c, 0x7fff);
            c.emit_op(Op::I32_EXTEND16_S, 0);
        })
        .as_i32(),
        32767
    );
}

#[test]
fn i32_extend16_s_sign_bit_set() {
    // 0x8000 — bit 15 set → sign-extends to -32768
    assert_eq!(
        run(|c| {
            push_i32(c, 0x8000);
            c.emit_op(Op::I32_EXTEND16_S, 0);
        })
        .as_i32(),
        -32768
    );
}

#[test]
fn i32_extend16_s_all_ones() {
    assert_eq!(
        run(|c| {
            push_i32(c, 0xffff);
            c.emit_op(Op::I32_EXTEND16_S, 0);
        })
        .as_i32(),
        -1
    );
}

#[test]
fn i32_extend16_s_high_bits_discarded() {
    assert_eq!(
        run(|c| {
            push_i32(c, 0x0123_0000);
            c.emit_op(Op::I32_EXTEND16_S, 0);
        })
        .as_i32(),
        0
    );
    assert_eq!(
        run(|c| {
            push_i32(c, 0xfedc_8000u32 as i32);
            c.emit_op(Op::I32_EXTEND16_S, 0);
        })
        .as_i32(),
        -0x8000
    );
}

#[test]
fn i32_extend16_s_minus_one_input() {
    assert_eq!(
        run(|c| {
            push_i32(c, -1);
            c.emit_op(Op::I32_EXTEND16_S, 0);
        })
        .as_i32(),
        -1
    );
}

// ── i64.extend8_s (0xC2) ─────────────────────────────────────────────

#[test]
fn i64_extend8_s_zero() {
    assert_eq!(
        run(|c| {
            push_i64(c, 0);
            c.emit_op(Op::I64_EXTEND8_S, 0);
        })
        .as_i64(),
        0
    );
}

#[test]
fn i64_extend8_s_positive_max() {
    assert_eq!(
        run(|c| {
            push_i64(c, 0x7f);
            c.emit_op(Op::I64_EXTEND8_S, 0);
        })
        .as_i64(),
        127
    );
}

#[test]
fn i64_extend8_s_sign_bit_set() {
    assert_eq!(
        run(|c| {
            push_i64(c, 0x80);
            c.emit_op(Op::I64_EXTEND8_S, 0);
        })
        .as_i64(),
        -128
    );
}

#[test]
fn i64_extend8_s_all_ones() {
    assert_eq!(
        run(|c| {
            push_i64(c, 0xff);
            c.emit_op(Op::I64_EXTEND8_S, 0);
        })
        .as_i64(),
        -1
    );
}

#[test]
fn i64_extend8_s_high_bits_discarded() {
    assert_eq!(
        run(|c| {
            push_i64(c, 0x0123456789abcd00);
            c.emit_op(Op::I64_EXTEND8_S, 0);
        })
        .as_i64(),
        0
    );
    assert_eq!(
        run(|c| {
            push_i64(c, -0x80_i64);
            c.emit_op(Op::I64_EXTEND8_S, 0);
        })
        .as_i64(),
        -0x80
    );
}

#[test]
fn i64_extend8_s_minus_one_input() {
    assert_eq!(
        run(|c| {
            push_i64(c, -1);
            c.emit_op(Op::I64_EXTEND8_S, 0);
        })
        .as_i64(),
        -1
    );
}

// ── i64.extend16_s (0xC3) ────────────────────────────────────────────

#[test]
fn i64_extend16_s_zero() {
    assert_eq!(
        run(|c| {
            push_i64(c, 0);
            c.emit_op(Op::I64_EXTEND16_S, 0);
        })
        .as_i64(),
        0
    );
}

#[test]
fn i64_extend16_s_positive_max() {
    assert_eq!(
        run(|c| {
            push_i64(c, 0x7fff);
            c.emit_op(Op::I64_EXTEND16_S, 0);
        })
        .as_i64(),
        32767
    );
}

#[test]
fn i64_extend16_s_sign_bit_set() {
    assert_eq!(
        run(|c| {
            push_i64(c, 0x8000);
            c.emit_op(Op::I64_EXTEND16_S, 0);
        })
        .as_i64(),
        -32768
    );
}

#[test]
fn i64_extend16_s_all_ones() {
    assert_eq!(
        run(|c| {
            push_i64(c, 0xffff);
            c.emit_op(Op::I64_EXTEND16_S, 0);
        })
        .as_i64(),
        -1
    );
}

#[test]
fn i64_extend16_s_high_bits_discarded() {
    assert_eq!(
        run(|c| {
            push_i64(c, 0x123456789abc0000_u64 as i64);
            c.emit_op(Op::I64_EXTEND16_S, 0);
        })
        .as_i64(),
        0
    );
    assert_eq!(
        run(|c| {
            push_i64(c, 0xfedcba9876548000_u64 as i64);
            c.emit_op(Op::I64_EXTEND16_S, 0);
        })
        .as_i64(),
        -0x8000
    );
}

#[test]
fn i64_extend16_s_minus_one_input() {
    assert_eq!(
        run(|c| {
            push_i64(c, -1);
            c.emit_op(Op::I64_EXTEND16_S, 0);
        })
        .as_i64(),
        -1
    );
}

// ── i64.extend32_s (0xC4) ────────────────────────────────────────────

#[test]
fn i64_extend32_s_zero() {
    assert_eq!(
        run(|c| {
            push_i64(c, 0);
            c.emit_op(Op::I64_EXTEND32_S, 0);
        })
        .as_i64(),
        0
    );
}

#[test]
fn i64_extend32_s_positive_values() {
    assert_eq!(
        run(|c| {
            push_i64(c, 0x7fff);
            c.emit_op(Op::I64_EXTEND32_S, 0);
        })
        .as_i64(),
        32767
    );
    assert_eq!(
        run(|c| {
            push_i64(c, 0x7fff_ffff);
            c.emit_op(Op::I64_EXTEND32_S, 0);
        })
        .as_i64(),
        i32::MAX as i64
    );
}

#[test]
fn i64_extend32_s_sign_bit_set() {
    // 0x80000000 — bit 31 set → sign-extends to i32::MIN
    assert_eq!(
        run(|c| {
            push_i64(c, 0x8000_0000);
            c.emit_op(Op::I64_EXTEND32_S, 0);
        })
        .as_i64(),
        i32::MIN as i64
    );
}

#[test]
fn i64_extend32_s_all_ones_32bit() {
    assert_eq!(
        run(|c| {
            push_i64(c, 0xffff_ffff);
            c.emit_op(Op::I64_EXTEND32_S, 0);
        })
        .as_i64(),
        -1
    );
}

#[test]
fn i64_extend32_s_high_bits_discarded() {
    assert_eq!(
        run(|c| {
            push_i64(c, 0x01234567_00000000);
            c.emit_op(Op::I64_EXTEND32_S, 0);
        })
        .as_i64(),
        0
    );
    assert_eq!(
        run(|c| {
            push_i64(c, 0xfedcba98_80000000_u64 as i64);
            c.emit_op(Op::I64_EXTEND32_S, 0);
        })
        .as_i64(),
        i32::MIN as i64
    );
}

#[test]
fn i64_extend32_s_minus_one_input() {
    assert_eq!(
        run(|c| {
            push_i64(c, -1);
            c.emit_op(Op::I64_EXTEND32_S, 0);
        })
        .as_i64(),
        -1
    );
}
