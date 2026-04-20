//! Relaxed-SIMD proposal opcodes.
//!
//! Spec: `proposals/spec/proposals/relaxed-simd/`.
//! The proposal assigns these to the `0xFD` SIMD prefix but at sub-values
//! `0x100..=0x113` — outside the u8 range used by the base `Op(u16)`
//! representation. We sidestep that by storing them under an internal
//! prefix byte (`0xDD`) with sub-values `0x00..=0x13`, and the WASM
//! emitter rewrites them back to spec-correct `0xFD + LEB128(0x100+sub)`
//! at binary time (see `wasm/code.rs`).
//!
//! The "relaxed" in relaxed-SIMD means implementations are allowed to
//! choose one of several behaviors on edge cases (NaN handling, lane
//! selection bit, FMA fusion). Our VM implements each op deterministically
//! so results are reproducible across platforms; real hardware codegen
//! would pick the lane-cheap variant.
//!
//! The first four (`*_relaxed_madd`, `*_relaxed_nmadd`) previously lived
//! on the 0xFD prefix with spec-wrong sub-values — those collided with
//! MVP convert opcodes. Migration preserves the `Op::*` constant names so
//! call sites don't change.

use super::Op;
use super::opcode_category;

impl Op {
    // Internal prefix: 0xDD. Sub values map 1:1 to spec sub minus 0x100.
    pub const I8X16_RELAXED_SWIZZLE: Op               = Op::new(0xDD, 0x00);
    pub const I32X4_RELAXED_TRUNC_F32X4_S: Op         = Op::new(0xDD, 0x01);
    pub const I32X4_RELAXED_TRUNC_F32X4_U: Op         = Op::new(0xDD, 0x02);
    pub const I32X4_RELAXED_TRUNC_F64X2_S_ZERO: Op    = Op::new(0xDD, 0x03);
    pub const I32X4_RELAXED_TRUNC_F64X2_U_ZERO: Op    = Op::new(0xDD, 0x04);
    pub const F32X4_RELAXED_MADD: Op                  = Op::new(0xDD, 0x05);
    pub const F32X4_RELAXED_NMADD: Op                 = Op::new(0xDD, 0x06);
    pub const F64X2_RELAXED_MADD: Op                  = Op::new(0xDD, 0x07);
    pub const F64X2_RELAXED_NMADD: Op                 = Op::new(0xDD, 0x08);
    pub const I8X16_RELAXED_LANESELECT: Op            = Op::new(0xDD, 0x09);
    pub const I16X8_RELAXED_LANESELECT: Op            = Op::new(0xDD, 0x0A);
    pub const I32X4_RELAXED_LANESELECT: Op            = Op::new(0xDD, 0x0B);
    pub const I64X2_RELAXED_LANESELECT: Op            = Op::new(0xDD, 0x0C);
    pub const F32X4_RELAXED_MIN: Op                   = Op::new(0xDD, 0x0D);
    pub const F32X4_RELAXED_MAX: Op                   = Op::new(0xDD, 0x0E);
    pub const F64X2_RELAXED_MIN: Op                   = Op::new(0xDD, 0x0F);
    pub const F64X2_RELAXED_MAX: Op                   = Op::new(0xDD, 0x10);
    pub const I16X8_RELAXED_Q15MULR_S: Op             = Op::new(0xDD, 0x11);
    pub const I16X8_RELAXED_DOT_I8X16_I7X16_S: Op     = Op::new(0xDD, 0x12);
    pub const I32X4_RELAXED_DOT_I8X16_I7X16_ADD_S: Op = Op::new(0xDD, 0x13);
}

/// Map an internal `0xDD` sub-value to the spec SIMD sub-opcode
/// (`0x100 + sub`). Used by the WASM emitter.
#[inline]
pub fn spec_sub(internal_sub: u8) -> u32 {
    0x100 + internal_sub as u32
}

opcode_category! {
    [0x00] i8x16_relaxed_swizzle                => None, "i8x16.relaxed_swizzle";
    [0x01] i32x4_relaxed_trunc_f32x4_s          => None, "i32x4.relaxed_trunc_f32x4_s";
    [0x02] i32x4_relaxed_trunc_f32x4_u          => None, "i32x4.relaxed_trunc_f32x4_u";
    [0x03] i32x4_relaxed_trunc_f64x2_s_zero     => None, "i32x4.relaxed_trunc_f64x2_s_zero";
    [0x04] i32x4_relaxed_trunc_f64x2_u_zero     => None, "i32x4.relaxed_trunc_f64x2_u_zero";
    [0x05] f32x4_relaxed_madd                   => None, "f32x4.relaxed_madd";
    [0x06] f32x4_relaxed_nmadd                  => None, "f32x4.relaxed_nmadd";
    [0x07] f64x2_relaxed_madd                   => None, "f64x2.relaxed_madd";
    [0x08] f64x2_relaxed_nmadd                  => None, "f64x2.relaxed_nmadd";
    [0x09] i8x16_relaxed_laneselect             => None, "i8x16.relaxed_laneselect";
    [0x0A] i16x8_relaxed_laneselect             => None, "i16x8.relaxed_laneselect";
    [0x0B] i32x4_relaxed_laneselect             => None, "i32x4.relaxed_laneselect";
    [0x0C] i64x2_relaxed_laneselect             => None, "i64x2.relaxed_laneselect";
    [0x0D] f32x4_relaxed_min                    => None, "f32x4.relaxed_min";
    [0x0E] f32x4_relaxed_max                    => None, "f32x4.relaxed_max";
    [0x0F] f64x2_relaxed_min                    => None, "f64x2.relaxed_min";
    [0x10] f64x2_relaxed_max                    => None, "f64x2.relaxed_max";
    [0x11] i16x8_relaxed_q15mulr_s              => None, "i16x8.relaxed_q15mulr_s";
    [0x12] i16x8_relaxed_dot_i8x16_i7x16_s      => None, "i16x8.relaxed_dot_i8x16_i7x16_s";
    [0x13] i32x4_relaxed_dot_i8x16_i7x16_add_s  => None, "i32x4.relaxed_dot_i8x16_i7x16_add_s";
}
