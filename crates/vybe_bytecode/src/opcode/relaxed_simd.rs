//! Relaxed-SIMD proposal opcodes.
//!
//! Spec: `proposals/spec/proposals/relaxed-simd/` and WASM 3.0
//! `5.3-binary.instructions.spectec`.
//!
//! These use the `0xFD` SIMD prefix with sub-values `0x100..=0x113`
//! (256..=275), matching the spec exactly. The u32 `Op` representation
//! can hold the full range.
//!
//! The "relaxed" in relaxed-SIMD means implementations are allowed to
//! choose one of several behaviors on edge cases (NaN handling, lane
//! selection bit, FMA fusion). Our VM implements each op deterministically
//! so results are reproducible across platforms.

use super::Op;
use super::opcode_category;

impl Op {
    pub const I8X16_RELAXED_SWIZZLE: Op = Op::new(0xFD, 256);
    pub const I32X4_RELAXED_TRUNC_F32X4_S: Op = Op::new(0xFD, 257);
    pub const I32X4_RELAXED_TRUNC_F32X4_U: Op = Op::new(0xFD, 258);
    pub const I32X4_RELAXED_TRUNC_F64X2_S_ZERO: Op = Op::new(0xFD, 259);
    pub const I32X4_RELAXED_TRUNC_F64X2_U_ZERO: Op = Op::new(0xFD, 260);
    pub const F32X4_RELAXED_MADD: Op = Op::new(0xFD, 261);
    pub const F32X4_RELAXED_NMADD: Op = Op::new(0xFD, 262);
    pub const F64X2_RELAXED_MADD: Op = Op::new(0xFD, 263);
    pub const F64X2_RELAXED_NMADD: Op = Op::new(0xFD, 264);
    pub const I8X16_RELAXED_LANESELECT: Op = Op::new(0xFD, 265);
    pub const I16X8_RELAXED_LANESELECT: Op = Op::new(0xFD, 266);
    pub const I32X4_RELAXED_LANESELECT: Op = Op::new(0xFD, 267);
    pub const I64X2_RELAXED_LANESELECT: Op = Op::new(0xFD, 268);
    pub const F32X4_RELAXED_MIN: Op = Op::new(0xFD, 269);
    pub const F32X4_RELAXED_MAX: Op = Op::new(0xFD, 270);
    pub const F64X2_RELAXED_MIN: Op = Op::new(0xFD, 271);
    pub const F64X2_RELAXED_MAX: Op = Op::new(0xFD, 272);
    pub const I16X8_RELAXED_Q15MULR_S: Op = Op::new(0xFD, 273);
    pub const I16X8_RELAXED_DOT_I8X16_I7X16_S: Op = Op::new(0xFD, 274);
    pub const I32X4_RELAXED_DOT_I8X16_I7X16_ADD_S: Op = Op::new(0xFD, 275);
}

opcode_category! {
    [256] i8x16_relaxed_swizzle                => None, "i8x16.relaxed_swizzle";
    [257] i32x4_relaxed_trunc_f32x4_s          => None, "i32x4.relaxed_trunc_f32x4_s";
    [258] i32x4_relaxed_trunc_f32x4_u          => None, "i32x4.relaxed_trunc_f32x4_u";
    [259] i32x4_relaxed_trunc_f64x2_s_zero     => None, "i32x4.relaxed_trunc_f64x2_s_zero";
    [260] i32x4_relaxed_trunc_f64x2_u_zero     => None, "i32x4.relaxed_trunc_f64x2_u_zero";
    [261] f32x4_relaxed_madd                   => None, "f32x4.relaxed_madd";
    [262] f32x4_relaxed_nmadd                  => None, "f32x4.relaxed_nmadd";
    [263] f64x2_relaxed_madd                   => None, "f64x2.relaxed_madd";
    [264] f64x2_relaxed_nmadd                  => None, "f64x2.relaxed_nmadd";
    [265] i8x16_relaxed_laneselect             => None, "i8x16.relaxed_laneselect";
    [266] i16x8_relaxed_laneselect             => None, "i16x8.relaxed_laneselect";
    [267] i32x4_relaxed_laneselect             => None, "i32x4.relaxed_laneselect";
    [268] i64x2_relaxed_laneselect             => None, "i64x2.relaxed_laneselect";
    [269] f32x4_relaxed_min                    => None, "f32x4.relaxed_min";
    [270] f32x4_relaxed_max                    => None, "f32x4.relaxed_max";
    [271] f64x2_relaxed_min                    => None, "f64x2.relaxed_min";
    [272] f64x2_relaxed_max                    => None, "f64x2.relaxed_max";
    [273] i16x8_relaxed_q15mulr_s              => None, "i16x8.relaxed_q15mulr_s";
    [274] i16x8_relaxed_dot_i8x16_i7x16_s      => None, "i16x8.relaxed_dot_i8x16_i7x16_s";
    [275] i32x4_relaxed_dot_i8x16_i7x16_add_s  => None, "i32x4.relaxed_dot_i8x16_i7x16_add_s";
}
