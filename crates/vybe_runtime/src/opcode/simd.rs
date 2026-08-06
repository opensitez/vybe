//! SIMD proposal opcodes (prefix 0xFD).
//! Byte values match the WASM SIMD specification (sub-opcodes are LEB128 u32).
//! Sub-values 0x00–0xFF live here; 0x100–0x113 (relaxed-SIMD) are in
//! `relaxed_simd.rs` under the SAME spec 0xFD prefix — the u16 sub in `Op`
//! holds the full spec range.

use super::Op;
use super::opcode_category;

impl Op {
    // ── Memory ops ──────────────────────────────────────────────────────────
    pub const V128_LOAD: Op = Op::new(0xFD, 0x00);
    pub const V128_LOAD8X8_S: Op = Op::new(0xFD, 0x01);
    pub const V128_LOAD8X8_U: Op = Op::new(0xFD, 0x02);
    pub const V128_LOAD16X4_S: Op = Op::new(0xFD, 0x03);
    pub const V128_LOAD16X4_U: Op = Op::new(0xFD, 0x04);
    pub const V128_LOAD32X2_S: Op = Op::new(0xFD, 0x05);
    pub const V128_LOAD32X2_U: Op = Op::new(0xFD, 0x06);
    pub const V128_LOAD8_SPLAT: Op = Op::new(0xFD, 0x07);
    pub const V128_LOAD16_SPLAT: Op = Op::new(0xFD, 0x08);
    pub const V128_LOAD32_SPLAT: Op = Op::new(0xFD, 0x09);
    pub const V128_LOAD64_SPLAT: Op = Op::new(0xFD, 0x0A);
    pub const V128_STORE: Op = Op::new(0xFD, 0x0B);
    pub const V128_CONST: Op = Op::new(0xFD, 0x0C);
    pub const I8X16_SHUFFLE: Op = Op::new(0xFD, 0x0D);
    pub const I8X16_SWIZZLE: Op = Op::new(0xFD, 0x0E);
    // ── Splat ────────────────────────────────────────────────────────────────
    pub const I8X16_SPLAT: Op = Op::new(0xFD, 0x0F);
    pub const I16X8_SPLAT: Op = Op::new(0xFD, 0x10);
    pub const I32X4_SPLAT: Op = Op::new(0xFD, 0x11);
    pub const I64X2_SPLAT: Op = Op::new(0xFD, 0x12);
    pub const F32X4_SPLAT: Op = Op::new(0xFD, 0x13);
    pub const F64X2_SPLAT: Op = Op::new(0xFD, 0x14);
    // ── Extract / replace lane ───────────────────────────────────────────────
    pub const I8X16_EXTRACT_LANE_S: Op = Op::new(0xFD, 0x15);
    pub const I8X16_EXTRACT_LANE_U: Op = Op::new(0xFD, 0x16);
    pub const I8X16_REPLACE_LANE: Op = Op::new(0xFD, 0x17);
    pub const I16X8_EXTRACT_LANE_S: Op = Op::new(0xFD, 0x18);
    pub const I16X8_EXTRACT_LANE_U: Op = Op::new(0xFD, 0x19);
    pub const I16X8_REPLACE_LANE: Op = Op::new(0xFD, 0x1A);
    pub const I32X4_EXTRACT_LANE: Op = Op::new(0xFD, 0x1B);
    pub const I32X4_REPLACE_LANE: Op = Op::new(0xFD, 0x1C);
    pub const I64X2_EXTRACT_LANE: Op = Op::new(0xFD, 0x1D);
    pub const I64X2_REPLACE_LANE: Op = Op::new(0xFD, 0x1E);
    pub const F32X4_EXTRACT_LANE: Op = Op::new(0xFD, 0x1F);
    pub const F32X4_REPLACE_LANE: Op = Op::new(0xFD, 0x20);
    pub const F64X2_EXTRACT_LANE: Op = Op::new(0xFD, 0x21);
    pub const F64X2_REPLACE_LANE: Op = Op::new(0xFD, 0x22);
    // ── i8x16 comparisons ────────────────────────────────────────────────────
    pub const I8X16_EQ: Op = Op::new(0xFD, 0x23);
    pub const I8X16_NE: Op = Op::new(0xFD, 0x24);
    pub const I8X16_LT_S: Op = Op::new(0xFD, 0x25);
    pub const I8X16_LT_U: Op = Op::new(0xFD, 0x26);
    pub const I8X16_GT_S: Op = Op::new(0xFD, 0x27);
    pub const I8X16_GT_U: Op = Op::new(0xFD, 0x28);
    pub const I8X16_LE_S: Op = Op::new(0xFD, 0x29);
    pub const I8X16_LE_U: Op = Op::new(0xFD, 0x2A);
    pub const I8X16_GE_S: Op = Op::new(0xFD, 0x2B);
    pub const I8X16_GE_U: Op = Op::new(0xFD, 0x2C);
    // ── i16x8 comparisons ────────────────────────────────────────────────────
    pub const I16X8_EQ: Op = Op::new(0xFD, 0x2D);
    pub const I16X8_NE: Op = Op::new(0xFD, 0x2E);
    pub const I16X8_LT_S: Op = Op::new(0xFD, 0x2F);
    pub const I16X8_LT_U: Op = Op::new(0xFD, 0x30);
    pub const I16X8_GT_S: Op = Op::new(0xFD, 0x31);
    pub const I16X8_GT_U: Op = Op::new(0xFD, 0x32);
    pub const I16X8_LE_S: Op = Op::new(0xFD, 0x33);
    pub const I16X8_LE_U: Op = Op::new(0xFD, 0x34);
    pub const I16X8_GE_S: Op = Op::new(0xFD, 0x35);
    pub const I16X8_GE_U: Op = Op::new(0xFD, 0x36);
    // ── i32x4 comparisons ────────────────────────────────────────────────────
    pub const I32X4_EQ: Op = Op::new(0xFD, 0x37);
    pub const I32X4_NE: Op = Op::new(0xFD, 0x38);
    pub const I32X4_LT_S: Op = Op::new(0xFD, 0x39);
    pub const I32X4_LT_U: Op = Op::new(0xFD, 0x3A);
    pub const I32X4_GT_S: Op = Op::new(0xFD, 0x3B);
    pub const I32X4_GT_U: Op = Op::new(0xFD, 0x3C);
    pub const I32X4_LE_S: Op = Op::new(0xFD, 0x3D);
    pub const I32X4_LE_U: Op = Op::new(0xFD, 0x3E);
    pub const I32X4_GE_S: Op = Op::new(0xFD, 0x3F);
    pub const I32X4_GE_U: Op = Op::new(0xFD, 0x40);
    // ── f32x4 comparisons ────────────────────────────────────────────────────
    pub const F32X4_EQ: Op = Op::new(0xFD, 0x41);
    pub const F32X4_NE: Op = Op::new(0xFD, 0x42);
    pub const F32X4_LT: Op = Op::new(0xFD, 0x43);
    pub const F32X4_GT: Op = Op::new(0xFD, 0x44);
    pub const F32X4_LE: Op = Op::new(0xFD, 0x45);
    pub const F32X4_GE: Op = Op::new(0xFD, 0x46);
    // ── f64x2 comparisons ────────────────────────────────────────────────────
    pub const F64X2_EQ: Op = Op::new(0xFD, 0x47);
    pub const F64X2_NE: Op = Op::new(0xFD, 0x48);
    pub const F64X2_LT: Op = Op::new(0xFD, 0x49);
    pub const F64X2_GT: Op = Op::new(0xFD, 0x4A);
    pub const F64X2_LE: Op = Op::new(0xFD, 0x4B);
    pub const F64X2_GE: Op = Op::new(0xFD, 0x4C);
    // ── v128 bitwise ─────────────────────────────────────────────────────────
    pub const V128_NOT: Op = Op::new(0xFD, 0x4D);
    pub const V128_AND: Op = Op::new(0xFD, 0x4E);
    pub const V128_ANDNOT: Op = Op::new(0xFD, 0x4F);
    pub const V128_OR: Op = Op::new(0xFD, 0x50);
    pub const V128_XOR: Op = Op::new(0xFD, 0x51);
    pub const V128_BITSELECT: Op = Op::new(0xFD, 0x52);
    pub const V128_ANY_TRUE: Op = Op::new(0xFD, 0x53);
    // ── Load/store lane, load zero ────────────────────────────────────────────
    pub const V128_LOAD8_LANE: Op = Op::new(0xFD, 0x54);
    pub const V128_LOAD16_LANE: Op = Op::new(0xFD, 0x55);
    pub const V128_LOAD32_LANE: Op = Op::new(0xFD, 0x56);
    pub const V128_LOAD64_LANE: Op = Op::new(0xFD, 0x57);
    pub const V128_STORE8_LANE: Op = Op::new(0xFD, 0x58);
    pub const V128_STORE16_LANE: Op = Op::new(0xFD, 0x59);
    pub const V128_STORE32_LANE: Op = Op::new(0xFD, 0x5A);
    pub const V128_STORE64_LANE: Op = Op::new(0xFD, 0x5B);
    pub const V128_LOAD32_ZERO: Op = Op::new(0xFD, 0x5C);
    pub const V128_LOAD64_ZERO: Op = Op::new(0xFD, 0x5D);
    // ── Promote / demote ─────────────────────────────────────────────────────
    pub const F32X4_DEMOTE_F64X2_ZERO: Op = Op::new(0xFD, 0x5E);
    pub const F64X2_PROMOTE_LOW_F32X4: Op = Op::new(0xFD, 0x5F);
    // ── i8x16 unary + test ───────────────────────────────────────────────────
    pub const I8X16_ABS: Op = Op::new(0xFD, 0x60);
    pub const I8X16_NEG: Op = Op::new(0xFD, 0x61);
    pub const I8X16_POPCNT: Op = Op::new(0xFD, 0x62);
    pub const I8X16_ALL_TRUE: Op = Op::new(0xFD, 0x63);
    pub const I8X16_BITMASK: Op = Op::new(0xFD, 0x64);
    pub const I8X16_NARROW_I16X8_S: Op = Op::new(0xFD, 0x65);
    pub const I8X16_NARROW_I16X8_U: Op = Op::new(0xFD, 0x66);
    // ── f32x4 unary ──────────────────────────────────────────────────────────
    pub const F32X4_CEIL: Op = Op::new(0xFD, 0x67);
    pub const F32X4_FLOOR: Op = Op::new(0xFD, 0x68);
    pub const F32X4_TRUNC: Op = Op::new(0xFD, 0x69);
    pub const F32X4_NEAREST: Op = Op::new(0xFD, 0x6A);
    // ── i8x16 shifts ─────────────────────────────────────────────────────────
    pub const I8X16_SHL: Op = Op::new(0xFD, 0x6B);
    pub const I8X16_SHR_S: Op = Op::new(0xFD, 0x6C);
    pub const I8X16_SHR_U: Op = Op::new(0xFD, 0x6D);
    // ── i8x16 arithmetic ─────────────────────────────────────────────────────
    pub const I8X16_ADD: Op = Op::new(0xFD, 0x6E);
    pub const I8X16_ADD_SAT_S: Op = Op::new(0xFD, 0x6F);
    pub const I8X16_ADD_SAT_U: Op = Op::new(0xFD, 0x70);
    pub const I8X16_SUB: Op = Op::new(0xFD, 0x71);
    pub const I8X16_SUB_SAT_S: Op = Op::new(0xFD, 0x72);
    pub const I8X16_SUB_SAT_U: Op = Op::new(0xFD, 0x73);
    // ── f64x2 unary ──────────────────────────────────────────────────────────
    pub const F64X2_CEIL: Op = Op::new(0xFD, 0x74);
    pub const F64X2_FLOOR: Op = Op::new(0xFD, 0x75);
    // ── i8x16 min/max/avgr ───────────────────────────────────────────────────
    pub const I8X16_MIN_S: Op = Op::new(0xFD, 0x76);
    pub const I8X16_MIN_U: Op = Op::new(0xFD, 0x77);
    pub const I8X16_MAX_S: Op = Op::new(0xFD, 0x78);
    pub const I8X16_MAX_U: Op = Op::new(0xFD, 0x79);
    pub const F64X2_TRUNC: Op = Op::new(0xFD, 0x7A);
    pub const I8X16_AVGR_U: Op = Op::new(0xFD, 0x7B);
    // ── extadd pairwise ──────────────────────────────────────────────────────
    pub const I16X8_EXTADD_PAIRWISE_I8X16_S: Op = Op::new(0xFD, 0x7C);
    pub const I16X8_EXTADD_PAIRWISE_I8X16_U: Op = Op::new(0xFD, 0x7D);
    pub const I32X4_EXTADD_PAIRWISE_I16X8_S: Op = Op::new(0xFD, 0x7E);
    pub const I32X4_EXTADD_PAIRWISE_I16X8_U: Op = Op::new(0xFD, 0x7F);
    // ── i16x8 unary + test ───────────────────────────────────────────────────
    pub const I16X8_ABS: Op = Op::new(0xFD, 0x80);
    pub const I16X8_NEG: Op = Op::new(0xFD, 0x81);
    pub const I16X8_Q15MULR_SAT_S: Op = Op::new(0xFD, 0x82);
    pub const I16X8_ALL_TRUE: Op = Op::new(0xFD, 0x83);
    pub const I16X8_BITMASK: Op = Op::new(0xFD, 0x84);
    pub const I16X8_NARROW_I32X4_S: Op = Op::new(0xFD, 0x85);
    pub const I16X8_NARROW_I32X4_U: Op = Op::new(0xFD, 0x86);
    pub const I16X8_EXTEND_LOW_I8X16_S: Op = Op::new(0xFD, 0x87);
    pub const I16X8_EXTEND_HIGH_I8X16_S: Op = Op::new(0xFD, 0x88);
    pub const I16X8_EXTEND_LOW_I8X16_U: Op = Op::new(0xFD, 0x89);
    pub const I16X8_EXTEND_HIGH_I8X16_U: Op = Op::new(0xFD, 0x8A);
    pub const I16X8_SHL: Op = Op::new(0xFD, 0x8B);
    pub const I16X8_SHR_S: Op = Op::new(0xFD, 0x8C);
    pub const I16X8_SHR_U: Op = Op::new(0xFD, 0x8D);
    pub const I16X8_ADD: Op = Op::new(0xFD, 0x8E);
    pub const I16X8_ADD_SAT_S: Op = Op::new(0xFD, 0x8F);
    pub const I16X8_ADD_SAT_U: Op = Op::new(0xFD, 0x90);
    pub const I16X8_SUB: Op = Op::new(0xFD, 0x91);
    pub const I16X8_SUB_SAT_S: Op = Op::new(0xFD, 0x92);
    pub const I16X8_SUB_SAT_U: Op = Op::new(0xFD, 0x93);
    pub const F64X2_NEAREST: Op = Op::new(0xFD, 0x94);
    pub const I16X8_MUL: Op = Op::new(0xFD, 0x95);
    pub const I16X8_MIN_S: Op = Op::new(0xFD, 0x96);
    pub const I16X8_MIN_U: Op = Op::new(0xFD, 0x97);
    pub const I16X8_MAX_S: Op = Op::new(0xFD, 0x98);
    pub const I16X8_MAX_U: Op = Op::new(0xFD, 0x99);
    pub const I16X8_AVGR_U: Op = Op::new(0xFD, 0x9B);
    pub const I16X8_EXTMUL_LOW_I8X16_S: Op = Op::new(0xFD, 0x9C);
    pub const I16X8_EXTMUL_HIGH_I8X16_S: Op = Op::new(0xFD, 0x9D);
    pub const I16X8_EXTMUL_LOW_I8X16_U: Op = Op::new(0xFD, 0x9E);
    pub const I16X8_EXTMUL_HIGH_I8X16_U: Op = Op::new(0xFD, 0x9F);
    // ── i32x4 unary + test ───────────────────────────────────────────────────
    pub const I32X4_ABS: Op = Op::new(0xFD, 0xA0);
    pub const I32X4_NEG: Op = Op::new(0xFD, 0xA1);
    pub const I32X4_ALL_TRUE: Op = Op::new(0xFD, 0xA3);
    pub const I32X4_BITMASK: Op = Op::new(0xFD, 0xA4);
    pub const I32X4_EXTEND_LOW_I16X8_S: Op = Op::new(0xFD, 0xA7);
    pub const I32X4_EXTEND_HIGH_I16X8_S: Op = Op::new(0xFD, 0xA8);
    pub const I32X4_EXTEND_LOW_I16X8_U: Op = Op::new(0xFD, 0xA9);
    pub const I32X4_EXTEND_HIGH_I16X8_U: Op = Op::new(0xFD, 0xAA);
    pub const I32X4_SHL: Op = Op::new(0xFD, 0xAB);
    pub const I32X4_SHR_S: Op = Op::new(0xFD, 0xAC);
    pub const I32X4_SHR_U: Op = Op::new(0xFD, 0xAD);
    pub const I32X4_ADD: Op = Op::new(0xFD, 0xAE);
    pub const I32X4_SUB: Op = Op::new(0xFD, 0xB1);
    pub const I32X4_MUL: Op = Op::new(0xFD, 0xB5);
    pub const I32X4_MIN_S: Op = Op::new(0xFD, 0xB6);
    pub const I32X4_MIN_U: Op = Op::new(0xFD, 0xB7);
    pub const I32X4_MAX_S: Op = Op::new(0xFD, 0xB8);
    pub const I32X4_MAX_U: Op = Op::new(0xFD, 0xB9);
    pub const I32X4_DOT_I16X8_S: Op = Op::new(0xFD, 0xBA);
    pub const I32X4_EXTMUL_LOW_I16X8_S: Op = Op::new(0xFD, 0xBC);
    pub const I32X4_EXTMUL_HIGH_I16X8_S: Op = Op::new(0xFD, 0xBD);
    pub const I32X4_EXTMUL_LOW_I16X8_U: Op = Op::new(0xFD, 0xBE);
    pub const I32X4_EXTMUL_HIGH_I16X8_U: Op = Op::new(0xFD, 0xBF);
    // ── i64x2 ────────────────────────────────────────────────────────────────
    pub const I64X2_ABS: Op = Op::new(0xFD, 0xC0);
    pub const I64X2_NEG: Op = Op::new(0xFD, 0xC1);
    pub const I64X2_ALL_TRUE: Op = Op::new(0xFD, 0xC3);
    pub const I64X2_BITMASK: Op = Op::new(0xFD, 0xC4);
    pub const I64X2_EXTEND_LOW_I32X4_S: Op = Op::new(0xFD, 0xC7);
    pub const I64X2_EXTEND_HIGH_I32X4_S: Op = Op::new(0xFD, 0xC8);
    pub const I64X2_EXTEND_LOW_I32X4_U: Op = Op::new(0xFD, 0xC9);
    pub const I64X2_EXTEND_HIGH_I32X4_U: Op = Op::new(0xFD, 0xCA);
    pub const I64X2_SHL: Op = Op::new(0xFD, 0xCB);
    pub const I64X2_SHR_S: Op = Op::new(0xFD, 0xCC);
    pub const I64X2_SHR_U: Op = Op::new(0xFD, 0xCD);
    pub const I64X2_ADD: Op = Op::new(0xFD, 0xCE);
    pub const I64X2_SUB: Op = Op::new(0xFD, 0xD1);
    pub const I64X2_MUL: Op = Op::new(0xFD, 0xD5);
    pub const I64X2_EQ: Op = Op::new(0xFD, 0xD6);
    pub const I64X2_NE: Op = Op::new(0xFD, 0xD7);
    pub const I64X2_LT_S: Op = Op::new(0xFD, 0xD8);
    pub const I64X2_GT_S: Op = Op::new(0xFD, 0xD9);
    pub const I64X2_LE_S: Op = Op::new(0xFD, 0xDA);
    pub const I64X2_GE_S: Op = Op::new(0xFD, 0xDB);
    pub const I64X2_EXTMUL_LOW_I32X4_S: Op = Op::new(0xFD, 0xDC);
    pub const I64X2_EXTMUL_HIGH_I32X4_S: Op = Op::new(0xFD, 0xDD);
    pub const I64X2_EXTMUL_LOW_I32X4_U: Op = Op::new(0xFD, 0xDE);
    pub const I64X2_EXTMUL_HIGH_I32X4_U: Op = Op::new(0xFD, 0xDF);
    // ── f32x4 ────────────────────────────────────────────────────────────────
    pub const F32X4_ABS: Op = Op::new(0xFD, 0xE0);
    pub const F32X4_NEG: Op = Op::new(0xFD, 0xE1);
    pub const F32X4_SQRT: Op = Op::new(0xFD, 0xE3);
    pub const F32X4_ADD: Op = Op::new(0xFD, 0xE4);
    pub const F32X4_SUB: Op = Op::new(0xFD, 0xE5);
    pub const F32X4_MUL: Op = Op::new(0xFD, 0xE6);
    pub const F32X4_DIV: Op = Op::new(0xFD, 0xE7);
    pub const F32X4_MIN: Op = Op::new(0xFD, 0xE8);
    pub const F32X4_MAX: Op = Op::new(0xFD, 0xE9);
    pub const F32X4_PMIN: Op = Op::new(0xFD, 0xEA);
    pub const F32X4_PMAX: Op = Op::new(0xFD, 0xEB);
    // ── f64x2 ────────────────────────────────────────────────────────────────
    pub const F64X2_ABS: Op = Op::new(0xFD, 0xEC);
    pub const F64X2_NEG: Op = Op::new(0xFD, 0xED);
    pub const F64X2_SQRT: Op = Op::new(0xFD, 0xEF);
    pub const F64X2_ADD: Op = Op::new(0xFD, 0xF0);
    pub const F64X2_SUB: Op = Op::new(0xFD, 0xF1);
    pub const F64X2_MUL: Op = Op::new(0xFD, 0xF2);
    pub const F64X2_DIV: Op = Op::new(0xFD, 0xF3);
    pub const F64X2_MIN: Op = Op::new(0xFD, 0xF4);
    pub const F64X2_MAX: Op = Op::new(0xFD, 0xF5);
    pub const F64X2_PMIN: Op = Op::new(0xFD, 0xF6);
    pub const F64X2_PMAX: Op = Op::new(0xFD, 0xF7);
    // ── Conversions ──────────────────────────────────────────────────────────
    pub const I32X4_TRUNC_SAT_F32X4_S: Op = Op::new(0xFD, 0xF8);
    pub const I32X4_TRUNC_SAT_F32X4_U: Op = Op::new(0xFD, 0xF9);
    pub const F32X4_CONVERT_I32X4_S: Op = Op::new(0xFD, 0xFA);
    pub const F32X4_CONVERT_I32X4_U: Op = Op::new(0xFD, 0xFB);
    pub const I32X4_TRUNC_SAT_F64X2_S_ZERO: Op = Op::new(0xFD, 0xFC);
    pub const I32X4_TRUNC_SAT_F64X2_U_ZERO: Op = Op::new(0xFD, 0xFD);
    pub const F64X2_CONVERT_LOW_I32X4_S: Op = Op::new(0xFD, 0xFE);
    pub const F64X2_CONVERT_LOW_I32X4_U: Op = Op::new(0xFD, 0xFF);
    // ── Relaxed-SIMD (spec sub-values 0x100–0x113) live under 0xDD prefix ───
    // See `opcode/relaxed_simd.rs`.
}

opcode_category! {
    // Memory
    // Spec: every v128 load/store carries a memarg. Internally it is the
    // OPTIONAL marker-tagged form (`SimdMemArg`) — compiler emissions omit
    // it (0 bytes), reader-translated modules carry it with the 0x80 marker.
    // Declaring `None` here made every operand walk desync on read modules.
    [0x00] v128_load => SimdMemArg, "v128.load";
    [0x01] v128_load8x8_s => SimdMemArg, "v128.load8x8_s";
    [0x02] v128_load8x8_u => SimdMemArg, "v128.load8x8_u";
    [0x03] v128_load16x4_s => SimdMemArg, "v128.load16x4_s";
    [0x04] v128_load16x4_u => SimdMemArg, "v128.load16x4_u";
    [0x05] v128_load32x2_s => SimdMemArg, "v128.load32x2_s";
    [0x06] v128_load32x2_u => SimdMemArg, "v128.load32x2_u";
    [0x07] v128_load8_splat => SimdMemArg, "v128.load8_splat";
    [0x08] v128_load16_splat => SimdMemArg, "v128.load16_splat";
    [0x09] v128_load32_splat => SimdMemArg, "v128.load32_splat";
    [0x0A] v128_load64_splat => SimdMemArg, "v128.load64_splat";
    [0x0B] v128_store => SimdMemArg, "v128.store";
    [0x0C] v128_const => V128Const, "v128.const";
    [0x0D] i8x16_shuffle => Shuffle, "i8x16.shuffle";
    [0x0E] i8x16_swizzle => None, "i8x16.swizzle";
    // Splat
    [0x0F] i8x16_splat => None, "i8x16.splat";
    [0x10] i16x8_splat => None, "i16x8.splat";
    [0x11] i32x4_splat => None, "i32x4.splat";
    [0x12] i64x2_splat => None, "i64x2.splat";
    [0x13] f32x4_splat => None, "f32x4.splat";
    [0x14] f64x2_splat => None, "f64x2.splat";
    // Extract / replace lane
    [0x15] i8x16_extract_lane_s => U8, "i8x16.extract_lane_s";
    [0x16] i8x16_extract_lane_u => U8, "i8x16.extract_lane_u";
    [0x17] i8x16_replace_lane => U8, "i8x16.replace_lane";
    [0x18] i16x8_extract_lane_s => U8, "i16x8.extract_lane_s";
    [0x19] i16x8_extract_lane_u => U8, "i16x8.extract_lane_u";
    [0x1A] i16x8_replace_lane => U8, "i16x8.replace_lane";
    [0x1B] i32x4_extract_lane => U8, "i32x4.extract_lane";
    [0x1C] i32x4_replace_lane => U8, "i32x4.replace_lane";
    [0x1D] i64x2_extract_lane => U8, "i64x2.extract_lane";
    [0x1E] i64x2_replace_lane => U8, "i64x2.replace_lane";
    [0x1F] f32x4_extract_lane => U8, "f32x4.extract_lane";
    [0x20] f32x4_replace_lane => U8, "f32x4.replace_lane";
    [0x21] f64x2_extract_lane => U8, "f64x2.extract_lane";
    [0x22] f64x2_replace_lane => U8, "f64x2.replace_lane";
    // i8x16 comparisons
    [0x23] i8x16_eq => None, "i8x16.eq";
    [0x24] i8x16_ne => None, "i8x16.ne";
    [0x25] i8x16_lt_s => None, "i8x16.lt_s";
    [0x26] i8x16_lt_u => None, "i8x16.lt_u";
    [0x27] i8x16_gt_s => None, "i8x16.gt_s";
    [0x28] i8x16_gt_u => None, "i8x16.gt_u";
    [0x29] i8x16_le_s => None, "i8x16.le_s";
    [0x2A] i8x16_le_u => None, "i8x16.le_u";
    [0x2B] i8x16_ge_s => None, "i8x16.ge_s";
    [0x2C] i8x16_ge_u => None, "i8x16.ge_u";
    // i16x8 comparisons
    [0x2D] i16x8_eq => None, "i16x8.eq";
    [0x2E] i16x8_ne => None, "i16x8.ne";
    [0x2F] i16x8_lt_s => None, "i16x8.lt_s";
    [0x30] i16x8_lt_u => None, "i16x8.lt_u";
    [0x31] i16x8_gt_s => None, "i16x8.gt_s";
    [0x32] i16x8_gt_u => None, "i16x8.gt_u";
    [0x33] i16x8_le_s => None, "i16x8.le_s";
    [0x34] i16x8_le_u => None, "i16x8.le_u";
    [0x35] i16x8_ge_s => None, "i16x8.ge_s";
    [0x36] i16x8_ge_u => None, "i16x8.ge_u";
    // i32x4 comparisons
    [0x37] i32x4_eq => None, "i32x4.eq";
    [0x38] i32x4_ne => None, "i32x4.ne";
    [0x39] i32x4_lt_s => None, "i32x4.lt_s";
    [0x3A] i32x4_lt_u => None, "i32x4.lt_u";
    [0x3B] i32x4_gt_s => None, "i32x4.gt_s";
    [0x3C] i32x4_gt_u => None, "i32x4.gt_u";
    [0x3D] i32x4_le_s => None, "i32x4.le_s";
    [0x3E] i32x4_le_u => None, "i32x4.le_u";
    [0x3F] i32x4_ge_s => None, "i32x4.ge_s";
    [0x40] i32x4_ge_u => None, "i32x4.ge_u";
    // f32x4 comparisons
    [0x41] f32x4_eq => None, "f32x4.eq";
    [0x42] f32x4_ne => None, "f32x4.ne";
    [0x43] f32x4_lt => None, "f32x4.lt";
    [0x44] f32x4_gt => None, "f32x4.gt";
    [0x45] f32x4_le => None, "f32x4.le";
    [0x46] f32x4_ge => None, "f32x4.ge";
    // f64x2 comparisons
    [0x47] f64x2_eq => None, "f64x2.eq";
    [0x48] f64x2_ne => None, "f64x2.ne";
    [0x49] f64x2_lt => None, "f64x2.lt";
    [0x4A] f64x2_gt => None, "f64x2.gt";
    [0x4B] f64x2_le => None, "f64x2.le";
    [0x4C] f64x2_ge => None, "f64x2.ge";
    // v128 bitwise
    [0x4D] v128_not => None, "v128.not";
    [0x4E] v128_and => None, "v128.and";
    [0x4F] v128_andnot => None, "v128.andnot";
    [0x50] v128_or => None, "v128.or";
    [0x51] v128_xor => None, "v128.xor";
    [0x52] v128_bitselect => None, "v128.bitselect";
    [0x53] v128_any_true => None, "v128.any_true";
    // Load/store lane, load zero
    [0x54] v128_load8_lane => MemLane, "v128.load8_lane";
    [0x55] v128_load16_lane => MemLane, "v128.load16_lane";
    [0x56] v128_load32_lane => MemLane, "v128.load32_lane";
    [0x57] v128_load64_lane => MemLane, "v128.load64_lane";
    [0x58] v128_store8_lane => MemLane, "v128.store8_lane";
    [0x59] v128_store16_lane => MemLane, "v128.store16_lane";
    [0x5A] v128_store32_lane => MemLane, "v128.store32_lane";
    [0x5B] v128_store64_lane => MemLane, "v128.store64_lane";
    [0x5C] v128_load32_zero => SimdMemArg, "v128.load32_zero";
    [0x5D] v128_load64_zero => SimdMemArg, "v128.load64_zero";
    // Promote / demote
    [0x5E] f32x4_demote_f64x2_zero => None, "f32x4.demote_f64x2_zero";
    [0x5F] f64x2_promote_low_f32x4 => None, "f64x2.promote_low_f32x4";
    // i8x16 unary
    [0x60] i8x16_abs => None, "i8x16.abs";
    [0x61] i8x16_neg => None, "i8x16.neg";
    [0x62] i8x16_popcnt => None, "i8x16.popcnt";
    [0x63] i8x16_all_true => None, "i8x16.all_true";
    [0x64] i8x16_bitmask => None, "i8x16.bitmask";
    [0x65] i8x16_narrow_i16x8_s => None, "i8x16.narrow_i16x8_s";
    [0x66] i8x16_narrow_i16x8_u => None, "i8x16.narrow_i16x8_u";
    // f32x4 unary
    [0x67] f32x4_ceil => None, "f32x4.ceil";
    [0x68] f32x4_floor => None, "f32x4.floor";
    [0x69] f32x4_trunc => None, "f32x4.trunc";
    [0x6A] f32x4_nearest => None, "f32x4.nearest";
    // i8x16 shifts + arithmetic
    [0x6B] i8x16_shl => None, "i8x16.shl";
    [0x6C] i8x16_shr_s => None, "i8x16.shr_s";
    [0x6D] i8x16_shr_u => None, "i8x16.shr_u";
    [0x6E] i8x16_add => None, "i8x16.add";
    [0x6F] i8x16_add_sat_s => None, "i8x16.add_sat_s";
    [0x70] i8x16_add_sat_u => None, "i8x16.add_sat_u";
    [0x71] i8x16_sub => None, "i8x16.sub";
    [0x72] i8x16_sub_sat_s => None, "i8x16.sub_sat_s";
    [0x73] i8x16_sub_sat_u => None, "i8x16.sub_sat_u";
    [0x74] f64x2_ceil => None, "f64x2.ceil";
    [0x75] f64x2_floor => None, "f64x2.floor";
    // i8x16 min/max/avgr
    [0x76] i8x16_min_s => None, "i8x16.min_s";
    [0x77] i8x16_min_u => None, "i8x16.min_u";
    [0x78] i8x16_max_s => None, "i8x16.max_s";
    [0x79] i8x16_max_u => None, "i8x16.max_u";
    [0x7A] f64x2_trunc => None, "f64x2.trunc";
    [0x7B] i8x16_avgr_u => None, "i8x16.avgr_u";
    [0x7C] i16x8_extadd_pairwise_i8x16_s => None, "i16x8.extadd_pairwise_i8x16_s";
    [0x7D] i16x8_extadd_pairwise_i8x16_u => None, "i16x8.extadd_pairwise_i8x16_u";
    [0x7E] i32x4_extadd_pairwise_i16x8_s => None, "i32x4.extadd_pairwise_i16x8_s";
    [0x7F] i32x4_extadd_pairwise_i16x8_u => None, "i32x4.extadd_pairwise_i16x8_u";
    // i16x8
    [0x80] i16x8_abs => None, "i16x8.abs";
    [0x81] i16x8_neg => None, "i16x8.neg";
    [0x82] i16x8_q15mulr_sat_s => None, "i16x8.q15mulr_sat_s";
    [0x83] i16x8_all_true => None, "i16x8.all_true";
    [0x84] i16x8_bitmask => None, "i16x8.bitmask";
    [0x85] i16x8_narrow_i32x4_s => None, "i16x8.narrow_i32x4_s";
    [0x86] i16x8_narrow_i32x4_u => None, "i16x8.narrow_i32x4_u";
    [0x87] i16x8_extend_low_i8x16_s => None, "i16x8.extend_low_i8x16_s";
    [0x88] i16x8_extend_high_i8x16_s => None, "i16x8.extend_high_i8x16_s";
    [0x89] i16x8_extend_low_i8x16_u => None, "i16x8.extend_low_i8x16_u";
    [0x8A] i16x8_extend_high_i8x16_u => None, "i16x8.extend_high_i8x16_u";
    [0x8B] i16x8_shl => None, "i16x8.shl";
    [0x8C] i16x8_shr_s => None, "i16x8.shr_s";
    [0x8D] i16x8_shr_u => None, "i16x8.shr_u";
    [0x8E] i16x8_add => None, "i16x8.add";
    [0x8F] i16x8_add_sat_s => None, "i16x8.add_sat_s";
    [0x90] i16x8_add_sat_u => None, "i16x8.add_sat_u";
    [0x91] i16x8_sub => None, "i16x8.sub";
    [0x92] i16x8_sub_sat_s => None, "i16x8.sub_sat_s";
    [0x93] i16x8_sub_sat_u => None, "i16x8.sub_sat_u";
    [0x94] f64x2_nearest => None, "f64x2.nearest";
    [0x95] i16x8_mul => None, "i16x8.mul";
    [0x96] i16x8_min_s => None, "i16x8.min_s";
    [0x97] i16x8_min_u => None, "i16x8.min_u";
    [0x98] i16x8_max_s => None, "i16x8.max_s";
    [0x99] i16x8_max_u => None, "i16x8.max_u";
    [0x9B] i16x8_avgr_u => None, "i16x8.avgr_u";
    [0x9C] i16x8_extmul_low_i8x16_s => None, "i16x8.extmul_low_i8x16_s";
    [0x9D] i16x8_extmul_high_i8x16_s => None, "i16x8.extmul_high_i8x16_s";
    [0x9E] i16x8_extmul_low_i8x16_u => None, "i16x8.extmul_low_i8x16_u";
    [0x9F] i16x8_extmul_high_i8x16_u => None, "i16x8.extmul_high_i8x16_u";
    // i32x4
    [0xA0] i32x4_abs => None, "i32x4.abs";
    [0xA1] i32x4_neg => None, "i32x4.neg";
    [0xA3] i32x4_all_true => None, "i32x4.all_true";
    [0xA4] i32x4_bitmask => None, "i32x4.bitmask";
    [0xA7] i32x4_extend_low_i16x8_s => None, "i32x4.extend_low_i16x8_s";
    [0xA8] i32x4_extend_high_i16x8_s => None, "i32x4.extend_high_i16x8_s";
    [0xA9] i32x4_extend_low_i16x8_u => None, "i32x4.extend_low_i16x8_u";
    [0xAA] i32x4_extend_high_i16x8_u => None, "i32x4.extend_high_i16x8_u";
    [0xAB] i32x4_shl => None, "i32x4.shl";
    [0xAC] i32x4_shr_s => None, "i32x4.shr_s";
    [0xAD] i32x4_shr_u => None, "i32x4.shr_u";
    [0xAE] i32x4_add => None, "i32x4.add";
    [0xB1] i32x4_sub => None, "i32x4.sub";
    [0xB5] i32x4_mul => None, "i32x4.mul";
    [0xB6] i32x4_min_s => None, "i32x4.min_s";
    [0xB7] i32x4_min_u => None, "i32x4.min_u";
    [0xB8] i32x4_max_s => None, "i32x4.max_s";
    [0xB9] i32x4_max_u => None, "i32x4.max_u";
    [0xBA] i32x4_dot_i16x8_s => None, "i32x4.dot_i16x8_s";
    [0xBC] i32x4_extmul_low_i16x8_s => None, "i32x4.extmul_low_i16x8_s";
    [0xBD] i32x4_extmul_high_i16x8_s => None, "i32x4.extmul_high_i16x8_s";
    [0xBE] i32x4_extmul_low_i16x8_u => None, "i32x4.extmul_low_i16x8_u";
    [0xBF] i32x4_extmul_high_i16x8_u => None, "i32x4.extmul_high_i16x8_u";
    // i64x2
    [0xC0] i64x2_abs => None, "i64x2.abs";
    [0xC1] i64x2_neg => None, "i64x2.neg";
    [0xC3] i64x2_all_true => None, "i64x2.all_true";
    [0xC4] i64x2_bitmask => None, "i64x2.bitmask";
    [0xC7] i64x2_extend_low_i32x4_s => None, "i64x2.extend_low_i32x4_s";
    [0xC8] i64x2_extend_high_i32x4_s => None, "i64x2.extend_high_i32x4_s";
    [0xC9] i64x2_extend_low_i32x4_u => None, "i64x2.extend_low_i32x4_u";
    [0xCA] i64x2_extend_high_i32x4_u => None, "i64x2.extend_high_i32x4_u";
    [0xCB] i64x2_shl => None, "i64x2.shl";
    [0xCC] i64x2_shr_s => None, "i64x2.shr_s";
    [0xCD] i64x2_shr_u => None, "i64x2.shr_u";
    [0xCE] i64x2_add => None, "i64x2.add";
    [0xD1] i64x2_sub => None, "i64x2.sub";
    [0xD5] i64x2_mul => None, "i64x2.mul";
    [0xD6] i64x2_eq => None, "i64x2.eq";
    [0xD7] i64x2_ne => None, "i64x2.ne";
    [0xD8] i64x2_lt_s => None, "i64x2.lt_s";
    [0xD9] i64x2_gt_s => None, "i64x2.gt_s";
    [0xDA] i64x2_le_s => None, "i64x2.le_s";
    [0xDB] i64x2_ge_s => None, "i64x2.ge_s";
    [0xDC] i64x2_extmul_low_i32x4_s => None, "i64x2.extmul_low_i32x4_s";
    [0xDD] i64x2_extmul_high_i32x4_s => None, "i64x2.extmul_high_i32x4_s";
    [0xDE] i64x2_extmul_low_i32x4_u => None, "i64x2.extmul_low_i32x4_u";
    [0xDF] i64x2_extmul_high_i32x4_u => None, "i64x2.extmul_high_i32x4_u";
    // f32x4
    [0xE0] f32x4_abs => None, "f32x4.abs";
    [0xE1] f32x4_neg => None, "f32x4.neg";
    [0xE3] f32x4_sqrt => None, "f32x4.sqrt";
    [0xE4] f32x4_add => None, "f32x4.add";
    [0xE5] f32x4_sub => None, "f32x4.sub";
    [0xE6] f32x4_mul => None, "f32x4.mul";
    [0xE7] f32x4_div => None, "f32x4.div";
    [0xE8] f32x4_min => None, "f32x4.min";
    [0xE9] f32x4_max => None, "f32x4.max";
    [0xEA] f32x4_pmin => None, "f32x4.pmin";
    [0xEB] f32x4_pmax => None, "f32x4.pmax";
    // f64x2
    [0xEC] f64x2_abs => None, "f64x2.abs";
    [0xED] f64x2_neg => None, "f64x2.neg";
    [0xEF] f64x2_sqrt => None, "f64x2.sqrt";
    [0xF0] f64x2_add => None, "f64x2.add";
    [0xF1] f64x2_sub => None, "f64x2.sub";
    [0xF2] f64x2_mul => None, "f64x2.mul";
    [0xF3] f64x2_div => None, "f64x2.div";
    [0xF4] f64x2_min => None, "f64x2.min";
    [0xF5] f64x2_max => None, "f64x2.max";
    [0xF6] f64x2_pmin => None, "f64x2.pmin";
    [0xF7] f64x2_pmax => None, "f64x2.pmax";
    // Conversions
    [0xF8] i32x4_trunc_sat_f32x4_s => None, "i32x4.trunc_sat_f32x4_s";
    [0xF9] i32x4_trunc_sat_f32x4_u => None, "i32x4.trunc_sat_f32x4_u";
    [0xFA] f32x4_convert_i32x4_s => None, "f32x4.convert_i32x4_s";
    [0xFB] f32x4_convert_i32x4_u => None, "f32x4.convert_i32x4_u";
    [0xFC] i32x4_trunc_sat_f64x2_s_zero => None, "i32x4.trunc_sat_f64x2_s_zero";
    [0xFD] i32x4_trunc_sat_f64x2_u_zero => None, "i32x4.trunc_sat_f64x2_u_zero";
    [0xFE] f64x2_convert_low_i32x4_s => None, "f64x2.convert_low_i32x4_s";
    [0xFF] f64x2_convert_low_i32x4_u => None, "f64x2.convert_low_i32x4_u";
}
