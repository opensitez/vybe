//! SIMD proposal opcodes (prefix 0xFD).
//! Byte values match the WASM SIMD specification.

use super::Op;
use super::opcode_category;

impl Op {
    pub const V128_LOAD: Op         = Op::new(0xFD, 0x00);
    pub const V128_STORE: Op        = Op::new(0xFD, 0x0B);
    pub const V128_CONST: Op        = Op::new(0xFD, 0x0C);
    pub const I8X16_SHUFFLE: Op     = Op::new(0xFD, 0x0D);
    pub const I8X16_SWIZZLE: Op     = Op::new(0xFD, 0x0E);
    pub const I8X16_SPLAT: Op       = Op::new(0xFD, 0x0F);
    pub const I16X8_SPLAT: Op       = Op::new(0xFD, 0x10);
    pub const I32X4_SPLAT: Op       = Op::new(0xFD, 0x11);
    pub const F32X4_SPLAT: Op       = Op::new(0xFD, 0x13);
    pub const F64X2_SPLAT: Op       = Op::new(0xFD, 0x14);
    pub const I8X16_EXTRACT_LANE_S: Op = Op::new(0xFD, 0x15);
    pub const I8X16_EXTRACT_LANE_U: Op = Op::new(0xFD, 0x16);
    pub const I8X16_REPLACE_LANE: Op = Op::new(0xFD, 0x17);
    pub const I16X8_EXTRACT_LANE_S: Op = Op::new(0xFD, 0x18);
    pub const I16X8_EXTRACT_LANE_U: Op = Op::new(0xFD, 0x19);
    pub const I16X8_REPLACE_LANE: Op = Op::new(0xFD, 0x1A);
    pub const I32X4_EXTRACT_LANE: Op = Op::new(0xFD, 0x1B);
    pub const I32X4_REPLACE_LANE: Op = Op::new(0xFD, 0x1C);
    pub const F32X4_EXTRACT_LANE: Op = Op::new(0xFD, 0x1F);
    pub const F32X4_REPLACE_LANE: Op = Op::new(0xFD, 0x20);
    pub const F64X2_EXTRACT_LANE: Op = Op::new(0xFD, 0x21);
    pub const F64X2_REPLACE_LANE: Op = Op::new(0xFD, 0x22);
    pub const I8X16_EQ: Op          = Op::new(0xFD, 0x23);
    pub const I32X4_EQ: Op          = Op::new(0xFD, 0x37);
    pub const I32X4_LT_S: Op        = Op::new(0xFD, 0x38);
    pub const I32X4_GT_S: Op        = Op::new(0xFD, 0x39);
    pub const F64X2_EQ: Op          = Op::new(0xFD, 0x47);
    pub const F64X2_LT: Op          = Op::new(0xFD, 0x48);
    pub const F64X2_LE: Op          = Op::new(0xFD, 0x4A);
    pub const V128_NOT: Op          = Op::new(0xFD, 0x4D);
    pub const V128_AND: Op          = Op::new(0xFD, 0x4E);
    pub const V128_ANDNOT: Op       = Op::new(0xFD, 0x4F);
    pub const V128_OR: Op           = Op::new(0xFD, 0x50);
    pub const V128_XOR: Op          = Op::new(0xFD, 0x51);
    pub const V128_BITSELECT: Op    = Op::new(0xFD, 0x52);
    pub const V128_ANY_TRUE: Op     = Op::new(0xFD, 0x53);
    pub const I8X16_ADD: Op         = Op::new(0xFD, 0x6E);
    pub const I8X16_SUB: Op         = Op::new(0xFD, 0x71);
    pub const I16X8_ADD: Op         = Op::new(0xFD, 0x8E);
    pub const I16X8_SUB: Op         = Op::new(0xFD, 0x91);
    pub const I16X8_MUL: Op         = Op::new(0xFD, 0x95);
    pub const I32X4_SHL: Op         = Op::new(0xFD, 0xAB);
    pub const I32X4_SHR_S: Op       = Op::new(0xFD, 0xAC);
    pub const I32X4_SHR_U: Op       = Op::new(0xFD, 0xAD);
    pub const I32X4_ADD: Op         = Op::new(0xFD, 0xAE);
    pub const I32X4_SUB: Op         = Op::new(0xFD, 0xB1);
    pub const I32X4_MUL: Op         = Op::new(0xFD, 0xB5);
    pub const F32X4_ADD: Op         = Op::new(0xFD, 0xE4);
    pub const F32X4_SUB: Op         = Op::new(0xFD, 0xE5);
    pub const F32X4_MUL: Op         = Op::new(0xFD, 0xE6);
    pub const F32X4_DIV: Op         = Op::new(0xFD, 0xE7);
    pub const F64X2_ABS: Op         = Op::new(0xFD, 0xEC);
    pub const F64X2_NEG: Op         = Op::new(0xFD, 0xED);
    pub const F64X2_SQRT: Op        = Op::new(0xFD, 0xEF);
    pub const F64X2_ADD: Op         = Op::new(0xFD, 0xF0);
    pub const F64X2_SUB: Op         = Op::new(0xFD, 0xF1);
    pub const F64X2_MUL: Op         = Op::new(0xFD, 0xF2);
    pub const F64X2_DIV: Op         = Op::new(0xFD, 0xF3);
    pub const F64X2_MIN: Op         = Op::new(0xFD, 0xF4);
    pub const F64X2_MAX: Op         = Op::new(0xFD, 0xF5);
    // Relaxed-SIMD proposal opcodes live at WASM sub-values 0x100..=0x113 —
    // outside the u8 range the `Op(u16)` representation allows in the 0xFD
    // prefix. They are declared under the internal prefix 0xDD
    // (see `opcode/relaxed_simd.rs`) and the WASM emitter rewrites them
    // back to `0xFD + LEB128(0x100 + sub)` at binary time.
}

opcode_category! {
    [0x00] v128_load => None, "v128.load";
    [0x0B] v128_store => None, "v128.store";
    [0x0C] v128_const => V128Const, "v128.const";
    [0x0D] i8x16_shuffle => Shuffle, "i8x16.shuffle";
    [0x0E] i8x16_swizzle => None, "i8x16.swizzle";
    [0x0F] i8x16_splat => None, "i8x16.splat";
    [0x10] i16x8_splat => None, "i16x8.splat";
    [0x11] i32x4_splat => None, "i32x4.splat";
    [0x13] f32x4_splat => None, "f32x4.splat";
    [0x14] f64x2_splat => None, "f64x2.splat";
    [0x15] i8x16_extract_lane_s => U8, "i8x16.extract_lane_s";
    [0x16] i8x16_extract_lane_u => U8, "i8x16.extract_lane_u";
    [0x17] i8x16_replace_lane => U8, "i8x16.replace_lane";
    [0x18] i16x8_extract_lane_s => U8, "i16x8.extract_lane_s";
    [0x19] i16x8_extract_lane_u => U8, "i16x8.extract_lane_u";
    [0x1A] i16x8_replace_lane => U8, "i16x8.replace_lane";
    [0x1B] i32x4_extract_lane => U8, "i32x4.extract_lane";
    [0x1C] i32x4_replace_lane => U8, "i32x4.replace_lane";
    [0x1F] f32x4_extract_lane => U8, "f32x4.extract_lane";
    [0x20] f32x4_replace_lane => U8, "f32x4.replace_lane";
    [0x21] f64x2_extract_lane => U8, "f64x2.extract_lane";
    [0x22] f64x2_replace_lane => U8, "f64x2.replace_lane";
    [0x23] i8x16_eq => None, "i8x16.eq";
    [0x37] i32x4_eq => None, "i32x4.eq";
    [0x38] i32x4_lt_s => None, "i32x4.lt_s";
    [0x39] i32x4_gt_s => None, "i32x4.gt_s";
    [0x47] f64x2_eq => None, "f64x2.eq";
    [0x48] f64x2_lt => None, "f64x2.lt";
    [0x4A] f64x2_le => None, "f64x2.le";
    [0x4D] v128_not => None, "v128.not";
    [0x4E] v128_and => None, "v128.and";
    [0x4F] v128_andnot => None, "v128.andnot";
    [0x50] v128_or => None, "v128.or";
    [0x51] v128_xor => None, "v128.xor";
    [0x52] v128_bitselect => None, "v128.bitselect";
    [0x53] v128_any_true => None, "v128.any_true";
    [0x6E] i8x16_add => None, "i8x16.add";
    [0x71] i8x16_sub => None, "i8x16.sub";
    [0x8E] i16x8_add => None, "i16x8.add";
    [0x91] i16x8_sub => None, "i16x8.sub";
    [0x95] i16x8_mul => None, "i16x8.mul";
    [0xAB] i32x4_shl => U8, "i32x4.shl";
    [0xAC] i32x4_shr_s => U8, "i32x4.shr_s";
    [0xAD] i32x4_shr_u => U8, "i32x4.shr_u";
    [0xAE] i32x4_add => None, "i32x4.add";
    [0xB1] i32x4_sub => None, "i32x4.sub";
    [0xB5] i32x4_mul => None, "i32x4.mul";
    [0xE4] f32x4_add => None, "f32x4.add";
    [0xE5] f32x4_sub => None, "f32x4.sub";
    [0xE6] f32x4_mul => None, "f32x4.mul";
    [0xE7] f32x4_div => None, "f32x4.div";
    [0xEC] f64x2_abs => None, "f64x2.abs";
    [0xED] f64x2_neg => None, "f64x2.neg";
    [0xEF] f64x2_sqrt => None, "f64x2.sqrt";
    [0xF0] f64x2_add => None, "f64x2.add";
    [0xF1] f64x2_sub => None, "f64x2.sub";
    [0xF2] f64x2_mul => None, "f64x2.mul";
    [0xF3] f64x2_div => None, "f64x2.div";
    [0xF4] f64x2_min => None, "f64x2.min";
    [0xF5] f64x2_max => None, "f64x2.max";
    // Sub-values 0xF8..=0xFF are reserved for the MVP convert/trunc-sat
    // family (f32x4.convert_i32x4_{s,u}, i32x4.trunc_sat_f64x2_*, …); we
    // haven't wired them up yet but DO NOT collide with them by placing
    // relaxed-SIMD entries in that range — those live under the 0xDD
    // internal prefix in `relaxed_simd.rs`.
}
