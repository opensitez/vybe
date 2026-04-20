//! Core WASM MVP opcodes (prefix 0x00).
//! Byte values match the WASM specification exactly.

use super::Op;
use super::opcode_category;

// ── Opcode constants ────────────────────────────────────────────
// Each is Op::new(0x00, <wasm_byte_value>).

impl Op {
    // Control
    pub const UNREACHABLE: Op       = Op::new(0x00, 0x00);
    pub const BLOCK: Op             = Op::new(0x00, 0x02);
    pub const LOOP: Op              = Op::new(0x00, 0x03);
    pub const THROW: Op             = Op::new(0x00, 0x08);
    pub const THROW_REF: Op         = Op::new(0x00, 0x0A);
    pub const END: Op               = Op::new(0x00, 0x0B);
    pub const BR: Op                = Op::new(0x00, 0x0C);
    pub const BR_IF_TRUE: Op        = Op::new(0x00, 0x0D);
    pub const BR_TABLE: Op          = Op::new(0x00, 0x0E);
    pub const RETURN: Op            = Op::new(0x00, 0x0F);
    pub const CALL: Op              = Op::new(0x00, 0x10);
    pub const CALL_INDIRECT: Op     = Op::new(0x00, 0x11);
    pub const RETURN_CALL: Op       = Op::new(0x00, 0x12);
    pub const RETURN_CALL_INDIRECT: Op = Op::new(0x00, 0x13);
    pub const CALL_REF: Op          = Op::new(0x00, 0x14);
    pub const RETURN_CALL_REF: Op   = Op::new(0x00, 0x15);
    pub const DROP: Op              = Op::new(0x00, 0x1A);
    pub const SELECT: Op            = Op::new(0x00, 0x1B);
    pub const TRY_TABLE: Op         = Op::new(0x00, 0x1F);
    // Variables
    pub const LOCAL_GET: Op         = Op::new(0x00, 0x20);
    pub const LOCAL_SET: Op         = Op::new(0x00, 0x21);
    pub const GLOBAL_GET: Op        = Op::new(0x00, 0x23);
    pub const GLOBAL_SET: Op        = Op::new(0x00, 0x24);
    // Memory load
    pub const I32_LOAD: Op          = Op::new(0x00, 0x28);
    pub const I64_LOAD: Op          = Op::new(0x00, 0x29);
    pub const F32_LOAD: Op          = Op::new(0x00, 0x2A);
    pub const F64_LOAD: Op          = Op::new(0x00, 0x2B);
    pub const I32_LOAD8_S: Op       = Op::new(0x00, 0x2C);
    pub const I32_LOAD8_U: Op       = Op::new(0x00, 0x2D);
    pub const I32_LOAD16_S: Op      = Op::new(0x00, 0x2E);
    pub const I32_LOAD16_U: Op      = Op::new(0x00, 0x2F);
    pub const I64_LOAD8_S: Op       = Op::new(0x00, 0x30);
    pub const I64_LOAD8_U: Op       = Op::new(0x00, 0x31);
    pub const I64_LOAD16_S: Op      = Op::new(0x00, 0x32);
    pub const I64_LOAD16_U: Op      = Op::new(0x00, 0x33);
    pub const I64_LOAD32_S: Op      = Op::new(0x00, 0x34);
    pub const I64_LOAD32_U: Op      = Op::new(0x00, 0x35);
    // Memory store
    pub const I32_STORE: Op         = Op::new(0x00, 0x36);
    pub const I64_STORE: Op         = Op::new(0x00, 0x37);
    pub const F32_STORE: Op         = Op::new(0x00, 0x38);
    pub const F64_STORE: Op         = Op::new(0x00, 0x39);
    pub const I32_STORE8: Op        = Op::new(0x00, 0x3A);
    pub const I32_STORE16: Op       = Op::new(0x00, 0x3B);
    pub const I64_STORE8: Op        = Op::new(0x00, 0x3C);
    pub const I64_STORE16: Op       = Op::new(0x00, 0x3D);
    pub const I64_STORE32: Op       = Op::new(0x00, 0x3E);
    // Memory
    pub const MEMORY_SIZE: Op       = Op::new(0x00, 0x3F);
    pub const MEMORY_GROW: Op       = Op::new(0x00, 0x40);
    // i32 comparisons
    pub const I32_EQZ: Op           = Op::new(0x00, 0x45);
    pub const EQ: Op                = Op::new(0x00, 0x46);
    pub const NE: Op                = Op::new(0x00, 0x47);
    // i64
    pub const I64_EQZ: Op           = Op::new(0x00, 0x50);
    // f64 comparisons
    pub const F64_LT: Op            = Op::new(0x00, 0x63);
    pub const F64_GT: Op            = Op::new(0x00, 0x64);
    pub const F64_LE: Op            = Op::new(0x00, 0x65);
    pub const F64_GE: Op            = Op::new(0x00, 0x66);
    // i32 arithmetic
    pub const I32_CLZ: Op           = Op::new(0x00, 0x67);
    pub const I32_CTZ: Op           = Op::new(0x00, 0x68);
    pub const I32_POPCNT: Op        = Op::new(0x00, 0x69);
    pub const I32_ADD: Op           = Op::new(0x00, 0x6A);
    pub const I32_SUB: Op           = Op::new(0x00, 0x6B);
    pub const I32_MUL: Op           = Op::new(0x00, 0x6C);
    pub const I32_DIV_S: Op         = Op::new(0x00, 0x6D);
    pub const I32_DIV_U: Op         = Op::new(0x00, 0x6E);
    pub const I32_REM_S: Op         = Op::new(0x00, 0x6F);
    pub const I32_REM_U: Op         = Op::new(0x00, 0x70);
    pub const I32_AND: Op           = Op::new(0x00, 0x71);
    pub const I32_OR: Op            = Op::new(0x00, 0x72);
    pub const I32_XOR: Op           = Op::new(0x00, 0x73);
    pub const I32_SHL: Op           = Op::new(0x00, 0x74);
    pub const I32_SHR_S: Op         = Op::new(0x00, 0x75);
    pub const I32_SHR_U: Op         = Op::new(0x00, 0x76);
    pub const I32_ROTL: Op          = Op::new(0x00, 0x77);
    pub const I32_ROTR: Op          = Op::new(0x00, 0x78);
    // i64 arithmetic
    pub const I64_CLZ: Op           = Op::new(0x00, 0x79);
    pub const I64_CTZ: Op           = Op::new(0x00, 0x7A);
    pub const I64_POPCNT: Op        = Op::new(0x00, 0x7B);
    pub const I64_ADD: Op           = Op::new(0x00, 0x7C);
    pub const I64_SUB: Op           = Op::new(0x00, 0x7D);
    pub const I64_MUL: Op           = Op::new(0x00, 0x7E);
    pub const I64_DIV_S: Op         = Op::new(0x00, 0x7F);
    pub const I64_DIV_U: Op         = Op::new(0x00, 0x80);
    pub const I64_REM_S: Op         = Op::new(0x00, 0x81);
    pub const I64_REM_U: Op         = Op::new(0x00, 0x82);
    pub const I64_AND: Op           = Op::new(0x00, 0x83);
    pub const I64_OR: Op            = Op::new(0x00, 0x84);
    pub const I64_XOR: Op           = Op::new(0x00, 0x85);
    pub const I64_SHL: Op           = Op::new(0x00, 0x86);
    pub const I64_SHR_S: Op         = Op::new(0x00, 0x87);
    pub const I64_SHR_U: Op         = Op::new(0x00, 0x88);
    pub const I64_ROTL: Op          = Op::new(0x00, 0x89);
    pub const I64_ROTR: Op          = Op::new(0x00, 0x8A);
    // f32 math
    pub const F32_ABS: Op           = Op::new(0x00, 0x8B);
    pub const F32_NEG: Op           = Op::new(0x00, 0x8C);
    pub const F32_CEIL: Op          = Op::new(0x00, 0x8D);
    pub const F32_FLOOR: Op         = Op::new(0x00, 0x8E);
    pub const F32_TRUNC: Op         = Op::new(0x00, 0x8F);
    pub const F32_NEAREST: Op       = Op::new(0x00, 0x90);
    pub const F32_SQRT: Op          = Op::new(0x00, 0x91);
    pub const F32_MIN: Op           = Op::new(0x00, 0x96);
    pub const F32_MAX: Op           = Op::new(0x00, 0x97);
    pub const F32_COPYSIGN: Op      = Op::new(0x00, 0x98);
    // f64 math
    pub const F64_ABS: Op           = Op::new(0x00, 0x99);
    pub const F64_NEG: Op           = Op::new(0x00, 0x9A);
    pub const F64_CEIL: Op          = Op::new(0x00, 0x9B);
    pub const F64_FLOOR: Op         = Op::new(0x00, 0x9C);
    pub const F64_TRUNC: Op         = Op::new(0x00, 0x9D);
    pub const F64_NEAREST: Op       = Op::new(0x00, 0x9E);
    pub const F64_SQRT: Op          = Op::new(0x00, 0x9F);
    pub const F64_ADD: Op           = Op::new(0x00, 0xA0);
    pub const F64_SUB: Op           = Op::new(0x00, 0xA1);
    pub const F64_MUL: Op           = Op::new(0x00, 0xA2);
    pub const F64_DIV: Op           = Op::new(0x00, 0xA3);
    pub const F64_MIN: Op           = Op::new(0x00, 0xA4);
    pub const F64_MAX: Op           = Op::new(0x00, 0xA5);
    pub const F64_COPYSIGN: Op      = Op::new(0x00, 0xA6);
    // Conversions
    pub const I32_WRAP_I64: Op      = Op::new(0x00, 0xA7);
    pub const I32_FROM_F64: Op      = Op::new(0x00, 0xAA);
    pub const I64_EXTEND_I32_S: Op  = Op::new(0x00, 0xAC);
    pub const I64_EXTEND_I32_U: Op  = Op::new(0x00, 0xAD);
    pub const I64_TRUNC_F64_S: Op   = Op::new(0x00, 0xB0);
    pub const I64_TRUNC_F64_U: Op   = Op::new(0x00, 0xB1);
    pub const F32_DEMOTE_F64: Op    = Op::new(0x00, 0xB6);
    pub const F64_FROM_I32: Op      = Op::new(0x00, 0xB7);
    pub const F64_PROMOTE_F32: Op   = Op::new(0x00, 0xBB);
    pub const I32_REINTERPRET_F32: Op = Op::new(0x00, 0xBC);
    pub const I64_REINTERPRET_F64: Op = Op::new(0x00, 0xBD);
    pub const F32_REINTERPRET_I32: Op = Op::new(0x00, 0xBE);
    pub const F64_REINTERPRET_I64: Op = Op::new(0x00, 0xBF);
    // Sign extension
    pub const I32_EXTEND8_S: Op     = Op::new(0x00, 0xC0);
    pub const I32_EXTEND16_S: Op    = Op::new(0x00, 0xC1);
    pub const I64_EXTEND8_S: Op     = Op::new(0x00, 0xC2);
    pub const I64_EXTEND16_S: Op    = Op::new(0x00, 0xC3);
    pub const I64_EXTEND32_S: Op    = Op::new(0x00, 0xC4);
    // References
    pub const NULL: Op              = Op::new(0x00, 0xD0);
    pub const REF_IS_NULL: Op       = Op::new(0x00, 0xD1);
    pub const REF_FUNC: Op          = Op::new(0x00, 0xD2);
    // GC proposal extensions to the core prefix.
    pub const REF_EQ: Op            = Op::new(0x00, 0xD3);
    pub const REF_AS_NON_NULL: Op   = Op::new(0x00, 0xD4);
    pub const BR_ON_NULL: Op        = Op::new(0x00, 0xD5);
    pub const BR_ON_NON_NULL: Op    = Op::new(0x00, 0xD6);
}

// ── Metadata (name + operand format) ────────────────────────────

opcode_category! {
    // Control
    [0x00] unreachable => None, "unreachable";
    // block / loop carry (u16 end_offset, u8 result_count). The VM
    // only needs end_offset for label_stack bookkeeping; the count
    // tells the WASM emitter whether to write `void` (0), single
    // externref (1), or a shared function-type blocktype (>=2).
    [0x02] block => U16_U8, "block";
    [0x03] r#loop => U16_U8, "loop";
    [0x08] throw => None, "throw";
    [0x0A] throw_ref => None, "throw_ref";
    [0x0B] end => None, "end";
    [0x0C] br => I16, "br";
    [0x0D] br_if_true => I16, "br_if";
    [0x0E] br_table => BrTable, "br_table";
    [0x0F] r#return => None, "return";
    [0x10] call => U8, "call";
    [0x11] call_indirect => U8, "call_indirect";
    [0x12] return_call => U8, "return_call";
    [0x13] return_call_indirect => U8, "return_call_indirect";
    [0x14] call_ref => U8, "call_ref";
    [0x15] return_call_ref => U8, "return_call_ref";
    [0x1A] drop => None, "drop";
    [0x1B] select => None, "select";
    [0x1F] try_table => TryTable, "try_table";
    // Variables
    [0x20] local_get => U16, "local.get";
    [0x21] local_set => U16, "local.set";
    [0x23] global_get => U16, "global.get";
    [0x24] global_set => U16, "global.set";
    // Memory load
    [0x28] i32_load => None, "i32.load";
    [0x29] i64_load => None, "i64.load";
    [0x2A] f32_load => None, "f32.load";
    [0x2B] f64_load => None, "f64.load";
    [0x2C] i32_load8_s => None, "i32.load8_s";
    [0x2D] i32_load8_u => None, "i32.load8_u";
    [0x2E] i32_load16_s => None, "i32.load16_s";
    [0x2F] i32_load16_u => None, "i32.load16_u";
    [0x30] i64_load8_s => None, "i64.load8_s";
    [0x31] i64_load8_u => None, "i64.load8_u";
    [0x32] i64_load16_s => None, "i64.load16_s";
    [0x33] i64_load16_u => None, "i64.load16_u";
    [0x34] i64_load32_s => None, "i64.load32_s";
    [0x35] i64_load32_u => None, "i64.load32_u";
    // Memory store
    [0x36] i32_store => None, "i32.store";
    [0x37] i64_store => None, "i64.store";
    [0x38] f32_store => None, "f32.store";
    [0x39] f64_store => None, "f64.store";
    [0x3A] i32_store8 => None, "i32.store8";
    [0x3B] i32_store16 => None, "i32.store16";
    [0x3C] i64_store8 => None, "i64.store8";
    [0x3D] i64_store16 => None, "i64.store16";
    [0x3E] i64_store32 => None, "i64.store32";
    // Memory
    [0x3F] memory_size => None, "memory.size";
    [0x40] memory_grow => U16, "memory.grow";
    // i32 comparisons
    [0x45] i32_eqz => None, "i32.eqz";
    [0x46] eq => None, "i32.eq";
    [0x47] ne => None, "i32.ne";
    // i64
    [0x50] i64_eqz => None, "i64.eqz";
    // f64 comparisons
    [0x63] f64_lt => None, "f64.lt";
    [0x64] f64_gt => None, "f64.gt";
    [0x65] f64_le => None, "f64.le";
    [0x66] f64_ge => None, "f64.ge";
    // i32 arithmetic
    [0x67] i32_clz => None, "i32.clz";
    [0x68] i32_ctz => None, "i32.ctz";
    [0x69] i32_popcnt => None, "i32.popcnt";
    [0x6A] i32_add => None, "i32.add";
    [0x6B] i32_sub => None, "i32.sub";
    [0x6C] i32_mul => None, "i32.mul";
    [0x6D] i32_div_s => None, "i32.div_s";
    [0x6E] i32_div_u => None, "i32.div_u";
    [0x6F] i32_rem_s => None, "i32.rem_s";
    [0x70] i32_rem_u => None, "i32.rem_u";
    [0x71] i32_and => None, "i32.and";
    [0x72] i32_or => None, "i32.or";
    [0x73] i32_xor => None, "i32.xor";
    [0x74] i32_shl => None, "i32.shl";
    [0x75] i32_shr_s => None, "i32.shr_s";
    [0x76] i32_shr_u => None, "i32.shr_u";
    [0x77] i32_rotl => None, "i32.rotl";
    [0x78] i32_rotr => None, "i32.rotr";
    // i64 arithmetic
    [0x79] i64_clz => None, "i64.clz";
    [0x7A] i64_ctz => None, "i64.ctz";
    [0x7B] i64_popcnt => None, "i64.popcnt";
    [0x7C] i64_add => None, "i64.add";
    [0x7D] i64_sub => None, "i64.sub";
    [0x7E] i64_mul => None, "i64.mul";
    [0x7F] i64_div_s => None, "i64.div_s";
    [0x80] i64_div_u => None, "i64.div_u";
    [0x81] i64_rem_s => None, "i64.rem_s";
    [0x82] i64_rem_u => None, "i64.rem_u";
    [0x83] i64_and => None, "i64.and";
    [0x84] i64_or => None, "i64.or";
    [0x85] i64_xor => None, "i64.xor";
    [0x86] i64_shl => None, "i64.shl";
    [0x87] i64_shr_s => None, "i64.shr_s";
    [0x88] i64_shr_u => None, "i64.shr_u";
    [0x89] i64_rotl => None, "i64.rotl";
    [0x8A] i64_rotr => None, "i64.rotr";
    // f32 math
    [0x8B] f32_abs => None, "f32.abs";
    [0x8C] f32_neg => None, "f32.neg";
    [0x8D] f32_ceil => None, "f32.ceil";
    [0x8E] f32_floor => None, "f32.floor";
    [0x8F] f32_trunc => None, "f32.trunc";
    [0x90] f32_nearest => None, "f32.nearest";
    [0x91] f32_sqrt => None, "f32.sqrt";
    [0x96] f32_min => None, "f32.min";
    [0x97] f32_max => None, "f32.max";
    [0x98] f32_copysign => None, "f32.copysign";
    // f64 math
    [0x99] f64_abs => None, "f64.abs";
    [0x9A] f64_neg => None, "f64.neg";
    [0x9B] f64_ceil => None, "f64.ceil";
    [0x9C] f64_floor => None, "f64.floor";
    [0x9D] f64_trunc => None, "f64.trunc";
    [0x9E] f64_nearest => None, "f64.nearest";
    [0x9F] f64_sqrt => None, "f64.sqrt";
    [0xA0] f64_add => None, "f64.add";
    [0xA1] f64_sub => None, "f64.sub";
    [0xA2] f64_mul => None, "f64.mul";
    [0xA3] f64_div => None, "f64.div";
    [0xA4] f64_min => None, "f64.min";
    [0xA5] f64_max => None, "f64.max";
    [0xA6] f64_copysign => None, "f64.copysign";
    // Conversions
    [0xA7] i32_wrap_i64 => None, "i32.wrap_i64";
    [0xAA] i32_from_f64 => None, "i32.trunc_f64_s";
    [0xAC] i64_extend_i32_s => None, "i64.extend_i32_s";
    [0xAD] i64_extend_i32_u => None, "i64.extend_i32_u";
    [0xB0] i64_trunc_f64_s => None, "i64.trunc_f64_s";
    [0xB1] i64_trunc_f64_u => None, "i64.trunc_f64_u";
    [0xB6] f32_demote_f64 => None, "f32.demote_f64";
    [0xB7] f64_from_i32 => None, "f64.convert_i32_s";
    [0xBB] f64_promote_f32 => None, "f64.promote_f32";
    [0xBC] i32_reinterpret_f32 => None, "i32.reinterpret_f32";
    [0xBD] i64_reinterpret_f64 => None, "i64.reinterpret_f64";
    [0xBE] f32_reinterpret_i32 => None, "f32.reinterpret_i32";
    [0xBF] f64_reinterpret_i64 => None, "f64.reinterpret_i64";
    // Sign extension
    [0xC0] i32_extend8_s => None, "i32.extend8_s";
    [0xC1] i32_extend16_s => None, "i32.extend16_s";
    [0xC2] i64_extend8_s => None, "i64.extend8_s";
    [0xC3] i64_extend16_s => None, "i64.extend16_s";
    [0xC4] i64_extend32_s => None, "i64.extend32_s";
    // References
    [0xD0] null => None, "ref.null";
    [0xD1] ref_is_null => None, "ref.is_null";
    [0xD2] ref_func => Closure, "ref.func";
    // GC proposal (core prefix extensions).
    [0xD3] ref_eq => None, "ref.eq";
    [0xD4] ref_as_non_null => None, "ref.as_non_null";
    [0xD5] br_on_null => I16, "br_on_null";
    [0xD6] br_on_non_null => I16, "br_on_non_null";
}
