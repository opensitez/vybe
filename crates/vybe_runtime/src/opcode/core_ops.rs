//! Core WASM MVP opcodes (prefix 0x00).
//! Byte values match the WASM specification exactly.

use super::Op;
use super::opcode_category;

// ── Opcode constants ────────────────────────────────────────────
// Each is Op::new(0x00, <wasm_byte_value>).

impl Op {
    // Control
    pub const UNREACHABLE: Op = Op::new(0x00, 0x00);
    pub const NOP: Op = Op::new(0x00, 0x01);
    pub const BLOCK: Op = Op::new(0x00, 0x02);
    pub const LOOP: Op = Op::new(0x00, 0x03);
    pub const IF: Op = Op::new(0x00, 0x04);
    pub const ELSE: Op = Op::new(0x00, 0x05);
    pub const THROW: Op = Op::new(0x00, 0x08);
    pub const RETHROW: Op = Op::new(0x00, 0x09);
    pub const THROW_REF: Op = Op::new(0x00, 0x0A);
    pub const END: Op = Op::new(0x00, 0x0B);
    pub const BR: Op = Op::new(0x00, 0x0C);
    pub const BR_IF: Op = Op::new(0x00, 0x0D);
    pub const BR_IF_TRUE: Op = Op::BR_IF;
    pub const BR_TABLE: Op = Op::new(0x00, 0x0E);
    pub const RETURN: Op = Op::new(0x00, 0x0F);
    pub const CALL: Op = Op::new(0x00, 0x10);
    pub const CALL_INDIRECT: Op = Op::new(0x00, 0x11);
    pub const RETURN_CALL: Op = Op::new(0x00, 0x12);
    pub const RETURN_CALL_INDIRECT: Op = Op::new(0x00, 0x13);
    pub const CALL_REF: Op = Op::new(0x00, 0x14);
    pub const RETURN_CALL_REF: Op = Op::new(0x00, 0x15);
    pub const DELEGATE: Op = Op::new(0x00, 0x18);
    pub const DROP: Op = Op::new(0x00, 0x1A);
    pub const SELECT: Op = Op::new(0x00, 0x1B);
    pub const SELECT_T: Op = Op::new(0x00, 0x1C);
    pub const TRY_TABLE: Op = Op::new(0x00, 0x1F);
    // Variables
    pub const LOCAL_GET: Op = Op::new(0x00, 0x20);
    pub const LOCAL_SET: Op = Op::new(0x00, 0x21);
    pub const LOCAL_TEE: Op = Op::new(0x00, 0x22);
    pub const GLOBAL_GET: Op = Op::new(0x00, 0x23);
    pub const GLOBAL_SET: Op = Op::new(0x00, 0x24);
    // Reference-types table access (core prefix).
    pub const TABLE_GET: Op = Op::new(0x00, 0x25);
    pub const TABLE_SET: Op = Op::new(0x00, 0x26);
    // Memory load
    pub const I32_LOAD: Op = Op::new(0x00, 0x28);
    pub const I64_LOAD: Op = Op::new(0x00, 0x29);
    pub const F32_LOAD: Op = Op::new(0x00, 0x2A);
    pub const F64_LOAD: Op = Op::new(0x00, 0x2B);
    pub const I32_LOAD8_S: Op = Op::new(0x00, 0x2C);
    pub const I32_LOAD8_U: Op = Op::new(0x00, 0x2D);
    pub const I32_LOAD16_S: Op = Op::new(0x00, 0x2E);
    pub const I32_LOAD16_U: Op = Op::new(0x00, 0x2F);
    pub const I64_LOAD8_S: Op = Op::new(0x00, 0x30);
    pub const I64_LOAD8_U: Op = Op::new(0x00, 0x31);
    pub const I64_LOAD16_S: Op = Op::new(0x00, 0x32);
    pub const I64_LOAD16_U: Op = Op::new(0x00, 0x33);
    pub const I64_LOAD32_S: Op = Op::new(0x00, 0x34);
    pub const I64_LOAD32_U: Op = Op::new(0x00, 0x35);
    // Memory store
    pub const I32_STORE: Op = Op::new(0x00, 0x36);
    pub const I64_STORE: Op = Op::new(0x00, 0x37);
    pub const F32_STORE: Op = Op::new(0x00, 0x38);
    pub const F64_STORE: Op = Op::new(0x00, 0x39);
    pub const I32_STORE8: Op = Op::new(0x00, 0x3A);
    pub const I32_STORE16: Op = Op::new(0x00, 0x3B);
    pub const I64_STORE8: Op = Op::new(0x00, 0x3C);
    pub const I64_STORE16: Op = Op::new(0x00, 0x3D);
    pub const I64_STORE32: Op = Op::new(0x00, 0x3E);
    // Memory
    pub const MEMORY_SIZE: Op = Op::new(0x00, 0x3F);
    pub const MEMORY_GROW: Op = Op::new(0x00, 0x40);
    // Numeric constants (WASM MVP 0x41–0x44)
    pub const I32_CONST: Op = Op::new(0x00, 0x41);
    pub const I64_CONST: Op = Op::new(0x00, 0x42);
    pub const F32_CONST: Op = Op::new(0x00, 0x43);
    pub const F64_CONST: Op = Op::new(0x00, 0x44);
    // i32 comparisons (WASM MVP 0x45–0x4F)
    pub const I32_EQZ: Op = Op::new(0x00, 0x45);
    pub const I32_EQ: Op = Op::new(0x00, 0x46);
    pub const I32_NE: Op = Op::new(0x00, 0x47);
    pub const I32_LT_S: Op = Op::new(0x00, 0x48);
    pub const I32_LT_U: Op = Op::new(0x00, 0x49);
    pub const I32_GT_S: Op = Op::new(0x00, 0x4A);
    pub const I32_GT_U: Op = Op::new(0x00, 0x4B);
    pub const I32_LE_S: Op = Op::new(0x00, 0x4C);
    pub const I32_LE_U: Op = Op::new(0x00, 0x4D);
    pub const I32_GE_S: Op = Op::new(0x00, 0x4E);
    pub const I32_GE_U: Op = Op::new(0x00, 0x4F);
    /// Alias for I32_EQ — kept for compatibility with existing emit sites.
    pub const EQ: Op = Op::new(0x00, 0x46);
    /// Alias for I32_NE — kept for compatibility with existing emit sites.
    pub const NE: Op = Op::new(0x00, 0x47);
    // f32 comparisons (WASM MVP 0x5B–0x60)
    pub const F32_EQ: Op = Op::new(0x00, 0x5B);
    pub const F32_NE: Op = Op::new(0x00, 0x5C);
    pub const F32_LT: Op = Op::new(0x00, 0x5D);
    pub const F32_GT: Op = Op::new(0x00, 0x5E);
    pub const F32_LE: Op = Op::new(0x00, 0x5F);
    pub const F32_GE: Op = Op::new(0x00, 0x60);
    // i64 comparisons (WASM MVP 0x50–0x5A)
    pub const I64_EQZ: Op = Op::new(0x00, 0x50);
    pub const I64_EQ: Op = Op::new(0x00, 0x51);
    pub const I64_NE: Op = Op::new(0x00, 0x52);
    pub const I64_LT_S: Op = Op::new(0x00, 0x53);
    pub const I64_LT_U: Op = Op::new(0x00, 0x54);
    pub const I64_GT_S: Op = Op::new(0x00, 0x55);
    pub const I64_GT_U: Op = Op::new(0x00, 0x56);
    pub const I64_LE_S: Op = Op::new(0x00, 0x57);
    pub const I64_LE_U: Op = Op::new(0x00, 0x58);
    pub const I64_GE_S: Op = Op::new(0x00, 0x59);
    pub const I64_GE_U: Op = Op::new(0x00, 0x5A);
    // f64 comparisons (WASM MVP 0x61–0x66)
    pub const F64_EQ: Op = Op::new(0x00, 0x61);
    pub const F64_NE: Op = Op::new(0x00, 0x62);
    pub const F64_LT: Op = Op::new(0x00, 0x63);
    pub const F64_GT: Op = Op::new(0x00, 0x64);
    pub const F64_LE: Op = Op::new(0x00, 0x65);
    pub const F64_GE: Op = Op::new(0x00, 0x66);
    // i32 arithmetic
    pub const I32_CLZ: Op = Op::new(0x00, 0x67);
    pub const I32_CTZ: Op = Op::new(0x00, 0x68);
    pub const I32_POPCNT: Op = Op::new(0x00, 0x69);
    pub const I32_ADD: Op = Op::new(0x00, 0x6A);
    pub const I32_SUB: Op = Op::new(0x00, 0x6B);
    pub const I32_MUL: Op = Op::new(0x00, 0x6C);
    pub const I32_DIV_S: Op = Op::new(0x00, 0x6D);
    pub const I32_DIV_U: Op = Op::new(0x00, 0x6E);
    pub const I32_REM_S: Op = Op::new(0x00, 0x6F);
    pub const I32_REM_U: Op = Op::new(0x00, 0x70);
    pub const I32_AND: Op = Op::new(0x00, 0x71);
    pub const I32_OR: Op = Op::new(0x00, 0x72);
    pub const I32_XOR: Op = Op::new(0x00, 0x73);
    pub const I32_SHL: Op = Op::new(0x00, 0x74);
    pub const I32_SHR_S: Op = Op::new(0x00, 0x75);
    pub const I32_SHR_U: Op = Op::new(0x00, 0x76);
    pub const I32_ROTL: Op = Op::new(0x00, 0x77);
    pub const I32_ROTR: Op = Op::new(0x00, 0x78);
    // i64 arithmetic
    pub const I64_CLZ: Op = Op::new(0x00, 0x79);
    pub const I64_CTZ: Op = Op::new(0x00, 0x7A);
    pub const I64_POPCNT: Op = Op::new(0x00, 0x7B);
    pub const I64_ADD: Op = Op::new(0x00, 0x7C);
    pub const I64_SUB: Op = Op::new(0x00, 0x7D);
    pub const I64_MUL: Op = Op::new(0x00, 0x7E);
    pub const I64_DIV_S: Op = Op::new(0x00, 0x7F);
    pub const I64_DIV_U: Op = Op::new(0x00, 0x80);
    pub const I64_REM_S: Op = Op::new(0x00, 0x81);
    pub const I64_REM_U: Op = Op::new(0x00, 0x82);
    pub const I64_AND: Op = Op::new(0x00, 0x83);
    pub const I64_OR: Op = Op::new(0x00, 0x84);
    pub const I64_XOR: Op = Op::new(0x00, 0x85);
    pub const I64_SHL: Op = Op::new(0x00, 0x86);
    pub const I64_SHR_S: Op = Op::new(0x00, 0x87);
    pub const I64_SHR_U: Op = Op::new(0x00, 0x88);
    pub const I64_ROTL: Op = Op::new(0x00, 0x89);
    pub const I64_ROTR: Op = Op::new(0x00, 0x8A);
    // f32 math
    pub const F32_ABS: Op = Op::new(0x00, 0x8B);
    pub const F32_NEG: Op = Op::new(0x00, 0x8C);
    pub const F32_CEIL: Op = Op::new(0x00, 0x8D);
    pub const F32_FLOOR: Op = Op::new(0x00, 0x8E);
    pub const F32_TRUNC: Op = Op::new(0x00, 0x8F);
    pub const F32_NEAREST: Op = Op::new(0x00, 0x90);
    pub const F32_SQRT: Op = Op::new(0x00, 0x91);
    pub const F32_ADD: Op = Op::new(0x00, 0x92);
    pub const F32_SUB: Op = Op::new(0x00, 0x93);
    pub const F32_MUL: Op = Op::new(0x00, 0x94);
    pub const F32_DIV: Op = Op::new(0x00, 0x95);
    pub const F32_MIN: Op = Op::new(0x00, 0x96);
    pub const F32_MAX: Op = Op::new(0x00, 0x97);
    pub const F32_COPYSIGN: Op = Op::new(0x00, 0x98);
    // f64 math
    pub const F64_ABS: Op = Op::new(0x00, 0x99);
    pub const F64_NEG: Op = Op::new(0x00, 0x9A);
    pub const F64_CEIL: Op = Op::new(0x00, 0x9B);
    pub const F64_FLOOR: Op = Op::new(0x00, 0x9C);
    pub const F64_TRUNC: Op = Op::new(0x00, 0x9D);
    pub const F64_NEAREST: Op = Op::new(0x00, 0x9E);
    pub const F64_SQRT: Op = Op::new(0x00, 0x9F);
    pub const F64_ADD: Op = Op::new(0x00, 0xA0);
    pub const F64_SUB: Op = Op::new(0x00, 0xA1);
    pub const F64_MUL: Op = Op::new(0x00, 0xA2);
    pub const F64_DIV: Op = Op::new(0x00, 0xA3);
    pub const F64_MIN: Op = Op::new(0x00, 0xA4);
    pub const F64_MAX: Op = Op::new(0x00, 0xA5);
    pub const F64_COPYSIGN: Op = Op::new(0x00, 0xA6);
    // Conversions
    pub const I32_WRAP_I64: Op = Op::new(0x00, 0xA7);
    pub const I32_TRUNC_F32_S: Op = Op::new(0x00, 0xA8);
    pub const I32_TRUNC_F32_U: Op = Op::new(0x00, 0xA9);
    pub const I32_FROM_F64: Op = Op::new(0x00, 0xAA);
    pub const I32_TRUNC_F64_U: Op = Op::new(0x00, 0xAB);
    pub const I64_EXTEND_I32_S: Op = Op::new(0x00, 0xAC);
    pub const I64_EXTEND_I32_U: Op = Op::new(0x00, 0xAD);
    pub const I64_TRUNC_F32_S: Op = Op::new(0x00, 0xAE);
    pub const I64_TRUNC_F32_U: Op = Op::new(0x00, 0xAF);
    pub const I64_TRUNC_F64_S: Op = Op::new(0x00, 0xB0);
    pub const I64_TRUNC_F64_U: Op = Op::new(0x00, 0xB1);
    pub const F32_CONVERT_I32_S: Op = Op::new(0x00, 0xB2);
    pub const F32_CONVERT_I32_U: Op = Op::new(0x00, 0xB3);
    pub const F32_CONVERT_I64_S: Op = Op::new(0x00, 0xB4);
    pub const F32_CONVERT_I64_U: Op = Op::new(0x00, 0xB5);
    pub const F32_DEMOTE_F64: Op = Op::new(0x00, 0xB6);
    pub const F64_FROM_I32: Op = Op::new(0x00, 0xB7);
    pub const F64_CONVERT_I32_U: Op = Op::new(0x00, 0xB8);
    pub const F64_CONVERT_I64_S: Op = Op::new(0x00, 0xB9);
    pub const F64_CONVERT_I64_U: Op = Op::new(0x00, 0xBA);
    pub const F64_PROMOTE_F32: Op = Op::new(0x00, 0xBB);
    pub const I32_REINTERPRET_F32: Op = Op::new(0x00, 0xBC);
    pub const I64_REINTERPRET_F64: Op = Op::new(0x00, 0xBD);
    pub const F32_REINTERPRET_I32: Op = Op::new(0x00, 0xBE);
    pub const F64_REINTERPRET_I64: Op = Op::new(0x00, 0xBF);
    // Sign extension
    pub const I32_EXTEND8_S: Op = Op::new(0x00, 0xC0);
    pub const I32_EXTEND16_S: Op = Op::new(0x00, 0xC1);
    pub const I64_EXTEND8_S: Op = Op::new(0x00, 0xC2);
    pub const I64_EXTEND16_S: Op = Op::new(0x00, 0xC3);
    pub const I64_EXTEND32_S: Op = Op::new(0x00, 0xC4);
    // References
    pub const NULL: Op = Op::new(0x00, 0xD0);
    pub const REF_IS_NULL: Op = Op::new(0x00, 0xD1);
    pub const REF_FUNC: Op = Op::new(0x00, 0xD2);
    // GC proposal extensions to the core prefix.
    pub const REF_EQ: Op = Op::new(0x00, 0xD3);
    pub const REF_AS_NON_NULL: Op = Op::new(0x00, 0xD4);
    pub const BR_ON_NULL: Op = Op::new(0x00, 0xD5);
    pub const BR_ON_NON_NULL: Op = Op::new(0x00, 0xD6);
    // Stack-switching proposal (WebAssembly/stack-switching). These are
    // real core-prefix WASM opcodes — spec bytes 0xE0..=0xE6, per
    // proposals/stack-switching/interpreter/binary/encode.ml — NOT
    // VM-internal opcodes.
    pub const CONT_NEW: Op = Op::new(0x00, 0xE0);
    pub const CONT_BIND: Op = Op::new(0x00, 0xE1);
    pub const SUSPEND: Op = Op::new(0x00, 0xE2);
    pub const RESUME: Op = Op::new(0x00, 0xE3);
    pub const RESUME_THROW: Op = Op::new(0x00, 0xE4);
    pub const RESUME_THROW_REF: Op = Op::new(0x00, 0xE5);
    pub const SWITCH: Op = Op::new(0x00, 0xE6);
}

// ── Metadata (name + operand format) ────────────────────────────

opcode_category! {
    // Control
    [0x00] unreachable => None, "unreachable";
    [0x01] nop => None, "nop";
    // Blocktype: TWO bytes — (param_count, result_count). The spec blocktype
    // is an s33 (0x40 empty / valtype / positive typeidx); internally both
    // counts are carried so the writer can reconstruct a full functype
    // blocktype and the VM knows each label's branch arity (spec: br to a
    // block/if label carries its RESULTS; br to a loop label carries its
    // PARAMS).
    [0x02] block => U8_U8, "block";
    [0x03] r#loop => U8_U8, "loop";
    [0x04] if_blk => U8_U8, "if";
    [0x05] else_blk => None, "else";
    [0x08] throw => U16, "throw";
    [0x09] rethrow => U32Leb, "rethrow";
    [0x0A] throw_ref => None, "throw_ref";
    [0x0B] end => None, "end";
    [0x0C] br => U32Leb, "br";
    [0x0D] br_if => U32Leb, "br_if";
    [0x0E] br_table => BrTable, "br_table";
    [0x0F] r#return => None, "return";
    // Spec `call funcidx` — a STATIC call. VM-internal immediates are
    // `u16 funcidx` + `u8 argc` (the argc byte is VM-internal, exactly like
    // the retired CALL_IMPORT's; the .wasm writer drops it and emits
    // `0x10` + LEB(funcidx)). Internally funcidx is chunk-scoped: it names
    // an entry in the frame chunk's import table — nothing emits local-
    // function indices here; those go through REF_FUNC + CALL_REF. The old
    // dynamic callee-on-stack `call` (byte-identical to `call_ref`) is
    // retired; see callimportretirement.md.
    [0x10] call => U16_U8, "call";
    [0x11] call_indirect => U8_U8_U8, "call_indirect";
    // Dynamic calls carry (argc, results) so the callee functype survives
    // round-trips exactly: compilers emit results=1 (the uniform boxed-value
    // ABI); the reader stamps the ingested functype's result count; the
    // writer keys the exact (params, results) functype off both bytes.
    [0x12] return_call => U8_U8, "return_call";
    [0x13] return_call_indirect => U8_U8_U8, "return_call_indirect";
    [0x14] call_ref => U8_U8, "call_ref";
    [0x15] return_call_ref => U8_U8, "return_call_ref";
    [0x1A] drop => None, "drop";
    [0x1B] select => None, "select";
    // Typed select carries a `vec(valtype)` — currently always 1 externref
    // for our uniform ABI, so the operand is encoded inline in the emitter
    // rather than carried by the bytecode.
    [0x1C] select_t => None, "select_t";
    [0x18] delegate => U32Leb, "delegate";
    [0x1F] try_table => TryTable, "try_table";
    // Variables
    [0x20] local_get => U16, "local.get";
    [0x21] local_set => U16, "local.set";
    [0x22] local_tee => U16, "local.tee";
    [0x23] global_get => U16, "global.get";
    [0x24] global_set => U16, "global.set";
    // Reference-types table access (core prefix). Operand is a u8 table
    // index — Vybe's single function-table is index 0.
    [0x25] table_get => U16, "table.get";
    [0x26] table_set => U16, "table.set";
    // Memory load
    // Loads/stores carry the OPTIONAL marker-tagged memarg (`SimdMemArg`
    // treatment): present iff the first LEB has 0x80; 0x100 = memory64
    // offset; 0x40 = explicit memidx follows (the spec multi-memory bit).
    // Absent = align natural, offset 0, memory 0. The spec binary always
    // writes a memarg; the writer materializes the defaults on the way out.
    [0x28] i32_load => SimdMemArg, "i32.load";
    [0x29] i64_load => SimdMemArg, "i64.load";
    [0x2A] f32_load => SimdMemArg, "f32.load";
    [0x2B] f64_load => SimdMemArg, "f64.load";
    [0x2C] i32_load8_s => SimdMemArg, "i32.load8_s";
    [0x2D] i32_load8_u => SimdMemArg, "i32.load8_u";
    [0x2E] i32_load16_s => SimdMemArg, "i32.load16_s";
    [0x2F] i32_load16_u => SimdMemArg, "i32.load16_u";
    [0x30] i64_load8_s => SimdMemArg, "i64.load8_s";
    [0x31] i64_load8_u => SimdMemArg, "i64.load8_u";
    [0x32] i64_load16_s => SimdMemArg, "i64.load16_s";
    [0x33] i64_load16_u => SimdMemArg, "i64.load16_u";
    [0x34] i64_load32_s => SimdMemArg, "i64.load32_s";
    [0x35] i64_load32_u => SimdMemArg, "i64.load32_u";
    // Memory store
    [0x36] i32_store => SimdMemArg, "i32.store";
    [0x37] i64_store => SimdMemArg, "i64.store";
    [0x38] f32_store => SimdMemArg, "f32.store";
    [0x39] f64_store => SimdMemArg, "f64.store";
    [0x3A] i32_store8 => SimdMemArg, "i32.store8";
    [0x3B] i32_store16 => SimdMemArg, "i32.store16";
    [0x3C] i64_store8 => SimdMemArg, "i64.store8";
    [0x3D] i64_store16 => SimdMemArg, "i64.store16";
    [0x3E] i64_store32 => SimdMemArg, "i64.store32";
    // Memory
    // Fixed u16 memidx immediate (multi-memory). The optional 0xEE selector
    // block is retired for these ops: an undeclared conditional immediate
    // desynced every format-driven walk (the memarg lesson, again).
    [0x3F] memory_size => U16, "memory.size";
    [0x40] memory_grow => U16, "memory.grow";
    // Numeric constants
    [0x41] i32_const => SlI32, "i32.const";
    [0x42] i64_const => SlI64, "i64.const";
    [0x43] f32_const => RawF32, "f32.const";
    [0x44] f64_const => RawF64, "f64.const";
    // i32 comparisons
    [0x45] i32_eqz => None, "i32.eqz";
    [0x46] eq      => None, "i32.eq";
    [0x47] ne      => None, "i32.ne";
    [0x48] i32_lt_s => None, "i32.lt_s";
    [0x49] i32_lt_u => None, "i32.lt_u";
    [0x4A] i32_gt_s => None, "i32.gt_s";
    [0x4B] i32_gt_u => None, "i32.gt_u";
    [0x4C] i32_le_s => None, "i32.le_s";
    [0x4D] i32_le_u => None, "i32.le_u";
    [0x4E] i32_ge_s => None, "i32.ge_s";
    [0x4F] i32_ge_u => None, "i32.ge_u";
    // i64
    [0x50] i64_eqz => None, "i64.eqz";
    [0x51] i64_eq  => None, "i64.eq";
    [0x52] i64_ne  => None, "i64.ne";
    [0x53] i64_lt_s => None, "i64.lt_s";
    [0x54] i64_lt_u => None, "i64.lt_u";
    [0x55] i64_gt_s => None, "i64.gt_s";
    [0x56] i64_gt_u => None, "i64.gt_u";
    [0x57] i64_le_s => None, "i64.le_s";
    [0x58] i64_le_u => None, "i64.le_u";
    [0x59] i64_ge_s => None, "i64.ge_s";
    [0x5A] i64_ge_u => None, "i64.ge_u";
    // f32 comparisons (WASM MVP 0x5B–0x60)
    [0x5B] f32_eq => None, "f32.eq";
    [0x5C] f32_ne => None, "f32.ne";
    [0x5D] f32_lt => None, "f32.lt";
    [0x5E] f32_gt => None, "f32.gt";
    [0x5F] f32_le => None, "f32.le";
    [0x60] f32_ge => None, "f32.ge";
    // f64 comparisons
    [0x61] f64_eq => None, "f64.eq";
    [0x62] f64_ne => None, "f64.ne";
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
    [0x92] f32_add => None, "f32.add";
    [0x93] f32_sub => None, "f32.sub";
    [0x94] f32_mul => None, "f32.mul";
    [0x95] f32_div => None, "f32.div";
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
    [0xA8] i32_trunc_f32_s => None, "i32.trunc_f32_s";
    [0xA9] i32_trunc_f32_u => None, "i32.trunc_f32_u";
    [0xAA] i32_from_f64 => None, "i32.trunc_f64_s";
    [0xAB] i32_trunc_f64_u => None, "i32.trunc_f64_u";
    [0xAC] i64_extend_i32_s => None, "i64.extend_i32_s";
    [0xAD] i64_extend_i32_u => None, "i64.extend_i32_u";
    [0xAE] i64_trunc_f32_s => None, "i64.trunc_f32_s";
    [0xAF] i64_trunc_f32_u => None, "i64.trunc_f32_u";
    [0xB0] i64_trunc_f64_s => None, "i64.trunc_f64_s";
    [0xB1] i64_trunc_f64_u => None, "i64.trunc_f64_u";
    [0xB2] f32_convert_i32_s => None, "f32.convert_i32_s";
    [0xB3] f32_convert_i32_u => None, "f32.convert_i32_u";
    [0xB4] f32_convert_i64_s => None, "f32.convert_i64_s";
    [0xB5] f32_convert_i64_u => None, "f32.convert_i64_u";
    [0xB6] f32_demote_f64 => None, "f32.demote_f64";
    [0xB7] f64_from_i32 => None, "f64.convert_i32_s";
    [0xB8] f64_convert_i32_u => None, "f64.convert_i32_u";
    [0xB9] f64_convert_i64_s => None, "f64.convert_i64_s";
    [0xBA] f64_convert_i64_u => None, "f64.convert_i64_u";
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
    [0xD0] null => U8, "ref.null";
    [0xD1] ref_is_null => None, "ref.is_null";
    [0xD2] ref_func => Closure, "ref.func";
    // GC proposal (core prefix extensions).
    [0xD3] ref_eq => None, "ref.eq";
    [0xD4] ref_as_non_null => None, "ref.as_non_null";
    [0xD5] br_on_null => I16, "br_on_null";
    [0xD6] br_on_non_null => I16, "br_on_non_null";
    // Stack-switching proposal (WebAssembly/stack-switching). Spec bytes
    // 0xE0..=0xE6. The in-memory operand widths below are Vybe's fixed
    // 2-byte encoding; the .wasm emitter/reader translate to/from the
    // spec's LEB var + resumetable forms.
    [0xE0] cont_new => None, "cont.new";
    [0xE1] cont_bind => U8, "cont.bind";
    [0xE2] suspend => U16, "suspend";
    [0xE3] resume => U16, "resume";
    [0xE4] resume_throw => U16, "resume_throw";
    [0xE5] resume_throw_ref => None, "resume_throw_ref";
    [0xE6] switch => U16, "switch";
}
