//! Misc proposal opcodes (prefix 0xFC).
//!
//! Includes the nontrapping-float-to-int-conversions proposal (0xFC 0x00–0x07):
//! saturating truncation that clamps instead of trapping on overflow / NaN.

use super::Op;
use super::opcode_category;

impl Op {
    // nontrapping-float-to-int-conversions proposal — 0xFC 0x00–0x07
    pub const I32_TRUNC_SAT_F32_S: Op = Op::new(0xFC, 0x00);
    pub const I32_TRUNC_SAT_F32_U: Op = Op::new(0xFC, 0x01);
    pub const I32_TRUNC_SAT_F64_S: Op = Op::new(0xFC, 0x02);
    pub const I32_TRUNC_SAT_F64_U: Op = Op::new(0xFC, 0x03);
    pub const I64_TRUNC_SAT_F32_S: Op = Op::new(0xFC, 0x04);
    pub const I64_TRUNC_SAT_F32_U: Op = Op::new(0xFC, 0x05);
    pub const I64_TRUNC_SAT_F64_S: Op = Op::new(0xFC, 0x06);
    pub const I64_TRUNC_SAT_F64_U: Op = Op::new(0xFC, 0x07);

    pub const MEMORY_INIT: Op = Op::new(0xFC, 0x08);
    // Reference-types table operations (reference-types proposal, 0xFC prefix).
    pub const DATA_DROP: Op = Op::new(0xFC, 0x09);
    pub const MEMORY_COPY: Op = Op::new(0xFC, 0x0A);
    pub const MEMORY_FILL: Op = Op::new(0xFC, 0x0B);
    pub const TABLE_INIT: Op = Op::new(0xFC, 0x0C);
    pub const ELEM_DROP: Op = Op::new(0xFC, 0x0D);
    pub const TABLE_COPY: Op = Op::new(0xFC, 0x0E);
    pub const TABLE_GROW: Op = Op::new(0xFC, 0x0F);
    pub const TABLE_SIZE: Op = Op::new(0xFC, 0x10);
    pub const TABLE_FILL: Op = Op::new(0xFC, 0x11);
}

opcode_category! {
    // nontrapping-float-to-int-conversions — no immediates
    [0x00] i32_trunc_sat_f32_s => None, "i32.trunc_sat_f32_s";
    [0x01] i32_trunc_sat_f32_u => None, "i32.trunc_sat_f32_u";
    [0x02] i32_trunc_sat_f64_s => None, "i32.trunc_sat_f64_s";
    [0x03] i32_trunc_sat_f64_u => None, "i32.trunc_sat_f64_u";
    [0x04] i64_trunc_sat_f32_s => None, "i64.trunc_sat_f32_s";
    [0x05] i64_trunc_sat_f32_u => None, "i64.trunc_sat_f32_u";
    [0x06] i64_trunc_sat_f64_s => None, "i64.trunc_sat_f64_s";
    [0x07] i64_trunc_sat_f64_u => None, "i64.trunc_sat_f64_u";
    // Internal immediates are fixed u16 BE (the VM's uniform index width —
    // spec indices are u32 LEBs; the reader rejects > u16::MAX loudly and
    // the writer re-serializes as LEB):
    //   memory.init   : u16 data_idx, u16 memidx
    //   data.drop     : u16 data_idx
    //   memory.copy   : u16 dst_mem, u16 src_mem
    //   memory.fill   : u16 memidx
    //   table.init    : u16 elem_idx, u16 table_idx
    //   elem.drop     : u16 elem_idx
    //   table.copy    : u16 dst_table, u16 src_table
    //   table.grow/size/fill : u16 table_idx
    // (The optional 0xEE memidx selector block is RETIRED — an undeclared
    // conditional immediate desynced every format-driven walk.)
    [0x08] memory_init => U32Leb_U32Leb, "memory.init";
    [0x09] data_drop   => U32Leb, "data.drop";
    [0x0A] memory_copy => U32Leb_U32Leb, "memory.copy";
    [0x0B] memory_fill => U32Leb, "memory.fill";
    [0x0C] table_init  => U32Leb_U32Leb, "table.init";
    [0x0D] elem_drop   => U32Leb, "elem.drop";
    [0x0E] table_copy  => U32Leb_U32Leb, "table.copy";
    [0x0F] table_grow  => U32Leb, "table.grow";
    [0x10] table_size  => U32Leb, "table.size";
    [0x11] table_fill  => U32Leb, "table.fill";
}
