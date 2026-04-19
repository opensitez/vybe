//! Misc proposal opcodes (prefix 0xFC).

use super::Op;
use super::opcode_category;

impl Op {
    pub const MEMORY_INIT: Op = Op::new(0xFC, 0x08);
    // Reference-types table operations (reference-types proposal, 0xFC prefix).
    pub const DATA_DROP: Op   = Op::new(0xFC, 0x09);
    pub const MEMORY_COPY: Op = Op::new(0xFC, 0x0A);
    pub const MEMORY_FILL: Op = Op::new(0xFC, 0x0B);
    pub const TABLE_INIT: Op  = Op::new(0xFC, 0x0C);
    pub const ELEM_DROP: Op   = Op::new(0xFC, 0x0D);
    pub const TABLE_COPY: Op  = Op::new(0xFC, 0x0E);
    pub const TABLE_GROW: Op  = Op::new(0xFC, 0x0F);
    pub const TABLE_SIZE: Op  = Op::new(0xFC, 0x10);
    pub const TABLE_FILL: Op  = Op::new(0xFC, 0x11);
}

opcode_category! {
    // Spec-correct operands:
    //   memory.init   : u8 data_idx, u8 memory_idx (we encode 0, 0)
    //   data.drop     : u8 data_idx
    //   memory.copy   : u8 dst_mem, u8 src_mem  (both 0 for our one memory)
    //   memory.fill   : u8 memory_idx
    //   table.init    : u8 elem_idx, u8 table_idx
    //   elem.drop     : u8 elem_idx
    //   table.copy    : u8 dst_table, u8 src_table
    //   table.grow/size/fill : u8 table_idx
    [0x08] memory_init => U8, "memory.init";       // data_idx; mem_idx appended by emitter
    [0x09] data_drop   => U8, "data.drop";
    [0x0A] memory_copy => None, "memory.copy";     // emitter appends 0 0
    [0x0B] memory_fill => None, "memory.fill";     // emitter appends 0
    [0x0C] table_init  => U8, "table.init";
    [0x0D] elem_drop   => U8, "elem.drop";
    [0x0E] table_copy  => U8, "table.copy";        // u8 = table index; emitter appends trailing 0
    [0x0F] table_grow  => U8, "table.grow";        // u8 = table index
    [0x10] table_size  => U8, "table.size";        // u8 = table index
    [0x11] table_fill  => U8, "table.fill";        // u8 = table index
}
