//! GC proposal opcodes (prefix 0xFB).
//! Byte values match the WASM GC specification.

use super::Op;
use super::opcode_category;

impl Op {
    pub const STRUCT_NEW: Op        = Op::new(0xFB, 0x00);
    pub const STRUCT_GET: Op        = Op::new(0xFB, 0x02);
    pub const STRUCT_SET: Op        = Op::new(0xFB, 0x05);
    pub const ARRAY_NEW: Op         = Op::new(0xFB, 0x06);
    pub const ARRAY_NEW_DEFAULT: Op = Op::new(0xFB, 0x07);
    pub const ARRAY_GET: Op         = Op::new(0xFB, 0x0B);
    pub const ARRAY_SET: Op         = Op::new(0xFB, 0x0E);
    pub const ARRAY_LENGTH: Op      = Op::new(0xFB, 0x0F);
    pub const ARRAY_FILL: Op        = Op::new(0xFB, 0x10);
    pub const ARRAY_COPY: Op        = Op::new(0xFB, 0x11);
    pub const REF_TEST: Op          = Op::new(0xFB, 0x14);
    pub const REF_CAST: Op          = Op::new(0xFB, 0x16);
    pub const BR_ON_CAST: Op        = Op::new(0xFB, 0x18);
    pub const BR_ON_CAST_FAIL: Op   = Op::new(0xFB, 0x19);
    pub const I31_NEW: Op           = Op::new(0xFB, 0x1C);
    pub const I31_GET_S: Op         = Op::new(0xFB, 0x1D);
    pub const I31_GET_U: Op         = Op::new(0xFB, 0x1E);
    // Custom Descriptors proposal
    pub const STRUCT_NEW_DESC: Op       = Op::new(0xFB, 0x20); // struct.new_desc $typeidx
    pub const STRUCT_NEW_DEFAULT_DESC: Op = Op::new(0xFB, 0x21); // struct.new_default_desc $typeidx
    pub const REF_GET_DESC: Op          = Op::new(0xFB, 0x22); // ref.get_desc $typeidx
}

opcode_category! {
    [0x00] struct_new => U16, "struct.new";
    [0x02] struct_get => U16, "struct.get";
    [0x05] struct_set => U16, "struct.set";
    [0x06] array_new => U16, "array.new_fixed";
    [0x07] array_new_default => None, "array.new_default";
    [0x0B] array_get => None, "array.get";
    [0x0E] array_set => None, "array.set";
    [0x0F] array_length => None, "array.len";
    [0x10] array_fill => None, "array.fill";
    [0x11] array_copy => None, "array.copy";
    [0x14] ref_test => U16, "ref.test";
    [0x16] ref_cast => U16, "ref.cast";
    [0x18] br_on_cast => U16_I16, "br_on_cast";
    [0x19] br_on_cast_fail => U16_I16, "br_on_cast_fail";
    [0x1C] i31_new => None, "ref.i31";
    [0x1D] i31_get_s => None, "i31.get_s";
    [0x1E] i31_get_u => None, "i31.get_u";
    // Custom Descriptors
    [0x20] struct_new_desc => U16, "struct.new_desc";
    [0x21] struct_new_default_desc => U16, "struct.new_default_desc";
    [0x22] ref_get_desc => U16, "ref.get_desc";
}
