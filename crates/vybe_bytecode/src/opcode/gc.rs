//! GC proposal opcodes (prefix 0xFB).
//! Byte values match the WASM GC specification.

use super::Op;
use super::opcode_category;

impl Op {
    // Struct ops (0x00..=0x05)
    pub const STRUCT_NEW: Op         = Op::new(0xFB, 0x00);
    pub const STRUCT_NEW_DEFAULT: Op = Op::new(0xFB, 0x01);
    pub const STRUCT_GET: Op         = Op::new(0xFB, 0x02);
    pub const STRUCT_GET_S: Op       = Op::new(0xFB, 0x03);
    pub const STRUCT_GET_U: Op       = Op::new(0xFB, 0x04);
    pub const STRUCT_SET: Op         = Op::new(0xFB, 0x05);
    // Array ops (0x06..=0x13)
    //
    // Spec splits array construction three ways:
    //   0x06 array.new         : [value, length] -> [array]
    //   0x07 array.new_default : [length]        -> [array]
    //   0x08 array.new_fixed N : [v1..vN]        -> [array]
    //
    // Our compilers pre-push N values then emit `ARRAY_NEW_FIXED` with an
    // N immediate — which matches `array.new_fixed` (0x08), NOT `array.new`.
    // The previous `Op::ARRAY_NEW` constant was at 0x06 but semantically
    // performed 0x08 and was named "array.new_fixed" in the category —
    // double-wrong. Fixed by moving to 0x08 and renaming the constant.
    pub const ARRAY_NEW: Op          = Op::new(0xFB, 0x06);
    pub const ARRAY_NEW_DEFAULT: Op  = Op::new(0xFB, 0x07);
    pub const ARRAY_NEW_FIXED: Op    = Op::new(0xFB, 0x08);
    pub const ARRAY_NEW_DATA: Op     = Op::new(0xFB, 0x09);
    pub const ARRAY_NEW_ELEM: Op     = Op::new(0xFB, 0x0A);
    pub const ARRAY_GET: Op          = Op::new(0xFB, 0x0B);
    pub const ARRAY_GET_S: Op        = Op::new(0xFB, 0x0C);
    pub const ARRAY_GET_U: Op        = Op::new(0xFB, 0x0D);
    pub const ARRAY_SET: Op          = Op::new(0xFB, 0x0E);
    pub const ARRAY_LENGTH: Op       = Op::new(0xFB, 0x0F);
    pub const ARRAY_FILL: Op         = Op::new(0xFB, 0x10);
    pub const ARRAY_COPY: Op         = Op::new(0xFB, 0x11);
    pub const ARRAY_INIT_DATA: Op    = Op::new(0xFB, 0x12);
    pub const ARRAY_INIT_ELEM: Op    = Op::new(0xFB, 0x13);
    // Reference tests / casts (0x14..=0x17)
    pub const REF_TEST: Op           = Op::new(0xFB, 0x14);
    pub const REF_TEST_NULL: Op      = Op::new(0xFB, 0x15);
    pub const REF_CAST: Op           = Op::new(0xFB, 0x16);
    pub const REF_CAST_NULL: Op      = Op::new(0xFB, 0x17);
    pub const BR_ON_CAST: Op         = Op::new(0xFB, 0x18);
    pub const BR_ON_CAST_FAIL: Op    = Op::new(0xFB, 0x19);
    // Extern <-> any conversion (0x1A, 0x1B)
    pub const ANY_CONVERT_EXTERN: Op = Op::new(0xFB, 0x1A);
    pub const EXTERN_CONVERT_ANY: Op = Op::new(0xFB, 0x1B);
    // i31 ops (0x1C..=0x1E)
    pub const I31_NEW: Op            = Op::new(0xFB, 0x1C);
    pub const I31_GET_S: Op          = Op::new(0xFB, 0x1D);
    pub const I31_GET_U: Op          = Op::new(0xFB, 0x1E);
    // Custom Descriptors proposal (extension, post-MVP).
    pub const STRUCT_NEW_DESC: Op        = Op::new(0xFB, 0x20); // struct.new_desc $typeidx
    pub const STRUCT_NEW_DEFAULT_DESC: Op = Op::new(0xFB, 0x21); // struct.new_default_desc $typeidx
    pub const REF_GET_DESC: Op           = Op::new(0xFB, 0x22); // ref.get_desc $typeidx
}

opcode_category! {
    // Struct ops (0x00..=0x05)
    [0x00] struct_new         => U16,    "struct.new";
    [0x01] struct_new_default => U16,    "struct.new_default";
    [0x02] struct_get         => U16,    "struct.get";
    [0x03] struct_get_s       => U16,    "struct.get_s";
    [0x04] struct_get_u       => U16,    "struct.get_u";
    [0x05] struct_set         => U16,    "struct.set";
    // Array ops (0x06..=0x13)
    //
    // Note on bytecode vs WASM binary: the spec assigns each of these
    // a `typeidx` immediate (and sometimes `dataidx`/`elemidx`/`N`).
    // In our internal bytecode the operand is carried ONLY for ops whose
    // VM behavior needs it (e.g. `array.new_fixed` needs the `N` count).
    // Other variants (new_default, fill, copy, get/set, len) use `None`
    // here because callers emit them with no operand — the emitter adds
    // the spec-required typeidx when lowering to the WASM binary.
    [0x06] array_new          => None,    "array.new";
    [0x07] array_new_default  => None,    "array.new_default";
    [0x08] array_new_fixed    => U16,     "array.new_fixed";
    [0x09] array_new_data     => U16,     "array.new_data";
    [0x0A] array_new_elem     => U16,     "array.new_elem";
    [0x0B] array_get          => None,    "array.get";
    [0x0C] array_get_s        => None,    "array.get_s";
    [0x0D] array_get_u        => None,    "array.get_u";
    [0x0E] array_set          => None,    "array.set";
    [0x0F] array_length       => None,    "array.len";
    [0x10] array_fill         => None,    "array.fill";
    [0x11] array_copy         => None,    "array.copy";
    [0x12] array_init_data    => U16,     "array.init_data";
    [0x13] array_init_elem    => U16,     "array.init_elem";
    // Reference tests / casts (0x14..=0x19)
    [0x14] ref_test           => U16,    "ref.test";
    [0x15] ref_test_null      => U16,    "ref.test_null";
    [0x16] ref_cast           => U16,    "ref.cast";
    [0x17] ref_cast_null      => U16,    "ref.cast_null";
    [0x18] br_on_cast         => U16_I16, "br_on_cast";
    [0x19] br_on_cast_fail    => U16_I16, "br_on_cast_fail";
    // Extern <-> any (0x1A..=0x1B)
    [0x1A] any_convert_extern => None,   "any.convert_extern";
    [0x1B] extern_convert_any => None,   "extern.convert_any";
    // i31 (0x1C..=0x1E)
    [0x1C] i31_new            => None,   "ref.i31";
    [0x1D] i31_get_s          => None,   "i31.get_s";
    [0x1E] i31_get_u          => None,   "i31.get_u";
    // Custom Descriptors (post-MVP extension)
    [0x20] struct_new_desc         => U16, "struct.new_desc";
    [0x21] struct_new_default_desc => U16, "struct.new_default_desc";
    [0x22] ref_get_desc            => U16, "ref.get_desc";
}
