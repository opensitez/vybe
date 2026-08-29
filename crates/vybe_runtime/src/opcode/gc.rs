//! GC proposal opcodes (prefix 0xFB).
//! Byte values match the WASM GC specification.

use super::Op;
use super::opcode_category;

impl Op {
    // Struct ops (0x00..=0x05)
    pub const STRUCT_NEW: Op = Op::new(0xFB, 0x00);
    pub const STRUCT_NEW_DEFAULT: Op = Op::new(0xFB, 0x01);
    pub const STRUCT_GET: Op = Op::new(0xFB, 0x02);
    pub const STRUCT_GET_S: Op = Op::new(0xFB, 0x03);
    pub const STRUCT_GET_U: Op = Op::new(0xFB, 0x04);
    pub const STRUCT_SET: Op = Op::new(0xFB, 0x05);
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
    pub const ARRAY_NEW: Op = Op::new(0xFB, 0x06);
    pub const ARRAY_NEW_DEFAULT: Op = Op::new(0xFB, 0x07);
    pub const ARRAY_NEW_FIXED: Op = Op::new(0xFB, 0x08);
    pub const ARRAY_NEW_DATA: Op = Op::new(0xFB, 0x09);
    pub const ARRAY_NEW_ELEM: Op = Op::new(0xFB, 0x0A);
    pub const ARRAY_GET: Op = Op::new(0xFB, 0x0B);
    pub const ARRAY_GET_S: Op = Op::new(0xFB, 0x0C);
    pub const ARRAY_GET_U: Op = Op::new(0xFB, 0x0D);
    pub const ARRAY_SET: Op = Op::new(0xFB, 0x0E);
    pub const ARRAY_LENGTH: Op = Op::new(0xFB, 0x0F);
    pub const ARRAY_FILL: Op = Op::new(0xFB, 0x10);
    pub const ARRAY_COPY: Op = Op::new(0xFB, 0x11);
    pub const ARRAY_INIT_DATA: Op = Op::new(0xFB, 0x12);
    pub const ARRAY_INIT_ELEM: Op = Op::new(0xFB, 0x13);
    // Reference tests / casts (0x14..=0x17)
    pub const REF_TEST: Op = Op::new(0xFB, 0x14);
    pub const REF_TEST_NULL: Op = Op::new(0xFB, 0x15);
    pub const REF_CAST: Op = Op::new(0xFB, 0x16);
    pub const REF_CAST_NULL: Op = Op::new(0xFB, 0x17);
    pub const BR_ON_CAST: Op = Op::new(0xFB, 0x18);
    pub const BR_ON_CAST_FAIL: Op = Op::new(0xFB, 0x19);
    // Extern <-> any conversion (0x1A, 0x1B)
    pub const ANY_CONVERT_EXTERN: Op = Op::new(0xFB, 0x1A);
    pub const EXTERN_CONVERT_ANY: Op = Op::new(0xFB, 0x1B);
    // i31 ops (0x1C..=0x1E)
    pub const I31_NEW: Op = Op::new(0xFB, 0x1C);
    pub const I31_GET_S: Op = Op::new(0xFB, 0x1D);
    pub const I31_GET_U: Op = Op::new(0xFB, 0x1E);
    // Custom Descriptors proposal (extension, post-MVP).
    pub const STRUCT_NEW_DESC: Op = Op::new(0xFB, 0x20); // struct.new_desc $typeidx
    pub const STRUCT_NEW_DEFAULT_DESC: Op = Op::new(0xFB, 0x21); // struct.new_default_desc $typeidx
    pub const REF_GET_DESC: Op = Op::new(0xFB, 0x22); // ref.get_desc $typeidx
    // Descriptor-comparing casts. The proposal encodes the two `ref.cast_desc_eq`
    // forms as separate opcodes (non-null / nullable result) rather than a flag,
    // mirroring `ref.cast` / `ref.cast_null`.
    pub const REF_CAST_DESC_EQ: Op = Op::new(0xFB, 0x23); // ref.cast_desc_eq (ref ht)
    pub const REF_CAST_DESC_EQ_NULL: Op = Op::new(0xFB, 0x24); // ref.cast_desc_eq (ref null ht)
    pub const BR_ON_CAST_DESC_EQ: Op = Op::new(0xFB, 0x25); // br_on_cast_desc_eq $l ht ht
    pub const BR_ON_CAST_DESC_EQ_FAIL: Op = Op::new(0xFB, 0x26); // br_on_cast_desc_eq_fail $l ht ht
    // Exact tests / casts — `(ref null? (exact $t))`. VM-INTERNAL: the spec
    // carries exactness inside the heaptype (`0x62 x:u32`), not in the opcode,
    // and our heaptype immediate is a single signed LEB with no room for a
    // third case. See the table entries below for the full reasoning; the
    // writer maps these back onto the spec's own 0x14..0x17.
    pub const REF_TEST_EXACT: Op = Op::new(0xFB, 0x27);
    pub const REF_TEST_EXACT_NULL: Op = Op::new(0xFB, 0x28);
    pub const REF_CAST_EXACT: Op = Op::new(0xFB, 0x29);
    pub const REF_CAST_EXACT_NULL: Op = Op::new(0xFB, 0x2A);
    /// VM-internal. `desc.set_proto $classname` — store the class's JS
    /// prototype into field 0 of that class's descriptor singleton.
    ///
    /// Not a spec opcode: the writer expands it to `local.set` / `global.get`
    /// / `local.get` / `struct.set`, so nothing VM-internal reaches the binary.
    /// It exists because the descriptor's field COUNT depends on the merged
    /// method list, which is only final after all emission — so a `struct.new`
    /// cannot be emitted on the class path, while a `struct.set` of field 0
    /// needs no count at all.
    pub const DESC_SET_PROTO: Op = Op::new(0xFB, 0x2B);

    // Stringref proposal (0x80..=0xB7). Byte values per proposals/stringref
    // Overview.md. Strings map onto `Value::String`. Ops carrying a `$mem`
    // immediate in the binary format use `operand_format: None` here — the VM
    // defaults to memory 0 (see the None-consistency note in the macro block).
    pub const STRING_NEW_UTF8: Op = Op::new(0xFB, 0x80);
    pub const STRING_MEASURE_UTF8: Op = Op::new(0xFB, 0x83);
    pub const STRING_MEASURE_WTF8: Op = Op::new(0xFB, 0x84);
    pub const STRING_MEASURE_WTF16: Op = Op::new(0xFB, 0x85);
    pub const STRING_ENCODE_UTF8: Op = Op::new(0xFB, 0x86);
    pub const STRING_ENCODE_WTF16: Op = Op::new(0xFB, 0x87);
    pub const STRING_CONCAT: Op = Op::new(0xFB, 0x88);
    pub const STRING_EQ: Op = Op::new(0xFB, 0x89);
    pub const STRING_NEW_WTF16: Op = Op::new(0xFB, 0x81);
    pub const STRING_IS_USV_SEQUENCE: Op = Op::new(0xFB, 0x8A);
    pub const STRING_NEW_LOSSY_UTF8: Op = Op::new(0xFB, 0x8B);
    pub const STRING_NEW_WTF8: Op = Op::new(0xFB, 0x8C);
    pub const STRING_ENCODE_LOSSY_UTF8: Op = Op::new(0xFB, 0x8D);
    pub const STRING_ENCODE_WTF8: Op = Op::new(0xFB, 0x8E);
    pub const STRING_AS_WTF8: Op = Op::new(0xFB, 0x90);
    pub const STRINGVIEW_WTF8_ADVANCE: Op = Op::new(0xFB, 0x91);
    pub const STRINGVIEW_WTF8_ENCODE_UTF8: Op = Op::new(0xFB, 0x92);
    pub const STRINGVIEW_WTF8_SLICE: Op = Op::new(0xFB, 0x93);
    pub const STRING_AS_WTF16: Op = Op::new(0xFB, 0x98);
    pub const STRINGVIEW_WTF16_LENGTH: Op = Op::new(0xFB, 0x99);
    pub const STRINGVIEW_WTF16_GET_CODEUNIT: Op = Op::new(0xFB, 0x9A);
    pub const STRINGVIEW_WTF16_ENCODE: Op = Op::new(0xFB, 0x9B);
    pub const STRINGVIEW_WTF16_SLICE: Op = Op::new(0xFB, 0x9C);
    pub const STRING_AS_ITER: Op = Op::new(0xFB, 0xA0);
    pub const STRINGVIEW_ITER_NEXT: Op = Op::new(0xFB, 0xA1);
    pub const STRINGVIEW_ITER_ADVANCE: Op = Op::new(0xFB, 0xA2);
    pub const STRINGVIEW_ITER_REWIND: Op = Op::new(0xFB, 0xA3);
    pub const STRINGVIEW_ITER_SLICE: Op = Op::new(0xFB, 0xA4);
    pub const STRING_NEW_UTF8_ARRAY: Op = Op::new(0xFB, 0xB0);
    pub const STRING_NEW_WTF16_ARRAY: Op = Op::new(0xFB, 0xB1);
    pub const STRING_ENCODE_UTF8_ARRAY: Op = Op::new(0xFB, 0xB2);
    pub const STRING_ENCODE_WTF16_ARRAY: Op = Op::new(0xFB, 0xB3);
    pub const STRING_NEW_LOSSY_UTF8_ARRAY: Op = Op::new(0xFB, 0xB4);
    pub const STRING_NEW_WTF8_ARRAY: Op = Op::new(0xFB, 0xB5);
    pub const STRING_ENCODE_LOSSY_UTF8_ARRAY: Op = Op::new(0xFB, 0xB6);
    pub const STRING_ENCODE_WTF8_ARRAY: Op = Op::new(0xFB, 0xB7);
}

opcode_category! {
    // Struct ops (0x00..=0x05)
    // `(typeidx, count)`, mirroring `array.new_fixed`. typeidx 0 = the
    // DYNAMIC object-literal form every language front end emits (count = k/v
    // pairs); non-zero is spec `struct.new $t`, whose field count comes from
    // $t itself, not from an immediate.
    [0x00] struct_new         => U16_U16, "struct.new";
    [0x01] struct_new_default => U16,    "struct.new_default";
    // `(typeidx, idx)` — same discrimination as `struct.new`. typeidx 0 means
    // `idx` is a constant-pool index for a field NAME (the dynamic path every
    // language front end emits); non-zero means `idx` is a spec `fieldidx`
    // into the instance's indexed storage.
    [0x02] struct_get         => U16_U16, "struct.get";
    [0x03] struct_get_s       => U16_U16, "struct.get_s";
    [0x04] struct_get_u       => U16_U16, "struct.get_u";
    [0x05] struct_set         => U16_U16, "struct.set";
    // Array ops (0x06..=0x13)
    //
    // Note on bytecode vs WASM binary: the spec assigns each of these
    // a `typeidx` immediate (and sometimes `dataidx`/`elemidx`/`N`).
    // In our internal bytecode the operand is carried ONLY for ops whose
    // VM behavior needs it (e.g. `array.new_fixed` needs the `N` count).
    // Other variants (new_default, fill, copy, get/set, len) use `None`
    // here because callers emit them with no operand — the emitter adds
    // the spec-required typeidx when lowering to the WASM binary.
    [0x06] array_new          => U16,     "array.new";         // typeidx (VM reads it)
    [0x07] array_new_default  => U16,     "array.new_default"; // typeidx (VM reads it)
    // BOTH immediates, `$t` and `N` — 4 bytes. Declaring this `U16` while
    // `Chunk::emit_array_new_fixed` writes four bytes and the dispatch reads
    // four leaves every operand_format-driven walk (the block-scanning
    // pre-pass, disassembly, codec skips) two bytes short, so the scan
    // resynchronises inside the NEXT instruction.
    [0x08] array_new_fixed    => U16_U16, "array.new_fixed";
    [0x09] array_new_data     => U16_U16, "array.new_data";
    [0x0A] array_new_elem     => U16_U16, "array.new_elem";
    [0x0B] array_get          => None,    "array.get";
    [0x0C] array_get_s        => U16,     "array.get_s"; // typeidx (VM + codec read it)
    [0x0D] array_get_u        => U16,     "array.get_u"; // typeidx (VM + codec read it)
    [0x0E] array_set          => None,    "array.set";
    [0x0F] array_length       => None,    "array.len";
    [0x10] array_fill         => None,    "array.fill";
    [0x11] array_copy         => None,    "array.copy";
    [0x12] array_init_data    => U16_U16, "array.init_data";
    [0x13] array_init_elem    => U16_U16, "array.init_elem";
    // Reference tests / casts (0x14..=0x19). The immediate is a HEAPTYPE, in
    // the spec's own encoding: one signed LEB where a negative value is an
    // abstract type (`any`, `struct`, `func`, …) and a non-negative value is
    // an index into the module's type section. Never a name — names live in
    // the type section, not in code.
    [0x14] ref_test           => SlI32,  "ref.test";
    [0x15] ref_test_null      => SlI32,  "ref.test_null";
    [0x16] ref_cast           => SlI32,  "ref.cast";
    [0x17] ref_cast_null      => SlI32,  "ref.cast_null";
    // br_on_cast / br_on_cast_fail use a structured label depth (u8)
    // matching core `br`'s encoding — the VM resolves depth via its
    // label_stack, and the WASM emitter writes it as a labelidx.
    [0x18] br_on_cast         => U16_U8, "br_on_cast";
    [0x19] br_on_cast_fail    => U16_U8, "br_on_cast_fail";
    // Extern <-> any (0x1A..=0x1B)
    [0x1A] any_convert_extern => None,   "any.convert_extern";
    [0x1B] extern_convert_any => None,   "extern.convert_any";
    // i31 (0x1C..=0x1E)
    [0x1C] i31_new            => None,   "ref.i31";
    [0x1D] i31_get_s          => None,   "i31.get_s";
    [0x1E] i31_get_u          => None,   "i31.get_u";
    // Custom Descriptors (post-MVP extension)
    [0x20] struct_new_desc         => U16_U16, "struct.new_desc";
    [0x21] struct_new_default_desc => U16, "struct.new_default_desc";
    [0x22] ref_get_desc            => U16, "ref.get_desc";
    // The descriptor-comparing casts carry the same operand shapes as their
    // plain counterparts: a u16 type-name constant index, plus a u8 label
    // depth for the branching forms.
    [0x23] ref_cast_desc_eq        => U16,    "ref.cast_desc_eq";
    [0x24] ref_cast_desc_eq_null   => U16,    "ref.cast_desc_eq_null";
    // ⛔ `U16_U16_U8`, not `U16_U8`. The spec form is
    //   `0xFB 37 (null_1?, null_2?):castflags l:labelidx ht_1 ht_2`
    // — BOTH heaptypes and BOTH null flags are this instruction's own
    // immediates. With one name the writer had nothing to say for `ht_1` and
    // emitted `HT_ANY` with castflags hardcoded `0x00`, so the emitted
    // instruction did not say what the source said.
    //
    // The two u16s are name-constant indices for `ht_2` (target) then `ht_1`
    // (source); NULLABILITY RIDES IN THE SPELLING (`ref null $t`), the same
    // convention `wasm_heap_type_ref_exact` already parses, which is why this
    // needs no new `OperandFormat` variant. The u8 is the label depth.
    [0x25] br_on_cast_desc_eq      => U16_U16_U8, "br_on_cast_desc_eq";
    [0x26] br_on_cast_desc_eq_fail => U16_U16_U8, "br_on_cast_desc_eq_fail";
    // Exact reference tests / casts — `(ref null? (exact $t))`.
    //
    // ⚠ THESE ARE VM-INTERNAL OPCODES WITH NO SPEC COUNTERPART, and that is
    // deliberate. In the binary format exactness lives INSIDE the heaptype
    // (`heaptype ::= … | 0x62 x:u32 => exact x`), so the spec spells an exact
    // test as an ordinary `ref.test` whose immediate happens to be exact. Our
    // heaptype immediate is a single signed LEB — `SlI32`, negative for
    // abstract and non-negative for a type index — which has no room for a
    // third case, and widening it would change the operand of four opcodes
    // every language already emits.
    //
    // Encoding the distinction in the OPCODE instead is not a new pattern
    // here: nullability is already an opcode choice (`ref.test` vs
    // `ref.test_null`) rather than an operand bit. The writer maps these back
    // to the spec's own `0xFB 0x14…0x17` with a `0x62`-prefixed heaptype, so
    // nothing VM-internal reaches the binary.
    //
    // The immediate is deliberately IDENTICAL to the inexact forms, so the
    // three-way operand contract (table / dispatch read / codec write) reads
    // the same bytes either way and only the comparison differs.
    [0x27] ref_test_exact          => SlI32,  "ref.test_exact";
    [0x28] ref_test_exact_null     => SlI32,  "ref.test_exact_null";
    [0x29] ref_cast_exact          => SlI32,  "ref.cast_exact";
    [0x2A] ref_cast_exact_null     => SlI32,  "ref.cast_exact_null";
    // VM-internal, expanded by the writer — see `DESC_SET_PROTO`. The
    // immediate is a constant-pool index naming the CLASS, because descriptor
    // type indices and descriptor global indices are both writer-side
    // concepts the compiler has no access to.
    [0x2B] desc_set_proto          => U16,    "desc.set_proto";

    // ── Stringref proposal (0x80..=0xB7) ──────────────────────────────────
    // operand_format is None for ALL of these: the ops that take a `$mem`
    // (new/encode) default to memory 0 in the VM, and the dispatch reads zero
    // immediate bytes. This keeps the three-way contract (table None / dispatch
    // reads nothing / codec writes nothing) internally consistent — the WASM
    // binary `$mem`/`$idx` uleb is a codec concern, not needed for text exec.
    [0x80] string_new_utf8            => None, "string.new_utf8";
    [0x81] string_new_wtf16           => None, "string.new_wtf16";
    [0x82] string_const               => None, "string.const";
    [0x83] string_measure_utf8        => None, "string.measure_utf8";
    [0x84] string_measure_wtf8        => None, "string.measure_wtf8";
    [0x85] string_measure_wtf16       => None, "string.measure_wtf16";
    [0x86] string_encode_utf8         => None, "string.encode_utf8";
    [0x87] string_encode_wtf16        => None, "string.encode_wtf16";
    [0x88] string_concat              => None, "string.concat";
    [0x89] string_eq                  => None, "string.eq";
    [0x8A] string_is_usv_sequence     => None, "string.is_usv_sequence";
    [0x8B] string_new_lossy_utf8      => None, "string.new_lossy_utf8";
    [0x8C] string_new_wtf8            => None, "string.new_wtf8";
    [0x8D] string_encode_lossy_utf8   => None, "string.encode_lossy_utf8";
    [0x8E] string_encode_wtf8         => None, "string.encode_wtf8";
    [0x90] string_as_wtf8             => None, "string.as_wtf8";
    [0x91] stringview_wtf8_advance    => None, "stringview_wtf8.advance";
    [0x92] stringview_wtf8_encode_utf8 => None, "stringview_wtf8.encode_utf8";
    [0x93] stringview_wtf8_slice      => None, "stringview_wtf8.slice";
    [0x98] string_as_wtf16            => None, "string.as_wtf16";
    [0x99] stringview_wtf16_length    => None, "stringview_wtf16.length";
    [0x9A] stringview_wtf16_get_codeunit => None, "stringview_wtf16.get_codeunit";
    [0x9B] stringview_wtf16_encode    => None, "stringview_wtf16.encode";
    [0x9C] stringview_wtf16_slice     => None, "stringview_wtf16.slice";
    [0xA0] string_as_iter             => None, "string.as_iter";
    [0xA1] stringview_iter_next       => None, "stringview_iter.next";
    [0xA2] stringview_iter_advance    => None, "stringview_iter.advance";
    [0xA3] stringview_iter_rewind     => None, "stringview_iter.rewind";
    [0xA4] stringview_iter_slice      => None, "stringview_iter.slice";
    [0xB0] string_new_utf8_array      => None, "string.new_utf8_array";
    [0xB1] string_new_wtf16_array     => None, "string.new_wtf16_array";
    [0xB2] string_encode_utf8_array   => None, "string.encode_utf8_array";
    [0xB3] string_encode_wtf16_array  => None, "string.encode_wtf16_array";
    [0xB4] string_new_lossy_utf8_array => None, "string.new_lossy_utf8_array";
    [0xB5] string_new_wtf8_array      => None, "string.new_wtf8_array";
    [0xB6] string_encode_lossy_utf8_array => None, "string.encode_lossy_utf8_array";
    [0xB7] string_encode_wtf8_array   => None, "string.encode_wtf8_array";
}
