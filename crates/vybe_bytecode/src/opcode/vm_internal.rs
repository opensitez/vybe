//! VM-internal opcodes (prefix 0xFF).
//! These are NOT part of the WASM specification.
//! They exist for VM performance and are lowered to standard WASM in .wasm binary output.

use super::Op;
use super::opcode_category;

impl Op {
    // Constants & stack
    pub const CONST: Op             = Op::new(0xFF, 0x00);
    pub const DUP: Op               = Op::new(0xFF, 0x01);
    pub const UPVALUE_GET: Op       = Op::new(0xFF, 0x02);
    pub const UPVALUE_SET: Op       = Op::new(0xFF, 0x03);
    pub const CALL_IMPORT: Op       = Op::new(0xFF, 0x04);
    // Branch variants
    pub const BR_IF_FALSE: Op       = Op::new(0xFF, 0x05);
    pub const BR_IF_NULL: Op        = Op::new(0xFF, 0x06);
    pub const BR_LABEL: Op          = Op::new(0xFF, 0x07);
    pub const BR_IF_LABEL: Op       = Op::new(0xFF, 0x08);
    // Immediate values
    pub const TRUE: Op              = Op::new(0xFF, 0x09);
    pub const FALSE: Op             = Op::new(0xFF, 0x0A);
    pub const I32_CONST_0: Op       = Op::new(0xFF, 0x0B);
    pub const I32_CONST_1: Op       = Op::new(0xFF, 0x0C);
    pub const F64_CONST_0: Op       = Op::new(0xFF, 0x0D);
    // Type checks
    pub const REF_IS_STRING: Op     = Op::new(0xFF, 0x0E);
    pub const REF_IS_NUMBER: Op     = Op::new(0xFF, 0x0F);
    pub const REF_IS_BOOL: Op       = Op::new(0xFF, 0x10);
    pub const REF_IS_OBJECT: Op     = Op::new(0xFF, 0x11);
    pub const REF_IS_FUNC: Op       = Op::new(0xFF, 0x12);
    pub const REF_TYPEOF: Op        = Op::new(0xFF, 0x13);
    pub const REF_IS_ARRAY: Op      = Op::new(0xFF, 0x14);
    // Dynamic dispatch
    pub const DYN_ADD: Op           = Op::new(0xFF, 0x15);
    pub const DYN_EQ: Op            = Op::new(0xFF, 0x16);
    pub const DYN_NE: Op            = Op::new(0xFF, 0x17);
    pub const DYN_LT: Op            = Op::new(0xFF, 0x18);
    pub const DYN_GT: Op            = Op::new(0xFF, 0x19);
    pub const DYN_LE: Op            = Op::new(0xFF, 0x1A);
    pub const DYN_GE: Op            = Op::new(0xFF, 0x1B);
    pub const DYN_NEG: Op           = Op::new(0xFF, 0x1C);
    pub const DYN_NOT: Op           = Op::new(0xFF, 0x1D);
    pub const DYN_TO_BOOL: Op       = Op::new(0xFF, 0x1E);
    // Exception handling
    pub const TRY_START: Op         = Op::new(0xFF, 0x1F);
    pub const TRY_END: Op           = Op::new(0xFF, 0x20);
    // Timers & spread
    pub const SET_TIMER: Op         = Op::new(0xFF, 0x21);
    pub const SPREAD: Op            = Op::new(0xFF, 0x22);
    // VM control
    pub const HALT: Op              = Op::new(0xFF, 0x23);
    // String builtins (wasm:js-string imports in .wasm output)
    pub const STR_CONCAT: Op        = Op::new(0xFF, 0x24);
    pub const STR_CONCAT_N: Op      = Op::new(0xFF, 0x25);
    pub const STR_LENGTH: Op        = Op::new(0xFF, 0x26);
    pub const STR_CHAR_CODE_AT: Op  = Op::new(0xFF, 0x27);
    pub const STR_FROM_CHAR_CODE: Op = Op::new(0xFF, 0x28);
    pub const STR_SUBSTRING: Op     = Op::new(0xFF, 0x29);
    pub const STR_INDEX_OF: Op      = Op::new(0xFF, 0x2A);
    pub const STR_LAST_INDEX_OF: Op = Op::new(0xFF, 0x2B);
    pub const STR_EQUALS: Op        = Op::new(0xFF, 0x2C);
    pub const STR_COMPARE: Op       = Op::new(0xFF, 0x2D);
    pub const STR_TO_UPPER: Op      = Op::new(0xFF, 0x2E);
    pub const STR_TO_LOWER: Op      = Op::new(0xFF, 0x2F);
    pub const STR_TRIM: Op          = Op::new(0xFF, 0x30);
    pub const STR_TRIM_START: Op    = Op::new(0xFF, 0x31);
    pub const STR_TRIM_END: Op      = Op::new(0xFF, 0x32);
    pub const STR_STARTS_WITH: Op   = Op::new(0xFF, 0x33);
    pub const STR_ENDS_WITH: Op     = Op::new(0xFF, 0x34);
    pub const STR_CONTAINS: Op      = Op::new(0xFF, 0x35);
    pub const STR_REPLACE: Op       = Op::new(0xFF, 0x36);
    pub const STR_SPLIT: Op         = Op::new(0xFF, 0x37);
    pub const STR_REPEAT: Op        = Op::new(0xFF, 0x38);
    pub const STR_PAD_START: Op     = Op::new(0xFF, 0x39);
    pub const STR_PAD_END: Op       = Op::new(0xFF, 0x3A);
    pub const STR_SLICE: Op         = Op::new(0xFF, 0x3B);
    pub const STR_CHAR_AT: Op       = Op::new(0xFF, 0x3C);
    pub const STR_REVERSE: Op       = Op::new(0xFF, 0x3D);
    pub const STR_FROM_CODE_POINT: Op = Op::new(0xFF, 0x3E);
    pub const STR_CODE_POINT_AT: Op = Op::new(0xFF, 0x3F);
    pub const STR_INTO_CHAR_CODES: Op = Op::new(0xFF, 0x40);
    pub const STR_FROM_CHAR_CODES: Op = Op::new(0xFF, 0x41);
    // Array builtins (host imports in .wasm output)
    // REMOVED (Phase E): the 9 `0xFF` ARRAY_* opcodes for dynamic array
    // mutation (push, pop, slice, join, reverse, contains, indexOf,
    // concat, shift — used to live at 0xFF 0x42–0x4A). They were Vybe-
    // specific and NOT in any WASM proposal. All callers now go through
    // `wasm:js-array.*` imports via `common::collections::*`. The opcode
    // IDs 0x42–0x4A are left vacant (not reused) so any legacy bytecode
    // that still carries them fails decode loudly rather than silently
    // aliasing to something else.
    // Stack switching (proposal not finalized)
    pub const CONT_NEW: Op          = Op::new(0xFF, 0x4B);
    pub const SUSPEND: Op           = Op::new(0xFF, 0x4C);
    pub const RESUME: Op            = Op::new(0xFF, 0x4D);
    pub const SWITCH: Op            = Op::new(0xFF, 0x4E);
    // JSPI
    pub const PROMISE_SUSPEND: Op   = Op::new(0xFF, 0x4F);
    // GC extensions
    pub const SET_TYPE_ID: Op       = Op::new(0xFF, 0x50);
    // Weak references
    pub const REF_MAKE_WEAK: Op     = Op::new(0xFF, 0x51);
    pub const REF_DEREF_WEAK: Op    = Op::new(0xFF, 0x52);
    pub const REF_IS_ALIVE: Op      = Op::new(0xFF, 0x53);
    pub const REF_REGISTER_FINALIZER: Op = Op::new(0xFF, 0x54);
    // Multi-memory
    pub const MEMORY_SELECT: Op     = Op::new(0xFF, 0x55);
    pub const MEMORY_COPY_CROSS: Op = Op::new(0xFF, 0x56);
    // Extended const
    pub const GLOBAL_INIT: Op       = Op::new(0xFF, 0x57);
    // Typed continuations
    pub const CONT_NEW_TYPED: Op    = Op::new(0xFF, 0x58);
    pub const SUSPEND_TYPED: Op     = Op::new(0xFF, 0x59);
    pub const RESUME_TYPED: Op      = Op::new(0xFF, 0x5A);
    // String references
    pub const STRING_AS_REF: Op     = Op::new(0xFF, 0x5B);
    pub const STRING_FROM_REF: Op   = Op::new(0xFF, 0x5C);
    pub const STRING_REF_EQ: Op     = Op::new(0xFF, 0x5D);
    // Shared GC objects
    pub const SHARED_NEW: Op        = Op::new(0xFF, 0x5E);
    pub const SHARED_STRUCT_GET: Op = Op::new(0xFF, 0x5F);
    pub const SHARED_STRUCT_SET: Op = Op::new(0xFF, 0x60);
    pub const SHARED_ARRAY_GET: Op  = Op::new(0xFF, 0x61);
    pub const SHARED_ARRAY_SET: Op  = Op::new(0xFF, 0x62);
    pub const SHARED_STRUCT_CAS: Op = Op::new(0xFF, 0x63);
    // Component model
    pub const CANON_LIFT: Op        = Op::new(0xFF, 0x64);
    pub const CANON_LOWER: Op       = Op::new(0xFF, 0x65);
    pub const TYPE_IMPORT: Op       = Op::new(0xFF, 0x66);
    pub const TYPE_EXPORT: Op       = Op::new(0xFF, 0x67);
    // Memory64
    pub const I64_MEMORY_SIZE: Op   = Op::new(0xFF, 0x68);
    pub const I64_MEMORY_GROW: Op   = Op::new(0xFF, 0x69);
    pub const I32_LOAD_64: Op       = Op::new(0xFF, 0x6A);
    pub const I64_LOAD_64: Op       = Op::new(0xFF, 0x6B);
    pub const F64_LOAD_64: Op       = Op::new(0xFF, 0x6C);
    pub const I32_STORE_64: Op      = Op::new(0xFF, 0x6D);
    pub const I64_STORE_64: Op      = Op::new(0xFF, 0x6E);
    pub const F64_STORE_64: Op      = Op::new(0xFF, 0x6F);
    // JS primitive creation / testing (js-primitive-builtins proposal)
    pub const UNDEFINED: Op         = Op::new(0xFF, 0x70);
    pub const SYMBOL: Op            = Op::new(0xFF, 0x71); // u16 const-idx (description)
    pub const BIGINT: Op            = Op::new(0xFF, 0x72); // u16 const-idx (Value::I64)
    pub const REF_IS_UNDEFINED: Op  = Op::new(0xFF, 0x73);
    pub const REF_IS_SYMBOL: Op     = Op::new(0xFF, 0x74);
    pub const REF_IS_BIGINT: Op     = Op::new(0xFF, 0x75);
    // Narrow numeric type tests + unsigned coercions + string formatting
    // (js-primitive-builtins wiring). These give compilers direct access
    // to the declared `wasm:js-*` imports for efficient interop.
    pub const REF_IS_I32: Op        = Op::new(0xFF, 0x76); // wasm:js-number.testI32
    pub const REF_IS_U32: Op        = Op::new(0xFF, 0x77); // wasm:js-number.testU32
    pub const NUM_BOX_U32: Op       = Op::new(0xFF, 0x78); // i32 → externref via fromU32
    pub const NUM_UNBOX_U32: Op     = Op::new(0xFF, 0x79); // externref → i32 via toU32
    pub const BOOL_CAST: Op         = Op::new(0xFF, 0x7A); // externref → i32 via js-boolean.cast
    pub const STR_CAST: Op          = Op::new(0xFF, 0x7B); // externref → externref (validates)
    pub const STR_FROM_I32: Op      = Op::new(0xFF, 0x7C);
    pub const STR_FROM_U32: Op      = Op::new(0xFF, 0x7D);
    pub const STR_FROM_I64: Op      = Op::new(0xFF, 0x7E);
    pub const STR_FROM_U64: Op      = Op::new(0xFF, 0x7F);
    pub const STR_FROM_F64: Op      = Op::new(0xFF, 0x80);
    pub const SYMBOL_EQ: Op         = Op::new(0xFF, 0x81); // (sym, sym) → bool via js-symbol.equals
}

opcode_category! {
    // Constants & stack
    [0x00] r#const => U16, "const";
    [0x01] dup => None, "dup";
    [0x02] upvalue_get => U8, "upvalue.get";
    [0x03] upvalue_set => U8, "upvalue.set";
    [0x04] call_import => U16_U8, "call_import";
    // Branch variants
    [0x05] br_if_false => I16, "br_if_false";
    [0x06] br_if_null => I16, "br_if_null";
    [0x07] br_label => U8, "br_label";
    [0x08] br_if_label => U8, "br_if_label";
    // Immediate values
    [0x09] r#true => None, "true";
    [0x0A] r#false => None, "false";
    [0x0B] i32_const_0 => None, "i32.const.0";
    [0x0C] i32_const_1 => None, "i32.const.1";
    [0x0D] f64_const_0 => None, "f64.const.0";
    // Type checks
    [0x0E] ref_is_string => None, "ref.is_string";
    [0x0F] ref_is_number => None, "ref.is_number";
    [0x10] ref_is_bool => None, "ref.is_bool";
    [0x11] ref_is_object => None, "ref.is_object";
    [0x12] ref_is_func => None, "ref.is_func";
    [0x13] ref_typeof => None, "ref.typeof";
    [0x14] ref_is_array => None, "ref.is_array";
    // Dynamic dispatch
    [0x15] dyn_add => None, "dyn.add";
    [0x16] dyn_eq => None, "dyn.eq";
    [0x17] dyn_ne => None, "dyn.ne";
    [0x18] dyn_lt => None, "dyn.lt";
    [0x19] dyn_gt => None, "dyn.gt";
    [0x1A] dyn_le => None, "dyn.le";
    [0x1B] dyn_ge => None, "dyn.ge";
    [0x1C] dyn_neg => None, "dyn.neg";
    [0x1D] dyn_not => None, "dyn.not";
    [0x1E] dyn_to_bool => None, "dyn.to_bool";
    // Exception handling
    [0x1F] try_start => U16_U16, "try_start";
    [0x20] try_end => None, "try_end";
    // Timers & spread
    [0x21] set_timer => None, "set_timer";
    [0x22] spread => None, "spread";
    // VM control
    [0x23] halt => None, "halt";
    // String builtins
    [0x24] str_concat => None, "string.concat";
    [0x25] str_concat_n => U8, "string.concat_n";
    [0x26] str_length => None, "string.length";
    [0x27] str_char_code_at => None, "string.charCodeAt";
    [0x28] str_from_char_code => None, "string.fromCharCode";
    [0x29] str_substring => None, "string.substring";
    [0x2A] str_index_of => None, "string.indexOf";
    [0x2B] str_last_index_of => None, "string.lastIndexOf";
    [0x2C] str_equals => None, "string.equals";
    [0x2D] str_compare => None, "string.compare";
    [0x2E] str_to_upper => None, "string.toUpperCase";
    [0x2F] str_to_lower => None, "string.toLowerCase";
    [0x30] str_trim => None, "string.trim";
    [0x31] str_trim_start => None, "string.trimStart";
    [0x32] str_trim_end => None, "string.trimEnd";
    [0x33] str_starts_with => None, "string.startsWith";
    [0x34] str_ends_with => None, "string.endsWith";
    [0x35] str_contains => None, "string.includes";
    [0x36] str_replace => None, "string.replace";
    [0x37] str_split => None, "string.split";
    [0x38] str_repeat => None, "string.repeat";
    [0x39] str_pad_start => None, "string.padStart";
    [0x3A] str_pad_end => None, "string.padEnd";
    [0x3B] str_slice => None, "string.slice";
    [0x3C] str_char_at => None, "string.charAt";
    [0x3D] str_reverse => None, "string.reverse";
    [0x3E] str_from_code_point => None, "string.fromCodePoint";
    [0x3F] str_code_point_at => None, "string.codePointAt";
    [0x40] str_into_char_codes => None, "string.intoCharCodes";
    [0x41] str_from_char_codes => None, "string.fromCharCodes";
    // Array builtins
    [0x42] array_push => None, "array.push";
    [0x43] array_pop => None, "array.pop";
    [0x44] array_slice => None, "array.slice";
    [0x45] array_join => None, "array.join";
    [0x46] array_reverse => None, "array.reverse";
    [0x47] array_contains => None, "array.contains";
    [0x48] array_index_of => None, "array.indexOf";
    [0x49] array_concat => None, "array.concat";
    [0x4A] array_shift => None, "array.shift";
    // Stack switching
    [0x4B] cont_new => None, "cont.new";
    [0x4C] suspend => U16, "suspend";
    [0x4D] resume => U16, "resume";
    [0x4E] switch => U16, "switch";
    // JSPI
    [0x4F] promise_suspend => None, "promise.suspend";
    // GC extensions
    [0x50] set_type_id => None, "set_type_id";
    // Weak references
    [0x51] ref_make_weak => None, "ref.make_weak";
    [0x52] ref_deref_weak => None, "ref.deref_weak";
    [0x53] ref_is_alive => None, "ref.is_alive";
    [0x54] ref_register_finalizer => None, "ref.register_finalizer";
    // Multi-memory
    [0x55] memory_select => U16, "memory.select";
    [0x56] memory_copy_cross => None, "memory.copy_cross";
    // Extended const
    [0x57] global_init => U16, "global.init";
    // Typed continuations
    [0x58] cont_new_typed => U16, "cont.new_typed";
    [0x59] suspend_typed => U16, "suspend.typed";
    [0x5A] resume_typed => U16, "resume.typed";
    // String references
    [0x5B] string_as_ref => None, "string.as_ref";
    [0x5C] string_from_ref => None, "string.from_ref";
    [0x5D] string_ref_eq => None, "string.ref_eq";
    // Shared GC objects
    [0x5E] shared_new => None, "shared.new";
    [0x5F] shared_struct_get => U16, "shared.struct_get";
    [0x60] shared_struct_set => U16, "shared.struct_set";
    [0x61] shared_array_get => None, "shared.array_get";
    [0x62] shared_array_set => None, "shared.array_set";
    [0x63] shared_struct_cas => U16, "shared.struct_cas";
    // Component model
    [0x64] canon_lift => U16, "canon.lift";
    [0x65] canon_lower => U16, "canon.lower";
    [0x66] type_import => U16, "type.import";
    [0x67] type_export => U16, "type.export";
    // Memory64
    [0x68] i64_memory_size => None, "i64.memory_size";
    [0x69] i64_memory_grow => None, "i64.memory_grow";
    [0x6A] i32_load_64 => None, "i32.load_64";
    [0x6B] i64_load_64 => None, "i64.load_64";
    [0x6C] f64_load_64 => None, "f64.load_64";
    [0x6D] i32_store_64 => None, "i32.store_64";
    [0x6E] i64_store_64 => None, "i64.store_64";
    [0x6F] f64_store_64 => None, "f64.store_64";
    // JS primitive creation / testing
    [0x70] undefined => None, "undefined";
    [0x71] symbol => U16, "symbol";
    [0x72] bigint => U16, "bigint";
    [0x73] ref_is_undefined => None, "ref.is_undefined";
    [0x74] ref_is_symbol => None, "ref.is_symbol";
    [0x75] ref_is_bigint => None, "ref.is_bigint";
    // Narrow numeric tests + unsigned coercions + string formatting
    [0x76] ref_is_i32 => None, "ref.is_i32";
    [0x77] ref_is_u32 => None, "ref.is_u32";
    [0x78] num_box_u32 => None, "num.box_u32";
    [0x79] num_unbox_u32 => None, "num.unbox_u32";
    [0x7A] bool_cast => None, "bool.cast";
    [0x7B] str_cast => None, "string.cast";
    [0x7C] str_from_i32 => None, "string.from_i32";
    [0x7D] str_from_u32 => None, "string.from_u32";
    [0x7E] str_from_i64 => None, "string.from_i64";
    [0x7F] str_from_u64 => None, "string.from_u64";
    [0x80] str_from_f64 => None, "string.from_f64";
    [0x81] symbol_eq => None, "symbol.eq";
}
