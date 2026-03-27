/// Bytecode opcodes — aligned with WebAssembly naming conventions.
///
/// WASM uses `f64.add`, `local.get`, `struct.new` — we use `f64_add`, `local_get`, `struct_new`
/// (dots aren't valid in Rust identifiers).
///
/// Operand encoding:
/// - `u16`: two bytes big-endian after the opcode.
/// - `u8`: one byte after the opcode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
#[allow(non_camel_case_types)]
pub enum Op {
    // -- Stack --
    r#const,         // push from constant pool: [const, hi, lo]
    drop,            // discard TOS
    dup,             // duplicate TOS

    // -- Variables --
    local_get,       // [local_get, hi, lo]
    local_set,       // [local_set, hi, lo]
    global_get,      // [global_get, hi, lo] (name from constant pool)
    global_set,      // [global_set, hi, lo]
    upvalue_get,     // [upvalue_get, u8]
    upvalue_set,     // [upvalue_set, u8]

    // -- Struct (WASM GC) --
    struct_get,      // [struct_get, hi, lo] name from constant pool; stack [obj] → [val]
    struct_set,      // [struct_set, hi, lo]; stack [obj, val] → [val]

    // -- Array (WASM GC) --
    array_get,       // stack [obj, key] → [val]
    array_set,       // stack [obj, key, val] → [val]

    // -- f64 arithmetic --
    f64_add,
    f64_sub,
    f64_mul,
    f64_div,
    f64_mod,
    f64_neg,

    // -- i32 arithmetic --
    i32_add,
    i32_sub,
    i32_mul,
    i32_div_s,
    i32_div_u,
    i32_rem_s,
    i32_rem_u,

    // -- String --
    str_concat,      // String + String → String
    str_concat_n,    // concatenate N from stack: [str_concat_n, u8 count]

    // -- i32 bitwise --
    i32_and,
    i32_or,
    i32_xor,
    i32_not,
    i32_shl,
    i32_shr_s,      // signed shift right
    i32_shr_u,      // unsigned shift right
    i32_rotl,        // rotate left
    i32_rotr,        // rotate right
    i32_clz,         // count leading zeros
    i32_ctz,         // count trailing zeros
    i32_popcnt,      // population count

    // -- Comparison --
    eq,              // same-type equality → Bool
    ne,
    f64_lt,
    f64_gt,
    f64_le,
    f64_ge,
    str_lt,
    str_gt,

    // -- Logical --
    bool_not,

    // -- Control flow --
    br,              // unconditional: [br, hi, lo] (signed i16 offset)
    br_if_false,     // branch if Bool(false), pops: [br_if_false, hi, lo]
    br_if_true,      // branch if Bool(true), pops
    br_if_null,      // branch if Null, pops

    // -- Functions --
    call,            // [call, u8 arg_count]
    r#return,
    ref_func,        // create closure: [ref_func, u16 chunk_idx, u8 upvalue_count, descriptors...]
    /// Direct call through a typed function reference (WASM typed function references).
    /// Faster than call_indirect — no table lookup, direct dispatch.
    call_ref,        // stack: [func_ref, args...] → [result]; operand: u8 arg_count

    // -- Imports --
    call_import,     // [call_import, u16 import_idx, u8 arg_count]

    // -- Construction (WASM GC) --
    struct_new,      // [struct_new, u16 prop_count]; stack [key, val, ...] → [obj]
    array_new,       // [array_new, u16 elem_count]; stack [elem, ...] → [arr]

    // -- Immediate values --
    null,
    r#true,
    r#false,
    i32_const_0,
    i32_const_1,
    f64_const_0,

    // -- Type checks (WASM GC: ref.is_null, ref.test) --
    ref_is_null,
    ref_is_string,
    ref_is_number,
    ref_is_bool,
    ref_is_object,
    ref_is_func,

    // -- Type test (WASM GC: ref.test, ref.cast, br_on_cast) --
    /// Test if object is of type (or subtype). Uses TypeRegistry.
    /// Stack: [value] → [bool]. Operand: u16 constant index (type name).
    ref_test,        // [ref_test, u16 type_name_idx]
    /// Cast value to type — trap if wrong type. Stack: [value] → [value].
    ref_cast,        // [ref_cast, u16 type_name_idx]
    /// Branch if value IS the given type (keeps value on stack).
    /// Combines ref_test + br_if_true in one op — avoids double dispatch.
    br_on_cast,      // [br_on_cast, u16 type_name_idx, i16 offset]
    /// Branch if value is NOT the given type.
    br_on_cast_fail, // [br_on_cast_fail, u16 type_name_idx, i16 offset]

    // -- i31ref (WASM GC: tagged small integers) --
    /// Box a 31-bit integer as a tagged i31ref. Avoids heap allocation.
    /// Stack: [i32] → [i31ref value]
    i31_new,
    /// Unbox i31ref → i32 (sign-extended).
    i31_get_s,       // [i31ref] → [i32]
    /// Unbox i31ref → i32 (zero-extended).
    i31_get_u,       // [i31ref] → [i32]

    // -- Conversions --
    f64_from_i32,    // WASM: f64.convert_i32_s
    i32_from_f64,    // WASM: i32.trunc_f64_s

    // -- Dynamic ops (type dispatch inline, no host call overhead) --
    // These handle mixed types in the VM loop directly.
    // Used by dynamic languages (JS, VB) for hot paths.
    dyn_add,         // number+number→number, string+any→string, any+string→string
    dyn_eq,          // same-type eq, null==null, type coercion for number/string
    dyn_ne,          // negation of dyn_eq
    dyn_lt,          // numeric or string comparison
    dyn_gt,
    dyn_le,
    dyn_ge,
    dyn_neg,         // -value (numeric)
    dyn_not,         // !value (truthy check → bool)
    dyn_to_bool,     // value → Bool (truthy conversion)

    // -- Exception handling --
    try_start,       // [try_start, u16 catch, u16 finally]
    try_end,
    throw,
    /// throw_ref: throw a value directly (WASM EH: throw_ref).
    throw_ref,
    /// try_table: modern block-based exception handling (WASM EH Phase 4).
    /// Combines try + typed catch handlers in one structured block.
    /// [try_table, u8 handler_count, then for each: u8 tag, u16 offset]
    try_table,

    // -- Async (WASI async proposal) --
    /// Await a Promise value. If pending, suspends the current fiber.
    /// Stack: [promise_or_value] → [resolved_value]
    r#await,
    /// Schedule a timer callback. Stack: [callback, ms] → [null]
    set_timer,

    // -- Iteration (future) --
    iter_get,
    iter_next,
    spread,

    // -- Class (future) --
    class_new,       // [class_new, u16 name]
    method_def,      // [method_def, u16 name]
    inherit,

    // -- Tail call (WASM tail call proposal) --
    /// return_call: reuses current frame for tail-call optimization.
    /// Prevents stack overflow on deep recursion.
    return_call,     // [return_call, u8 arg_count]
    /// Tail call through function table index.
    return_call_indirect, // stack: [fn_table_idx, args...]; operand: u8 arg_count
    /// Tail call through a typed function reference.
    return_call_ref, // stack: [func_ref, args...]; operand: u8 arg_count

    // -- i64 arithmetic --
    i64_add,
    i64_sub,
    i64_mul,
    i64_div_s,
    i64_div_u,
    i64_rem_s,
    i64_rem_u,
    i64_and,
    i64_or,
    i64_xor,
    i64_shl,
    i64_shr_s,
    i64_shr_u,
    i64_rotl,
    i64_rotr,
    i64_clz,
    i64_ctz,
    i64_popcnt,

    // -- f64 math --
    f64_abs,
    f64_ceil,
    f64_floor,
    f64_trunc,
    f64_nearest,
    f64_sqrt,
    f64_min,
    f64_max,
    f64_copysign,

    // -- f32 (promoted to f64 in our VM) --
    f32_abs,
    f32_neg,
    f32_ceil,
    f32_floor,
    f32_trunc,
    f32_nearest,
    f32_sqrt,
    f32_min,
    f32_max,
    f32_copysign,

    // -- WASM select --
    select,          // [val1, val2, cond] → val1 if cond!=0 else val2

    // -- Linear memory (WASM MVP) --
    /// Memory operations on the VM's byte buffer.
    memory_size,     // → [i32 page_count]
    memory_grow,     // [pages] → [old_size or -1]
    i32_load,        // [addr] → [i32 value]
    i32_store,       // [addr, value] → []
    i64_load,        // [addr] → [i64 value]
    i64_store,       // [addr, value] → []
    f64_load,        // [addr] → [f64 value]
    f64_store,       // [addr, value] → []
    f32_load,        // [addr] → [f32 as f64]
    f32_store,       // [addr, f64 as f32] → []
    i32_load8_s,     // [addr] → [i32 sign-extended byte]
    i32_load8_u,     // [addr] → [i32 byte]
    i32_load16_s,    // [addr] → [i32 sign-extended i16]
    i32_load16_u,    // [addr] → [i32 zero-extended u16]
    i32_store16,     // [addr, value] → []
    i32_store8,      // [addr, byte] → []
    i64_load8_s,
    i64_load8_u,
    i64_load16_s,
    i64_load16_u,
    i64_load32_s,
    i64_load32_u,
    i64_store8,
    i64_store16,
    i64_store32,

    // -- Conversions --
    i32_wrap_i64,
    i64_extend_i32_s,
    i64_extend_i32_u,
    i64_trunc_f64_s,
    i64_trunc_f64_u,
    f64_promote_f32,
    f32_demote_f64,
    i32_reinterpret_f32,
    i64_reinterpret_f64,
    f32_reinterpret_i32,
    f64_reinterpret_i64,
    i32_extend8_s,
    i32_extend16_s,
    i64_extend8_s,
    i64_extend16_s,
    i64_extend32_s,

    // -- WASM eqz --
    i32_eqz,
    i64_eqz,

    // -- Multi-value (WASM multi-value) --
    /// Pack N values from stack into an array.
    pack,            // [pack, u8 count] → [array]
    /// Unpack array onto stack.
    unpack,          // [array] → [val, val, ...]

    // -- Block/loop structured control (WASM MVP) --
    /// Begin a block with a label. break jumps to end.
    block,           // [block, u16 end_offset]
    /// Begin a loop. continue jumps to start.
    r#loop,          // [loop, u16 body_size]
    /// End of block/loop.
    end,
    /// Branch to enclosing label: [br_label, u8 depth]
    /// depth 0 = innermost block, 1 = next outer, etc.
    br_label,
    /// Conditional branch to label.
    br_if_label,     // [br_if_label, u8 depth]
    /// Branch table (switch): [br_table, u8 count, u8 default, u8 label0, u8 label1, ...]
    br_table,

    // -- Function tables / call_indirect (WASM MVP) --
    /// Call a function from the function table by index.
    call_indirect,   // [call_indirect, u8 arg_count]; stack: [fn_table_idx, args...] → [result]

    // -- Component Model (WASM Component Model) --
    /// Lift a value to a component interface type.
    canon_lift,      // [canon_lift, u16 interface_idx]
    /// Lower a component interface type to core value.
    canon_lower,     // [canon_lower, u16 interface_idx]

    // -- JS String Builtins (wasm:js-string proposal) --
    // These match the WASM JS String Builtins proposal import names.
    // All operate on Value::String directly in the VM — no host call overhead.
    str_length,      // [str] → [i32]
    str_char_code_at,// [str, i32 index] → [i32 code]
    str_from_char_code, // [i32 code] → [str]  (Chr in VB)
    str_substring,   // [str, i32 start, i32 end] → [str]
    str_index_of,    // [str, str needle] → [i32 pos] (-1 if not found)
    str_last_index_of, // [str, str needle] → [i32 pos]
    str_equals,      // [str, str] → [bool]
    str_compare,     // [str, str] → [i32]  (-1, 0, 1)
    str_to_upper,    // [str] → [str]
    str_to_lower,    // [str] → [str]
    str_trim,        // [str] → [str]
    str_trim_start,  // [str] → [str]
    str_trim_end,    // [str] → [str]
    str_starts_with, // [str, str prefix] → [bool]
    str_ends_with,   // [str, str suffix] → [bool]
    str_contains,    // [str, str needle] → [bool]
    str_replace,     // [str, str old, str new] → [str]
    str_split,       // [str, str delim] → [array of str]
    str_repeat,      // [str, i32 count] → [str]
    str_pad_start,   // [str, i32 len, str fill] → [str]
    str_pad_end,     // [str, i32 len, str fill] → [str]
    str_slice,       // [str, i32 start, i32 end] → [str] (same as substring)
    str_char_at,     // [str, i32 index] → [str single-char]
    str_reverse,     // [str] → [str]

    // -- Array builtins (WASM GC array.* + extras) --
    array_length,    // [array] → [i32]                 (GC: array.len)
    array_push,      // [array, value] → [array]
    array_pop,       // [array] → [value]
    array_slice,     // [array, i32 start, i32 end] → [array]
    array_join,      // [array, str delim] → [str]
    array_reverse,   // [array] → [array]
    array_contains,  // [array, value] → [bool]
    array_index_of,  // [array, value] → [i32]
    // WASM GC spec ops
    array_new_default, // [i32 len] → [array of nulls]    (GC: array.new_default)
    array_fill,      // [array, value, i32 start, i32 len] → []  (GC: array.fill)
    array_copy,      // [dst, dst_off, src, src_off, len] → []   (GC: array.copy)
    array_concat,    // [array, array] → [array]
    array_shift,     // [array] → [value] (remove first)

    // -- Stack Switching (wasm stack-switching proposal) --
    // Enables async/await, generators, coroutines without CPS transform.
    /// Create a new continuation (fiber/coroutine).
    /// [ref_func] → [continuation]
    cont_new,
    /// Suspend the current continuation, yielding a value.
    /// [value] → suspends; resumes with [value]
    suspend,         // [suspend, u16 tag_idx]
    /// Resume a suspended continuation, passing a value.
    /// [continuation, value] → [result]
    resume,          // [resume, u16 tag_idx]
    /// Resume a continuation and immediately suspend the current one (symmetric switch).
    /// [continuation, value] → suspends; other resumes with [value]
    switch,          // [switch, u16 tag_idx]

    halt,
}

impl Op {
    pub fn from_byte(byte: u8) -> Option<Op> {
        if byte <= Op::halt as u8 {
            Some(unsafe { std::mem::transmute(byte) })
        } else {
            None
        }
    }
}
