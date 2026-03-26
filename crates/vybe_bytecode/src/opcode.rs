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
