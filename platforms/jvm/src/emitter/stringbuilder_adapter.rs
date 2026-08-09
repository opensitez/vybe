//! JVM `java.lang.StringBuilder` adapter.

use vybe_compiler::primitives::{instructions::core_wasm, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const BUFFER_KEY: &str = "__buffer";

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

pub fn emit_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let initial = chunks[current].alloc_scratch(1);
    if argc > 0 {
        set(&mut chunks[current], initial, line);
        for _ in 1..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    } else {
        chunks[current].emit_string_const("", line);
        set(&mut chunks[current], initial, line);
    }

    chunks[current].emit_struct_new(0, 0, line);
    core_wasm::dup(&mut chunks[current], line);
    get(&mut chunks[current], initial, line);
    let key = chunks[current].add_constant(vybe_runtime::Value::String(BUFFER_KEY.into()));
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
}

pub fn emit_append(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let sb = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    if argc > 2 {
        for _ in 2..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    }
    set(&mut chunks[current], sb, line);
    append_slot(chunks, current, sb, value, false, line);
}

pub fn emit_append_line(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let sb = chunks[current].alloc_scratch(1);
    if argc > 1 {
        set(&mut chunks[current], value, line);
        if argc > 2 {
            for _ in 2..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
        }
    } else {
        chunks[current].emit_string_const("", line);
        set(&mut chunks[current], value, line);
    }
    set(&mut chunks[current], sb, line);
    append_slot(chunks, current, sb, value, true, line);
}

fn append_slot(
    chunks: &mut [Chunk],
    current: usize,
    sb: u16,
    value: u16,
    newline: bool,
    line: u32,
) {
    let key = chunks[current].add_constant(vybe_runtime::Value::String(BUFFER_KEY.into()));
    // Appending a CHAR ARRAY appends its characters, not "x,y,z".
    get(&mut chunks[current], value, line);
    let is_arr = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_arr, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_string_const("", line);
    vybe_compiler::primitives::collections::emit_join(chunks, current, line);
    set(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
    // Appending another BUILDER appends its text, not "[object StringBuilder]".
    get(&mut chunks[current], value, line);
    let tof = chunks[current].add_import("ecma:value", "typeof");
    chunks[current].emit_call(tof, 1, line);
    chunks[current].emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    let other = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], other, line);
    get(&mut chunks[current], other, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], other, line);
    set(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], sb, line);
    get(&mut chunks[current], sb, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    get(&mut chunks[current], value, line);
    strings::emit_str_concat_coercing(&mut chunks[current], line);
    if newline {
        chunks[current].emit_string_const("\n", line);
        strings::emit_str_concat(&mut chunks[current], line);
    }
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    get(&mut chunks[current], sb, line);
}

pub fn emit_to_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 1 {
        for _ in 1..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    }
    let key = chunks[current].add_constant(vybe_runtime::Value::String(BUFFER_KEY.into()));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
}

// ── The rest of the `java.lang.StringBuilder` surface ───────────────────────
//
// All of it manipulates the `__buffer` string through the shared string
// primitives and `ecma:string` host fns — no new host surface.

fn buffer_get(chunks: &mut [Chunk], current: usize, sb: u16, line: u32) {
    let key = chunks[current].add_constant(vybe_runtime::Value::String(BUFFER_KEY.into()));
    get(&mut chunks[current], sb, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
}

/// `[new_buffer]` → stores it, leaves the BUILDER (for chaining).
fn buffer_set(chunks: &mut [Chunk], current: usize, sb: u16, line: u32) {
    let tmp = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], tmp, line);
    let key = chunks[current].add_constant(vybe_runtime::Value::String(BUFFER_KEY.into()));
    get(&mut chunks[current], sb, line);
    get(&mut chunks[current], tmp, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    get(&mut chunks[current], sb, line);
}

fn host(chunks: &mut [Chunk], current: usize, module: &str, func: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import(module, func);
    chunks[current].emit_call(idx, argc, line);
}

/// `buffer[0..i]` — slice helper. Stack: pushes the slice from locals.
fn buffer_slice(
    chunks: &mut [Chunk],
    current: usize,
    sb: u16,
    from: u16,
    to: Option<u16>,
    line: u32,
) {
    buffer_get(chunks, current, sb, line);
    get(&mut chunks[current], from, line);
    if let Some(to) = to {
        get(&mut chunks[current], to, line);
        host(chunks, current, "ecma:string", "slice", 3, line);
    } else {
        host(chunks, current, "ecma:string", "slice", 2, line);
    }
}

/// `sb.length()` / `sb.capacity()` — the buffer's length.
pub fn emit_length(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let key = chunks[current].add_constant(vybe_runtime::Value::String(BUFFER_KEY.into()));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    strings::emit_length(&mut chunks[current], line);
}

/// `sb.insert(i, v)`.
pub fn emit_insert(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let sb = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], v, line);
    set(&mut chunks[current], i, line);
    if argc > 3 {
        for _ in 3..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    }
    set(&mut chunks[current], sb, line);
    buffer_slice(chunks, current, sb, i, None, line);
    let tail = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], tail, line);
    let zero = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], zero, line);
    buffer_slice(chunks, current, sb, zero, Some(i), line);
    get(&mut chunks[current], v, line);
    strings::emit_str_concat_coercing(&mut chunks[current], line);
    get(&mut chunks[current], tail, line);
    strings::emit_str_concat(&mut chunks[current], line);
    buffer_set(chunks, current, sb, line);
}

/// `sb.delete(a, b)` — end-exclusive. `sb.deleteAt(i)` / `deleteCharAt(i)`
/// pass `b = a + 1`.
pub fn emit_delete(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let b = chunks[current].alloc_scratch(1);
    let a = chunks[current].alloc_scratch(1);
    let sb = chunks[current].alloc_scratch(1);
    if argc >= 3 {
        set(&mut chunks[current], b, line);
        set(&mut chunks[current], a, line);
    } else {
        // deleteAt(i): [i] → a = i, b = i + 1
        set(&mut chunks[current], a, line);
        get(&mut chunks[current], a, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        set(&mut chunks[current], b, line);
    }
    set(&mut chunks[current], sb, line);
    let zero = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], zero, line);
    buffer_slice(chunks, current, sb, zero, Some(a), line);
    buffer_slice(chunks, current, sb, b, None, line);
    strings::emit_str_concat(&mut chunks[current], line);
    buffer_set(chunks, current, sb, line);
}

/// `sb.reverse()`.
pub fn emit_reverse(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let sb = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sb, line);
    buffer_get(chunks, current, sb, line);
    chunks[current].emit_string_const("", line);
    host(chunks, current, "ecma:string", "split", 2, line);
    vybe_compiler::primitives::collections::emit_reverse(chunks, current, line);
    chunks[current].emit_string_const("", line);
    vybe_compiler::primitives::collections::emit_join(chunks, current, line);
    buffer_set(chunks, current, sb, line);
}

/// `sb.setLength(n)` — truncates (the grow-with-NULs case has no consumer).
pub fn emit_set_length(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let n = chunks[current].alloc_scratch(1);
    let sb = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], n, line);
    set(&mut chunks[current], sb, line);
    let zero = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], zero, line);
    buffer_slice(chunks, current, sb, zero, Some(n), line);
    buffer_set(chunks, current, sb, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `sb.clear()`.
pub fn emit_clear(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let sb = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sb, line);
    chunks[current].emit_string_const("", line);
    buffer_set(chunks, current, sb, line);
}

/// `sb.setCharAt(i, c)` / `sb[i] = c`.
pub fn emit_set_char_at(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let c = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let sb = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], c, line);
    set(&mut chunks[current], i, line);
    set(&mut chunks[current], sb, line);
    let next = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], next, line);
    let zero = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], zero, line);
    buffer_slice(chunks, current, sb, zero, Some(i), line);
    get(&mut chunks[current], c, line);
    strings::emit_str_concat_coercing(&mut chunks[current], line);
    buffer_slice(chunks, current, sb, next, None, line);
    strings::emit_str_concat(&mut chunks[current], line);
    buffer_set(chunks, current, sb, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `sb.get(i)` / `sb.charAt(i)`.
pub fn emit_char_at(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let i = chunks[current].alloc_scratch(1);
    let sb = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], i, line);
    set(&mut chunks[current], sb, line);
    let next = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], next, line);
    buffer_slice(chunks, current, sb, i, Some(next), line);
}

/// `sb.appendCodePoint(cp)`.
pub fn emit_append_code_point(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let cp = chunks[current].alloc_scratch(1);
    let sb = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], cp, line);
    set(&mut chunks[current], sb, line);
    buffer_get(chunks, current, sb, line);
    get(&mut chunks[current], cp, line);
    host(chunks, current, "ecma:string", "fromCodePoint", 1, line);
    strings::emit_str_concat(&mut chunks[current], line);
    buffer_set(chunks, current, sb, line);
}

/// `sb.isEmpty()` / `sb.isNotEmpty()`.
pub fn emit_is_empty(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_length(chunks, current, 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_is_not_empty(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_is_empty(chunks, current, argc, line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `sb.substring(from[, to])` / `sb.subSequence(from, to)`.
pub fn emit_substring(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 3 {
        let to = chunks[current].alloc_scratch(1);
        let from = chunks[current].alloc_scratch(1);
        let sb = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], to, line);
        set(&mut chunks[current], from, line);
        set(&mut chunks[current], sb, line);
        buffer_slice(chunks, current, sb, from, Some(to), line);
    } else {
        let from = chunks[current].alloc_scratch(1);
        let sb = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], from, line);
        set(&mut chunks[current], sb, line);
        buffer_slice(chunks, current, sb, from, None, line);
    }
}

/// `sb.replace(from, to, replacement)` — JDK end-exclusive splice.
pub fn emit_replace(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let rep = chunks[current].alloc_scratch(1);
    let to = chunks[current].alloc_scratch(1);
    let from = chunks[current].alloc_scratch(1);
    let sb = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rep, line);
    set(&mut chunks[current], to, line);
    set(&mut chunks[current], from, line);
    set(&mut chunks[current], sb, line);
    let zero = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], zero, line);
    buffer_slice(chunks, current, sb, zero, Some(from), line);
    get(&mut chunks[current], rep, line);
    strings::emit_str_concat_coercing(&mut chunks[current], line);
    buffer_slice(chunks, current, sb, to, None, line);
    strings::emit_str_concat(&mut chunks[current], line);
    buffer_set(chunks, current, sb, line);
}
