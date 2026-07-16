//! .NET `System.Text.StringBuilder` adapter — bytecode-only.
//!
//! StringBuilder has no direct ECMA-262 mirror (JS uses immutable
//! strings + array-of-pieces + `join('')`). The wrapper materializes
//! a plain Object with a `__buffer` string field; `Append` / `Clear`
//! mutate the buffer via `Op::DYN_ADD` (string concat) + `STRUCT_SET`.
//! Pure WASM, zero host fns.
//!
//! Multi-arity constructor support:
//!   `new StringBuilder()` (argc=0)         — empty buffer
//!   `new StringBuilder("initial")` (argc=1) — initial-keyed buffer
//!
//! The dispatch layer threads `argc` so a single Common emit name
//! handles both shapes — the .NET wrapper exposes one `ConstructorDef`
//! and the emit decides at compile time which body to lay out.

use vybe_emitter::instructions::{core_wasm, host};
use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

const BUFFER_KEY: &str = "__buffer";
const CAPACITY_KEY: &str = "Capacity";

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

/// Stack-based scratch slot — bumps `local_count` and returns the new slot.
fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn emit_format_append_value(chunk: &mut Chunk, value_slot: u16, line: u32) {
    let formatted_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(chunk, "wasm:js-boolean", "test", 1, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(chunk, "wasm:js-boolean", "cast", 1, line);
    chunk.emit_if(line);
    push_const(chunk, Value::String(Arc::from("True")), line);
    chunk.emit_else(line);
    push_const(chunk, Value::String(Arc::from("False")), line);
    chunk.emit_end(line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);

    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, formatted_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, formatted_slot, line);
}

fn emit_update_capacity_from_buffer(chunk: &mut Chunk, sb_slot: u16, line: u32) {
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    let capacity_key = chunk.add_constant(Value::String(Arc::from(CAPACITY_KEY)));
    let len_slot = reserve_slot(chunk);
    let capacity_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, capacity_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, capacity_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, capacity_slot, line);
    vybe_emitter::ops::emit_dyn_gt(chunk, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, capacity_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line);
}

/// `new StringBuilder()` / `new StringBuilder(initial)` — multi-arity
/// constructor. Branches on `argc` to pick the right shape.
///
/// Stack on entry: `[]` (argc=0) or `[initial]` (argc=1)
/// Stack on exit: `[obj]` — Object with `__buffer` field set
pub fn emit_string_builder_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    let capacity_key = chunk.add_constant(Value::String(Arc::from(CAPACITY_KEY)));
    match argc {
        0 => {
            // [] → STRUCT_NEW → [obj]
            chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
            // [obj] → DUP → [obj, obj] → push "" → [obj, obj, ""]
            core_wasm::dup(chunk, line);
            push_const(chunk, Value::String(Arc::from("")), line);
            // STRUCT_SET pops [obj, val] and pushes [val]; we then drop val.
            chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
            chunk.emit_op(Op::DROP, line);
            core_wasm::dup(chunk, line);
            push_const(chunk, Value::I32(16), line);
            chunk.emit_op_u16(Op::STRUCT_SET, capacity_key, line);
            chunk.emit_op(Op::DROP, line);
        }
        _ => {
            for _ in 1..argc {
                chunk.emit_op(Op::DROP, line);
            }
            let init_slot = reserve_slot(chunk);
            let obj_slot = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, init_slot, line);

            chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
            chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

            chunk.emit_op_u16(Op::LOCAL_GET, init_slot, line);
            host::emit(chunk, "wasm:js-number", "test", 1, line);
            chunk.emit_if(line);

            // StringBuilder(int capacity): empty buffer, explicit capacity.
            chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
            push_const(chunk, Value::String(Arc::from("")), line);
            chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, init_slot, line);
            chunk.emit_op_u16(Op::STRUCT_SET, capacity_key, line);
            chunk.emit_op(Op::DROP, line);

            chunk.emit_else(line);

            // StringBuilder(string value): seed buffer, default capacity.
            chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, init_slot, line);
            chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
            push_const(chunk, Value::I32(16), line);
            chunk.emit_op_u16(Op::STRUCT_SET, capacity_key, line);
            chunk.emit_op(Op::DROP, line);

            chunk.emit_end(line);
            chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
        }
    }
}

/// `sb.Append(s)` — concat `s` onto `sb.__buffer`.
///
/// Stack on entry: `[sb, s]`
/// Stack on exit: `[sb]`
pub fn emit_sb_append(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    let s_slot = reserve_slot(chunk);
    let sb_slot = reserve_slot(chunk);

    // [sb, s] → stash s → [sb]
    chunk.emit_op_u16(Op::LOCAL_SET, s_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    // [sb, buffer] → push formatted s → [sb, buffer, s] → DYN_ADD → [sb, buffer+s]
    emit_format_append_value(chunk, s_slot, line);
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    // STRUCT_SET pops [sb, buffer+s], pushes [buffer+s]; drop it.
    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    emit_update_capacity_from_buffer(chunk, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
}

/// `sb.AppendLine(s)` — same as `Append` but with `"\n"` appended.
///
/// Stack on entry: `[sb]` or `[sb, s]` ; Stack on exit: `[sb]`
pub fn emit_sb_append_line(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    let s_slot = reserve_slot(chunk);
    let sb_slot = reserve_slot(chunk);

    if argc > 1 {
        chunk.emit_op_u16(Op::LOCAL_SET, s_slot, line);
    } else {
        push_const(chunk, Value::String(Arc::from("")), line);
        chunk.emit_op_u16(Op::LOCAL_SET, s_slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    emit_format_append_value(chunk, s_slot, line);
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    push_const(chunk, Value::String(Arc::from("\r\n")), line);
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    emit_update_capacity_from_buffer(chunk, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
}

/// `sb.AppendFormat(fmt, values...)` — append .NET composite-format output.
///
/// Stack on entry: `[sb, fmt, arg0, ...]` ; Stack on exit: `[sb]`
pub fn emit_sb_append_format(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let value_argc = argc.saturating_sub(1);
    let slots: Vec<u16> = (0..value_argc)
        .map(|_| reserve_slot(&mut chunks[current]))
        .collect();
    let sb_slot = reserve_slot(&mut chunks[current]);
    let formatted_slot = reserve_slot(&mut chunks[current]);

    {
        let chunk = &mut chunks[current];
        for slot in &slots {
            chunk.emit_op_u16(Op::LOCAL_SET, *slot, line);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, sb_slot, line);
        for slot in slots.iter().rev() {
            chunk.emit_op_u16(Op::LOCAL_GET, *slot, line);
        }
    }

    super::string_format_adapter::emit_string_format(chunks, current, value_argc, line);

    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    chunk.emit_op_u16(Op::LOCAL_SET, formatted_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, formatted_slot, line);
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    emit_update_capacity_from_buffer(chunk, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
}

/// `sb.ToString()` / `sb.ToString(start, count)` — return buffer text.
///
/// Stack on entry: `[sb]` or `[sb, start, count]` ; Stack on exit: `[buffer_string]`
pub fn emit_sb_to_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    if argc == 3 {
        let count_slot = reserve_slot(chunk);
        let start_slot = reserve_slot(chunk);
        let sb_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, start_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, sb_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
        chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
        chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
        chunk.emit_op(Op::F64_ADD, line);
        host::emit(chunk, "wasm:js-string", "substring", 3, line);
    } else {
        chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    }
}

/// `sb.Clear()` — reset `sb.__buffer` to `""`.
///
/// Stack on entry: `[sb]` ; Stack on exit: `[null]`
pub fn emit_sb_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    let sb_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
}

/// `sb.Length` (property-style read of buffer length).
///
/// Stack on entry: `[sb]` ; Stack on exit: `[length:i32]`
pub fn emit_sb_length(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
}

/// `sb.Capacity` — tracked as a numeric field, defaulting to 16.
pub fn emit_sb_capacity(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let capacity_key = chunk.add_constant(Value::String(Arc::from(CAPACITY_KEY)));
    chunk.emit_op_u16(Op::STRUCT_GET, capacity_key, line);
}

/// `sb.Insert(idx, text)` — splice `text` into `sb.__buffer` at `idx`.
///
/// Stack on entry: `[sb, idx, text]` ; Stack on exit: `[null]`
///
/// Implemented as `before = buf.substring(0, idx)`, `after =
/// buf.substring(idx, len)`, `buf' = before + text + after` via
/// `wasm:js-string.substring` / `.length` imports + `emit_dyn_add`.
pub fn emit_sb_insert(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    let text_slot = reserve_slot(chunk);
    let idx_slot = reserve_slot(chunk);
    let buf_slot = reserve_slot(chunk);
    let sb_slot = reserve_slot(chunk);

    // Stash text and idx args.
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, sb_slot, line);

    // Read sb.__buffer once and stash.
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, buf_slot, line);

    // before = buf.substring(0, idx)
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    push_const(chunk, Value::I32(0), line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);
    // Stack: [sb, before]

    // before + text
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    // Stack: [sb, before+text]

    // after = buf.substring(idx, buf.length)
    chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);
    // Stack: [sb, before+text, after]

    // (before + text) + after
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    // Stack: [sb, full]

    // sb.__buffer = full
    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    emit_update_capacity_from_buffer(chunk, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
}

/// `sb.Remove(start, count)` — remove `count` chars from `start`.
///
/// Stack on entry: `[sb, start, count]` ; Stack on exit: `[null]`
pub fn emit_sb_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    let count_slot = reserve_slot(chunk);
    let start_slot = reserve_slot(chunk);
    let buf_slot = reserve_slot(chunk);
    let sb_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, start_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, sb_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, buf_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    push_const(chunk, Value::I32(0), line);
    chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);

    chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);

    vybe_emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    emit_update_capacity_from_buffer(chunk, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
}

/// `sb[index]` — read a character from the backing buffer.
///
/// Stack on entry: `[sb, index]` ; Stack on exit: `[char_string]`
pub fn emit_sb_index_get(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    let index_slot = reserve_slot(chunk);
    let sb_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    host::emit(chunk, "ecma:string", "charAt", 2, line);
}

/// `sb[index] = value` — replace one character in the backing buffer.
///
/// Stack on entry: `[sb, index, value]` ; Stack on exit: `[sb]`
pub fn emit_sb_index_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    let value_slot = reserve_slot(chunk);
    let index_slot = reserve_slot(chunk);
    let sb_slot = reserve_slot(chunk);
    let buf_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, sb_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, buf_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    push_const(chunk, Value::I32(0), line);
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);

    emit_format_append_value(chunk, value_slot, line);
    push_const(chunk, Value::I32(0), line);
    host::emit(chunk, "ecma:string", "charAt", 2, line);
    vybe_emitter::ops::emit_dyn_add(chunk, line);

    chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);
    vybe_emitter::ops::emit_dyn_add(chunk, line);

    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
}

/// `sb.Replace(oldText, newText)` — replace all occurrences in buffer.
///
/// Stack on entry: `[sb, old, new]` ; Stack on exit: `[null]`
///
/// Lowers to `ecma:string.replaceAll(buf, old, new)` (ECMA-262
/// §22.1.3.18). The .NET `StringBuilder.Replace` always does
/// replace-all (no first-only variant) — matches `replaceAll` exactly.
pub fn emit_sb_replace(chunks: &mut [Chunk], current: usize, line: u32) {
    let replace_idx = chunks[0].add_import("ecma:string", "replaceAll");
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    let new_slot = reserve_slot(chunk);
    let old_slot = reserve_slot(chunk);
    let sb_slot = reserve_slot(chunk);

    // Stash old, new
    chunk.emit_op_u16(Op::LOCAL_SET, new_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, old_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, sb_slot, line);

    // [sb, sb, buf, old, new] — buffer + replace args
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, old_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, new_slot, line);

    // ecma:string.replaceAll(buf, old, new) → replaced
    chunk.emit_op_u16(Op::CALL_IMPORT, replace_idx, line);
    chunk.emit(3, line);
    // Stack: [sb, replaced]

    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
}
