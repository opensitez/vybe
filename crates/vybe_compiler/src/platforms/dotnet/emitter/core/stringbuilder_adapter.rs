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

use crate::emitter::instructions::{core_wasm, host};
use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

const BUFFER_KEY: &str = "__buffer";

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
    let slot = chunk.local_count;
    chunk.local_count = slot + 1;
    slot
}

/// `new StringBuilder()` / `new StringBuilder(initial)` — multi-arity
/// constructor. Branches on `argc` to pick the right shape.
///
/// Stack on entry: `[]` (argc=0) or `[initial]` (argc=1)
/// Stack on exit: `[obj]` — Object with `__buffer` field set
pub fn emit_string_builder_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
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
        }
        _ => {
            // Defensive: drop any extra args beyond the first (per
            // .NET overload, only the 1-arg form is supported here).
            for _ in 1..argc {
                chunk.emit_op(Op::DROP, line);
            }
            // [initial] → stash to scratch slot
            let init_slot = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, init_slot, line);
            chunk.emit_op(Op::DROP, line);
            // [] → STRUCT_NEW → [obj] → DUP → [obj, obj] → load initial → [obj, obj, initial]
            chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
            core_wasm::dup(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_GET, init_slot, line);
            chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
            chunk.emit_op(Op::DROP, line);
        }
    }
}

/// `sb.Append(s)` — concat `s` onto `sb.__buffer`.
///
/// Stack on entry: `[sb, s]`
/// Stack on exit: `[null]` (void return; .NET returns `sb` for chaining
/// but tests don't rely on that today)
pub fn emit_sb_append(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    let s_slot = reserve_slot(chunk);

    // [sb, s] → stash s → [sb]
    chunk.emit_op_u16(Op::LOCAL_SET, s_slot, line);
    chunk.emit_op(Op::DROP, line);
    // [sb] → DUP → [sb, sb] → STRUCT_GET __buffer → [sb, buffer]
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    // [sb, buffer] → push s → [sb, buffer, s] → DYN_ADD → [sb, buffer+s]
    chunk.emit_op_u16(Op::LOCAL_GET, s_slot, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    // STRUCT_SET pops [sb, buffer+s], pushes [buffer+s]; drop it.
    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::NULL, line);
}

/// `sb.AppendLine(s)` — same as `Append` but with `"\n"` appended.
///
/// Stack on entry: `[sb, s]` ; Stack on exit: `[null]`
pub fn emit_sb_append_line(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    let s_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, s_slot, line);
    chunk.emit_op(Op::DROP, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s_slot, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    push_const(chunk, Value::String(Arc::from("\n")), line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::NULL, line);
}

/// `sb.ToString()` — return `sb.__buffer`.
///
/// Stack on entry: `[sb]` ; Stack on exit: `[buffer_string]`
pub fn emit_sb_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
}

/// `sb.Clear()` — reset `sb.__buffer` to `""`.
///
/// Stack on entry: `[sb]` ; Stack on exit: `[null]`
pub fn emit_sb_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::NULL, line);
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

/// `sb.Insert(idx, text)` — splice `text` into `sb.__buffer` at `idx`.
///
/// Stack on entry: `[sb, idx, text]` ; Stack on exit: `[null]`
///
/// Implemented as `before = buf.substring(0, idx)`, `after =
/// buf.substring(idx, len)`, `buf' = before + text + after`. Pure
/// bytecode: `Op::STR_SUBSTRING` + `Op::STR_LENGTH` + `Op::DYN_ADD`.
pub fn emit_sb_insert(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = chunk.add_constant(Value::String(Arc::from(BUFFER_KEY)));
    let text_slot = reserve_slot(chunk);
    let idx_slot = reserve_slot(chunk);
    let buf_slot = reserve_slot(chunk);

    // Stash text and idx args.
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunk.emit_op(Op::DROP, line);

    // Read sb.__buffer once and stash.
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, buf_slot, line);
    chunk.emit_op(Op::DROP, line);
    // Stack now: [sb]

    // before = buf.substring(0, idx)
    chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    push_const(chunk, Value::I32(0), line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);
    // Stack: [sb, before]

    // before + text
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    // Stack: [sb, before+text]

    // after = buf.substring(idx, buf.length)
    chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);
    // Stack: [sb, before+text, after]

    // (before + text) + after
    crate::emitter::ops::emit_dyn_add(chunk, line);
    // Stack: [sb, full]

    // sb.__buffer = full
    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::NULL, line);
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

    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, start_slot, line);
    chunk.emit_op(Op::DROP, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, buf_slot, line);
    chunk.emit_op(Op::DROP, line);

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

    crate::emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::NULL, line);
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

    // Stash old, new
    chunk.emit_op_u16(Op::LOCAL_SET, new_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, old_slot, line);
    chunk.emit_op(Op::DROP, line);
    // Stack: [sb]

    // [sb, sb, buf, old, new] — buffer + replace args
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, old_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, new_slot, line);

    // ecma:string.replaceAll(buf, old, new) → replaced
    chunk.emit_op_u16(Op::CALL_IMPORT, replace_idx, line);
    chunk.emit(3, line);
    // Stack: [sb, replaced]

    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::NULL, line);
}
