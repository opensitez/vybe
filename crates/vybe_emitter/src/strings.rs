//! String interpolation helpers — shared bytecode patterns for string building.
//!
//! All languages (Python f-strings, Dart $interpolation, JS template literals,
//! C# $strings, VB string concat) emit the same pattern: compile parts,
//! toString each expression, concatenate.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

/// Emit toString conversion on TOS via host call.
/// Stack before: [value]  Stack after: [string]
/// Emit toString conversion. Adds import to the given chunk.
/// Use `emit_to_string_with_import` if your compiler requires imports in chunk 0.
pub fn emit_to_string(chunk: &mut Chunk, line: u32) {
    let to_str = chunk.add_import("ecma:string", "String");
    chunk.emit_call(to_str, 1, line);
}

/// Emit toString using a pre-resolved import index.
pub fn emit_to_string_with_import(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 1, line);
}

/// Emit concatenation of N string parts already on the stack.
/// If N == 0, pushes empty string. If N == 1, no-op.
/// Stack before: [part1, part2, ..., partN]  Stack after: [concatenated_string]
pub fn emit_concat(chunk: &mut Chunk, part_count: usize, line: u32) {
    if part_count == 0 {
        chunk.emit_string_const("", line);
    } else if part_count > 1 {
        let concat_idx = chunk.add_import("wasm:js-string", "concat");
        let base = chunk.local_count;
        chunk.local_count = chunk
            .local_count
            .checked_add(part_count as u16)
            .expect("emit_concat: local slot overflow");
        for i in (0..part_count).rev() {
            chunk.emit_op_u16(Op::LOCAL_SET, base + i as u16, line);
        }
        chunk.emit_op_u16(Op::LOCAL_GET, base, line);
        for i in 1..part_count {
            chunk.emit_op_u16(Op::LOCAL_GET, base + i as u16, line);
            chunk.emit_call(concat_idx, 2, line);
        }
    }
    // part_count == 1: string is already on stack, nothing to do
}

/// Emit a complete string interpolation: convert expression part to string.
/// Call this after compiling each expression part (between literal parts).
/// Stack before: [value]  Stack after: [string]
pub fn emit_interpolation_part(chunk: &mut Chunk, line: u32) {
    emit_to_string(chunk, line);
}

/// Emit a string literal part.
/// Stack: [] → [string]
pub fn emit_literal_part(chunk: &mut Chunk, text: &str, line: u32) {
    chunk.emit_string_const(text, line);
}

// ── String operations ──────────────────────────────────────────────────
// Single-opcode wrappers for consistency across all compilers.

/// String length. Stack: [string] → [i32]
pub fn emit_length(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(idx, 1, line);
}

/// Substring. Stack: [string, start, length] → [string]
pub fn emit_substring(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "substring");
    chunk.emit_call(idx, 3, line);
}

/// Index of substring. Stack: [haystack, needle] → [i32]
pub fn emit_index_of(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "indexOf");
    chunk.emit_call(idx, 2, line);
}

/// Last index of substring. Stack: [haystack, needle] → [i32]
pub fn emit_last_index_of(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "lastIndexOf");
    chunk.emit_call(idx, 2, line);
}

/// Replace. Stack: [string, search, replace] → [string]
pub fn emit_replace(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "replaceAll");
    chunk.emit_call(idx, 3, line);
}

/// Split. Stack: [string, delimiter] → [array]
pub fn emit_split(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "split");
    chunk.emit_call(idx, 2, line);
}

/// To lowercase. Stack: [string] → [string]
pub fn emit_to_lower(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "toLowerCase");
    chunk.emit_call(idx, 1, line);
}

/// To uppercase. Stack: [string] → [string]
pub fn emit_to_upper(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "toUpperCase");
    chunk.emit_call(idx, 1, line);
}

/// Trim whitespace. Stack: [string] → [string]
pub fn emit_trim(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "trim");
    chunk.emit_call(idx, 1, line);
}

/// Trim start. Stack: [string] → [string]
pub fn emit_trim_start(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "trimStart");
    chunk.emit_call(idx, 1, line);
}

/// Trim end. Stack: [string] → [string]
pub fn emit_trim_end(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "trimEnd");
    chunk.emit_call(idx, 1, line);
}

/// Repeat string. Stack: [string, count] → [string]
pub fn emit_repeat(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "repeat");
    chunk.emit_call(idx, 2, line);
}

/// Pairwise concatenation. Stack: [a, b] → [ab]
pub fn emit_str_concat(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "concat");
    chunk.emit_call(idx, 2, line);
}

/// Concat with ToString coercion of both operands (VB `&`, C# string
/// concat, …): `wasm:js-string.concat` is spec-strict and traps on
/// non-string args, so operands that may be numbers/booleans must go
/// through `ecma:string.String` first.
/// Stack: [l, r] → [String(l) + String(r)]
pub fn emit_str_concat_coercing(chunk: &mut Chunk, line: u32) {
    let to_str = chunk.add_import("ecma:string", "String");
    let concat = chunk.add_import("wasm:js-string", "concat");
    let r = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, r, line); // [l]
    chunk.emit_call(to_str, 1, line); // [String(l)]
    chunk.emit_op_u16(Op::LOCAL_GET, r, line); // [String(l), r]
    chunk.emit_call(to_str, 1, line); // [String(l), String(r)]
    chunk.emit_call(concat, 2, line);
}

/// Reverse string. Stack: [string] → [reversed]
/// Composed: split("") → reverse() → join("")
pub fn emit_str_reverse(chunk: &mut Chunk, line: u32) {
    chunk.emit_string_const("", line);
    let split = chunk.add_import("ecma:string", "split");
    chunk.emit_call(split, 2, line);
    let reverse = chunk.add_import("ecma:array", "reverse");
    chunk.emit_call(reverse, 1, line);
    chunk.emit_string_const("", line);
    let join = chunk.add_import("ecma:array", "join");
    chunk.emit_call(join, 2, line);
}
