//! String interpolation helpers — shared bytecode patterns for string building.
//!
//! All languages (Python f-strings, Dart $interpolation, JS template literals,
//! C# $strings, VB string concat) emit the same pattern: compile parts,
//! toString each expression, concatenate.

use std::sync::Arc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

/// Emit toString conversion on TOS via host call.
/// Stack before: [value]  Stack after: [string]
/// Emit toString conversion. Adds import to the given chunk.
/// Use `emit_to_string_with_import` if your compiler requires imports in chunk 0.
pub fn emit_to_string(chunk: &mut Chunk, line: u32) {
    let to_str = chunk.add_import("vybe:convert", "toString");
    chunk.emit_op_u16(Op::call_import, to_str, line);
    chunk.emit(1, line);
}

/// Emit toString using a pre-resolved import index.
pub fn emit_to_string_with_import(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_op_u16(Op::call_import, import_idx, line);
    chunk.emit(1, line);
}

/// Emit concatenation of N string parts already on the stack.
/// If N == 0, pushes empty string. If N == 1, no-op.
/// Stack before: [part1, part2, ..., partN]  Stack after: [concatenated_string]
pub fn emit_concat(chunk: &mut Chunk, part_count: usize, line: u32) {
    if part_count == 0 {
        let c = chunk.add_constant(Value::String(Arc::from("")));
        chunk.emit_op_u16(Op::r#const, c, line);
    } else if part_count > 1 {
        // Use str_concat_n if available (more efficient for many parts)
        if part_count <= 255 {
            chunk.emit_op_u8(Op::str_concat_n, part_count as u8, line);
        } else {
            // Fallback: pairwise concat
            for _ in 1..part_count {
                chunk.emit_op(Op::str_concat, line);
            }
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
    let c = chunk.add_constant(Value::String(Arc::from(text)));
    chunk.emit_op_u16(Op::r#const, c, line);
}

// ── String operations ──────────────────────────────────────────────────
// Single-opcode wrappers for consistency across all compilers.

/// String length. Stack: [string] → [i32]
pub fn emit_length(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::str_length, line); }

/// Substring. Stack: [string, start, length] → [string]
pub fn emit_substring(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::str_substring, line); }

/// Index of substring. Stack: [haystack, needle] → [i32]
pub fn emit_index_of(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::str_index_of, line); }

/// Last index of substring. Stack: [haystack, needle] → [i32]
pub fn emit_last_index_of(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::str_last_index_of, line); }

/// Replace. Stack: [string, search, replace] → [string]
pub fn emit_replace(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::str_replace, line); }

/// Split. Stack: [string, delimiter] → [array]
pub fn emit_split(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::str_split, line); }

/// To lowercase. Stack: [string] → [string]
pub fn emit_to_lower(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::str_to_lower, line); }

/// To uppercase. Stack: [string] → [string]
pub fn emit_to_upper(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::str_to_upper, line); }

/// Trim whitespace. Stack: [string] → [string]
pub fn emit_trim(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::str_trim, line); }

/// Trim start. Stack: [string] → [string]
pub fn emit_trim_start(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::str_trim_start, line); }

/// Trim end. Stack: [string] → [string]
pub fn emit_trim_end(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::str_trim_end, line); }

/// Repeat string. Stack: [string, count] → [string]
pub fn emit_repeat(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::str_repeat, line); }

/// Pairwise concatenation. Stack: [a, b] → [ab]
pub fn emit_str_concat(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::str_concat, line); }
