//! String interpolation helpers — shared bytecode patterns for string building.
//!
//! All languages (Python f-strings, Dart $interpolation, JS template literals,
//! C# $strings, VB string concat) emit the same pattern: compile parts,
//! toString each expression, concatenate.

use std::rc::Rc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

/// Emit toString conversion on TOS via host call.
/// Stack before: [value]  Stack after: [string]
pub fn emit_to_string(chunk: &mut Chunk, line: u32) {
    let to_str = chunk.add_import("vybe:convert", "toString");
    chunk.emit_op_u16(Op::call_import, to_str, line);
    chunk.emit(1, line);
}

/// Emit concatenation of N string parts already on the stack.
/// If N == 0, pushes empty string. If N == 1, no-op.
/// Stack before: [part1, part2, ..., partN]  Stack after: [concatenated_string]
pub fn emit_concat(chunk: &mut Chunk, part_count: usize, line: u32) {
    if part_count == 0 {
        let c = chunk.add_constant(Value::String(Rc::from("")));
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
    let c = chunk.add_constant(Value::String(Rc::from(text)));
    chunk.emit_op_u16(Op::r#const, c, line);
}
