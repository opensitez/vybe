//! Type conversion compilation — maps language-specific casts to WASM ops + host imports.
//!
//! WASM has: i32.trunc_f64_s, f64.convert_i32_s, i32.wrap_i64, etc.
//! String conversion uses host imports (not in WASM spec).

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use crate::Target;

// ── Direct WASM opcodes ─────────────────────────────────────

/// float → int (truncate). Stack: [f64] → [i32]
pub fn emit_to_int(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::i32_from_f64, line);
}

/// int → float. Stack: [i32] → [f64]
pub fn emit_to_float(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::f64_from_i32, line);
}

/// i64 → i32. Stack: [i64] → [i32]
pub fn emit_i32_wrap(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::i32_wrap_i64, line);
}

/// i32 → i64. Stack: [i32] → [i64]
pub fn emit_i64_extend(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::i64_extend_i32_s, line);
}

// ── Host imports (string conversions) ───────────────────────

/// Any value → string representation. Stack: [value] → [string]
pub fn emit_to_string(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:convert", "toString");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

/// String → int (parse). Stack: [string] → [i32]
pub fn emit_parse_int(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:convert", "cint");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

/// String → float (parse). Stack: [string] → [f64]
pub fn emit_parse_float(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:convert", "cdbl");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

/// Check if value is numeric. Stack: [value] → [bool]
pub fn emit_is_numeric(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:convert", "isNumeric");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

/// Dynamic truthiness conversion. Stack: [value] → [bool]
pub fn emit_to_bool(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::dyn_to_bool, line);
}

// ── Target-aware variants ───────────────────────────────────

/// Target-aware toString. On Vybe, uses host import. On pure WASM, emits
/// a type-switch that handles common cases inline.
pub fn emit_to_string_targeted(chunk: &mut Chunk, target: &Target, line: u32) {
    if target.has_module("vybe:convert") {
        emit_to_string(chunk, line);
    } else {
        // Fallback: import from a generic "env" module.
        // The embedder must provide env/toString.
        let idx = chunk.add_import("env", "toString");
        chunk.emit_op_u16(Op::call_import, idx, line);
        chunk.emit(1, line);
    }
}
