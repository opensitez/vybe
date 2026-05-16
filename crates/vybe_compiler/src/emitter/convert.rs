//! Type conversion compilation — maps language-specific casts to WASM ops + host imports.
//!
//! WASM has: i32.trunc_f64_s, f64.convert_i32_s, i32.wrap_i64, etc.
//! String conversion uses host imports (not in WASM spec).

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use crate::emitter::Target;

// ── Direct WASM opcodes ─────────────────────────────────────

/// float → int (truncate). Stack: [f64] → [i32]
pub fn emit_to_int(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_FROM_F64, line);
}

/// int → float. Stack: [i32] → [f64]
pub fn emit_to_float(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_FROM_I32, line);
}

/// i64 → i32. Stack: [i64] → [i32]
pub fn emit_i32_wrap(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_WRAP_I64, line);
}

/// i32 → i64. Stack: [i32] → [i64]
pub fn emit_i64_extend(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I64_EXTEND_I32_S, line);
}

// ── Host imports (string conversions) ───────────────────────

/// Any value → string representation. Stack: [value] → [string]
pub fn emit_to_string(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:string", "String");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}

/// String → int (parse). Stack: [string] → [i32]
///
/// Uses `ecma:number.parseInt` (§19.2.5) — stops at the first non-digit
/// so `parseInt("3.7") = 3`, matching VB `CInt`/Python `int`/PHP `intval`
/// semantics for string parsing. `Number(x)` would return 3.7 here.
pub fn emit_parse_int(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:number", "parseInt");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}

/// String → float (parse). Stack: [string] → [f64]
pub fn emit_parse_float(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("ecma:number", "Number");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}

/// Dynamic truthiness conversion. Stack: [value] → [bool]
pub fn emit_to_bool(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::DYN_TO_BOOL, line);
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
        chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
        chunk.emit(1, line);
    }
}
