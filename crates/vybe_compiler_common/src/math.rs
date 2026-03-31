//! Math compilation — maps language-specific math to WASM opcodes + host imports.
//!
//! WASM has: abs, ceil, floor, trunc, nearest, sqrt, min, max, copysign, neg
//! WASM does NOT have: pow, log, sin, cos, tan, atan2, exp, random
//! Those use host imports (standard across all languages).

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

// ── Direct WASM opcodes (no host call) ──────────────────────

pub fn emit_abs(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::f64_abs, line); }
pub fn emit_floor(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::f64_floor, line); }
pub fn emit_ceil(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::f64_ceil, line); }
pub fn emit_trunc(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::f64_trunc, line); }
pub fn emit_round(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::f64_nearest, line); }
pub fn emit_sqrt(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::f64_sqrt, line); }
pub fn emit_min(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::f64_min, line); }
pub fn emit_max(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::f64_max, line); }
pub fn emit_neg(chunk: &mut Chunk, line: u32) { chunk.emit_op(Op::f64_neg, line); }

// ── Host imports (standard math, same across all languages) ──

/// Stack: [base, exponent] → [result]
pub fn emit_pow(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:math", "pow");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(2, line);
}

/// Stack: [value] → [result]
pub fn emit_log(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:math", "log");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

pub fn emit_sin(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:math", "sin");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

pub fn emit_cos(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:math", "cos");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

pub fn emit_tan(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:math", "tan");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

pub fn emit_exp(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:math", "exp");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

/// Stack: [] → [f64 random 0..1]
pub fn emit_random(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:math", "random");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(0, line);
}
