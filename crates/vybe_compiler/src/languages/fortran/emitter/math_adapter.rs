//! Fortran math helpers — Rust inline opcode emitters.
//!
//! Implements `max(a, b, c, ...)` / `min(a, b, c, ...)` as variadic
//! intrinsics. Composes pure WASM `f64.max` / `f64.min` opcodes —
//! no host calls.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

/// Fortran `max(a, b, c, ...)` — variadic.
/// Stack on entry: `[arg0, arg1, ..., argN-1]` (argc args).
/// Stack on exit: `[largest]`.
///
/// Composes: chained `f64.max` (one fewer than argc).
pub fn emit_fortran_max(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        chunk.emit_op(Op::NULL, line);
        return;
    }
    for _ in 1..argc {
        chunk.emit_op(Op::F64_MAX, line);
    }
}

/// Fortran `min(a, b, c, ...)` — variadic.
/// Stack on entry: `[arg0, arg1, ..., argN-1]` (argc args).
/// Stack on exit: `[smallest]`.
pub fn emit_fortran_min(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        chunk.emit_op(Op::NULL, line);
        return;
    }
    for _ in 1..argc {
        chunk.emit_op(Op::F64_MIN, line);
    }
}
