//! Fortran string helpers — Rust inline opcode emitters.
//!
//! Implements Fortran-faithful intrinsics that aren't a single
//! WASM opcode or single ECMA host fn:
//!
//! - `len_trim(s)` — `s.trimEnd().length`
//! - `adjustl(s)` — `s.trimStart() + spaces(N - s.trimStart().length)`
//!   (stub: trimStart for now; full adjustl needs declared length).
//! - `adjustr(s)` — symmetric.

use vybe_bytecode::Chunk;
use vybe_emitter::instructions::host;

/// Fortran `len_trim(s)` — length of string after stripping trailing
/// blanks. Composes `ecma:string.trimEnd` + `wasm:js-string.length`.
///
/// Stack on entry: `[s]`. Stack on exit: `[length_i32]`.
pub fn emit_fortran_len_trim(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    host::emit(chunk, "ecma:string", "trimEnd", 1, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
}

/// Fortran `adjustl(s)` — left-justify by moving leading blanks to
/// the end. Approximated as `s.trimStart()` for now (does not pad
/// to declared length — that's a fixed-len-string concern).
pub fn emit_fortran_adjustl(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    host::emit(chunk, "ecma:string", "trimStart", 1, line);
}
