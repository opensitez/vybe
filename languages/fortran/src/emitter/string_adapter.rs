//! Fortran string helpers — Rust inline opcode emitters.
//!
//! Implements Fortran-faithful intrinsics that aren't a single
//! WASM opcode or single ECMA host fn:
//!
//! - `len_trim(s)` — `s.trimEnd().length`
//! - `adjustl(s)` — `s.trimStart() + spaces(N - s.trimStart().length)`
//!   (stub: trimStart for now; full adjustl needs declared length).
//! - `adjustr(s)` — symmetric.

use vybe_compiler::primitives::instructions::host;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// "Did this argument arrive as a whole array, or as one value?"
///
/// `merge(1, 0, values < 3)` has to ask it of the mask and of each source. It
/// cannot be asked in the AST: `Array.isArray(x)` is unreachable because
/// Fortran is case-insensitive and `Array` folds to a name that resolves to
/// nothing, and `rank(x) > 0` is a COMPARISON whose left operand looks array-ish
/// to the elementwise lowering, which then maps over a scalar.
///
/// Stack on entry: `[value]`. Stack on exit: `[boolean]`.
pub fn emit_fortran_is_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    host::emit(chunk, "ecma:array", "isArray", 1, line);
}

/// Render one item of a `print`/`write` list.
///
/// A LOGICAL writes `T`/`F`; everything else takes the general rendering. The
/// choice is made by ASKING THE VALUE — `wasm:js-boolean.test` — not by
/// inspecting the expression's shape in the walker, which could only ever see
/// literals and comparisons and left a `logical ::` variable rendering as
/// `true`. That works because a comparison now MATERIALIZES a boolean
/// (`materialize_bool_results` in the profile); before that, the value reaching
/// here was a raw i32 and there was nothing to ask.
///
/// Stack on entry: `[value]`. Stack on exit: `[string]`.
pub fn emit_fortran_io_str(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(chunk, "wasm:js-boolean", "test", 1, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("T", line);
    chunk.emit_else(line);
    chunk.emit_string_const("F", line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    // A COMPLEX writes `(real,imag)`. It is a `{real, imag}` object — the shared
    // complex shape — and `ObjectKind::Ordinary` cannot tell it apart from a
    // derived-type value, so the object carries a `__type` stamp and this asks
    // for it. Same shape and same question as Python's `repr`, different
    // spelling, which is why the spelling lives here and not in the primitive.
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_struct_get(chunk, "__type", line);
    chunk.emit_string_const("complex", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);

    chunk.emit_string_const("(", line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_struct_get(chunk, "real", line);
    vybe_compiler::primitives::strings::emit_to_string(chunk, line);
    emit_concat(chunk, line);
    chunk.emit_string_const(",", line);
    emit_concat(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_struct_get(chunk, "imag", line);
    vybe_compiler::primitives::strings::emit_to_string(chunk, line);
    emit_concat(chunk, line);
    chunk.emit_string_const(")", line);
    emit_concat(chunk, line);

    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    vybe_compiler::primitives::strings::emit_to_string(chunk, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_struct_get(chunk: &mut Chunk, key: &str, line: u32) {
    let key = chunk.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(key)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
}

fn emit_concat(chunk: &mut Chunk, line: u32) {
    host::emit(chunk, "wasm:js-string", "concat", 2, line);
}

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
