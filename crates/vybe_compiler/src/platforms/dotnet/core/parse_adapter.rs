//! .NET `<Type>.Parse(s)` / `TryParse(s, out)` — bytecode emitters.
//!
//! `int.Parse("42")` returns `42`; `int.Parse("abc")` throws
//! `FormatException` per ECMA-335. JS `Number(s)` returns `NaN` on
//! failure, which is what every `intrinsic:cint` emit currently does.
//! These adapters wrap the JS coercion with a NaN check that throws a
//! .NET-shape error so try/catch around the parse picks it up.
//!
//! Wired into the C# / VB profiles via `common:dotnet.parse_*`.

use std::sync::Arc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let s = chunk.local_count;
    chunk.local_count = s + 1;
    s
}

/// `int.Parse(s)` — `Number(s)` then `Math.floor`. If the result is
/// NaN, throw `Error("Input string was not in a correct format.")`
/// (the .NET `FormatException` message). Stack: `[s]` → `[i32]`.
pub fn emit_parse_int(chunks: &mut [Chunk], current: usize, line: u32) {
    let number_idx = chunks[0].add_import("ecma:number", "Number");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, number_idx, line);
    chunk.emit(1, line);
    let result = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_op(Op::DROP, line);

    // NaN check: `r !== r` is the canonical NaN test.
    let if_block = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_br_if(0, line);
    // NaN — throw FormatException-shaped object so `e.Message` works.
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op(Op::DUP, line);
    let msg = chunk.add_constant(Value::String(Arc::from(
        "Input string was not in a correct format.",
    )));
    chunk.emit_op_u16(Op::CONST, msg, line);
    crate::emitter::errors::emit_exception_new_finalize(chunk, "FormatException", line);
    crate::emitter::errors::emit_throw(chunk, line);
    chunk.emit_end(line);
    chunk.patch_block(if_block);

    // Floor for integer semantics (matches `intrinsic:cint`).
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    chunk.emit_op(Op::F64_FLOOR, line);
}

/// `double.Parse(s)` — `Number(s)` with NaN guard. Stack: `[s]` → `[f64]`.
pub fn emit_parse_double(chunks: &mut [Chunk], current: usize, line: u32) {
    let number_idx = chunks[0].add_import("ecma:number", "Number");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, number_idx, line);
    chunk.emit(1, line);
    let result = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_op(Op::DROP, line);

    let if_block = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_br_if(0, line);
    let msg = chunk.add_constant(Value::String(Arc::from(
        "Input string was not in a correct format.",
    )));
    chunk.emit_op_u16(Op::CONST, msg, line);
    crate::emitter::errors::emit_throw(chunk, line);
    chunk.emit_end(line);
    chunk.patch_block(if_block);

    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
}

/// `bool.Parse(s)` — accepts `"true"` / `"false"` (case-insensitive),
/// throws on anything else. Stack: `[s]` → `[bool]`.
///
/// Inline-emits `s.toLowerCase() === "true"` as the truthy path and
/// throws if neither `"true"` nor `"false"` was given. Matches
/// .NET `Boolean.Parse` semantics per ECMA-335.
pub fn emit_parse_bool(chunks: &mut [Chunk], current: usize, line: u32) {
    let lower_idx = chunks[0].add_import("ecma:string", "toLowerCase");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, lower_idx, line);
    chunk.emit(1, line);
    let lc = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, lc, line);
    chunk.emit_op(Op::DROP, line);

    // If lc === "true" → push true and return.
    let true_str = chunk.add_constant(Value::String(Arc::from("true")));
    let false_str = chunk.add_constant(Value::String(Arc::from("false")));

    let outer = chunk.emit_block(line);
    // Branch 1: true
    let not_true = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, lc, line);
    chunk.emit_op_u16(Op::CONST, true_str, line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::TRUE, line);
    chunk.emit_br(1, line);
    chunk.emit_end(line);
    chunk.patch_block(not_true);
    // Branch 2: false
    let not_false = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, lc, line);
    chunk.emit_op_u16(Op::CONST, false_str, line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::FALSE, line);
    chunk.emit_br(1, line);
    chunk.emit_end(line);
    chunk.patch_block(not_false);
    // Neither — throw FormatException-shape object.
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op(Op::DUP, line);
    let msg = chunk.add_constant(Value::String(Arc::from(
        "String was not recognized as a valid Boolean.",
    )));
    chunk.emit_op_u16(Op::CONST, msg, line);
    crate::emitter::errors::emit_exception_new_finalize(chunk, "FormatException", line);
    crate::emitter::errors::emit_throw(chunk, line);
    chunk.emit_end(line);
    chunk.patch_block(outer);
}

/// `char.Parse(s)` — require a single-character string and return it.
pub fn emit_parse_char(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op(Op::DROP, line);

    let ok_block = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    chunk.emit_op(Op::I32_CONST_1, line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_br_if(0, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op(Op::DUP, line);
    let msg = chunk.add_constant(Value::String(Arc::from(
        "String must be exactly one character long.",
    )));
    chunk.emit_op_u16(Op::CONST, msg, line);
    crate::emitter::errors::emit_exception_new_finalize(chunk, "FormatException", line);
    crate::emitter::errors::emit_throw(chunk, line);
    chunk.emit_end(line);
    chunk.patch_block(ok_block);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
}
