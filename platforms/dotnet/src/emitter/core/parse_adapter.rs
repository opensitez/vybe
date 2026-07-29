//! .NET `<Type>.Parse(s)` / `TryParse(s, out)` — bytecode emitters.
//!
//! `int.Parse("42")` returns `42`; `int.Parse("abc")` throws
//! `FormatException` per ECMA-335. JS `Number(s)` returns `NaN` on
//! failure, which is what every `intrinsic:cint` emit currently does.
//! These adapters wrap the JS coercion with a NaN check that throws a
//! .NET-shape error so try/catch around the parse picks it up.
//!
//! Wired into the C# / VB profiles via `common:dotnet.parse_*`.

use vybe_bytecode::opcode::Op;
use vybe_bytecode::Chunk;
use vybe_compiler::compiler::instructions::{core_wasm, host};

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

/// `int.Parse(s)` — `Number(s)` then `Math.floor`. If the result is
/// NaN, throw `Error("Input string was not in a correct format.")`
/// (the .NET `FormatException` message). Stack: `[s]` → `[i32]`.
pub fn emit_parse_int(chunks: &mut [Chunk], current: usize, line: u32) {
    let number_idx = chunks[current].add_import("ecma:number", "Number");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, number_idx, line);
    chunk.emit(1, line);
    let result = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);

    // NaN check: `r !== r` is the canonical NaN test.
    let if_block = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    chunk.emit_br_if(0, line);
    // NaN — throw FormatException-shaped object so `e.Message` works.
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_string_const("Input string was not in a correct format.", line);
    vybe_compiler::compiler::errors::emit_exception_new_finalize(chunk, "FormatException", line);
    vybe_compiler::compiler::errors::emit_throw(chunk, line);
    chunk.emit_end(line);
    chunk.patch_block(if_block);

    // Floor for integer semantics (matches `intrinsic:cint`).
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    chunk.emit_op(Op::F64_FLOOR, line);
}

/// `int.TryParse(s)` core for the walker's out-param desugar: like
/// `emit_parse_int` but yields `null` instead of throwing on failure, so
/// `(r = int.TryParse(s)) != null` is the success bool. The desugar restores
/// .NET's zero-on-failure out value with a `|| ((r = 0) == null)` fallback.
/// Stack: `[s]` → `[i32 | null]`.
pub fn emit_try_parse_int(chunks: &mut [Chunk], current: usize, line: u32) {
    let number_idx = chunks[current].add_import("ecma:number", "Number");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, number_idx, line);
    chunk.emit(1, line);
    let result = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);

    // `r === r` is the canonical NaN test: NaN fails self-equality.
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    chunk.emit_else(line);
    chunk.emit_op(Op::NULL, line);
    chunk.emit_end(line);
}

/// `double.Parse(s)` — `Number(s)` with NaN guard. Stack: `[s]` → `[f64]`.
pub fn emit_parse_double(chunks: &mut [Chunk], current: usize, line: u32) {
    let number_idx = chunks[current].add_import("ecma:number", "Number");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, number_idx, line);
    chunk.emit(1, line);
    let result = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);

    let if_block = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    chunk.emit_br_if(0, line);
    chunk.emit_string_const("Input string was not in a correct format.", line);
    vybe_compiler::compiler::errors::emit_throw(chunk, line);
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
    let lower_idx = chunks[current].add_import("ecma:string", "toLowerCase");
    let chunk = &mut chunks[current];
    chunk.emit_call(lower_idx, 1, line);
    let lc = alloc_local(chunk);
    let result = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, lc, line);

    // A plain `()->()` block cannot carry a value out on the stack, so each arm
    // stages its boolean in `result` and the value is reloaded after the blocks
    // close — the same staging the console adapter uses.
    let outer = chunk.emit_block(line);
    // Branch 1: "true"
    let not_true = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, lc, line);
    chunk.emit_string_const("true", line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::compiler::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);
    core_wasm::bool_const(chunk, line, true);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_br(1, line);
    chunk.emit_end(line);
    chunk.patch_block(not_true);
    // Branch 2: "false"
    let not_false = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, lc, line);
    chunk.emit_string_const("false", line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::compiler::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);
    core_wasm::bool_const(chunk, line, false);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_br(1, line);
    chunk.emit_end(line);
    chunk.patch_block(not_false);
    // Neither — throw a FormatException-shape object.
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_string_const("String was not recognized as a valid Boolean.", line);
    vybe_compiler::compiler::errors::emit_exception_new_finalize(chunk, "FormatException", line);
    vybe_compiler::compiler::errors::emit_throw(chunk, line);
    chunk.emit_end(line);
    chunk.patch_block(outer);

    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
}

/// `char.Parse(s)` — require a single-character string and return it.
pub fn emit_parse_char(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    let ok_block = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    core_wasm::i32_const(chunk, line, 1);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    chunk.emit_br_if(0, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_string_const("String must be exactly one character long.", line);
    vybe_compiler::compiler::errors::emit_exception_new_finalize(chunk, "FormatException", line);
    vybe_compiler::compiler::errors::emit_throw(chunk, line);
    chunk.emit_end(line);
    chunk.patch_block(ok_block);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
}
