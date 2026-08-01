//! Kotlin `Double`/`Float` classification predicates.

use vybe_compiler::primitives::ops;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// `println(x)` where `x` is statically a `Double`.
///
/// Kotlin's `Double.toString()` always carries a fraction — `2.0 * 3` prints
/// `6.0` — and an integral f64 renders as `6` through every generic path,
/// because at runtime `6.0` and `6` are the same value. Only the static type
/// distinguishes them, which is why this is reached through a
/// `[builtin_slots.double]` binding rather than by a test on the value.
///
/// Stack: `[value]` → `[]`.
pub fn emit_print_double(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
    } else {
        emit_double_to_string(chunks, current, line);
    }
    let log = chunks[current].add_import("wasi:logging/logging", "log");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, log, line);
    chunks[current].emit(1, line);
}

/// Kotlin `Double.toString()`. Stack: `[value]` → `[string]`.
///
/// `6.0` → `"6.0"`, `0.125` → `"0.125"`, `-2.0` → `"-2.0"`. Appends `.0` only
/// when the ECMA rendering produced no fraction and no exponent — `1e21`
/// renders as `1e+21` in both languages and must not gain a suffix.
pub fn emit_double_to_string(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let text = chunks[current].alloc_scratch(1);

    let to_str = chunks[current].add_import("ecma:string", "String");
    chunks[current].emit_call(to_str, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
    emit_contains_non_integral_mark(chunks, current, text, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const(".0", line);
    chunks[current].emit_end(line);
    vybe_compiler::primitives::strings::emit_concat(&mut chunks[current], 2, line);
}

/// Push i32 `1` when the rendered number already shows a fraction or exponent
/// (`.`, `e`, `E`) or is not a finite decimal at all (`NaN`, `Infinity`).
fn emit_contains_non_integral_mark(chunks: &mut Vec<Chunk>, current: usize, text: u16, line: u32) {
    let mut first = true;
    for mark in [".", "e", "E", "N", "I"] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
        chunks[current].emit_string_const(mark, line);
        let idx = chunks[current].add_import("ecma:string", "includes");
        chunks[current].emit_call(idx, 2, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        if first {
            first = false;
        } else {
            chunks[current].emit_op(Op::I32_OR, line);
        }
    }
}

/// Kotlin `Double.isInfinite()`. Stack: `[value]` → `[bool]`.
///
/// ECMA has `Number.isFinite` and `Number.isNaN` but nothing for "infinite",
/// which is neither of them: `isInfinite(x)` is `!isFinite(x) && !isNaN(x)`.
/// Composed from the two host predicates rather than adding a host function.
pub fn emit_is_infinite(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    let is_finite = chunks[current].add_import("ecma:number", "isFinite");
    chunks[current].emit_call(is_finite, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    // finite → not infinite
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    // not finite → infinite unless it is NaN
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    let is_nan = chunks[current].add_import("ecma:number", "isNaN");
    chunks[current].emit_call(is_nan, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}
