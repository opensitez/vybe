//! `kotlin.time.Duration` — a Duration IS a number of MILLISECONDS
//! (`Duration.INFINITE` is the f64 infinity), so `+ - * /` and every
//! comparison are ordinary numeric ops. The walker lowers the spellings
//! (`toDuration`, `inWhole*`, `toLong(unit)`) to arithmetic; only the two
//! pieces that need bytecode live here.

use vybe_compiler::primitives::instructions::{core_wasm, host};
use vybe_compiler::primitives::ops;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// `trunc(v * num / den)` — every `inWholeX`/`toLong(unit)`.
///
/// Stack: `[v, num, den] → [whole]`. Truncation is toward ZERO, Kotlin's
/// contract for the whole-unit accessors (`(-1501).ms.inWholeSeconds == -1`).
pub fn emit_duration_whole(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let den = chunks[current].alloc_scratch(1);
    let num = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], den, line);
    set(&mut chunks[current], num, line);
    get(&mut chunks[current], num, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    get(&mut chunks[current], den, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_TRUNC, line);
}

/// Append `n + suffix` to the string in `out` when `n > 0`.
fn append_component(
    chunks: &mut Vec<Chunk>,
    current: usize,
    out: u16,
    n: u16,
    suffix: &str,
    line: u32,
) {
    get(&mut chunks[current], n, line);
    core_wasm::f64_const(&mut chunks[current], line, 0.0);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], n, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_string_const(suffix, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], out, line);
    chunks[current].emit_end(line);
}

/// Kotlin's `Duration.toString()` shape: `"2.5s"`, `"1m 30s"`,
/// `"1d 2h 3m 4s"`, `"0s"`, `"Infinity"`; negative durations lead with `-`.
///
/// Stack: `[ms] → [string]`.
pub fn emit_duration_str(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], v, line);

    get(&mut chunks[current], v, line);
    let is_finite = chunks[current].add_import("ecma:number", "isFinite");
    chunks[current].emit_call(is_finite, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    {
        // ±Infinity spell themselves.
        get(&mut chunks[current], v, line);
        core_wasm::f64_const(&mut chunks[current], line, 0.0);
        chunks[current].emit_op(Op::F64_LT, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const("-Infinity", line);
        chunks[current].emit_else(line);
        chunks[current].emit_string_const("Infinity", line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_else(line);
    {
        let neg = chunks[current].alloc_scratch(1);
        get(&mut chunks[current], v, line);
        core_wasm::f64_const(&mut chunks[current], line, 0.0);
        chunks[current].emit_op(Op::F64_LT, line);
        set(&mut chunks[current], neg, line);
        get(&mut chunks[current], v, line);
        chunks[current].emit_op(Op::F64_ABS, line);
        set(&mut chunks[current], v, line);

        // Split |ms| into d/h/m and fractional seconds.
        let part = |chunks: &mut Vec<Chunk>, v: u16, per: f64, line: u32| {
            get(&mut chunks[current], v, line);
            core_wasm::f64_const(&mut chunks[current], line, per);
            chunks[current].emit_op(Op::F64_DIV, line);
            chunks[current].emit_op(Op::F64_TRUNC, line);
            let slot = chunks[current].alloc_scratch(1);
            set(&mut chunks[current], slot, line);
            get(&mut chunks[current], v, line);
            get(&mut chunks[current], slot, line);
            core_wasm::f64_const(&mut chunks[current], line, per);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_SUB, line);
            set(&mut chunks[current], v, line);
            slot
        };
        let d = part(chunks, v, 86_400_000.0, line);
        let h = part(chunks, v, 3_600_000.0, line);
        let m = part(chunks, v, 60_000.0, line);
        let s = chunks[current].alloc_scratch(1);
        get(&mut chunks[current], v, line);
        core_wasm::f64_const(&mut chunks[current], line, 1000.0);
        chunks[current].emit_op(Op::F64_DIV, line);
        set(&mut chunks[current], s, line);

        let out = chunks[current].alloc_scratch(1);
        chunks[current].emit_string_const("", line);
        set(&mut chunks[current], out, line);
        append_component(chunks, current, out, d, "d ", line);
        append_component(chunks, current, out, h, "h ", line);
        append_component(chunks, current, out, m, "m ", line);
        append_component(chunks, current, out, s, "s", line);

        // Nothing appended (a true zero) → "0s"; else drop a trailing space.
        get(&mut chunks[current], out, line);
        chunks[current].emit_string_const("", line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        chunks[current].emit_string_const("0s", line);
        set(&mut chunks[current], out, line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], out, line);
        host::emit(&mut chunks[current], "ecma:string", "trim", 1, line);
        set(&mut chunks[current], out, line);
        chunks[current].emit_end(line);

        get(&mut chunks[current], neg, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const("-", line);
        get(&mut chunks[current], out, line);
        ops::emit_dyn_add(&mut chunks[current], line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], out, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);
}
