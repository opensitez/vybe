//! libc `common:libc.*` emit dispatch.
//!
//! C stdio formatting is owned here under the libc platform, not by the
//! generic common formatter.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn call_math(chunks: &mut [Chunk], current: usize, func: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import("ecma:math", func);
    chunks[current].emit_call(idx, argc, line);
}

fn emit_erf(chunks: &mut [Chunk], current: usize, line: u32) {
    let x = chunks[current].alloc_scratch(1);
    let t = chunks[current].alloc_scratch(1);
    let poly = chunks[current].alloc_scratch(1);
    let y = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, x, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_NE, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_f64_const(f64::INFINITY, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_f64_const(f64::NEG_INFINITY, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(-1.0, line);
    chunks[current].emit_else(line);

    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    chunks[current].emit_f64_const(0.3275911, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t, line);

    chunks[current].emit_f64_const(1.061405429, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_f64_const(-1.453152027, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_f64_const(1.421413741, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_f64_const(-0.284496736, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_f64_const(0.254829592, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, poly, line);

    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, poly, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_NEG, line);
    call_math(chunks, current, "exp", 1, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, y, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, y, line);
    chunks[current].emit_op(Op::F64_NEG, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, y, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_erfc(chunks: &mut [Chunk], current: usize, line: u32) {
    let erf = chunks[current].alloc_scratch(1);
    emit_erf(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, erf, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, erf, line);
    chunks[current].emit_op(Op::F64_SUB, line);
}

fn emit_stirling_gamma(chunks: &mut [Chunk], current: usize, x: u16, line: u32) {
    chunks[current].emit_f64_const(2.0 * std::f64::consts::PI, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_SQRT, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_f64_const(std::f64::consts::E, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    call_math(chunks, current, "pow", 2, line);
    chunks[current].emit_op(Op::F64_MUL, line);

    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_f64_const(12.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_ADD, line);

    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_f64_const(288.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_ADD, line);

    chunks[current].emit_f64_const(139.0, line);
    chunks[current].emit_f64_const(51840.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_SUB, line);

    chunks[current].emit_op(Op::F64_MUL, line);
}

fn emit_tgamma(chunks: &mut [Chunk], current: usize, line: u32) {
    let x = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, x, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_f64_const(0.5, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(std::f64::consts::PI, line);
    chunks[current].emit_op(Op::F64_SQRT, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_f64_const(-0.5, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(-2.0 * std::f64::consts::PI.sqrt(), line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(f64::INFINITY, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    call_math(chunks, current, "floor", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_else(line);
    emit_stirling_gamma(chunks, current, x, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_lgamma(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_tgamma(chunks, current, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    call_math(chunks, current, "log", 1, line);
}

pub fn emit_math(name: &str, chunks: &mut [Chunk], current: usize, line: u32) -> bool {
    match name {
        "libc.math.erf" => {
            emit_erf(chunks, current, line);
            true
        }
        "libc.math.erfc" => {
            emit_erfc(chunks, current, line);
            true
        }
        "libc.math.tgamma" | "libc.math.gamma" => {
            emit_tgamma(chunks, current, line);
            true
        }
        "libc.math.lgamma" => {
            emit_lgamma(chunks, current, line);
            true
        }
        _ => false,
    }
}

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    if super::sdl::emit_sdl(name, chunks, current, argc, line) {
        return true;
    }
    if emit_math(name, chunks, current, line) {
        return true;
    }
    match name {
        "libc.stdio.printf" => {
            super::stdio_format::emit_sprintf(chunks, current, argc, line);
            let idx = chunks[current].add_import("wasi:logging/logging", "log");
            chunks[current].emit_call(idx, 1, line);
            true
        }
        "libc.stdio.sprintf" => {
            super::stdio_format::emit_sprintf(chunks, current, argc, line);
            true
        }
        "libc.stdio.vsprintf" => {
            super::stdio_format::emit_sprintf_from_array(chunks, current, line);
            true
        }
        _ => false,
    }
}
