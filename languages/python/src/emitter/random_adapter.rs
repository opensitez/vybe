use vybe_compiler::primitives::{collections, globals, random, tuples};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::opcode::heaptype::HT_EXTERN;

const RNG_GLOBAL: &str = "__vybe_rng";

fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].local_count;
    chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + i as u16, line);
    }
    base
}

fn call_import(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn emit_slot_f64(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:value", "toNumber", 1, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
}

fn emit_slot_i32(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    emit_slot_f64(chunks, current, slot, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
}

fn emit_value_error(chunks: &mut [Chunk], current: usize, message: &str, line: u32) {
    chunks[current].emit_string_const(message, line);
    crate::emitter::runtime_adapter::emit_py_exception(chunks, current, 1, "ValueError", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
}

pub fn emit_r(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    random::emit_next_unit(chunks, current, line);
}

pub fn emit_uniform(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        chunks[current].emit_f64_const(0.0, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    emit_slot_f64(chunks, current, base, line);
    emit_slot_f64(chunks, current, base + 1, line);
    emit_slot_f64(chunks, current, base, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    random::emit_next_unit(chunks, current, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

pub fn emit_expovariate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 1 {
        chunks[current].emit_f64_const(0.0, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    chunks[current].emit_f64_const(1.0, line);
    random::emit_next_unit(chunks, current, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    call_import(chunks, current, "ecma:math", "log", 1, line);
    chunks[current].emit_op(Op::F64_NEG, line);
    emit_slot_f64(chunks, current, base, line);
    chunks[current].emit_op(Op::F64_DIV, line);
}

pub fn emit_gauss(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        chunks[current].emit_f64_const(0.0, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let x2pi = chunks[current].alloc_scratch(1);
    let g2rad = chunks[current].alloc_scratch(1);

    random::emit_next_unit(chunks, current, line);
    chunks[current].emit_f64_const(2.0 * std::f64::consts::PI, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, x2pi, line);

    chunks[current].emit_f64_const(1.0, line);
    random::emit_next_unit(chunks, current, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    call_import(chunks, current, "ecma:math", "log", 1, line);
    chunks[current].emit_f64_const(-2.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SQRT, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, g2rad, line);

    emit_slot_f64(chunks, current, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x2pi, line);
    call_import(chunks, current, "ecma:math", "cos", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, g2rad, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    emit_slot_f64(chunks, current, base + 1, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

pub fn emit_lognormvariate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_gauss(chunks, current, argc, line);
    call_import(chunks, current, "ecma:math", "exp", 1, line);
}

pub fn emit_triangular(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        chunks[current].emit_f64_const(0.0, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    emit_slot_f64(chunks, current, base, line);
    emit_slot_f64(chunks, current, base + 1, line);
    emit_slot_f64(chunks, current, base, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    random::emit_next_unit(chunks, current, line);
    chunks[current].emit_op(Op::F64_SQRT, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

pub fn emit_paretovariate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 1 {
        chunks[current].emit_f64_const(1.0, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let denom = chunks[current].alloc_scratch(1);
    chunks[current].emit_f64_const(1.0, line);
    random::emit_next_unit(chunks, current, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    call_import(chunks, current, "ecma:math", "log", 1, line);
    emit_slot_f64(chunks, current, base, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    call_import(chunks, current, "ecma:math", "exp", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, denom, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, denom, line);
    chunks[current].emit_op(Op::F64_DIV, line);
}

pub fn emit_weibullvariate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        chunks[current].emit_f64_const(0.0, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    emit_slot_f64(chunks, current, base, line);
    chunks[current].emit_f64_const(1.0, line);
    random::emit_next_unit(chunks, current, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    call_import(chunks, current, "ecma:math", "log", 1, line);
    chunks[current].emit_op(Op::F64_NEG, line);
    call_import(chunks, current, "ecma:math", "log", 1, line);
    emit_slot_f64(chunks, current, base + 1, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    call_import(chunks, current, "ecma:math", "exp", 1, line);
    chunks[current].emit_op(Op::F64_MUL, line);
}

pub fn emit_vonmisesvariate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        chunks[current].emit_f64_const(0.0, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    emit_slot_f64(chunks, current, base, line);
    random::emit_next_unit(chunks, current, line);
    chunks[current].emit_f64_const(0.5, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    emit_slot_f64(chunks, current, base + 1, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

pub fn emit_gammavariate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        chunks[current].emit_f64_const(0.0, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let total = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, total, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(3, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, total, line);
    chunks[current].emit_f64_const(1.0, line);
    random::emit_next_unit(chunks, current, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    call_import(chunks, current, "ecma:math", "log", 1, line);
    chunks[current].emit_op(Op::F64_NEG, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, total, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, total, line);
    emit_slot_f64(chunks, current, base + 1, line);
    chunks[current].emit_op(Op::F64_MUL, line);
}

pub fn emit_betavariate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        chunks[current].emit_f64_const(0.0, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let y1 = chunks[current].alloc_scratch(1);
    let y2 = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    chunks[current].emit_f64_const(1.0, line);
    emit_gammavariate(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, y1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    chunks[current].emit_f64_const(1.0, line);
    emit_gammavariate(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, y2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, y1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, y1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, y2, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op(Op::F64_DIV, line);
}

pub fn emit_getrandbits(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 1 {
        chunks[current].emit_i32_const(0, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let k = chunks[current].alloc_scratch(1);
    let total = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    emit_slot_i32(chunks, current, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, k, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, total, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, k, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, total, line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op(Op::I32_MUL, line);
    random::emit_next_unit(chunks, current, line);
    chunks[current].emit_f64_const(0.5, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, total, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, total, line);
}

pub fn emit_randbytes(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 1 {
        chunks[current].emit_array_new_fixed(0, 0, line);
        call_import(chunks, current, "ecma:uint8array", "from", 1, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let n = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    emit_slot_i32(chunks, current, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    chunks[current].emit_array_new_fixed(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    random::emit_next_unit(chunks, current, line);
    chunks[current].emit_f64_const(256.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_U, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    call_import(chunks, current, "ecma:uint8array", "from", 1, line);
}

pub fn emit_randint(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        chunks[current].emit_i32_const(0, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let start = chunks[current].alloc_scratch(1);
    let stop = chunks[current].alloc_scratch(1);
    emit_slot_i32(chunks, current, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
    emit_slot_i32(chunks, current, base + 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stop, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, stop, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    emit_value_error(chunks, current, "empty range for randrange", line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stop, line);
    random::emit_rand_int_inclusive(chunks, current, line);
}

pub fn emit_randrange(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 3 {
        chunks[current].emit_i32_const(0, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let start = chunks[current].alloc_scratch(1);
    let stop = chunks[current].alloc_scratch(1);
    let step = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    emit_slot_i32(chunks, current, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);
    emit_slot_i32(chunks, current, base + 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stop, line);
    emit_slot_i32(chunks, current, base + 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, step, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, step, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    emit_value_error(chunks, current, "randrange step must not be zero", line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, step, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stop, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, step, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, step, line);
    chunks[current].emit_op(Op::I32_DIV_S, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stop, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, step, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, step, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op(Op::I32_DIV_S, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_if(line);
    emit_value_error(chunks, current, "empty range for randrange", line);
    chunks[current].emit_end(line);

    random::emit_next_unit(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, step, line);
    chunks[current].emit_op(Op::I32_MUL, line);
    chunks[current].emit_op(Op::I32_ADD, line);
}

pub fn emit_choices(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 4 {
        chunks[current].emit_array_new_fixed(0, 0, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let pop = base;
    let weights = base + 1;
    let cum = base + 2;
    let k_arg = base + 3;
    let result = chunks[current].alloc_scratch(1);
    let cums = chunks[current].alloc_scratch(1);
    let total = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let target = chunks[current].alloc_scratch(1);
    let pick = chunks[current].alloc_scratch(1);
    let chosen = chunks[current].alloc_scratch(1);
    let k = chunks[current].alloc_scratch(1);

    chunks[current].emit_array_new_fixed(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    chunks[current].emit_array_new_fixed(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cums, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, total, line);
    emit_slot_i32(chunks, current, k_arg, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, k, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cum, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    fill_cumulative_from_array(chunks, current, cums, cum, total, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, weights, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    fill_cumulative_from_weights(chunks, current, cums, weights, total, line);
    chunks[current].emit_else(line);
    fill_cumulative_uniform(chunks, current, cums, pop, total, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    let outer = chunks[current].emit_block(line);
    let (outer_lp, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, k, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    random::emit_next_unit(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, total, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, target, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pop, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pick, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, chosen, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    let scan = chunks[current].emit_block(line);
    let (scan_lp, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cums, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, chosen, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, target, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cums, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    collections::emit_get(chunks, current, line);
    call_import(chunks, current, "ecma:value", "toNumber", 1, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pick, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, chosen, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(scan_lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(scan);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pop, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pick, line);
    collections::emit_get(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}

fn fill_cumulative_from_array(
    chunks: &mut [Chunk],
    current: usize,
    cums: u16,
    source: u16,
    total: u16,
    line: u32,
) {
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    call_import(chunks, current, "ecma:value", "toNumber", 1, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, total, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cums, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, total, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
}

fn fill_cumulative_from_weights(
    chunks: &mut [Chunk],
    current: usize,
    cums: u16,
    source: u16,
    total: u16,
    line: u32,
) {
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, total, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    call_import(chunks, current, "ecma:value", "toNumber", 1, line);
    call_import(chunks, current, "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, total, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cums, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, total, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
}

fn fill_cumulative_uniform(
    chunks: &mut [Chunk],
    current: usize,
    cums: u16,
    pop: u16,
    total: u16,
    line: u32,
) {
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pop, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cums, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, total, line);
}

pub fn emit_getstate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    globals::emit_read(&mut chunks[current], RNG_GLOBAL, line);
    tuples::emit_tuple(chunks, current, 1, line);
}

pub fn emit_setstate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 1 {
        chunks[current].emit_ref_null(HT_EXTERN, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    globals::emit_write(&mut chunks[current], RNG_GLOBAL, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
}
