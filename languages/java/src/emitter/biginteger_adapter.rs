//! Java `BigInteger` adapters backed by the existing ECMA BigInt runtime.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use vybe_emitter::instructions::host;

pub fn emit_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:bigint", "toString", 1, line);
}

pub fn emit_binary(chunks: &mut [Chunk], current: usize, op: &'static str, line: u32) {
    host::emit(&mut chunks[current], "ecma:bigint", op, 2, line);
}

pub fn emit_unary(chunks: &mut [Chunk], current: usize, op: &'static str, line: u32) {
    host::emit(&mut chunks[current], "ecma:bigint", op, 1, line);
}

pub fn emit_abs(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_bigint_i32(chunks, current, 0, line);
    host::emit(&mut chunks[current], "ecma:bigint", "lt", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(&mut chunks[current], "ecma:bigint", "neg", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_compare_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let other_slot = chunks[current].alloc_scratch(1);
    let self_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, other_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, self_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, other_slot, line);
    host::emit(&mut chunks[current], "ecma:bigint", "lt", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, other_slot, line);
    host::emit(&mut chunks[current], "ecma:bigint", "gt", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_signum(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_bigint_i32(chunks, current, 0, line);
    host::emit(&mut chunks[current], "ecma:bigint", "lt", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_bigint_i32(chunks, current, 0, line);
    host::emit(&mut chunks[current], "ecma:bigint", "gt", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_min_max(chunks: &mut [Chunk], current: usize, want_min: bool, line: u32) {
    let other_slot = chunks[current].alloc_scratch(1);
    let self_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, other_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, self_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, other_slot, line);
    host::emit(
        &mut chunks[current],
        "ecma:bigint",
        if want_min { "lt" } else { "gt" },
        2,
        line,
    );
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, self_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, other_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_bit_length(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_i32_const(2, line);
    host::emit(
        &mut chunks[current],
        "ecma:bigint",
        "toStringRadix",
        2,
        line,
    );
    vybe_emitter::strings::emit_length(&mut chunks[current], line);
}

pub fn emit_test_bit(chunks: &mut [Chunk], current: usize, line: u32) {
    let bit_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bit_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_bigint_i32(chunks, current, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bit_slot, line);
    host::emit(&mut chunks[current], "ecma:bigint", "shl", 2, line);
    host::emit(&mut chunks[current], "ecma:bigint", "and", 2, line);
    emit_bigint_i32(chunks, current, 0, line);
    host::emit(&mut chunks[current], "ecma:bigint", "ne", 2, line);
}

pub fn emit_gcd(chunks: &mut [Chunk], current: usize, line: u32) {
    let b_slot = chunks[current].alloc_scratch(1);
    let a_slot = chunks[current].alloc_scratch(1);
    let tmp_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a_slot, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
    emit_bigint_i32(chunks, current, 0, line);
    host::emit(&mut chunks[current], "ecma:bigint", "eq", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
    host::emit(&mut chunks[current], "ecma:bigint", "rem", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, tmp_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tmp_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
    emit_abs(chunks, current, line);
}

pub fn emit_is_probable_prime(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    emit_to_number(chunks, current, line);
    emit_number_is_prime(chunks, current, line);
}

pub fn emit_next_probable_prime(chunks: &mut [Chunk], current: usize, line: u32) {
    let n_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_slot, line);
    emit_to_number(chunks, current, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n_slot, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_slot, line);
    emit_number_is_prime(chunks, current, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, n_slot, line);
    host::emit(&mut chunks[current], "ecma:bigint", "BigInt", 1, line);
}

fn emit_number_is_prime(chunks: &mut [Chunk], current: usize, line: u32) {
    let n_slot = chunks[current].alloc_scratch(1);
    let i_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n_slot, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, n_slot, line);
    chunks[current].emit_f64_const(2.0, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_slot, line);
    vybe_emitter::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, n_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    vybe_emitter::math::emit_c_fmod(&mut chunks[current], line);
    chunks[current].emit_f64_const(0.0, line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

fn emit_to_number(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:bigint", "toString", 1, line);
    host::emit(&mut chunks[current], "ecma:number", "Number", 1, line);
}

fn emit_bigint_i32(chunks: &mut [Chunk], current: usize, value: i32, line: u32) {
    chunks[current].emit_i32_const(value, line);
    host::emit(&mut chunks[current], "ecma:bigint", "BigInt", 1, line);
}
