use crate::emitter::exceptions::emit_jvm_exception_throw;
use vybe_compiler::primitives::collections;
use vybe_compiler::primitives::instructions::host;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const MULT: i64 = 0x5DEECE66D;
const ADD: i64 = 0xB;
const MASK: i64 = (1i64 << 48) - 1;

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

pub fn emit_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_i64_const(1, line);
    }
    emit_scramble_seed(chunks, current, line);
    collections::emit_array_new(chunks, current, 1, line);
}

pub fn emit_set_seed(chunks: &mut [Chunk], current: usize, line: u32) {
    let seed = chunks[current].alloc_scratch(1);
    let rng = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], seed, line);
    set(&mut chunks[current], rng, line);
    get(&mut chunks[current], seed, line);
    emit_scramble_seed(chunks, current, line);
    let state = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], state, line);
    get(&mut chunks[current], rng, line);
    chunks[current].emit_i32_const(0, line);
    get(&mut chunks[current], state, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_next_int(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 2 {
        // `nextInt(bound)` — JLS: bound <= 0 throws IllegalArgumentException.
        let bound = chunks[current].alloc_scratch(1);
        let rng = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], bound, line);
        set(&mut chunks[current], rng, line);
        get(&mut chunks[current], bound, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::I32_LE_S, line);
        chunks[current].emit_if_value(line);
        emit_jvm_exception_throw(chunks, current, "IllegalArgumentException", line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], rng, line);
        emit_next_bits(chunks, current, 31, line);
        get(&mut chunks[current], bound, line);
        chunks[current].emit_op(Op::I32_REM_U, line);
        chunks[current].emit_end(line);
        return;
    }
    if argc == 3 {
        // `nextInt(origin, bound)` — origin >= bound throws.
        let bound = chunks[current].alloc_scratch(1);
        let origin = chunks[current].alloc_scratch(1);
        let rng = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], bound, line);
        set(&mut chunks[current], origin, line);
        set(&mut chunks[current], rng, line);
        get(&mut chunks[current], bound, line);
        get(&mut chunks[current], origin, line);
        chunks[current].emit_op(Op::I32_LE_S, line);
        chunks[current].emit_if_value(line);
        emit_jvm_exception_throw(chunks, current, "IllegalArgumentException", line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], rng, line);
        emit_next_bits(chunks, current, 31, line);
        get(&mut chunks[current], bound, line);
        get(&mut chunks[current], origin, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        chunks[current].emit_op(Op::I32_REM_U, line);
        get(&mut chunks[current], origin, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_end(line);
        return;
    }
    emit_next_bits(chunks, current, 32, line);
}

fn emit_next_long_core(chunks: &mut [Chunk], current: usize, line: u32) {
    let rng = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rng, line);
    get(&mut chunks[current], rng, line);
    emit_next_bits(chunks, current, 32, line);
    chunks[current].emit_op(Op::I64_EXTEND_I32_S, line);
    chunks[current].emit_i64_const(32, line);
    chunks[current].emit_op(Op::I64_SHL, line);
    get(&mut chunks[current], rng, line);
    emit_next_bits(chunks, current, 32, line);
    chunks[current].emit_op(Op::I64_EXTEND_I32_S, line);
    chunks[current].emit_op(Op::I64_ADD, line);
}

pub fn emit_next_long(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 2 {
        // `nextLong(bound)` — bound <= 0 throws IllegalArgumentException.
        let bound = chunks[current].alloc_scratch(1);
        let rng = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], bound, line);
        set(&mut chunks[current], rng, line);
        get(&mut chunks[current], bound, line);
        chunks[current].emit_i64_const(0, line);
        chunks[current].emit_op(Op::I64_LE_S, line);
        chunks[current].emit_if_value(line);
        emit_jvm_exception_throw(chunks, current, "IllegalArgumentException", line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], rng, line);
        emit_next_long_core(chunks, current, line);
        get(&mut chunks[current], bound, line);
        chunks[current].emit_op(Op::I64_REM_U, line);
        chunks[current].emit_end(line);
        return;
    }
    if argc == 3 {
        // `nextLong(origin, bound)` — origin >= bound throws.
        let bound = chunks[current].alloc_scratch(1);
        let origin = chunks[current].alloc_scratch(1);
        let rng = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], bound, line);
        set(&mut chunks[current], origin, line);
        set(&mut chunks[current], rng, line);
        get(&mut chunks[current], bound, line);
        get(&mut chunks[current], origin, line);
        chunks[current].emit_op(Op::I64_LE_S, line);
        chunks[current].emit_if_value(line);
        emit_jvm_exception_throw(chunks, current, "IllegalArgumentException", line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], rng, line);
        emit_next_long_core(chunks, current, line);
        get(&mut chunks[current], bound, line);
        get(&mut chunks[current], origin, line);
        chunks[current].emit_op(Op::I64_SUB, line);
        chunks[current].emit_op(Op::I64_REM_U, line);
        get(&mut chunks[current], origin, line);
        chunks[current].emit_op(Op::I64_ADD, line);
        chunks[current].emit_end(line);
        return;
    }
    emit_next_long_core(chunks, current, line);
}

pub fn emit_next_bool(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_next_bits(chunks, current, 1, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_NE, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_next_float(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_next_bits(chunks, current, 24, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    chunks[current].emit_f64_const((1u64 << 24) as f64, line);
    chunks[current].emit_op(Op::F64_DIV, line);
}

fn emit_next_double_core(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_next_bits(chunks, current, 31, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    chunks[current].emit_f64_const((1u64 << 31) as f64, line);
    chunks[current].emit_op(Op::F64_DIV, line);
}

pub fn emit_next_double(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 2 {
        // `nextDouble(bound)` — bound <= 0.0 throws IllegalArgumentException.
        let bound = chunks[current].alloc_scratch(1);
        let rng = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], bound, line);
        set(&mut chunks[current], rng, line);
        get(&mut chunks[current], bound, line);
        chunks[current].emit_f64_const(0.0, line);
        chunks[current].emit_op(Op::F64_LE, line);
        chunks[current].emit_if_value(line);
        emit_jvm_exception_throw(chunks, current, "IllegalArgumentException", line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], rng, line);
        emit_next_double_core(chunks, current, line);
        get(&mut chunks[current], bound, line);
        chunks[current].emit_op(Op::F64_MUL, line);
        chunks[current].emit_end(line);
        return;
    }
    if argc == 3 {
        // `nextDouble(origin, bound)` — origin >= bound throws.
        let bound = chunks[current].alloc_scratch(1);
        let origin = chunks[current].alloc_scratch(1);
        let rng = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], bound, line);
        set(&mut chunks[current], origin, line);
        set(&mut chunks[current], rng, line);
        get(&mut chunks[current], bound, line);
        get(&mut chunks[current], origin, line);
        chunks[current].emit_op(Op::F64_LE, line);
        chunks[current].emit_if_value(line);
        emit_jvm_exception_throw(chunks, current, "IllegalArgumentException", line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], rng, line);
        emit_next_double_core(chunks, current, line);
        get(&mut chunks[current], bound, line);
        get(&mut chunks[current], origin, line);
        chunks[current].emit_op(Op::F64_SUB, line);
        chunks[current].emit_op(Op::F64_MUL, line);
        get(&mut chunks[current], origin, line);
        chunks[current].emit_op(Op::F64_ADD, line);
        chunks[current].emit_end(line);
        return;
    }
    emit_next_double_core(chunks, current, line);
}

/// Fill `arr[0..len]` with signed bytes (-128..127) drawn from `rng`.
/// `len` must hold an i32.
fn emit_fill_signed_bytes(
    chunks: &mut [Chunk],
    current: usize,
    rng: u16,
    arr: u16,
    len: u16,
    line: u32,
) {
    let i = chunks[current].alloc_scratch(1);
    let byte = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], i, line);
    let block = chunks[current].emit_block(line);
    let (loop_pos, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], rng, line);
    emit_next_bits(chunks, current, 32, line);
    chunks[current].emit_i32_const(255, line);
    chunks[current].emit_op(Op::I32_AND, line);
    set(&mut chunks[current], byte, line);
    get(&mut chunks[current], arr, line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], byte, line);
    chunks[current].emit_i32_const(128, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], byte, line);
    chunks[current].emit_i32_const(256, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], byte, line);
    chunks[current].emit_end(line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_pos);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
}

pub fn emit_next_bytes(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    // Stack: `[rng, arg]`. Two signatures share this leaf:
    //   JDK  `void nextBytes(byte[] bytes)` — fills `bytes` in place,
    //   Kotlin `nextBytes(size: Int): ByteArray` — builds a fresh array.
    // The argument's shape (array vs number) picks the form at runtime.
    let arg = chunks[current].alloc_scratch(1);
    let rng = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], arg, line);
    set(&mut chunks[current], rng, line);
    let to_i32 = chunks[current].add_import("wasm:js-number", "toI32");
    get(&mut chunks[current], arg, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    // JDK form: fill the caller's array, evaluate to null (void).
    get(&mut chunks[current], arg, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_call(to_i32, 1, line);
    set(&mut chunks[current], len, line);
    emit_fill_signed_bytes(chunks, current, rng, arg, len, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    // Kotlin form: build and return a fresh array of `arg` bytes.
    get(&mut chunks[current], arg, line);
    chunks[current].emit_call(to_i32, 1, line);
    set(&mut chunks[current], len, line);
    get(&mut chunks[current], len, line);
    collections::emit_new_with_length(chunks, current, line);
    set(&mut chunks[current], out, line);
    emit_fill_signed_bytes(chunks, current, rng, out, len, line);
    get(&mut chunks[current], out, line);
    chunks[current].emit_end(line);
}

pub fn emit_split(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_i64_const(0x1234ABCD, line);
    collections::emit_array_new(chunks, current, 1, line);
}

pub fn emit_ints(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_stream(chunks, current, argc, line, 1);
}

pub fn emit_longs(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_stream(chunks, current, argc, line, 1);
}

pub fn emit_doubles(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_stream(chunks, current, argc, line, 0);
}

fn emit_stream(chunks: &mut [Chunk], current: usize, argc: u8, line: u32, value: i32) {
    if argc == 2 {
        let count_slot = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], count_slot, line);
        chunks[current].emit_op(Op::DROP, line);
        get(&mut chunks[current], count_slot, line);
        emit_repeated_array(chunks, current, value, line);
    } else if argc == 4 {
        let count_slot = chunks[current].alloc_scratch(1);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op(Op::DROP, line);
        set(&mut chunks[current], count_slot, line);
        chunks[current].emit_op(Op::DROP, line);
        get(&mut chunks[current], count_slot, line);
        emit_repeated_array(chunks, current, value, line);
    } else {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        for _ in 0..64 {
            chunks[current].emit_i32_const(value, line);
        }
        collections::emit_array_new(chunks, current, 64, line);
    }
}

fn emit_repeated_array(chunks: &mut [Chunk], current: usize, value: i32, line: u32) {
    let count = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], count, line);
    collections::emit_array_new(chunks, current, 0, line);
    set(&mut chunks[current], out, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], i, line);
    let block = chunks[current].emit_block(line);
    let (loop_pos, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], count, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], out, line);
    chunks[current].emit_i32_const(value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_pos);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    get(&mut chunks[current], out, line);
}

fn emit_scramble_seed(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::I64_EXTEND_I32_S, line);
    chunks[current].emit_i64_const(MULT, line);
    chunks[current].emit_op(Op::I64_XOR, line);
    chunks[current].emit_i64_const(MASK, line);
    chunks[current].emit_op(Op::I64_AND, line);
}

fn emit_next_bits(chunks: &mut [Chunk], current: usize, bits: i64, line: u32) {
    let rng = chunks[current].alloc_scratch(1);
    let state = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rng, line);
    get(&mut chunks[current], rng, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], state, line);
    get(&mut chunks[current], state, line);
    chunks[current].emit_i64_const(MULT, line);
    chunks[current].emit_op(Op::I64_MUL, line);
    chunks[current].emit_i64_const(ADD, line);
    chunks[current].emit_op(Op::I64_ADD, line);
    chunks[current].emit_i64_const(MASK, line);
    chunks[current].emit_op(Op::I64_AND, line);
    set(&mut chunks[current], state, line);
    get(&mut chunks[current], rng, line);
    chunks[current].emit_i32_const(0, line);
    get(&mut chunks[current], state, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], state, line);
    chunks[current].emit_i64_const(48 - bits, line);
    chunks[current].emit_op(Op::I64_SHR_U, line);
    chunks[current].emit_op(Op::I32_WRAP_I64, line);
}
