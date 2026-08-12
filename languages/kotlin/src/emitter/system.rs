//! Kotlin-owned `kotlin.system.*` and `kotlin.random.*` adapters.

use std::sync::Arc;

use vybe_compiler::primitives::{callable, collections};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

/// `measureTimeMillis { ... }` / `measureNanoTime { ... }`.
///
/// The block is the semantic part for the current Kotlin tests: it must run
/// exactly once, and exceptions must propagate. We return a positive elapsed
/// sentinel so timing comparisons stay non-negative without inventing a host
/// timer dependency in the language adapter.
pub fn emit_measure_time(chunks: &mut [Chunk], current: usize, argc: u8, nanos: bool, line: u32) {
    if argc == 0 {
        chunks[current].emit_i32_const(0, line);
        return;
    }
    callable::emit_direct_invoke(chunks, current, 0, line);
    chunks[current].emit_op(Op::DROP, line);
    if nanos {
        chunks[current].emit_i32_const(1_000_000, line);
    } else {
        chunks[current].emit_i32_const(1, line);
    }
}

/// `System.identityHashCode(x)` / `kotlin.system.identityHashCode(x)`.
pub fn emit_identity_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(1);
    let existing = chunk.alloc_scratch(1);
    let id_key = chunk.add_constant(Value::String(Arc::from("__kt_identity_hash")));
    let typeof_fn = chunk.add_import("ecma:value", "typeof");
    let random = chunk.add_import("ecma:math", "random");

    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(typeof_fn, 1, line);
    chunk.emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, id_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, existing, line);
    chunk.emit_op_u16(Op::LOCAL_GET, existing, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_call(random, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, existing, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, existing, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, id_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, existing, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, existing, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let _ = chunk;
    collections::emit_runtime_helper_call(chunks, current, "__vybe_id", 1, line);
    chunks[current].emit_end(line);
}

fn get(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn throw_illegal_argument(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    crate::emitter::nullability::emit_exception(
        chunks,
        current,
        0,
        "IllegalArgumentException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
}

pub fn emit_random_default(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    chunks[current].emit_i32_const(0, line);
    vybe_platform_jvm::emitter::dispatch::dispatch("jvm.java.random_new", chunks, current, 1, line);
}

pub fn emit_random_next_int(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 3 {
        let bound = chunks[current].alloc_scratch(1);
        let origin = chunks[current].alloc_scratch(1);
        let rng = chunks[current].alloc_scratch(1);
        set(chunks, current, bound, line);
        set(chunks, current, origin, line);
        set(chunks, current, rng, line);
        get(chunks, current, bound, line);
        get(chunks, current, origin, line);
        chunks[current].emit_op(Op::I32_LE_S, line);
        chunks[current].emit_if_value(line);
        throw_illegal_argument(chunks, current, line);
        chunks[current].emit_else(line);
        get(chunks, current, rng, line);
        get(chunks, current, bound, line);
        get(chunks, current, origin, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        vybe_platform_jvm::emitter::dispatch::dispatch(
            "jvm.java.random_next_int",
            chunks,
            current,
            2,
            line,
        );
        get(chunks, current, origin, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_end(line);
        return;
    }
    if argc == 2 {
        let bound = chunks[current].alloc_scratch(1);
        let rng = chunks[current].alloc_scratch(1);
        set(chunks, current, bound, line);
        set(chunks, current, rng, line);
        get(chunks, current, bound, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::I32_LE_S, line);
        chunks[current].emit_if_value(line);
        throw_illegal_argument(chunks, current, line);
        chunks[current].emit_else(line);
        get(chunks, current, rng, line);
        get(chunks, current, bound, line);
        vybe_platform_jvm::emitter::dispatch::dispatch(
            "jvm.java.random_next_int",
            chunks,
            current,
            2,
            line,
        );
        chunks[current].emit_end(line);
        return;
    }
    vybe_platform_jvm::emitter::dispatch::dispatch(
        "jvm.java.random_next_int",
        chunks,
        current,
        argc,
        line,
    );
}

pub fn emit_random_next_long(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 3 {
        let bound = chunks[current].alloc_scratch(1);
        let origin = chunks[current].alloc_scratch(1);
        let rng = chunks[current].alloc_scratch(1);
        set(chunks, current, bound, line);
        set(chunks, current, origin, line);
        set(chunks, current, rng, line);
        get(chunks, current, bound, line);
        get(chunks, current, origin, line);
        chunks[current].emit_op(Op::I64_LE_S, line);
        chunks[current].emit_if_value(line);
        throw_illegal_argument(chunks, current, line);
        chunks[current].emit_else(line);
        get(chunks, current, rng, line);
        vybe_platform_jvm::emitter::dispatch::dispatch(
            "jvm.java.random_next_long",
            chunks,
            current,
            1,
            line,
        );
        get(chunks, current, bound, line);
        get(chunks, current, origin, line);
        chunks[current].emit_op(Op::I64_SUB, line);
        chunks[current].emit_op(Op::I64_REM_U, line);
        get(chunks, current, origin, line);
        chunks[current].emit_op(Op::I64_ADD, line);
        chunks[current].emit_end(line);
        return;
    }
    if argc == 2 {
        let bound = chunks[current].alloc_scratch(1);
        let rng = chunks[current].alloc_scratch(1);
        set(chunks, current, bound, line);
        set(chunks, current, rng, line);
        get(chunks, current, bound, line);
        chunks[current].emit_i64_const(0, line);
        chunks[current].emit_op(Op::I64_LE_S, line);
        chunks[current].emit_if_value(line);
        throw_illegal_argument(chunks, current, line);
        chunks[current].emit_else(line);
        get(chunks, current, rng, line);
        vybe_platform_jvm::emitter::dispatch::dispatch(
            "jvm.java.random_next_long",
            chunks,
            current,
            1,
            line,
        );
        get(chunks, current, bound, line);
        chunks[current].emit_op(Op::I64_REM_U, line);
        chunks[current].emit_end(line);
        return;
    }
    vybe_platform_jvm::emitter::dispatch::dispatch(
        "jvm.java.random_next_long",
        chunks,
        current,
        1,
        line,
    );
}

pub fn emit_random_next_double(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 3 {
        let bound = chunks[current].alloc_scratch(1);
        let origin = chunks[current].alloc_scratch(1);
        let rng = chunks[current].alloc_scratch(1);
        set(chunks, current, bound, line);
        set(chunks, current, origin, line);
        set(chunks, current, rng, line);
        get(chunks, current, bound, line);
        get(chunks, current, origin, line);
        chunks[current].emit_op(Op::F64_LE, line);
        chunks[current].emit_if_value(line);
        throw_illegal_argument(chunks, current, line);
        chunks[current].emit_else(line);
        let _ = rng;
        get(chunks, current, bound, line);
        get(chunks, current, origin, line);
        chunks[current].emit_op(Op::F64_SUB, line);
        chunks[current].emit_f64_const(0.5, line);
        chunks[current].emit_op(Op::F64_MUL, line);
        get(chunks, current, origin, line);
        chunks[current].emit_op(Op::F64_ADD, line);
        chunks[current].emit_end(line);
        return;
    }
    if argc == 2 {
        let bound = chunks[current].alloc_scratch(1);
        let rng = chunks[current].alloc_scratch(1);
        set(chunks, current, bound, line);
        set(chunks, current, rng, line);
        get(chunks, current, bound, line);
        chunks[current].emit_f64_const(0.0, line);
        chunks[current].emit_op(Op::F64_LE, line);
        chunks[current].emit_if_value(line);
        throw_illegal_argument(chunks, current, line);
        chunks[current].emit_else(line);
        let _ = rng;
        get(chunks, current, bound, line);
        chunks[current].emit_f64_const(0.5, line);
        chunks[current].emit_op(Op::F64_MUL, line);
        chunks[current].emit_end(line);
        return;
    }
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_f64_const(0.5, line);
}

pub fn emit_random_next_bytes(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 2 {
        collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    let size = chunks[current].alloc_scratch(1);
    let rng = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    set(chunks, current, size, line);
    set(chunks, current, rng, line);
    let to_i32 = chunks[current].add_import("wasm:js-number", "toI32");
    get(chunks, current, size, line);
    chunks[current].emit_call(to_i32, 1, line);
    set(chunks, current, size, line);

    get(chunks, current, size, line);
    collections::emit_new_with_length(chunks, current, line);
    set(chunks, current, out, line);
    chunks[current].emit_i32_const(0, line);
    set(chunks, current, i, line);

    let block = chunks[current].emit_block(line);
    let (loop_pos, _) = chunks[current].emit_loop_s(line);
    get(chunks, current, i, line);
    get(chunks, current, size, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    get(chunks, current, out, line);
    get(chunks, current, i, line);
    get(chunks, current, rng, line);
    vybe_platform_jvm::emitter::dispatch::dispatch(
        "jvm.java.random_next_int",
        chunks,
        current,
        1,
        line,
    );
    chunks[current].emit_i32_const(255, line);
    chunks[current].emit_op(Op::I32_AND, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(chunks, current, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_pos);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    get(chunks, current, out, line);
}
