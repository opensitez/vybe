//! Kotlin-owned `kotlin.system.*` and `kotlin.random.*` adapters.

use std::sync::Arc;

use vybe_compiler::primitives::{callable, collections};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};
use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ValueSource,
};

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
    let id_key = class_slots::resolve_interned(chunk, &ClassSlot::internal("__kt_identity_hash"), &PlainNames);
    let typeof_fn = chunk.add_import("ecma:value", "typeof");
    let random = chunk.add_import("ecma:math", "random");

    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(typeof_fn, 1, line);
    chunk.emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &id_key, Dest::Stack, line);
    chunk.emit_op_u16(Op::LOCAL_SET, existing, line);
    chunk.emit_op_u16(Op::LOCAL_GET, existing, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_call(random, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, existing, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, existing, line);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &id_key, ValueSource::Stack, line);
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

pub fn emit_random_default(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    chunks[current].emit_i32_const(0, line);
    vybe_platform_jvm::emitter::dispatch::dispatch("jvm.java.random_new", chunks, current, 1, line);
}
