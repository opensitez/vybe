//! Linear libc allocations with an eight-byte private header.

use vybe_compiler::primitives::pointers;
use vybe_runtime::{Chunk, Op};

fn get(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn set(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
}
fn tee(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_TEE, slot, line);
}

/// Caller has validated the requested byte count. Reuse its local and the
/// result local instead of repeating the public allocation range checks.
fn emit_valid_alloc(chunks: &mut [Chunk], current: usize, line: u32, size: u16, base: u16) {
    let c = &mut chunks[current];
    get(c, size, line);
    c.emit_i32_const(8, line);
    c.emit_op(Op::I32_ADD, line);
    pointers::emit_linear_alloc(chunks, current, line);
    let c = &mut chunks[current];
    tee(c, base, line);
    get(c, size, line);
    c.emit_op(Op::I32_STORE, line);
    get(c, base, line);
    c.emit_i32_const(8, line);
    c.emit_op(Op::I32_ADD, line);
}

/// [byte count] -> [payload address]. The header retains the requested size,
/// so realloc never reads beyond the original allocation.
pub fn emit_alloc(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let first = c.alloc_scratch(2);
    let (size, base) = (first, first + 1);
    tee(c, size, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_GE, line);
    get(c, size, line);
    c.emit_f64_const((u32::MAX - 15) as f64, line);
    c.emit_op(Op::F64_LE, line);
    c.emit_op(Op::I32_AND, line);
    // Both arms produce a wasm32 address, including the null failure result.
    c.emit_if_i32(line);
    emit_valid_alloc(chunks, current, line, size, base);
    let c = &mut chunks[current];
    c.emit_else(line);
    c.emit_i32_const(0, line);
    c.emit_end(line);
}

/// [old payload address, new byte count] -> [new payload address].
pub fn emit_realloc(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let first = c.alloc_scratch(4);
    let (old, size, new, count) = (first, first + 1, first + 2, first + 3);
    set(c, size, line);
    set(c, old, line);
    // Zero requests must not convert or dereference the old pointer.
    get(c, size, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_EQ, line);
    c.emit_if_value(line);
    c.emit_i32_const(0, line);
    c.emit_else(line);
    get(c, size, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_GE, line);
    get(c, size, line);
    c.emit_f64_const((u32::MAX - 15) as f64, line);
    c.emit_op(Op::F64_LE, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_if_value(line);
    get(c, old, line);
    c.emit_op(Op::I32_EQZ, line);
    // Null has no retained prefix. Select a native i32 count, then share
    // the same allocation path with growth instead of emitting it twice.
    c.emit_if_i32(line);
    c.emit_i32_const(0, line);
    c.emit_else(line);
    get(c, old, line);
    c.emit_i32_const(8, line);
    c.emit_op(Op::I32_SUB, line);
    c.emit_op(Op::I32_LOAD, line);
    c.emit_end(line);
    tee(c, count, line);
    c.emit_op(Op::F64_CONVERT_I32_U, line);
    get(c, size, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_if_value(line);
    // Shrinking needs neither allocation nor a payload copy. Retain the
    // logical requested size so a later growth cannot copy discarded bytes.
    get(c, old, line);
    c.emit_i32_const(8, line);
    c.emit_op(Op::I32_SUB, line);
    get(c, size, line);
    c.emit_op(Op::I32_STORE, line);
    get(c, old, line);
    c.emit_else(line);
    emit_valid_alloc(chunks, current, line, size, new);
    let c = &mut chunks[current];
    tee(c, new, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if_value(line);
    get(c, new, line);
    c.emit_else(line);
    // This branch grows the block, so the original requested size is the
    // copy length. It was loaded once, before selecting shrink versus grow.
    get(c, new, line);
    get(c, old, line);
    get(c, count, line);
    c.emit_op_idx(Op::MEMORY_COPY, 0u32, line);
    c.emit_leb_u32(0u32, line);
    get(c, new, line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_else(line);
    // Match the old rejected-request path's pointer conversion without
    // dereferencing it, allocating a block, or touching memory.
    get(c, old, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_op(Op::DROP, line);
    c.emit_i32_const(0, line);
    c.emit_end(line);
    c.emit_end(line);
}
