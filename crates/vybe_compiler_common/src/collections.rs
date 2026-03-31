//! Collection operations — arrays, sets, sorting, range.
//!
//! Array ops use direct WASM GC opcodes where possible.
//! Higher-level ops (range, sorted, enumerate, zip) use host imports.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

// ── Direct WASM GC opcodes (no host call) ───────────────────

/// Create array from N stack values. Stack: [v1, v2, ..., vN] → [array]
pub fn emit_array_new(chunk: &mut Chunk, count: u16, line: u32) {
    chunk.emit_op_u16(Op::array_new, count, line);
}

/// Array length. Stack: [array] → [i32]
pub fn emit_len(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_length, line);
}

/// Array push. Stack: [array, value] → [array]
pub fn emit_push(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_push, line);
}

/// Array pop. Stack: [array] → [value]
pub fn emit_pop(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_pop, line);
}

/// Array get. Stack: [array, index] → [value]
pub fn emit_get(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_get, line);
}

/// Array set. Stack: [array, index, value] → [value]
pub fn emit_set(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_set, line);
}

/// Array slice. Stack: [array, start, end] → [array]
pub fn emit_slice(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_slice, line);
}

/// Array join. Stack: [array, delimiter] → [string]
pub fn emit_join(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_join, line);
}

/// Array reverse (in-place). Stack: [array] → [array]
pub fn emit_reverse(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_reverse, line);
}

/// Array contains. Stack: [array, value] → [bool]
pub fn emit_contains(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_contains, line);
}

/// Array indexOf. Stack: [array, value] → [i32]
pub fn emit_index_of(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_index_of, line);
}

/// Array concat. Stack: [array, array] → [array]
pub fn emit_concat(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_concat, line);
}

/// Array shift (remove first). Stack: [array] → [value]
pub fn emit_shift(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_shift, line);
}

// ── Host imports (higher-level operations) ──────────────────

/// range(stop) or range(start, stop) or range(start, stop, step).
/// Stack: [args...] → [array]
pub fn emit_range(chunk: &mut Chunk, arg_count: u8, line: u32) {
    let idx = chunk.add_import("vybe:array", "range");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(arg_count, line);
}

/// sorted(iterable). Stack: [array] → [sorted_array]
pub fn emit_sorted(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:array", "sorted");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

/// reversed(iterable). Stack: [array] → [reversed_array]
pub fn emit_reversed(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:array", "reversed");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

/// enumerate(iterable). Stack: [array] → [array of [i, val] pairs]
pub fn emit_enumerate(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:array", "enumerate");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

/// zip(a, b). Stack: [array, array] → [array of [a_i, b_i] pairs]
pub fn emit_zip(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:array", "zip");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(2, line);
}

/// sum(iterable). Stack: [array] → [number]
pub fn emit_sum(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("vybe:array", "sum");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

/// min(args...). Stack: [args...] → [min_value]
pub fn emit_min(chunk: &mut Chunk, arg_count: u8, line: u32) {
    let idx = chunk.add_import("vybe:array", "pymin");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(arg_count, line);
}

/// max(args...). Stack: [args...] → [max_value]
pub fn emit_max(chunk: &mut Chunk, arg_count: u8, line: u32) {
    let idx = chunk.add_import("vybe:array", "pymax");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(arg_count, line);
}
