//! Previous full-sort selection retained as a fixed benchmark reference.
use vybe_compiler::primitives::{collections, instructions::core_wasm};
use vybe_runtime::{Chunk, Op};
fn emit_n_of(chunks: &mut [Chunk], current: usize, argc: u8, largest: bool, line: u32) {
    let n = chunks[current].alloc_scratch(1);
    let data = chunks[current].alloc_scratch(1);
    let key_fn = chunks[current].alloc_scratch(1);
    if argc >= 3 {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, key_fn, line);
        chunk.emit_op_u16(Op::LOCAL_SET, data, line);
        chunk.emit_op_u16(Op::LOCAL_SET, n, line);
        chunk.emit_op_u16(Op::LOCAL_GET, data, line);
        chunk.emit_op_u16(Op::LOCAL_GET, key_fn, line);
        let _ = chunk;
        collections::emit_sort_by_key_in_place(chunks, current, line);
    } else {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, data, line);
        chunk.emit_op_u16(Op::LOCAL_SET, n, line);
        chunk.emit_op_u16(Op::LOCAL_GET, data, line);
        let sorted = chunk.add_import("ecma:array", "toSorted");
        chunk.emit_call(sorted, 1, line);
    }

    let chunk = &mut chunks[current];
    if largest {
        let rev = chunk.add_import("ecma:array", "toReversed");
        chunk.emit_call(rev, 1, line);
    }
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    let slice = chunk.add_import("ecma:array", "slice");
    chunk.emit_call(slice, 3, line);
}

/// Return the `n` smallest items from data. Stack: `[n, data]` -> `[array]`.
pub fn emit_nsmallest(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_n_of(chunks, current, argc, false, line);
}

/// Return the `n` largest items from data. Stack: `[n, data]` -> `[array]`.
pub fn emit_nlargest(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_n_of(chunks, current, argc, true, line);
}
