use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use crate::emitter::adapter_util::{lget, stash_args};
use crate::emitter::collections_adapter::{emit_copy, emit_update};

pub fn emit_dict_ior(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 2 {
        return;
    }
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let src = base + 1;
    lget(&mut chunks[current], recv, line);
    lget(&mut chunks[current], src, line);
    emit_update(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], recv, line);
}

pub fn emit_dict_or(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 2 {
        return;
    }
    let base = stash_args(chunks, current, 2, line);
    let left = base;
    let right = base + 1;
    let out = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], left, line);
    emit_copy(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    lget(&mut chunks[current], out, line);
    lget(&mut chunks[current], right, line);
    emit_update(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], out, line);
}

fn emit_dict_update_if_present(
    chunks: &mut [Chunk],
    current: usize,
    recv: u16,
    src: u16,
    line: u32,
) {
    lget(&mut chunks[current], src, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], recv, line);
    lget(&mut chunks[current], src, line);
    emit_update(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

pub fn emit_dict_update(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 3 {
        return;
    }
    let base = stash_args(chunks, current, 3, line);
    let recv = base;
    emit_dict_update_if_present(chunks, current, recv, base + 1, line);
    emit_dict_update_if_present(chunks, current, recv, base + 2, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}
