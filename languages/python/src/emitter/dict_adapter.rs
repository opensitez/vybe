use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use crate::emitter::adapter_util::{lget, stash_args};
use crate::emitter::collections_adapter::{emit_copy, emit_update};
use vybe_compiler::primitives::collections;

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

pub fn emit_dict_subclass_init(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 2 {
        return;
    }
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let src = base + 1;
    emit_slot_is_map(chunks, current, src, line);
    chunks[current].emit_if(line);
    emit_map_source_into_object(chunks, current, recv, src, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], recv, line);
    lget(&mut chunks[current], src, line);
    emit_update(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    lget(&mut chunks[current], recv, line);
}

fn emit_slot_is_map(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    lget(&mut chunks[current], slot, line);
    let tag = chunks[current].add_import("ecma:object", "toStringTag");
    chunks[current].emit_call(tag, 1, line);
    chunks[current].emit_string_const("[object Map]", line);
    let eq = chunks[current].add_import("wasm:js-string", "equals");
    chunks[current].emit_call(eq, 2, line);
}

fn emit_map_source_into_object(
    chunks: &mut [Chunk],
    current: usize,
    recv: u16,
    src: u16,
    line: u32,
) {
    let keys = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], src, line);
    let keys_fn = chunks[current].add_import("ecma:map", "keys");
    chunks[current].emit_call(keys_fn, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);
    lget(&mut chunks[current], keys, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    lget(&mut chunks[current], i, line);
    lget(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    lget(&mut chunks[current], keys, line);
    lget(&mut chunks[current], i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);

    lget(&mut chunks[current], src, line);
    lget(&mut chunks[current], key, line);
    let get_fn = chunks[current].add_import("ecma:map", "get");
    chunks[current].emit_call(get_fn, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);

    lget(&mut chunks[current], recv, line);
    lget(&mut chunks[current], key, line);
    lget(&mut chunks[current], value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
}
