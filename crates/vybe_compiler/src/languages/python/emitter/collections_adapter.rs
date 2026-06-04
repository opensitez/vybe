use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

use crate::emitter::{collections, dict, strings};

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[0].add_import(module, name);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(argc, line);
}

fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].local_count;
    chunks[current].local_count += argc as u16;
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    base
}

pub fn emit_extend(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let src = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op(Op::DUP, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    collections::emit_insert_range(chunks, current, line);
}

pub fn emit_get(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let recv = base;
    let key = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    dict::emit_get(chunks, current, line);

    chunks[current].emit_op(Op::DUP, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::DROP, line);
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 2, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);
}

pub fn emit_index(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let recv = base;
    let needle = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op(Op::REF_IS_STRING, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle, line);
    strings::emit_index_of(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle, line);
    collections::emit_index_of(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let recv = base;
    let keys_key =
        chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__keys")));

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);
    collections::emit_remove_range(chunks, current, line);

    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:set", "clear", 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    dict::emit_method_clear_stack(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    call_import(chunks, current, "ecma:set", "add", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

fn emit_remove_impl(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let value = base + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    crate::emitter::ops::emit_dyn_ne(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_import(chunks, current, "ecma:array", "removeValue", 2, line);

    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_import(chunks, current, "ecma:set", "delete", 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_remove_impl(chunks, current, line);
}

pub fn emit_discard(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_remove_impl(chunks, current, line);
}

pub fn emit_copy(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let recv = base;
    let keys_key =
        chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__keys")));

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    crate::emitter::ops::emit_dyn_ne(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);
    collections::emit_slice(chunks, current, line);

    chunks[current].emit_else(line);
    call_import(chunks, current, "ecma:set", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:set", "union", 2, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    dict::emit_new(chunks, current, line);
    let out = chunks[current].local_count;
    chunks[current].local_count += 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:object", "assign", 2, line);
    chunks[current].emit_end(line);
}

pub fn emit_update(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    let recv = base;
    let src = base + 1;
    let keys_key =
        chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__keys")));

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    call_import(chunks, current, "ecma:set", "unionWith", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    call_import(chunks, current, "ecma:object", "assign", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::NULL, line);
}

fn emit_set_update_call(chunks: &mut [Chunk], current: usize, func: &str, line: u32) {
    let base = stash_args(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
    call_import(chunks, current, "ecma:set", func, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_intersection_update(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_set_update_call(chunks, current, "intersectWith", line);
}

pub fn emit_difference_update(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_set_update_call(chunks, current, "exceptWith", line);
}

pub fn emit_symmetric_difference_update(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_set_update_call(chunks, current, "symmetricExceptWith", line);
}

pub fn emit_pop(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(chunks, current, argc, line);
    let recv = base;
    let value_slot = chunks[current].local_count;
    chunks[current].local_count += 1;

    if argc == 1 {
        let index_slot = chunks[current].local_count;
        chunks[current].local_count += 1;

        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        collections::emit_len(chunks, current, line);
        chunks[current].emit_op(Op::I32_CONST_1, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);
        chunks[current].emit_op(Op::DROP, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
        chunks[current].emit_op(Op::DROP, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
        collections::emit_remove_at(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        return;
    } else {
        let keys_key = chunks[current]
            .add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__keys")));
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);

        let index = base + 1;
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
        chunks[current].emit_op(Op::DROP, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
        collections::emit_remove_at(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);

        chunks[current].emit_else(line);

        let key = base + 1;
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
        dict::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
        chunks[current].emit_op(Op::DROP, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        if argc >= 3 {
            chunks[current].emit_op_u16(Op::LOCAL_GET, base + 2, line);
        } else {
            chunks[current].emit_op(Op::NULL, line);
        }

        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
        dict::emit_method_delete(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    }
}

pub fn emit_length(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let recv = base;
    let keys_key =
        chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("__keys")));
    let size_key =
        chunks[current].add_constant(vybe_bytecode::Value::String(std::sync::Arc::from("size")));

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op(Op::REF_IS_STRING, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    strings::emit_length(&mut chunks[current], line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    chunks[current].emit_op(Op::DUP, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    crate::emitter::ops::emit_dyn_ne(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, size_key, line);
    chunks[current].emit_op(Op::DUP, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
}
