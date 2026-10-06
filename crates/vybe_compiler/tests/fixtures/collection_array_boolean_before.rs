#![allow(dead_code)]
use vybe_compiler::primitives::collections::*;
use vybe_runtime::{Chunk, opcode::Op};
pub(crate) fn emit_import_call(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    let c = &mut chunks[current];
    c.emit_call(idx, argc, line);
}

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

pub fn emit_rank(chunks: &mut [Chunk], current: usize, line: u32) {
    fn emit_is_array(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        emit_import_call(chunks, current, "ecma:array", "isArray", 1, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    }

    fn emit_array_get_zero(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        chunks[current].emit_i32_const(0, line);
        emit_import_call(chunks, current, "ecma:array", "get", 2, line);
    }

    let value_slot = alloc_local(&mut chunks[current]);
    let first_slot = alloc_local(&mut chunks[current]);
    let second_slot = alloc_local(&mut chunks[current]);

    lset(&mut chunks[current], value_slot, line);
    emit_is_array(chunks, current, value_slot, line);
    chunks[current].emit_if_value(line);

    emit_array_get_zero(chunks, current, value_slot, line);
    lset(&mut chunks[current], first_slot, line);
    emit_is_array(chunks, current, first_slot, line);
    chunks[current].emit_if_value(line);

    emit_array_get_zero(chunks, current, first_slot, line);
    lset(&mut chunks[current], second_slot, line);
    emit_is_array(chunks, current, second_slot, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(3, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
}

pub fn emit_pop(chunks: &mut [Chunk], current: usize, line: u32) {
    let src = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, src, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_array, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    emit_import_call(chunks, current, "ecma:array", "pop", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_else(line);

    // Keyed: take the last insertion-ordered key, read it, drop it.
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    emit_import_call(chunks, current, "ecma:map", "keys", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    emit_import_call(chunks, current, "ecma:map", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    emit_import_call(chunks, current, "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_iter_natural(chunks: &mut [Chunk], current: usize, line: u32) {
    let slot = chunks[current].alloc_scratch(1);
    let tag_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);

    // Array → iterate elements directly.
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    {
        let idx = chunks[current].add_import("ecma:array", "isArray");
        chunks[current].emit_call(idx, 1, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_else(line);

    // String → characters (via the shared materialize path).
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    {
        let idx = chunks[current].add_import("wasm:js-string", "test");
        chunks[current].emit_call(idx, 1, line);
    }
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    emit_iter_for_of(chunks, current, line);
    chunks[current].emit_else(line);

    // tag = Object.prototype.toString(obj) → "[object Map]" / "[object Object]" / …
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    {
        let idx = chunks[current].add_import("ecma:object", "toStringTag");
        chunks[current].emit_call(idx, 1, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, tag_slot, line);

    // Map / plain Object → keys.
    chunks[current].emit_op_u16(Op::LOCAL_GET, tag_slot, line);
    chunks[current].emit_string_const("[object Map]", line);
    {
        let idx = chunks[current].add_import("wasm:js-string", "equals");
        chunks[current].emit_call(idx, 2, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, tag_slot, line);
    chunks[current].emit_string_const("[object Object]", line);
    {
        let idx = chunks[current].add_import("wasm:js-string", "equals");
        chunks[current].emit_call(idx, 2, line);
    }
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    emit_iter_keys(chunks, current, line);
    chunks[current].emit_else(line);
    // Set / anything else → values.
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    emit_iter_for_of(chunks, current, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_promote_empty_array_for_string_key(
    chunks: &mut [Chunk],
    current: usize,
    obj_slot: u16,
    key_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    {
        let idx = chunks[current].add_import("ecma:array", "isArray");
        chunks[current].emit_call(idx, 1, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    vybe_compiler::primitives::instructions::host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "test",
        1,
        line,
    );
    // js-string.test already returns the i32 condition required by IF.
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    // GC array length is an i32, so emptiness needs no dynamic equality.
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);

    emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_shift(chunks: &mut [Chunk], current: usize, line: u32) {
    // Polymorphic for the same reason as `emit_pop` — see the note there.
    // `ecma:array.shift` (§23.1.3.25) understands only a packed list; a keyed
    // collection loses its FIRST insertion-ordered entry instead.
    let src = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, src, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_array, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    emit_import_call(chunks, current, "ecma:array", "shift", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    emit_import_call(chunks, current, "ecma:map", "keys", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    emit_import_call(chunks, current, "ecma:map", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    emit_import_call(chunks, current, "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_clear_keyed(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = alloc_local(&mut chunks[current]);
    let keys_slot = alloc_local(&mut chunks[current]);
    let len_slot = alloc_local(&mut chunks[current]);
    let idx_slot = alloc_local(&mut chunks[current]);
    let key_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], value_slot, line);

    lget(&mut chunks[current], value_slot, line);
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_array, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);

    lget(&mut chunks[current], value_slot, line);
    emit_import_call(chunks, current, "ecma:object", "keys", 1, line);
    lset(&mut chunks[current], keys_slot, line);
    lget(&mut chunks[current], keys_slot, line);
    emit_len(chunks, current, line);
    lset(&mut chunks[current], len_slot, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], idx_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    lget(&mut chunks[current], idx_slot, line);
    lget(&mut chunks[current], len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    lget(&mut chunks[current], keys_slot, line);
    lget(&mut chunks[current], idx_slot, line);
    emit_get(chunks, current, line);
    lset(&mut chunks[current], key_slot, line);
    lget(&mut chunks[current], value_slot, line);
    lget(&mut chunks[current], key_slot, line);
    emit_import_call(chunks, current, "ecma:object", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], idx_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], idx_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}
