use vybe_compiler::primitives::{callable, class_context, collections, dict, ops, reflection};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "go.maps_keys" if argc == 1 => emit_maps_keys(chunks, current, line),
        "go.maps_values" | "go.slices_values" if argc == 1 => {
            emit_maps_values(chunks, current, line)
        }
        "go.slices_delete_func" if argc == 2 => emit_slices_delete_func(chunks, current, line),
        "go.slices_compact_func" if argc == 2 => emit_slices_compact_func(chunks, current, line),
        "go.maps_equal" if argc == 2 => emit_maps_equal(chunks, current, false, line),
        "go.maps_equal_func" if argc == 3 => emit_maps_equal(chunks, current, true, line),
        _ => return false,
    }
    true
}

fn emit_slices_delete_func(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let slice = base;
    let pred = base + 1;
    let result = base + 2;
    let index = base + 3;
    let value = base + 4;

    lset(chunks, current, pred, line);
    lset(chunks, current, slice, line);

    collections::emit_array_new(chunks, current, 0, line);
    lset(chunks, current, result, line);

    let state =
        vybe_compiler::primitives::loops::emit_for_in_start(chunks, current, slice, index, line);
    lset(chunks, current, value, line);

    lget(chunks, current, pred, line);
    let abi = class_context::module_receiver_abi(chunks);
    let recv = callable::emit_callback_receiver(&mut chunks[current], abi, line);
    lget(chunks, current, value, line);
    callable::emit_direct_invoke_chunk(&mut chunks[current], 1 + recv, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    lget(chunks, current, result, line);
    lget(chunks, current, value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    vybe_compiler::primitives::loops::emit_for_in_end(chunks, current, index, state, line);
    lget(chunks, current, result, line);
}

fn emit_slices_compact_func(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let slice = base;
    let eq = base + 1;
    let result = base + 2;
    let index = base + 3;
    let value = base + 4;

    lset(chunks, current, eq, line);
    lset(chunks, current, slice, line);

    collections::emit_array_new(chunks, current, 0, line);
    lset(chunks, current, result, line);

    let state =
        vybe_compiler::primitives::loops::emit_for_in_start(chunks, current, slice, index, line);
    lset(chunks, current, value, line);

    lget(chunks, current, index, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    lget(chunks, current, result, line);
    lget(chunks, current, value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    lget(chunks, current, eq, line);
    let abi = class_context::module_receiver_abi(chunks);
    let recv = callable::emit_callback_receiver(&mut chunks[current], abi, line);
    lget(chunks, current, value, line);
    lget(chunks, current, slice, line);
    lget(chunks, current, index, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
    callable::emit_direct_invoke_chunk(&mut chunks[current], 2 + recv, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    lget(chunks, current, result, line);
    lget(chunks, current, value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    vybe_compiler::primitives::loops::emit_for_in_end(chunks, current, index, state, line);
    lget(chunks, current, result, line);
}

fn emit_maps_equal(chunks: &mut Vec<Chunk>, current: usize, with_func: bool, line: u32) {
    let base = chunks[current].alloc_scratch(8);
    let left = base;
    let right = base + 1;
    let eq_fn = base + 2;
    let keys = base + 3;
    let len = base + 4;
    let index = base + 5;
    let key = base + 6;
    let right_value = base + 7;
    let result = chunks[current].alloc_scratch(1);

    if with_func {
        lset(chunks, current, eq_fn, line);
    }
    lset(chunks, current, right, line);
    lset(chunks, current, left, line);

    chunks[current].emit_bool_const(true, line);
    lset(chunks, current, result, line);

    lget(chunks, current, left, line);
    emit_map_len_nil_zero(chunks, current, line);
    lget(chunks, current, right, line);
    emit_map_len_nil_zero(chunks, current, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    lget(chunks, current, left, line);
    emit_map_keys_nil_empty(chunks, current, line);
    lset(chunks, current, keys, line);

    lget(chunks, current, keys, line);
    collections::emit_len(chunks, current, line);
    lset(chunks, current, len, line);

    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);

    lget(chunks, current, index, line);
    lget(chunks, current, len, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    lget(chunks, current, keys, line);
    lget(chunks, current, index, line);
    collections::emit_get(chunks, current, line);
    lset(chunks, current, key, line);

    lget(chunks, current, right, line);
    lget(chunks, current, key, line);
    reflection::emit_has_own(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    lset(chunks, current, result, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    lget(chunks, current, right, line);
    lget(chunks, current, key, line);
    collections::emit_get(chunks, current, line);
    lset(chunks, current, right_value, line);

    if with_func {
        lget(chunks, current, eq_fn, line);
        let abi = class_context::module_receiver_abi(chunks);
        let recv = callable::emit_callback_receiver(&mut chunks[current], abi, line);
        lget(chunks, current, left, line);
        lget(chunks, current, key, line);
        collections::emit_get(chunks, current, line);
        lget(chunks, current, right_value, line);
        callable::emit_direct_invoke_chunk(&mut chunks[current], 2 + recv, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
    } else {
        lget(chunks, current, left, line);
        lget(chunks, current, key, line);
        collections::emit_get(chunks, current, line);
        lget(chunks, current, right_value, line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
    }
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    lset(chunks, current, result, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    lget(chunks, current, index, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, index, line);
    chunks[current].emit_br(0, line);

    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);

    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    lset(chunks, current, result, line);
    chunks[current].emit_end(line);

    lget(chunks, current, result, line);
}

fn emit_maps_keys(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    lset(chunks, current, value, line);
    lget(chunks, current, value, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    let keys = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    emit_filtered_keys_from_local(chunks, current, value, keys, result, index, key, line);
    chunks[current].emit_end(line);
}

fn emit_maps_values(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    lset(chunks, current, value, line);
    lget(chunks, current, value, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    let keys = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    emit_filtered_values_from_local(chunks, current, value, keys, result, index, key, line);
    chunks[current].emit_end(line);
}

fn emit_filtered_keys_from_local(
    chunks: &mut [Chunk],
    current: usize,
    map_slot: u16,
    keys_slot: u16,
    result_slot: u16,
    idx_slot: u16,
    key_slot: u16,
    line: u32,
) {
    lget(chunks, current, map_slot, line);
    dict::emit_keys(chunks, current, line);
    lset(chunks, current, keys_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    lset(chunks, current, result_slot, line);

    let state = vybe_compiler::primitives::loops::emit_for_in_start(
        chunks, current, keys_slot, idx_slot, line,
    );
    lset(chunks, current, key_slot, line);

    lget(chunks, current, map_slot, line);
    lget(chunks, current, key_slot, line);
    reflection::emit_has_own(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(chunks, current, result_slot, line);
    lget(chunks, current, key_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    vybe_compiler::primitives::loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    lget(chunks, current, result_slot, line);
}

fn emit_filtered_values_from_local(
    chunks: &mut [Chunk],
    current: usize,
    map_slot: u16,
    keys_slot: u16,
    result_slot: u16,
    idx_slot: u16,
    key_slot: u16,
    line: u32,
) {
    lget(chunks, current, map_slot, line);
    dict::emit_keys(chunks, current, line);
    lset(chunks, current, keys_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    lset(chunks, current, result_slot, line);

    let state = vybe_compiler::primitives::loops::emit_for_in_start(
        chunks, current, keys_slot, idx_slot, line,
    );
    lset(chunks, current, key_slot, line);

    lget(chunks, current, map_slot, line);
    lget(chunks, current, key_slot, line);
    reflection::emit_has_own(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(chunks, current, result_slot, line);
    lget(chunks, current, map_slot, line);
    lget(chunks, current, key_slot, line);
    collections::emit_get(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    vybe_compiler::primitives::loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    lget(chunks, current, result_slot, line);
}

fn emit_map_len_nil_zero(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    lset(chunks, current, value, line);
    lget(chunks, current, value, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_else(line);
    lget(chunks, current, value, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_map_keys_nil_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    lset(chunks, current, value, line);
    lget(chunks, current, value, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_else(line);
    lget(chunks, current, value, line);
    dict::emit_keys(chunks, current, line);
    chunks[current].emit_end(line);
}

fn lget(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}
