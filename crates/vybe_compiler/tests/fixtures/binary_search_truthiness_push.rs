use vybe_compiler::primitives::collections::{emit_array_new, emit_get, emit_len, emit_push};
use vybe_runtime::{Chunk, Op};
fn alloc_local(c: &mut Chunk) -> u16 {
    c.alloc_scratch(1)
}
fn lget(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_GET, s, l);
}
fn lset(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_SET, s, l);
}
pub fn emit_binary_search_pair(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_binary_search_pair_impl(chunks, current, line, false);
}

/// Binary search using comparator(array_value, target). Returns [index, found].
/// Stack: [array, target, comparator] -> [pair_array].
pub fn emit_binary_search_func_pair(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_binary_search_pair_impl(chunks, current, line, true);
}

fn emit_binary_search_pair_impl(chunks: &mut [Chunk], current: usize, line: u32, has_cmp: bool) {
    let arr_slot = alloc_local(&mut chunks[current]);
    let target_slot = alloc_local(&mut chunks[current]);
    let cmp_slot = alloc_local(&mut chunks[current]);
    let len_slot = alloc_local(&mut chunks[current]);
    let lo_slot = alloc_local(&mut chunks[current]);
    let hi_slot = alloc_local(&mut chunks[current]);
    let mid_slot = alloc_local(&mut chunks[current]);
    let found_slot = alloc_local(&mut chunks[current]);
    let out_slot = alloc_local(&mut chunks[current]);

    if has_cmp {
        lset(&mut chunks[current], cmp_slot, line);
    }
    lset(&mut chunks[current], target_slot, line);
    lset(&mut chunks[current], arr_slot, line);

    lget(&mut chunks[current], arr_slot, line);
    emit_len(chunks, current, line);
    lset(&mut chunks[current], len_slot, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], lo_slot, line);
    lget(&mut chunks[current], len_slot, line);
    lset(&mut chunks[current], hi_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    lget(&mut chunks[current], lo_slot, line);
    lget(&mut chunks[current], hi_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    lget(&mut chunks[current], lo_slot, line);
    lget(&mut chunks[current], hi_slot, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SHR_U, line);
    lset(&mut chunks[current], mid_slot, line);

    if has_cmp {
        let __abi = vybe_compiler::primitives::class_context::module_receiver_abi(chunks);
        lget(&mut chunks[current], cmp_slot, line);
        let __recv = vybe_compiler::primitives::callable::emit_callback_receiver(
            &mut chunks[current],
            __abi,
            line,
        );
        lget(&mut chunks[current], arr_slot, line);
        lget(&mut chunks[current], mid_slot, line);
        emit_get(chunks, current, line);
        lget(&mut chunks[current], target_slot, line);
        vybe_compiler::primitives::callable::emit_direct_invoke_chunk(
            &mut chunks[current],
            2 + __recv,
            line,
        );
        chunks[current].emit_i32_const(0, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    } else {
        lget(&mut chunks[current], arr_slot, line);
        lget(&mut chunks[current], mid_slot, line);
        emit_get(chunks, current, line);
        lget(&mut chunks[current], target_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], mid_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], lo_slot, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], mid_slot, line);
    lset(&mut chunks[current], hi_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    chunks[current].emit_bool_const(false, line);
    lset(&mut chunks[current], found_slot, line);
    lget(&mut chunks[current], lo_slot, line);
    lget(&mut chunks[current], len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    if has_cmp {
        let __abi = vybe_compiler::primitives::class_context::module_receiver_abi(chunks);
        lget(&mut chunks[current], cmp_slot, line);
        let __recv = vybe_compiler::primitives::callable::emit_callback_receiver(
            &mut chunks[current],
            __abi,
            line,
        );
        lget(&mut chunks[current], arr_slot, line);
        lget(&mut chunks[current], lo_slot, line);
        emit_get(chunks, current, line);
        lget(&mut chunks[current], target_slot, line);
        vybe_compiler::primitives::callable::emit_direct_invoke_chunk(
            &mut chunks[current],
            2 + __recv,
            line,
        );
        chunks[current].emit_i32_const(0, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    } else {
        lget(&mut chunks[current], arr_slot, line);
        lget(&mut chunks[current], lo_slot, line);
        emit_get(chunks, current, line);
        lget(&mut chunks[current], target_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(true, line);
    lset(&mut chunks[current], found_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_TEE, out_slot, line);
    lget(&mut chunks[current], lo_slot, line);
    emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], out_slot, line);
    lget(&mut chunks[current], found_slot, line);
    emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], out_slot, line);
}
