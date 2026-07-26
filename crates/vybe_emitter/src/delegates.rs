//! Shared delegate runtime emitters.
//!
//! Stack contracts:
//! - combine: [current, handler] -> [delegate]
//! - remove:  [current, handler] -> [delegate]

use crate::collections;
use crate::instructions::core_wasm;
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

fn emit_slot_is_nullish(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let idx = chunks[current].add_import("wasm:js-undefined", "test");
    chunks[current].emit_call(idx, 1, line);
    chunks[current].emit_op(Op::I32_OR, line);
}

/// Delegate combine semantics compatible with .NET multicast delegates.
pub fn emit_combine(chunks: &mut [Chunk], current: usize, line: u32) {
    let cur_slot = chunks[current].alloc_scratch(2);
    let handler_slot = cur_slot + 1;

    chunks[current].emit_op_u16(Op::LOCAL_SET, handler_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur_slot, line);

    emit_slot_is_nullish(chunks, current, cur_slot, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    chunks[current].emit_else(line);

    emit_slot_is_nullish(chunks, current, handler_slot, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    {
        let idx = chunks[current].add_import("ecma:array", "isArray");
        chunks[current].emit_call(idx, 1, line);
    }
    crate::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_else(line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// Invoke a (possibly multicast) delegate. `argc` counts every value on the
/// stack: the delegate plus its `argc - 1` handler arguments.
pub fn emit_invoke(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let n = (argc as u16).saturating_sub(1);
    let base = chunks[current].alloc_scratch(5 + n);
    let delegate_slot = base;
    let result_slot = base + 1;
    let len_slot = base + 2;
    let i_slot = base + 3;
    let handler_slot = base + 4;
    let arg_base = base + 5;

    let mut k = n;
    while k > 0 {
        k -= 1;
        chunks[current].emit_op_u16(Op::LOCAL_SET, arg_base + k, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, delegate_slot, line);

    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    emit_slot_is_nullish(chunks, current, delegate_slot, line);
    crate::ops::emit_dyn_not(&mut chunks[current], line);
    crate::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, delegate_slot, line);
    {
        let idx = chunks[current].add_import("ecma:array", "isArray");
        chunks[current].emit_call(idx, 1, line);
    }
    crate::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, delegate_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let block_p = chunks[current].emit_block(line);
    let (loop_p, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    crate::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, delegate_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, handler_slot, line);
    emit_slot_is_nullish(chunks, current, handler_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    for j in 0..n {
        chunks[current].emit_op_u16(Op::LOCAL_GET, arg_base + j, line);
    }
    chunks[current].emit_op_u8(Op::CALL_REF, n as u8, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    crate::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_p);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_p);

    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delegate_slot, line);
    for j in 0..n {
        chunks[current].emit_op_u16(Op::LOCAL_GET, arg_base + j, line);
    }
    chunks[current].emit_op_u8(Op::CALL_REF, n as u8, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `.GetInvocationList()` normalizes a delegate to the array of handlers.
pub fn emit_get_invocation_list(chunks: &mut [Chunk], current: usize, line: u32) {
    let delegate_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, delegate_slot, line);

    emit_slot_is_nullish(chunks, current, delegate_slot, line);
    chunks[current].emit_if_value(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, delegate_slot, line);
    {
        let idx = chunks[current].add_import("ecma:array", "isArray");
        chunks[current].emit_call(idx, 1, line);
    }
    crate::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delegate_slot, line);
    chunks[current].emit_else(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delegate_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
}

/// Delegate remove semantics compatible with .NET multicast delegates.
pub fn emit_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let cur_slot = chunks[current].alloc_scratch(6);
    let handler_slot = cur_slot + 1;
    let idx_slot = cur_slot + 2;
    let len_slot = cur_slot + 3;
    let elem_slot = cur_slot + 4;
    let loop_counter = cur_slot + 5;

    chunks[current].emit_op_u16(Op::LOCAL_SET, handler_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur_slot, line);

    emit_slot_is_nullish(chunks, current, cur_slot, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    {
        let idx = chunks[current].add_import("ecma:array", "isArray");
        chunks[current].emit_call(idx, 1, line);
    }
    crate::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, loop_counter, line);

    let block_patch = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, loop_counter, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    crate::ops::emit_dyn_ge(&mut chunks[current], line);
    crate::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, loop_counter, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    crate::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, loop_counter, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, loop_counter, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, loop_counter, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_patch);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    crate::ops::emit_dyn_ge(&mut chunks[current], line);
    crate::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    crate::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    crate::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    crate::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}
