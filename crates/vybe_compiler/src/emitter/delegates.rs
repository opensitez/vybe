//! Shared delegate runtime emitters.
//!
//! Stack contracts:
//! - combine: [current, handler] -> [delegate]
//! - remove:  [current, handler] -> [delegate]

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

use crate::emitter::collections;

/// Delegate combine semantics compatible with .NET multicast delegates.
pub fn emit_combine(chunks: &mut [Chunk], current: usize, line: u32) {
    let cur_slot = {
        let s = chunks[current].local_count;
        chunks[current].local_count = s + 1;
        s
    };
    let handler_slot = cur_slot + 1;
    chunks[current].local_count = handler_slot + 1;

    chunks[current].emit_op_u16(Op::LOCAL_SET, handler_slot, line); chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur_slot, line); chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op(Op::REF_IS_ARRAY, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_else(line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op(Op::DUP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::DUP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// Delegate remove semantics compatible with .NET multicast delegates.
pub fn emit_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let cur_slot = {
        let s = chunks[current].local_count;
        chunks[current].local_count = s + 1;
        s
    };
    let handler_slot = cur_slot + 1;
    let idx_slot = cur_slot + 2;
    let len_slot = cur_slot + 3;
    let elem_slot = cur_slot + 4;
    let loop_counter = cur_slot + 5;
    chunks[current].local_count = loop_counter + 1;

    chunks[current].emit_op_u16(Op::LOCAL_SET, handler_slot, line); chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur_slot, line); chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op(Op::REF_IS_ARRAY, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    // Array case: walk backwards and compare each candidate with the
    // handler. This stays entirely in bytecode; no host helper is
    // required beyond the existing array length/get/remove ops.
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line); chunks[current].emit_op(Op::DROP, line);

    let minus_one_idx = chunks[current].add_constant(Value::F64(-1.0));
    chunks[current].emit_op_u16(Op::CONST, minus_one_idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line); chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_CONST_1, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, loop_counter, line); chunks[current].emit_op(Op::DROP, line);

    let block_patch = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, loop_counter, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    crate::emitter::ops::emit_dyn_ge(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, loop_counter, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line); chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, loop_counter, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line); chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, loop_counter, line);
    chunks[current].emit_op(Op::I32_CONST_1, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, loop_counter, line); chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_patch);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    crate::emitter::ops::emit_dyn_ge(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    // Check final length and return appropriate delegate
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line); chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_CONST_1, line);
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}
