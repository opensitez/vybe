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
    let cur_not_null = chunks[current].emit_jump(Op::BR_IF_FALSE, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    let done = chunks[current].emit_jump(Op::BR, line);
    chunks[current].patch_jump(cur_not_null);

    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    let handler_not_null = chunks[current].emit_jump(Op::BR_IF_FALSE, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    let done2 = chunks[current].emit_jump(Op::BR, line);
    chunks[current].patch_jump(handler_not_null);

    let is_array_idx = chunks[0].add_import("ecma:array", "isArray");
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, is_array_idx, line);
    chunks[current].emit(1, line);
    let not_array = chunks[current].emit_jump(Op::BR_IF_FALSE, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    let done3 = chunks[current].emit_jump(Op::BR, line);

    chunks[current].patch_jump(not_array);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op(Op::DUP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::DUP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].patch_jump(done3);
    chunks[current].patch_jump(done2);
    chunks[current].patch_jump(done);
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
    chunks[current].local_count = len_slot + 1;

    chunks[current].emit_op_u16(Op::LOCAL_SET, handler_slot, line); chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur_slot, line); chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    let cur_not_null = chunks[current].emit_jump(Op::BR_IF_FALSE, line);
    chunks[current].emit_op(Op::NULL, line);
    let done = chunks[current].emit_jump(Op::BR, line);
    chunks[current].patch_jump(cur_not_null);

    let is_array_idx = chunks[0].add_import("ecma:array", "isArray");
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, is_array_idx, line);
    chunks[current].emit(1, line);
    let not_array = chunks[current].emit_jump(Op::BR_IF_FALSE, line);

    // Array case: use findIndex to search backwards and remove
    // Stack setup: [array, startIndex, searchHandler] for custom search
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line); chunks[current].emit_op(Op::DROP, line);

    // Initialize idx = -1 (not found)
    let const_idx = chunks[current].add_constant(Value::F64(-1.0));
    chunks[current].emit_op_u16(Op::CONST, const_idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line); chunks[current].emit_op(Op::DROP, line);

    // Manual iteration from len-1 down to 0
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_CONST_1, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    let loop_counter = cur_slot + 4;
    chunks[current].local_count = loop_counter + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, loop_counter, line); chunks[current].emit_op(Op::DROP, line);

    // Loop backwards
    let loop_start = chunks[current].current_offset();
    chunks[current].emit_op_u16(Op::LOCAL_GET, loop_counter, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    chunks[current].emit_op(Op::DYN_GE, line);
    let loop_end = chunks[current].emit_jump(Op::BR_IF_FALSE, line);

    // array[i] === handler? (using function name comparison as fallback)
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, loop_counter, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    chunks[current].emit_op(Op::DYN_EQ, line);
    let elem_matches = chunks[current].emit_jump(Op::BR_IF_TRUE, line);

    // No match: decrement counter and continue loop
    chunks[current].emit_op_u16(Op::LOCAL_GET, loop_counter, line);
    chunks[current].emit_op(Op::I32_CONST_1, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, loop_counter, line); chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_loop(loop_start, line);

    chunks[current].patch_jump(elem_matches);
    // Found match: set idx and break
    chunks[current].emit_op_u16(Op::LOCAL_GET, loop_counter, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line); chunks[current].emit_op(Op::DROP, line);
    chunks[current].patch_jump(loop_end);

    // Check if we found anything
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    chunks[current].emit_op(Op::DYN_GE, line);
    let no_remove = chunks[current].emit_jump(Op::BR_IF_FALSE, line);
    
    // Remove the element at idx
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].patch_jump(no_remove);

    // Check final length and return appropriate delegate
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line); chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    chunks[current].emit_op(Op::DYN_EQ, line);
    let len_not_zero = chunks[current].emit_jump(Op::BR_IF_FALSE, line);
    chunks[current].emit_op(Op::NULL, line);
    let done2 = chunks[current].emit_jump(Op::BR, line);
    chunks[current].patch_jump(len_not_zero);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_CONST_1, line);
    chunks[current].emit_op(Op::DYN_EQ, line);
    let len_not_one = chunks[current].emit_jump(Op::BR_IF_FALSE, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    collections::emit_get(chunks, current, line);
    let done3 = chunks[current].emit_jump(Op::BR, line);
    chunks[current].patch_jump(len_not_one);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    let done4 = chunks[current].emit_jump(Op::BR, line);

    chunks[current].patch_jump(not_array);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    chunks[current].emit_op(Op::DYN_EQ, line);
    let neq = chunks[current].emit_jump(Op::BR_IF_FALSE, line);
    chunks[current].emit_op(Op::NULL, line);
    let done5 = chunks[current].emit_jump(Op::BR, line);
    chunks[current].patch_jump(neq);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_slot, line);

    chunks[current].patch_jump(done5);
    chunks[current].patch_jump(done4);
    chunks[current].patch_jump(done3);
    chunks[current].patch_jump(done2);
    chunks[current].patch_jump(done);
}

