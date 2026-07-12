//! Java array utility adapters.
//!
//! `java.util.Arrays` has overloads whose argument order and defaults do not
//! always match the lower-level ECMA array helpers, so keep those translations
//! in the Java frontend.

use crate::emitter::collections;
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

pub fn emit_fill(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    match argc {
        2 => {
            let value_slot = chunk.alloc_scratch(1);
            let array_slot = chunk.alloc_scratch(1);
            chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, array_slot, line);

            chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
            chunk.emit_i32_const(0, line);
            chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
            collections::emit_len(chunks, current, line);
            collections::emit_fill(chunks, current, line);
        }
        4 => {
            let value_slot = chunk.alloc_scratch(1);
            let to_slot = chunk.alloc_scratch(1);
            let from_slot = chunk.alloc_scratch(1);
            let array_slot = chunk.alloc_scratch(1);
            chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, to_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, from_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, array_slot, line);

            chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, from_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, to_slot, line);
            collections::emit_fill(chunks, current, line);
        }
        _ => {}
    }
}

pub fn emit_copy_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let new_len_slot = chunks[current].alloc_scratch(1);
    let source_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    let source_len_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, new_len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, new_len_slot, line);
    collections::emit_new_with_length(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, new_len_slot, line);
    collections::emit_fill(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source_len_slot, line);

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, new_len_slot, line);
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source_len_slot, line);
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_get(chunks, current, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_i32_const(1, line);
    crate::emitter::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunks[current].emit_br(0, line);

    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_copy_of_range(chunks: &mut [Chunk], current: usize, line: u32) {
    let to_slot = chunks[current].alloc_scratch(1);
    let from_slot = chunks[current].alloc_scratch(1);
    let source_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, to_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, from_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, from_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, to_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, from_slot, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get_range(chunks, current, line);
}

pub fn emit_binary_search(chunks: &mut [Chunk], current: usize, line: u32) {
    let key_slot = chunks[current].alloc_scratch(1);
    let array_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    crate::emitter::ops::emit_dyn_gt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunks[current].emit_br(0, line);

    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let joined_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_string_const(", ", line);
    collections::emit_join(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, joined_slot, line);

    chunks[current].emit_string_const("[", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, joined_slot, line);
    crate::emitter::strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_string_const("]", line);
    crate::emitter::strings::emit_str_concat(&mut chunks[current], line);
}

pub fn emit_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_sequence_equal(chunks, current, line);
}

pub fn emit_new_int_2d(chunks: &mut [Chunk], current: usize, line: u32) {
    let cols_slot = chunks[current].alloc_scratch(1);
    let rows_slot = chunks[current].alloc_scratch(1);
    let outer_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);
    let row_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, cols_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rows_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, rows_slot, line);
    collections::emit_new_with_length(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, outer_slot, line);

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rows_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cols_slot, line);
    collections::emit_new_with_length(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, row_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, row_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cols_slot, line);
    collections::emit_fill(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, outer_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, row_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunks[current].emit_br(0, line);

    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, outer_slot, line);
}
