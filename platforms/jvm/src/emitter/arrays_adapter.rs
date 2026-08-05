//! JVM `java.util.Arrays` utility adapters.
//!
//! `java.util.Arrays` has overloads whose argument order and defaults do not
//! always match the lower-level ECMA array helpers, so keep those translations
//! in the JVM platform adapter.

use vybe_compiler::primitives::{collections, instructions::host, object, ops, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

pub fn emit_sort(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 3 {
        let to = chunks[current].alloc_scratch(1);
        let from = chunks[current].alloc_scratch(1);
        let array = chunks[current].alloc_scratch(1);
        let slice = chunks[current].alloc_scratch(1);
        let count = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], to, line);
        set(&mut chunks[current], from, line);
        set(&mut chunks[current], array, line);
        get(&mut chunks[current], to, line);
        get(&mut chunks[current], from, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        set(&mut chunks[current], count, line);
        get(&mut chunks[current], array, line);
        get(&mut chunks[current], from, line);
        get(&mut chunks[current], to, line);
        collections::emit_slice(chunks, current, line);
        collections::emit_sort(chunks, current, line);
        set(&mut chunks[current], slice, line);
        get(&mut chunks[current], array, line);
        get(&mut chunks[current], from, line);
        get(&mut chunks[current], count, line);
        collections::emit_remove_range(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        get(&mut chunks[current], array, line);
        get(&mut chunks[current], from, line);
        get(&mut chunks[current], slice, line);
        collections::emit_insert_range(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    } else if argc == 2 {
        let comparator = chunks[current].alloc_scratch(1);
        let array = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], comparator, line);
        set(&mut chunks[current], array, line);
        get(&mut chunks[current], comparator, line);
        chunks[current].emit_string_const("__java_reverse_order", line);
        host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], array, line);
        collections::emit_sort(chunks, current, line);
        collections::emit_reverse(chunks, current, line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], array, line);
        get(&mut chunks[current], comparator, line);
        collections::emit_sort_with_comparator(chunks, current, line);
        chunks[current].emit_end(line);
    } else {
        collections::emit_sort(chunks, current, line);
    }
}

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
    emit_copy_of_ex(chunks, current, 2, line);
}

/// `argc` INCLUDES the receiver: Kotlin's no-arg member `arr.copyOf()` is a
/// same-length clone (the old fixed-arity pop read the RECEIVER as the
/// length and crashed).
pub fn emit_copy_of_ex(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let new_len_slot = chunks[current].alloc_scratch(1);
    let source_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    let source_len_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);

    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, new_len_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, source_slot, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_SET, source_slot, line);
        get(&mut chunks[current], source_slot, line);
        collections::emit_len(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, new_len_slot, line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, new_len_slot, line);
    collections::emit_new_with_length(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    get(&mut chunks[current], source_slot, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("number", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    let default_slot = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], default_slot, line);
    get(&mut chunks[current], result_slot, line);
    get(&mut chunks[current], default_slot, line);
    chunks[current].emit_i32_const(0, line);
    get(&mut chunks[current], new_len_slot, line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source_len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
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

    // `Arrays.copyOfRange` THROWS: IllegalArgumentException when from > to,
    // ArrayIndexOutOfBoundsException when from is outside the source
    // (javadoc-specified); the plain slice silently clamped.
    chunks[current].emit_op_u16(Op::LOCAL_GET, from_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, to_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("fromIndex > toIndex", line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        &mut chunks[current],
        "IllegalArgumentException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, from_slot, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, from_slot, line);
    {
        let c = &mut chunks[current];
        c.emit_op_u16(Op::LOCAL_GET, source_slot, line);
    }
    collections::emit_len(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("fromIndex out of range", line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        &mut chunks[current],
        "IndexOutOfBoundsException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, from_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, to_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, from_slot, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get_range(chunks, current, line);
}

pub fn emit_binary_search(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let key_slot = chunks[current].alloc_scratch(1);
    let array_slot = chunks[current].alloc_scratch(1);
    let from_slot = chunks[current].alloc_scratch(1);
    let to_slot = chunks[current].alloc_scratch(1);
    if argc == 4 {
        set(&mut chunks[current], key_slot, line);
        set(&mut chunks[current], to_slot, line);
        set(&mut chunks[current], from_slot, line);
        set(&mut chunks[current], array_slot, line);
    } else {
        set(&mut chunks[current], key_slot, line);
        set(&mut chunks[current], array_slot, line);
        chunks[current].emit_i32_const(0, line);
        set(&mut chunks[current], from_slot, line);
        get(&mut chunks[current], array_slot, line);
        collections::emit_len(chunks, current, line);
        set(&mut chunks[current], to_slot, line);
    }
    emit_binary_search_range(
        chunks, current, array_slot, from_slot, to_slot, key_slot, line,
    );
}

fn emit_binary_search_range(
    chunks: &mut [Chunk],
    current: usize,
    array_slot: u16,
    from_slot: u16,
    to_slot: u16,
    key_slot: u16,
    line: u32,
) {
    let index_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], from_slot, line);
    set(&mut chunks[current], index_slot, line);

    get(&mut chunks[current], from_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0, line);
    get(&mut chunks[current], to_slot, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    get(&mut chunks[current], to_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    get(&mut chunks[current], to_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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

pub fn emit_mismatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let right = chunks[current].alloc_scratch(1);
    let left = chunks[current].alloc_scratch(1);
    let left_len = chunks[current].alloc_scratch(1);
    let right_len = chunks[current].alloc_scratch(1);
    let min_len = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], right, line);
    set(&mut chunks[current], left, line);
    get(&mut chunks[current], left, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], left_len, line);
    get(&mut chunks[current], right, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], right_len, line);
    get(&mut chunks[current], left_len, line);
    get(&mut chunks[current], right_len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], left_len, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], right_len, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], min_len, line);
    chunks[current].emit_i32_const(-1, line);
    set(&mut chunks[current], result, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], min_len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], left, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    get(&mut chunks[current], right, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], index, line);
    set(&mut chunks[current], result, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);

    get(&mut chunks[current], result, line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], left_len, line);
    get(&mut chunks[current], right_len, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], min_len, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], result, line);
    chunks[current].emit_end(line);
}

pub fn emit_compare(chunks: &mut [Chunk], current: usize, line: u32) {
    let right = chunks[current].alloc_scratch(1);
    let left = chunks[current].alloc_scratch(1);
    let mismatch = chunks[current].alloc_scratch(1);
    let left_len = chunks[current].alloc_scratch(1);
    let right_len = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], right, line);
    set(&mut chunks[current], left, line);
    get(&mut chunks[current], left, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], left_len, line);
    get(&mut chunks[current], right, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], right_len, line);
    get(&mut chunks[current], left, line);
    get(&mut chunks[current], right, line);
    emit_mismatch(chunks, current, line);
    set(&mut chunks[current], mismatch, line);
    get(&mut chunks[current], mismatch, line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], mismatch, line);
    get(&mut chunks[current], left_len, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], mismatch, line);
    get(&mut chunks[current], right_len, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], left, line);
    get(&mut chunks[current], mismatch, line);
    collections::emit_get(chunks, current, line);
    get(&mut chunks[current], right, line);
    get(&mut chunks[current], mismatch, line);
    collections::emit_get(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_unsigned_byte_at(chunks: &mut [Chunk], current: usize, array: u16, index: u16, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], array, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_i32_const(256, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
}

pub fn emit_compare_unsigned(chunks: &mut [Chunk], current: usize, line: u32) {
    let right = chunks[current].alloc_scratch(1);
    let left = chunks[current].alloc_scratch(1);
    let mismatch = chunks[current].alloc_scratch(1);
    let left_len = chunks[current].alloc_scratch(1);
    let right_len = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], right, line);
    set(&mut chunks[current], left, line);
    get(&mut chunks[current], left, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], left_len, line);
    get(&mut chunks[current], right, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], right_len, line);
    get(&mut chunks[current], left, line);
    get(&mut chunks[current], right, line);
    emit_mismatch(chunks, current, line);
    set(&mut chunks[current], mismatch, line);
    get(&mut chunks[current], mismatch, line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], mismatch, line);
    get(&mut chunks[current], left_len, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], mismatch, line);
    get(&mut chunks[current], right_len, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_else(line);
    emit_unsigned_byte_at(chunks, current, left, mismatch, line);
    emit_unsigned_byte_at(chunks, current, right, mismatch, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_set_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let mapper = chunks[current].alloc_scratch(1);
    let array = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], mapper, line);
    set(&mut chunks[current], array, line);
    get(&mut chunks[current], array, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], mapper, line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], array, line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_parallel_prefix(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let operator = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], operator, line);
    let array = chunks[current].alloc_scratch(1);
    if argc == 3 {
        set(&mut chunks[current], array, line);
        chunks[current].emit_op(Op::DROP, line);
    } else {
        set(&mut chunks[current], array, line);
    }
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let previous = chunks[current].alloc_scratch(1);
    let current_value = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], array, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    chunks[current].emit_i32_const(1, line);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], array, line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], previous, line);
    get(&mut chunks[current], array, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], current_value, line);
    get(&mut chunks[current], operator, line);
    get(&mut chunks[current], previous, line);
    get(&mut chunks[current], current_value, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], array, line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_deep_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let right = chunks[current].alloc_scratch(1);
    let left = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    let left_row = chunks[current].alloc_scratch(1);
    let right_row = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], right, line);
    set(&mut chunks[current], left, line);
    get(&mut chunks[current], left, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    get(&mut chunks[current], len, line);
    get(&mut chunks[current], right, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    set(&mut chunks[current], result, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], result, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], left, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], left_row, line);
    get(&mut chunks[current], right, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], right_row, line);
    get(&mut chunks[current], left_row, line);
    get(&mut chunks[current], right_row, line);
    collections::emit_sequence_equal(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], result, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], result, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_deep_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let json_slot = chunks[current].alloc_scratch(1);
    host::emit(&mut chunks[current], "ecma:json", "stringify", 1, line);
    set(&mut chunks[current], json_slot, line);
    get(&mut chunks[current], json_slot, line);
    chunks[current].emit_string_const(",", line);
    chunks[current].emit_string_const(", ", line);
    host::emit(&mut chunks[current], "ecma:string", "replaceAll", 3, line);
}

pub fn emit_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let items = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], items, line);
    get(&mut chunks[current], items, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    chunks[current].emit_string_const("[", line);
    set(&mut chunks[current], out, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], index, line);
    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], out, line);
    chunks[current].emit_string_const(", ", line);
    strings::emit_str_concat(&mut chunks[current], line);
    set(&mut chunks[current], out, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], items, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], elem, line);
    get(&mut chunks[current], elem, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("null", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("", line);
    get(&mut chunks[current], elem, line);
    strings::emit_str_concat_coercing(&mut chunks[current], line);
    chunks[current].emit_end(line);
    let elem_text = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], elem_text, line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], elem_text, line);
    strings::emit_str_concat(&mut chunks[current], line);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], out, line);
    chunks[current].emit_string_const("]", line);
    strings::emit_str_concat(&mut chunks[current], line);
}

pub fn emit_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_sequence_equal(chunks, current, line);
}

pub fn emit_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let items = chunks[current].alloc_scratch(3);
    let hash = items + 1;
    let index = items + 2;
    set(&mut chunks[current], items, line);
    chunks[current].emit_i32_const(1, line);
    set(&mut chunks[current], hash, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], items, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_i32_const(31, line);
    get(&mut chunks[current], hash, line);
    chunks[current].emit_op(Op::I32_MUL, line);
    get(&mut chunks[current], items, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    let elem = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], elem, line);
    get(&mut chunks[current], elem, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("number", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], elem, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], elem, line);
    object::emit_hash_code(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], hash, line);

    get(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], hash, line);
}

pub fn emit_deep_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let json = chunks[current].alloc_scratch(1);
    host::emit(&mut chunks[current], "ecma:json", "stringify", 1, line);
    set(&mut chunks[current], json, line);
    get(&mut chunks[current], json, line);
    object::emit_hash_code(&mut chunks[current], line);
}

pub fn emit_arrays_as_list(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 1 {
        collections::emit_array_new(chunks, current, argc as u16, line);
        return;
    }

    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:array", "length", 1, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    collections::emit_array_new(chunks, current, 1, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
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
