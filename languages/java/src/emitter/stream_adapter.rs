//! Java primitive stream adapters.
//!
//! Primitive streams are represented as plain arrays. Most intermediate
//! operations naturally return another array-backed stream; terminal optional
//! operations reuse the Java Optional pair representation.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use vybe_compiler::compiler::collections;
use vybe_compiler::compiler::instructions::host;

const STREAM_GENERATED_LIMIT: i32 = 128;

pub fn emit_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_array_new(chunks, current, 0, line);
}

pub fn emit_of(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    collections::emit_array_new(chunks, current, argc as u16, line);
}

pub fn emit_builder(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_array_new(chunks, current, 0, line);
}

pub fn emit_builder_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let builder = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, builder, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, builder, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, builder, line);
}

pub fn emit_range(chunks: &mut [Chunk], current: usize, inclusive: bool, line: u32) {
    let end_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

    if inclusive {
        chunks[current].emit_op_u16(Op::LOCAL_GET, end_slot, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, end_slot, line);
    }

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::compiler::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunks[current].emit_br(0, line);

    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_generate(chunks: &mut [Chunk], current: usize, line: u32) {
    let supplier_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, supplier_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_i32_const(STREAM_GENERATED_LIMIT, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, supplier_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    increment_index(chunks, current, index_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_iterate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_iterate_with_predicate_timing(chunks, current, argc, false, line);
}

pub fn emit_iterate_strict(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_iterate_with_predicate_timing(chunks, current, argc, true, line);
}

fn emit_iterate_with_predicate_timing(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    predicate_before_push: bool,
    line: u32,
) {
    let next_slot = chunks[current].alloc_scratch(1);
    let predicate_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, next_slot, line);
    if argc == 3 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, predicate_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_i32_const(STREAM_GENERATED_LIMIT, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    if predicate_before_push && argc == 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, predicate_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
        vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
        vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_br_if(1, line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    if !predicate_before_push && argc == 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, predicate_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
        vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
        vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_br_if(1, line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, next_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    increment_index(chunks, current, index_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_count(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_len(chunks, current, line);
}

pub fn emit_to_array(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 1 {
        let generator_slot = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, generator_slot, line);
    }
}

pub fn emit_sum(chunks: &mut [Chunk], current: usize, line: u32) {
    // `Stream.sum()` — fold the array with `+`, via the shared for-in scaffold.
    let arr = chunks[current].alloc_scratch(1);
    let total = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr, line);
    vybe_compiler::compiler::instructions::core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, total, line);
    let state = vybe_compiler::compiler::loops::emit_for_in_start(chunks, current, arr, idx, line);
    // total += element  (element is on top from for_in_start)
    chunks[current].emit_op_u16(Op::LOCAL_GET, total, line);
    vybe_compiler::compiler::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, total, line);
    vybe_compiler::compiler::loops::emit_for_in_end(chunks, current, idx, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, total, line);
}

pub fn emit_map(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:array", "map", 2, line);
}

pub fn emit_filter(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:array", "filter", 2, line);
}

pub fn emit_peek(chunks: &mut [Chunk], current: usize, line: u32) {
    let fn_slot = chunks[current].alloc_scratch(1);
    let array_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    host::emit(&mut chunks[current], "ecma:array", "forEach", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
}

pub fn emit_flat_map(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:array", "flatMap", 2, line);
}

pub fn emit_sorted(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 2 {
        collections::emit_sort_with_comparator(chunks, current, line);
    } else {
        collections::emit_sort(chunks, current, line);
    }
}

pub fn emit_for_each(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:array", "forEach", 2, line);
}

pub fn emit_limit(chunks: &mut [Chunk], current: usize, line: u32) {
    let limit_slot = chunks[current].alloc_scratch(1);
    let array_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, limit_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, limit_slot, line);
    collections::emit_slice(chunks, current, line);
}

pub fn emit_skip(chunks: &mut [Chunk], current: usize, line: u32) {
    let skip_slot = chunks[current].alloc_scratch(1);
    let array_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, skip_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, skip_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    collections::emit_len(chunks, current, line);
    collections::emit_slice(chunks, current, line);
}

pub fn emit_take_while(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_while_slice(chunks, current, true, line);
}

pub fn emit_drop_while(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_while_slice(chunks, current, false, line);
}

pub fn emit_min(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_min_max(chunks, current, true, line);
}

pub fn emit_max(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_min_max(chunks, current, false, line);
}

pub fn emit_extreme_value(chunks: &mut [Chunk], current: usize, argc: u8, is_min: bool, line: u32) {
    let comparator_slot = chunks[current].alloc_scratch(1);
    let array_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);
    let best_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);

    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, comparator_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    super::optional_adapter::emit_empty(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, comparator_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, best_slot, line);
        chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
        chunks[current].emit_i32_const(0, line);
        if is_min {
            vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
        } else {
            vybe_compiler::compiler::ops::emit_dyn_gt(&mut chunks[current], line);
        }
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, best_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        if is_min {
            vybe_compiler::compiler::ops::emit_dyn_gt(&mut chunks[current], line);
        } else {
            vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
        }
    }
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::compiler::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunks[current].emit_br(0, line);

    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, best_slot, line);
    super::optional_adapter::emit_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_max_value(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_max(chunks, current, line);
    emit_get_optional_value(chunks, current, line);
}

pub fn emit_average_value(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_average(chunks, current, line);
    emit_get_optional_value(chunks, current, line);
}

pub fn emit_average(chunks: &mut [Chunk], current: usize, line: u32) {
    let array_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    super::optional_adapter::emit_empty(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    emit_sum(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    super::optional_adapter::emit_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_find_first(chunks: &mut [Chunk], current: usize, line: u32) {
    let array_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    super::optional_adapter::emit_empty(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    super::optional_adapter::emit_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_get_optional_value(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
}

pub fn emit_concat(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_concat(chunks, current, line);
}

pub fn emit_collectors_joining(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let delimiter_slot = chunks[current].alloc_scratch(1);
    let prefix_slot = chunks[current].alloc_scratch(1);
    let suffix_slot = chunks[current].alloc_scratch(1);

    match argc {
        0 => {
            chunks[current].emit_string_const("", line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, delimiter_slot, line);
            chunks[current].emit_string_const("", line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, prefix_slot, line);
            chunks[current].emit_string_const("", line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, suffix_slot, line);
        }
        1 => {
            chunks[current].emit_op_u16(Op::LOCAL_SET, delimiter_slot, line);
            chunks[current].emit_string_const("", line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, prefix_slot, line);
            chunks[current].emit_string_const("", line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, suffix_slot, line);
        }
        _ => {
            chunks[current].emit_op_u16(Op::LOCAL_SET, suffix_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, prefix_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, delimiter_slot, line);
        }
    }

    chunks[current].emit_string_const("joining", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delimiter_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, prefix_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, suffix_slot, line);
    collections::emit_array_new(chunks, current, 4, line);
}

pub fn emit_collectors_to_list(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_string_const("toList", line);
    collections::emit_array_new(chunks, current, 1, line);
}

pub fn emit_collector_tag(chunks: &mut [Chunk], current: usize, tag: &str, argc: u8, line: u32) {
    let second_slot = chunks[current].alloc_scratch(1);
    let first_slot = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, second_slot, line);
    }
    if argc >= 1 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, first_slot, line);
    }
    chunks[current].emit_string_const(tag, line);
    if argc >= 1 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, first_slot, line);
    }
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, second_slot, line);
    }
    collections::emit_array_new(chunks, current, argc as u16 + 1, line);
}

pub fn emit_collector_tag_with_default_downstream(
    chunks: &mut [Chunk],
    current: usize,
    tag: &str,
    argc: u8,
    line: u32,
) {
    if argc == 1 {
        let classifier_slot = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, classifier_slot, line);
        chunks[current].emit_string_const(tag, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, classifier_slot, line);
        chunks[current].emit_string_const("toList", line);
        collections::emit_array_new(chunks, current, 1, line);
        collections::emit_array_new(chunks, current, 3, line);
    } else {
        emit_collector_tag(chunks, current, tag, argc, line);
    }
}

pub fn emit_collect(chunks: &mut [Chunk], current: usize, line: u32) {
    let collector_slot = chunks[current].alloc_scratch(1);
    let array_slot = chunks[current].alloc_scratch(1);
    let joined_slot = chunks[current].alloc_scratch(1);
    let prefix_slot = chunks[current].alloc_scratch(1);
    let suffix_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, collector_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_string_const("joining", line);
    vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    collections::emit_join(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, joined_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(2, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, prefix_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(3, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, suffix_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, prefix_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, joined_slot, line);
    super::string_adapter::emit_concat(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, suffix_slot, line);
    super::string_adapter::emit_concat(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    emit_collect_non_joining(
        chunks,
        current,
        collector_slot,
        array_slot,
        result_slot,
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

fn emit_collect_non_joining(
    chunks: &mut [Chunk],
    current: usize,
    collector_slot: u16,
    array_slot: u16,
    result_slot: u16,
    line: u32,
) {
    emit_collector_kind_eq(chunks, current, collector_slot, "toSet", line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    emit_distinct(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);

    emit_collector_kind_eq(chunks, current, collector_slot, "counting", line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);

    emit_collector_kind_eq(chunks, current, collector_slot, "summingInt", line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    emit_map(chunks, current, line);
    emit_sum(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);

    emit_collector_kind_eq(chunks, current, collector_slot, "averagingInt", line);
    chunks[current].emit_if(line);
    emit_collect_averaging_int(
        chunks,
        current,
        collector_slot,
        array_slot,
        result_slot,
        line,
    );
    chunks[current].emit_else(line);

    emit_collector_kind_eq(chunks, current, collector_slot, "toMap", line);
    chunks[current].emit_if(line);
    emit_collect_to_map(
        chunks,
        current,
        collector_slot,
        array_slot,
        result_slot,
        line,
    );
    chunks[current].emit_else(line);

    emit_collector_kind_eq(chunks, current, collector_slot, "toCollection", line);
    chunks[current].emit_if(line);
    emit_collect_to_collection(
        chunks,
        current,
        collector_slot,
        array_slot,
        result_slot,
        line,
    );
    chunks[current].emit_else(line);

    emit_collector_kind_eq(chunks, current, collector_slot, "mapping", line);
    chunks[current].emit_if(line);
    emit_collect_mapping(
        chunks,
        current,
        collector_slot,
        array_slot,
        result_slot,
        line,
    );
    chunks[current].emit_else(line);

    emit_collector_kind_eq(chunks, current, collector_slot, "filtering", line);
    chunks[current].emit_if(line);
    emit_collect_filtering(
        chunks,
        current,
        collector_slot,
        array_slot,
        result_slot,
        line,
    );
    chunks[current].emit_else(line);

    emit_collector_kind_eq(chunks, current, collector_slot, "collectingAndThen", line);
    chunks[current].emit_if(line);
    emit_collect_collecting_and_then(
        chunks,
        current,
        collector_slot,
        array_slot,
        result_slot,
        line,
    );
    chunks[current].emit_else(line);

    emit_collector_kind_eq(chunks, current, collector_slot, "reducing", line);
    chunks[current].emit_if(line);
    emit_collect_reducing(
        chunks,
        current,
        collector_slot,
        array_slot,
        result_slot,
        line,
    );
    chunks[current].emit_else(line);

    emit_collector_kind_eq(chunks, current, collector_slot, "minBy", line);
    chunks[current].emit_if(line);
    emit_collect_min_max_by(
        chunks,
        current,
        collector_slot,
        array_slot,
        result_slot,
        true,
        line,
    );
    chunks[current].emit_else(line);

    emit_collector_kind_eq(chunks, current, collector_slot, "maxBy", line);
    chunks[current].emit_if(line);
    emit_collect_min_max_by(
        chunks,
        current,
        collector_slot,
        array_slot,
        result_slot,
        false,
        line,
    );
    chunks[current].emit_else(line);

    emit_collector_kind_eq(chunks, current, collector_slot, "partitioningBy", line);
    chunks[current].emit_if(line);
    emit_collect_grouping(
        chunks,
        current,
        collector_slot,
        array_slot,
        result_slot,
        true,
        line,
    );
    chunks[current].emit_else(line);

    emit_collector_kind_eq(chunks, current, collector_slot, "groupingBy", line);
    chunks[current].emit_if(line);
    emit_collect_grouping(
        chunks,
        current,
        collector_slot,
        array_slot,
        result_slot,
        false,
        line,
    );
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_collector_kind_eq(
    chunks: &mut [Chunk],
    current: usize,
    collector_slot: u16,
    tag: &str,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_string_const(tag, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
}

fn emit_collect_averaging_int(
    chunks: &mut [Chunk],
    current: usize,
    collector_slot: u16,
    array_slot: u16,
    result_slot: u16,
    line: u32,
) {
    let mapped_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    emit_map(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mapped_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_slot, line);
    emit_sum(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
}

fn emit_collect_to_map(
    chunks: &mut [Chunk],
    current: usize,
    collector_slot: u16,
    array_slot: u16,
    result_slot: u16,
    line: u32,
) {
    let key_mapper_slot = chunks[current].alloc_scratch(1);
    let value_mapper_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_mapper_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(2, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_mapper_slot, line);

    collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_mapper_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_mapper_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    increment_index(chunks, current, index_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);
}

fn emit_collect_to_collection(
    chunks: &mut [Chunk],
    current: usize,
    collector_slot: u16,
    array_slot: u16,
    result_slot: u16,
    line: u32,
) {
    let supplier_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, supplier_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, supplier_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    super::list_adapter::emit_add(chunks, current, 2, line);
    chunks[current].emit_op(Op::DROP, line);

    increment_index(chunks, current, index_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);
}

fn emit_collect_mapping(
    chunks: &mut [Chunk],
    current: usize,
    collector_slot: u16,
    array_slot: u16,
    result_slot: u16,
    line: u32,
) {
    let mapped_slot = chunks[current].alloc_scratch(1);
    let downstream_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(2, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, downstream_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    emit_map(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mapped_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, downstream_slot, line);
    emit_collect_downstream_simple(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
}

fn emit_collect_filtering(
    chunks: &mut [Chunk],
    current: usize,
    collector_slot: u16,
    array_slot: u16,
    result_slot: u16,
    line: u32,
) {
    let filtered_slot = chunks[current].alloc_scratch(1);
    let downstream_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(2, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, downstream_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    emit_filter(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, filtered_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, filtered_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, downstream_slot, line);
    emit_collect_downstream_simple(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
}

fn emit_collect_collecting_and_then(
    chunks: &mut [Chunk],
    current: usize,
    collector_slot: u16,
    array_slot: u16,
    result_slot: u16,
    line: u32,
) {
    let downstream_slot = chunks[current].alloc_scratch(1);
    let finisher_slot = chunks[current].alloc_scratch(1);
    let intermediate_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, downstream_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(2, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, finisher_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, downstream_slot, line);
    emit_collect_downstream_simple(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, intermediate_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, finisher_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, intermediate_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
}

fn emit_collect_reducing(
    chunks: &mut [Chunk],
    current: usize,
    collector_slot: u16,
    array_slot: u16,
    result_slot: u16,
    line: u32,
) {
    let first_slot = chunks[current].alloc_scratch(1);
    let second_slot = chunks[current].alloc_scratch(1);
    let third_slot = chunks[current].alloc_scratch(1);
    let mapped_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(2, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, first_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, first_slot, line);
    emit_reduce(chunks, current, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(3, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, first_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(2, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, second_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, first_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, second_slot, line);
    emit_reduce(chunks, current, 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, first_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(2, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, second_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(3, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, third_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, second_slot, line);
    emit_map(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mapped_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, first_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, third_slot, line);
    emit_reduce(chunks, current, 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_collect_min_max_by(
    chunks: &mut [Chunk],
    current: usize,
    collector_slot: u16,
    array_slot: u16,
    result_slot: u16,
    is_min: bool,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    emit_extreme_value(chunks, current, 2, is_min, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
}

fn emit_collect_grouping(
    chunks: &mut [Chunk],
    current: usize,
    collector_slot: u16,
    array_slot: u16,
    result_slot: u16,
    preseed_partitions: bool,
    line: u32,
) {
    if preseed_partitions {
        emit_collect_partitioning(
            chunks,
            current,
            collector_slot,
            array_slot,
            result_slot,
            line,
        );
        return;
    }

    let classifier_slot = chunks[current].alloc_scratch(1);
    let downstream_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    let key_slot = chunks[current].alloc_scratch(1);
    let bucket_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, classifier_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(2, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, downstream_slot, line);

    collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, classifier_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    if preseed_partitions {
        vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_bool_const(true, line);
        chunks[current].emit_else(line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);

    emit_downstream_is_counting(chunks, current, downstream_slot, line);
    chunks[current].emit_if(line);
    emit_group_count_step(chunks, current, result_slot, key_slot, bucket_slot, line);
    chunks[current].emit_else(line);
    emit_group_list_step(
        chunks,
        current,
        result_slot,
        downstream_slot,
        key_slot,
        value_slot,
        bucket_slot,
        line,
    );
    chunks[current].emit_end(line);

    increment_index(chunks, current, index_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);
}

fn emit_collect_partitioning(
    chunks: &mut [Chunk],
    current: usize,
    collector_slot: u16,
    array_slot: u16,
    result_slot: u16,
    line: u32,
) {
    let classifier_slot = chunks[current].alloc_scratch(1);
    let downstream_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    let true_slot = chunks[current].alloc_scratch(1);
    let false_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, classifier_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(2, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, downstream_slot, line);

    emit_downstream_is_counting(chunks, current, downstream_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, true_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, false_slot, line);
    chunks[current].emit_else(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, true_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, false_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, classifier_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_partition_bucket_add(
        chunks,
        current,
        downstream_slot,
        true_slot,
        value_slot,
        line,
    );
    chunks[current].emit_else(line);
    emit_partition_bucket_add(
        chunks,
        current,
        downstream_slot,
        false_slot,
        value_slot,
        line,
    );
    chunks[current].emit_end(line);

    increment_index(chunks, current, index_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    emit_collector_kind_eq(chunks, current, downstream_slot, "toSet", line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, true_slot, line);
    emit_distinct(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, true_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, false_slot, line);
    emit_distinct(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, false_slot, line);
    chunks[current].emit_end(line);

    collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, true_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, false_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_partition_bucket_add(
    chunks: &mut [Chunk],
    current: usize,
    downstream_slot: u16,
    bucket_slot: u16,
    value_slot: u16,
    line: u32,
) {
    emit_downstream_is_counting(chunks, current, downstream_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bucket_slot, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::compiler::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bucket_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bucket_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

fn emit_downstream_is_counting(
    chunks: &mut [Chunk],
    current: usize,
    downstream_slot: u16,
    line: u32,
) {
    emit_collector_kind_eq(chunks, current, downstream_slot, "counting", line);
}

fn emit_group_count_step(
    chunks: &mut [Chunk],
    current: usize,
    result_slot: u16,
    key_slot: u16,
    bucket_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bucket_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bucket_slot, line);
    host::emit(&mut chunks[current], "wasm:js-undefined", "test", 1, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bucket_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bucket_slot, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::compiler::ops::emit_dyn_add(&mut chunks[current], line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_group_list_step(
    chunks: &mut [Chunk],
    current: usize,
    result_slot: u16,
    downstream_slot: u16,
    key_slot: u16,
    value_slot: u16,
    bucket_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bucket_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bucket_slot, line);
    host::emit(&mut chunks[current], "wasm:js-undefined", "test", 1, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bucket_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bucket_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    emit_collector_kind_eq(chunks, current, downstream_slot, "filtering", line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, downstream_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bucket_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);

    emit_collector_kind_eq(chunks, current, downstream_slot, "mapping", line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bucket_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, downstream_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bucket_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_collect_downstream_simple(chunks: &mut [Chunk], current: usize, line: u32) {
    let collector_slot = chunks[current].alloc_scratch(1);
    let array_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, collector_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);

    emit_collector_kind_eq(chunks, current, collector_slot, "toSet", line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    emit_distinct(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);

    emit_collector_kind_eq(chunks, current, collector_slot, "counting", line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);

    emit_collector_kind_eq(chunks, current, collector_slot, "summingInt", line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, collector_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    emit_map(chunks, current, line);
    emit_sum(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_any_match(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:array", "some", 2, line);
}

pub fn emit_all_match(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:array", "every", 2, line);
}

pub fn emit_none_match(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_any_match(chunks, current, line);
    vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_reduce(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 3 {
        let fn_slot = chunks[current].alloc_scratch(1);
        let identity_slot = chunks[current].alloc_scratch(1);
        let array_slot = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, identity_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, identity_slot, line);
        host::emit(&mut chunks[current], "ecma:array", "reduce", 3, line);
        return;
    }

    let fn_slot = chunks[current].alloc_scratch(1);
    let array_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    super::optional_adapter::emit_empty(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    collections::emit_len(chunks, current, line);
    collections::emit_slice(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    host::emit(&mut chunks[current], "ecma:array", "reduce", 3, line);
    super::optional_adapter::emit_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_distinct(chunks: &mut [Chunk], current: usize, line: u32) {
    let array_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_contains(chunks, current, line);
    vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::compiler::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

fn emit_min_max(chunks: &mut [Chunk], current: usize, is_min: bool, line: u32) {
    let array_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);
    let best_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    super::optional_adapter::emit_empty(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, best_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    if is_min {
        vybe_compiler::compiler::ops::emit_dyn_gt(&mut chunks[current], line);
    } else {
        vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
    }
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::compiler::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunks[current].emit_br(0, line);

    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, best_slot, line);
    super::optional_adapter::emit_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

fn emit_while_slice(chunks: &mut [Chunk], current: usize, keep_prefix: bool, line: u32) {
    let predicate_slot = chunks[current].alloc_scratch(1);
    let array_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let index_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, predicate_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, predicate_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    vybe_compiler::compiler::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    increment_index(chunks, current, index_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
    if keep_prefix {
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    }
    collections::emit_slice(chunks, current, line);
}

fn increment_index(chunks: &mut [Chunk], current: usize, index_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::compiler::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);
}
