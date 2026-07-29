use vybe_bytecode::opcode::Op;
use vybe_bytecode::Chunk;
use vybe_compiler::compiler::instructions::{core_wasm, host};

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn emit_throw_dotnet_exception(chunk: &mut Chunk, exception_name: &str, message: &str, line: u32) {
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_dup(line);
    chunk.emit_string_const(message, line);
    vybe_compiler::compiler::errors::emit_exception_new_finalize(chunk, exception_name, line);
    vybe_compiler::compiler::errors::emit_throw(chunk, line);
}

fn emit_ignore_case_flag(chunk: &mut Chunk, comparison_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, comparison_slot, line);
    chunk.emit_bool_const(true, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);

    chunk.emit_op_u16(Op::LOCAL_GET, comparison_slot, line);
    chunk.emit_string_const("__dotnet_stringcomparison_ordinalignorecase", line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    chunk.emit_op(Op::I32_OR, line);

    chunk.emit_op_u16(Op::LOCAL_GET, comparison_slot, line);
    chunk.emit_string_const("__dotnet_stringcomparison_invariantignorecase", line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn emit_load_maybe_lowercase(chunk: &mut Chunk, value_slot: u16, ignore_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, ignore_slot, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    vybe_compiler::compiler::strings::emit_to_lower(chunk, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_end(line);
}

fn emit_string_len(chunk: &mut Chunk, value_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
}

fn emit_string_substr_from_slots(
    chunk: &mut Chunk,
    value_slot: u16,
    start_slot: u16,
    len_slot: u16,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    host::emit(chunk, "ecma:string", "substr", 3, line);
}

fn emit_string_index_of_slots(
    chunk: &mut Chunk,
    haystack_slot: u16,
    needle_slot: u16,
    start_slot: Option<u16>,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, haystack_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, needle_slot, line);
    if let Some(start_slot) = start_slot {
        chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
        host::emit(chunk, "ecma:string", "indexOf", 3, line);
    } else {
        host::emit(chunk, "ecma:string", "indexOf", 2, line);
    }
}

fn emit_string_last_index_of_slots(
    chunk: &mut Chunk,
    haystack_slot: u16,
    needle_slot: u16,
    start_slot: Option<u16>,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, haystack_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, needle_slot, line);
    if let Some(start_slot) = start_slot {
        chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
        host::emit(chunk, "ecma:string", "lastIndexOf", 3, line);
    } else {
        host::emit(chunk, "ecma:string", "lastIndexOf", 2, line);
    }
}

fn emit_value_type_is_numeric(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_string_const("number", line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_string_const("i32", line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_string_const("i64", line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn stash_string_compare_args(chunk: &mut Chunk, argc: u8, line: u32) -> (u16, u16, u16) {
    let left_slot = chunk.alloc_scratch(3);
    let right_slot = left_slot + 1;
    let ignore_slot = left_slot + 2;

    if argc >= 3 {
        let comparison_slot = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, comparison_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
        emit_ignore_case_flag(chunk, comparison_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, ignore_slot, line);
    } else {
        chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, ignore_slot, line);
    }

    (left_slot, right_slot, ignore_slot)
}

pub fn emit_string_compare(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 6 {
        emit_string_compare_substrings(chunks, current, line);
        return;
    }
    let (left_slot, right_slot, ignore_slot) =
        stash_string_compare_args(&mut chunks[current], argc, line);
    emit_string_compare_slots(
        &mut chunks[current],
        left_slot,
        right_slot,
        ignore_slot,
        line,
    );
}

fn emit_string_compare_slots(
    chunk: &mut Chunk,
    left_slot: u16,
    right_slot: u16,
    ignore_slot: u16,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(-1, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_else(line);
    emit_load_maybe_lowercase(chunk, left_slot, ignore_slot, line);
    emit_load_maybe_lowercase(chunk, right_slot, ignore_slot, line);
    let compare_idx = chunk.add_import("ecma:string", "localeCompare");
    chunk.emit_call(compare_idx, 2, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_string_compare_substrings(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let left_slot = chunk.alloc_scratch(6);
    let left_index_slot = left_slot + 1;
    let right_slot = left_slot + 2;
    let right_index_slot = left_slot + 3;
    let length_slot = left_slot + 4;
    let ignore_slot = left_slot + 5;
    let comparison_slot = chunk.alloc_scratch(1);
    let substr_idx = chunk.add_import("ecma:string", "substr");

    chunk.emit_op_u16(Op::LOCAL_SET, comparison_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, length_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, right_index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);

    emit_ignore_case_flag(chunk, comparison_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ignore_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, left_index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, length_slot, line);
    chunk.emit_call(substr_idx, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, length_slot, line);
    chunk.emit_call(substr_idx, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);

    emit_string_compare_slots(chunk, left_slot, right_slot, ignore_slot, line);
}

pub fn emit_string_equals(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (left_slot, right_slot, ignore_slot) =
        stash_string_compare_args(&mut chunks[current], argc, line);
    emit_load_maybe_lowercase(&mut chunks[current], left_slot, ignore_slot, line);
    emit_load_maybe_lowercase(&mut chunks[current], right_slot, ignore_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_string_contains(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (left_slot, right_slot, ignore_slot) =
        stash_string_compare_args(&mut chunks[current], argc, line);
    emit_load_maybe_lowercase(&mut chunks[current], left_slot, ignore_slot, line);
    emit_load_maybe_lowercase(&mut chunks[current], right_slot, ignore_slot, line);
    vybe_compiler::compiler::strings::emit_index_of(&mut chunks[current], line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    vybe_compiler::compiler::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_string_starts_with(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (left_slot, right_slot, ignore_slot) =
        stash_string_compare_args(&mut chunks[current], argc, line);
    emit_load_maybe_lowercase(&mut chunks[current], left_slot, ignore_slot, line);
    emit_load_maybe_lowercase(&mut chunks[current], right_slot, ignore_slot, line);
    let idx = chunks[current].add_import("ecma:string", "startsWith");
    chunks[current].emit_call(idx, 2, line);
}

pub fn emit_string_ends_with(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (left_slot, right_slot, ignore_slot) =
        stash_string_compare_args(&mut chunks[current], argc, line);
    emit_load_maybe_lowercase(&mut chunks[current], left_slot, ignore_slot, line);
    emit_load_maybe_lowercase(&mut chunks[current], right_slot, ignore_slot, line);
    let idx = chunks[current].add_import("ecma:string", "endsWith");
    chunks[current].emit_call(idx, 2, line);
}

pub fn emit_string_index_of(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = reserve_slot(chunk);
    let needle_slot = reserve_slot(chunk);
    let start_slot = reserve_slot(chunk);
    let count_slot = reserve_slot(chunk);
    let comparison_slot = reserve_slot(chunk);
    let ignore_slot = reserve_slot(chunk);
    let slice_slot = reserve_slot(chunk);
    let found_slot = reserve_slot(chunk);

    match argc {
        4 => {
            chunk.emit_op_u16(Op::LOCAL_SET, comparison_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, start_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, needle_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
            emit_value_type_is_numeric(chunk, comparison_slot, line);
            chunk.emit_if(line);
            chunk.emit_op_u16(Op::LOCAL_GET, comparison_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
            chunk.emit_i32_const(0, line);
            chunk.emit_op_u16(Op::LOCAL_SET, ignore_slot, line);
            emit_string_substr_from_slots(chunk, value_slot, start_slot, count_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, slice_slot, line);
            emit_string_index_of_slots(chunk, slice_slot, needle_slot, None, line);
            chunk.emit_op_u16(Op::LOCAL_SET, found_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, found_slot, line);
            chunk.emit_i32_const(0, line);
            chunk.emit_op(Op::I32_LT_S, line);
            chunk.emit_if_value(line);
            chunk.emit_i32_const(-1, line);
            chunk.emit_else(line);
            chunk.emit_op_u16(Op::LOCAL_GET, found_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
            chunk.emit_op(Op::I32_ADD, line);
            chunk.emit_end(line);
            chunk.emit_else(line);
            emit_ignore_case_flag(chunk, comparison_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, ignore_slot, line);
            emit_load_maybe_lowercase(chunk, value_slot, ignore_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
            emit_load_maybe_lowercase(chunk, needle_slot, ignore_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, needle_slot, line);
            emit_string_index_of_slots(chunk, value_slot, needle_slot, Some(start_slot), line);
            chunk.emit_end(line);
        }
        3 => {
            chunk.emit_op_u16(Op::LOCAL_SET, comparison_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, needle_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
            emit_value_type_is_numeric(chunk, comparison_slot, line);
            chunk.emit_if(line);
            emit_string_index_of_slots(chunk, value_slot, needle_slot, Some(comparison_slot), line);
            chunk.emit_else(line);
            emit_ignore_case_flag(chunk, comparison_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, ignore_slot, line);
            emit_load_maybe_lowercase(chunk, value_slot, ignore_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
            emit_load_maybe_lowercase(chunk, needle_slot, ignore_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, needle_slot, line);
            emit_string_index_of_slots(chunk, value_slot, needle_slot, None, line);
            chunk.emit_end(line);
        }
        _ => {
            chunk.emit_op_u16(Op::LOCAL_SET, needle_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
            emit_string_index_of_slots(chunk, value_slot, needle_slot, None, line);
        }
    }
}

pub fn emit_string_last_index_of(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = reserve_slot(chunk);
    let needle_slot = reserve_slot(chunk);
    let start_slot = reserve_slot(chunk);
    let count_slot = reserve_slot(chunk);
    let begin_slot = reserve_slot(chunk);
    let slice_slot = reserve_slot(chunk);
    let found_slot = reserve_slot(chunk);

    match argc {
        4 => {
            chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, start_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, needle_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
            chunk.emit_op(Op::I32_SUB, line);
            chunk.emit_i32_const(1, line);
            chunk.emit_op(Op::I32_ADD, line);
            chunk.emit_op_u16(Op::LOCAL_SET, begin_slot, line);
            emit_string_substr_from_slots(chunk, value_slot, begin_slot, count_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, slice_slot, line);
            emit_string_last_index_of_slots(chunk, slice_slot, needle_slot, None, line);
            chunk.emit_op_u16(Op::LOCAL_SET, found_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, found_slot, line);
            chunk.emit_i32_const(0, line);
            chunk.emit_op(Op::I32_LT_S, line);
            chunk.emit_if_value(line);
            chunk.emit_i32_const(-1, line);
            chunk.emit_else(line);
            chunk.emit_op_u16(Op::LOCAL_GET, found_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, begin_slot, line);
            chunk.emit_op(Op::I32_ADD, line);
            chunk.emit_end(line);
        }
        3 => {
            chunk.emit_op_u16(Op::LOCAL_SET, start_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, needle_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
            emit_value_type_is_numeric(chunk, start_slot, line);
            chunk.emit_if(line);
            emit_string_last_index_of_slots(chunk, value_slot, needle_slot, Some(start_slot), line);
            chunk.emit_else(line);
            emit_ignore_case_flag(chunk, start_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
            emit_load_maybe_lowercase(chunk, value_slot, count_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
            emit_load_maybe_lowercase(chunk, needle_slot, count_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, needle_slot, line);
            emit_string_last_index_of_slots(chunk, value_slot, needle_slot, None, line);
            chunk.emit_end(line);
        }
        _ => {
            chunk.emit_op_u16(Op::LOCAL_SET, needle_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
            emit_string_last_index_of_slots(chunk, value_slot, needle_slot, None, line);
        }
    }
}

pub fn emit_string_index_of_any(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let start_slot = reserve_slot(&mut chunks[current]);
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, start_slot, line);
    } else {
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, start_slot, line);
    }
    let targets_slot = reserve_slot(&mut chunks[current]);
    let value_slot = reserve_slot(&mut chunks[current]);
    let i_slot = reserve_slot(&mut chunks[current]);
    let len_slot = reserve_slot(&mut chunks[current]);
    let candidate_slot = reserve_slot(&mut chunks[current]);
    let best_slot = reserve_slot(&mut chunks[current]);
    let char_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, targets_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, targets_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let state = vybe_compiler::compiler::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::compiler::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, targets_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, char_slot, line);
    emit_string_index_of_slots(
        &mut chunks[current],
        value_slot,
        char_slot,
        Some(start_slot),
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, candidate_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, candidate_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, best_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, candidate_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, best_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, candidate_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_compiler::compiler::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, best_slot, line);
}

pub fn emit_string_last_index_of_any(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let targets_slot = reserve_slot(&mut chunks[current]);
    let value_slot = reserve_slot(&mut chunks[current]);
    let i_slot = reserve_slot(&mut chunks[current]);
    let len_slot = reserve_slot(&mut chunks[current]);
    let candidate_slot = reserve_slot(&mut chunks[current]);
    let best_slot = reserve_slot(&mut chunks[current]);
    let char_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, targets_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, targets_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let state = vybe_compiler::compiler::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::compiler::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, targets_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, char_slot, line);
    emit_string_last_index_of_slots(&mut chunks[current], value_slot, char_slot, None, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, candidate_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, candidate_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, best_slot, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, candidate_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_compiler::compiler::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, best_slot, line);
}

pub fn emit_string_substring(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc >= 3 {
        let len_slot = reserve_slot(chunk);
        let start_slot = reserve_slot(chunk);
        let value_slot = reserve_slot(chunk);
        let str_len_slot = reserve_slot(chunk);
        let end_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, start_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
        emit_string_len(chunk, value_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, str_len_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
        vybe_compiler::compiler::ops::emit_dyn_add(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_SET, end_slot, line);
        emit_substring_bounds_check(chunk, start_slot, len_slot, end_slot, str_len_slot, line);
        emit_string_substr_from_slots(chunk, value_slot, start_slot, len_slot, line);
    } else {
        let start_slot = reserve_slot(chunk);
        let value_slot = reserve_slot(chunk);
        let str_len_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, start_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
        emit_string_len(chunk, value_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, str_len_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
        chunk.emit_i32_const(0, line);
        vybe_compiler::compiler::ops::emit_dyn_ge(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, str_len_slot, line);
        vybe_compiler::compiler::ops::emit_dyn_le(chunk, line);
        chunk.emit_op(Op::I32_AND, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        emit_throw_dotnet_exception(
            chunk,
            "ArgumentOutOfRangeException",
            "startIndex cannot be larger than length of string.",
            line,
        );
        chunk.emit_end(line);
        chunk.emit_op_u16(Op::LOCAL_GET, str_len_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
        chunk.emit_op(Op::I32_SUB, line);
        let len_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
        emit_string_substr_from_slots(chunk, value_slot, start_slot, len_slot, line);
    }
}

pub fn emit_string_char_at_checked(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let index_slot = reserve_slot(chunk);
    let value_slot = reserve_slot(chunk);
    let len_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    emit_string_len(chunk, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunk.emit_i32_const(0, line);
    vybe_compiler::compiler::ops::emit_dyn_ge(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    emit_throw_dotnet_exception(
        chunk,
        "IndexOutOfRangeException",
        "Index was outside the bounds of the string.",
        line,
    );
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunk.emit_i32_const(1, line);
    vybe_compiler::compiler::ops::emit_dyn_add(chunk, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);
}

fn emit_substring_bounds_check(
    chunk: &mut Chunk,
    start_slot: u16,
    len_slot: u16,
    end_slot: u16,
    str_len_slot: u16,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunk.emit_i32_const(0, line);
    vybe_compiler::compiler::ops::emit_dyn_ge(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_i32_const(0, line);
    vybe_compiler::compiler::ops::emit_dyn_ge(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, str_len_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_le(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, str_len_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_le(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    emit_throw_dotnet_exception(
        chunk,
        "ArgumentOutOfRangeException",
        "Index and length must refer to a location within the string.",
        line,
    );
    chunk.emit_end(line);
}

pub fn emit_string_pad_left(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_string_pad(chunks, current, argc, line, true);
}

pub fn emit_string_pad_right(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_string_pad(chunks, current, argc, line, false);
}

fn emit_string_pad(chunks: &mut [Chunk], current: usize, argc: u8, line: u32, left: bool) {
    let chunk = &mut chunks[current];
    let pad_slot = reserve_slot(chunk);
    let width_slot = reserve_slot(chunk);
    let value_slot = reserve_slot(chunk);

    if argc >= 3 {
        chunk.emit_op_u16(Op::LOCAL_SET, pad_slot, line);
    } else {
        chunk.emit_string_const(" ", line);
        chunk.emit_op_u16(Op::LOCAL_SET, pad_slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, width_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, width_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, pad_slot, line);
    host::emit(
        chunk,
        "ecma:string",
        if left { "padStart" } else { "padEnd" },
        3,
        line,
    );
}

pub fn emit_string_replace(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    host::emit(&mut chunks[current], "ecma:string", "replaceAll", 3, line);
}

pub fn emit_string_concat(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    vybe_compiler::compiler::strings::emit_concat(&mut chunks[current], argc as usize, line);
}

pub fn emit_string_split(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = reserve_slot(chunk);
    let delims_slot = reserve_slot(chunk);
    let remove_empty_slot = reserve_slot(chunk);

    match argc {
        0 | 1 => {
            chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
            chunk.emit_string_const(" ", line);
            vybe_compiler::compiler::collections::emit_array_new(chunks, current, 1, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, delims_slot, line);
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, remove_empty_slot, line);
        }
        2 => {
            let arg_slot = reserve_slot(&mut chunks[current]);
            chunks[current].emit_op_u16(Op::LOCAL_SET, arg_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arg_slot, line);
            host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
            vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arg_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, delims_slot, line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arg_slot, line);
            vybe_compiler::compiler::collections::emit_array_new(chunks, current, 1, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, delims_slot, line);
            chunks[current].emit_end(line);
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, remove_empty_slot, line);
        }
        3 => {
            let second_slot = reserve_slot(&mut chunks[current]);
            let first_slot = reserve_slot(&mut chunks[current]);
            chunks[current].emit_op_u16(Op::LOCAL_SET, second_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, first_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

            chunks[current].emit_op_u16(Op::LOCAL_GET, second_slot, line);
            chunks[current].emit_string_const("__dotnet_stringsplit_removeemptyentries", line);
            vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_bool_const(true, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, remove_empty_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, first_slot, line);
            host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
            vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, first_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, delims_slot, line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, first_slot, line);
            vybe_compiler::compiler::collections::emit_array_new(chunks, current, 1, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, delims_slot, line);
            chunks[current].emit_end(line);
            chunks[current].emit_else(line);
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, remove_empty_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, first_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, second_slot, line);
            vybe_compiler::compiler::collections::emit_array_new(chunks, current, 2, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, delims_slot, line);
            chunks[current].emit_end(line);
        }
        _ => {
            let arg_base = chunks[current].alloc_scratch(argc as u16);
            for i in (0..argc).rev() {
                chunks[current].emit_op_u16(Op::LOCAL_SET, arg_base + i as u16, line);
            }
            chunks[current].emit_op_u16(Op::LOCAL_GET, arg_base, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
            for i in 1..argc {
                chunks[current].emit_op_u16(Op::LOCAL_GET, arg_base + i as u16, line);
            }
            vybe_compiler::compiler::collections::emit_array_new(
                chunks,
                current,
                (argc - 1).into(),
                line,
            );
            chunks[current].emit_op_u16(Op::LOCAL_SET, delims_slot, line);
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, remove_empty_slot, line);
        }
    }

    emit_string_split_slots(
        chunks,
        current,
        value_slot,
        delims_slot,
        remove_empty_slot,
        line,
    );
}

pub fn emit_vb_strings_left(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let len_slot = reserve_slot(&mut chunks[current]);
    let value_slot = reserve_slot(&mut chunks[current]);
    let start_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_slot, line);
    emit_string_substr_from_slots(&mut chunks[current], value_slot, start_slot, len_slot, line);
}

pub fn emit_vb_strings_right(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let take_slot = reserve_slot(&mut chunks[current]);
    let value_slot = reserve_slot(&mut chunks[current]);
    let len_slot = reserve_slot(&mut chunks[current]);
    let start_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, take_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    emit_string_len(&mut chunks[current], value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, take_slot, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_slot, line);
    emit_string_substr_from_slots(
        &mut chunks[current],
        value_slot,
        start_slot,
        take_slot,
        line,
    );
}

pub fn emit_vb_strings_mid(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let len_slot = reserve_slot(&mut chunks[current]);
    let start_one_slot = reserve_slot(&mut chunks[current]);
    let value_slot = reserve_slot(&mut chunks[current]);
    let start_slot = reserve_slot(&mut chunks[current]);
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_one_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_one_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_slot, line);
    if argc < 3 {
        emit_string_len(&mut chunks[current], value_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    }
    emit_string_substr_from_slots(&mut chunks[current], value_slot, start_slot, len_slot, line);
}

fn emit_string_split_slots(
    chunks: &mut [Chunk],
    current: usize,
    value_slot: u16,
    delims_slot: u16,
    remove_empty_slot: u16,
    line: u32,
) {
    let source_slot = reserve_slot(&mut chunks[current]);
    let delim0_slot = reserve_slot(&mut chunks[current]);
    let len_slot = reserve_slot(&mut chunks[current]);
    let i_slot = reserve_slot(&mut chunks[current]);
    let delim_slot = reserve_slot(&mut chunks[current]);
    let parts_slot = reserve_slot(&mut chunks[current]);
    let result_slot = reserve_slot(&mut chunks[current]);
    let part_slot = reserve_slot(&mut chunks[current]);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delims_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const(" ", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delims_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, delim0_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let normalize_loop = vybe_compiler::compiler::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::compiler::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delims_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, delim_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delim_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delim0_slot, line);
    host::emit(&mut chunks[current], "ecma:string", "replaceAll", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_compiler::compiler::loops::emit_loop_end(chunks, current, normalize_loop, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delim0_slot, line);
    host::emit(&mut chunks[current], "ecma:string", "split", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parts_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, remove_empty_slot, line);
    chunks[current].emit_if(line);
    vybe_compiler::compiler::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let filter_loop = vybe_compiler::compiler::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parts_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::compiler::loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parts_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, part_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, part_slot, line);
    chunks[current].emit_string_const("", line);
    vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, part_slot, line);
    vybe_compiler::compiler::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_compiler::compiler::loops::emit_loop_end(chunks, current, filter_loop, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parts_slot, line);
    chunks[current].emit_end(line);
}

fn emit_char_set_contains(chunk: &mut Chunk, set_slot: u16, char_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, set_slot, line);
    host::emit(chunk, "ecma:array", "isArray", 1, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, set_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, char_slot, line);
    host::emit(chunk, "ecma:array", "includes", 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, char_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, set_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    chunk.emit_end(line);
}

pub fn emit_string_trim_chars(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_string_trim_chars_mode(chunks, current, argc, line, true, true);
}

pub fn emit_string_trim_start_chars(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_string_trim_chars_mode(chunks, current, argc, line, true, false);
}

pub fn emit_string_trim_end_chars(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_string_trim_chars_mode(chunks, current, argc, line, false, true);
}

fn emit_string_trim_chars_mode(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
    trim_start: bool,
    trim_end: bool,
) {
    let set_slot = reserve_slot(&mut chunks[current]);
    if argc <= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, set_slot, line);
    } else {
        vybe_compiler::compiler::collections::emit_array_new(
            chunks,
            current,
            (argc - 1).into(),
            line,
        );
        chunks[current].emit_op_u16(Op::LOCAL_SET, set_slot, line);
    }
    let chunk = &mut chunks[current];
    let value_slot = reserve_slot(chunk);
    let len_slot = reserve_slot(chunk);
    let start_slot = reserve_slot(chunk);
    let end_slot = reserve_slot(chunk);
    let char_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    emit_string_len(chunk, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, start_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, end_slot, line);

    if trim_start {
        let state = vybe_compiler::compiler::loops::emit_loop_start(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, end_slot, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        vybe_compiler::compiler::loops::emit_loop_cond(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
        host::emit(&mut chunks[current], "ecma:string", "charAt", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, char_slot, line);
        emit_char_set_contains(&mut chunks[current], set_slot, char_slot, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, start_slot, line);
        chunks[current].emit_else(line);
        chunks[current].emit_br(2, line);
        chunks[current].emit_end(line);
        vybe_compiler::compiler::loops::emit_loop_end(chunks, current, state, line);
    }

    if trim_end {
        let state = vybe_compiler::compiler::loops::emit_loop_start(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, end_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
        chunks[current].emit_op(Op::I32_GT_S, line);
        vybe_compiler::compiler::loops::emit_loop_cond(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, end_slot, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        host::emit(&mut chunks[current], "ecma:string", "charAt", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, char_slot, line);
        emit_char_set_contains(&mut chunks[current], set_slot, char_slot, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, end_slot, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, end_slot, line);
        chunks[current].emit_else(line);
        chunks[current].emit_br(2, line);
        chunks[current].emit_end(line);
        vybe_compiler::compiler::loops::emit_loop_end(chunks, current, state, line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_slot, line);
    host::emit(&mut chunks[current], "wasm:js-string", "substring", 3, line);
}

pub fn emit_string_from_chars(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
        return;
    }
    if argc >= 3 {
        let len_slot = reserve_slot(&mut chunks[current]);
        let start_slot = reserve_slot(&mut chunks[current]);
        let chars_slot = reserve_slot(&mut chunks[current]);
        let end_slot = reserve_slot(&mut chunks[current]);
        let slice_slot = reserve_slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, start_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, chars_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, end_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, chars_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, end_slot, line);
        vybe_compiler::compiler::collections::emit_slice(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slice_slot, line);
        emit_char_array_to_string(chunks, current, slice_slot, line);
        return;
    }
    let chunk = &mut chunks[current];
    let chars_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, chars_slot, line);
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }

    let string_test_idx = chunks[current].add_import("wasm:js-string", "test");
    chunks[current].emit_op_u16(Op::LOCAL_GET, chars_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, string_test_idx, line);
    chunks[current].emit(1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, chars_slot, line);
    chunks[current].emit_else(line);
    emit_char_array_to_string(chunks, current, chars_slot, line);
    chunks[current].emit_end(line);
}

fn emit_char_array_to_string(chunks: &mut [Chunk], current: usize, chars_slot: u16, line: u32) {
    let char_code_idx = chunks[current].add_import("wasm:js-string", "charCodeAt");
    let from_chars_idx = chunks[current].add_import("wasm:js-string", "fromCharCodeArray");
    let string_test_idx = chunks[current].add_import("wasm:js-string", "test");
    let chunk = &mut chunks[current];
    let units_slot = reserve_slot(chunk);
    let len_slot = reserve_slot(chunk);
    let i_slot = reserve_slot(chunk);
    let elem_slot = reserve_slot(chunk);

    vybe_compiler::compiler::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, units_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, chars_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let state = vybe_compiler::compiler::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::compiler::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, chars_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, units_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, string_test_idx, line);
    chunks[current].emit(1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    host::emit(&mut chunks[current], "ecma:string", "String", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, char_code_idx, line);
    chunks[current].emit(2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_end(line);
    vybe_compiler::compiler::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_compiler::compiler::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, units_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, units_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, from_chars_idx, line);
    chunks[current].emit(3, line);
}
