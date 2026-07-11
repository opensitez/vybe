//! Java collection overloads composed from the shared ECMA array surface.

use crate::emitter::{collections, instructions::core_wasm};
use vybe_bytecode::opcode::Op;
use vybe_bytecode::Chunk;

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

pub fn emit_arrays_as_list(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 1 {
        collections::emit_array_new(chunks, current, argc as u16, line);
        return;
    }

    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    let length = chunks[current].add_import("ecma:array", "length");
    chunks[current].emit_call(length, 1, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    collections::emit_array_new(chunks, current, 1, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
}

pub fn emit_n_copies(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    collections::emit_new_with_length(chunks, current, line);
    get(&mut chunks[current], value, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_i32_const(i32::MAX, line);
    collections::emit_fill(chunks, current, line);
}

pub fn emit_add(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    if argc == 2 {
        let list = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], list, line);
        get(&mut chunks[current], list, line);
        get(&mut chunks[current], value, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_bool_const(true, line);
    } else {
        let index = chunks[current].alloc_scratch(1);
        let list = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], index, line);
        set(&mut chunks[current], list, line);
        get(&mut chunks[current], list, line);
        get(&mut chunks[current], index, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        get(&mut chunks[current], value, line);
        collections::emit_insert(chunks, current, line);
    }
}

pub fn emit_add_all(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);
    let index = if argc == 3 {
        let slot = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], slot, line);
        Some(slot)
    } else {
        None
    };
    let list = chunks[current].alloc_scratch(1);
    let changed = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);

    get(&mut chunks[current], source, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    crate::emitter::ops::emit_dyn_gt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    set(&mut chunks[current], changed, line);

    get(&mut chunks[current], list, line);
    if let Some(index) = index {
        get(&mut chunks[current], index, line);
    } else {
        get(&mut chunks[current], list, line);
        collections::emit_len(chunks, current, line);
    }
    get(&mut chunks[current], source, line);
    collections::emit_insert_range(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], changed, line);
}

pub fn emit_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    let previous = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], index, line);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], previous, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], previous, line);
}

pub fn emit_sort(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 1 {
        collections::emit_sort(chunks, current, line);
        return;
    }
    let comparator = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], comparator, line);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], comparator, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], list, line);
    collections::emit_sort(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], comparator, line);
    collections::emit_runtime_helper_call(chunks, current, "__vybe_sort_with_comparator", 2, line);
    chunks[current].emit_end(line);
}

pub fn emit_remove_all(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_filter_members(chunks, current, line, false);
}

pub fn emit_retain_all(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_filter_members(chunks, current, line, true);
}

fn emit_filter_members(chunks: &mut [Chunk], current: usize, line: u32, retain: bool) {
    let members = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    let snapshot = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let length = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let changed = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], members, line);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    collections::emit_clone(chunks, current, line);
    set(&mut chunks[current], snapshot, line);
    get(&mut chunks[current], snapshot, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], length, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], changed, line);

    let outer = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], length, line);
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], snapshot, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], members, line);
    get(&mut chunks[current], value, line);
    collections::emit_contains(chunks, current, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    if retain {
        chunks[current].emit_op(Op::I32_EQZ, line);
    }
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    collections::emit_remove_value(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], changed, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    crate::emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], changed, line);
}
