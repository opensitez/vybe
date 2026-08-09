//! Java BitSet backed by `ecma:set` members containing set bit indexes.

use vybe_compiler::primitives::{
    collections,
    instructions::{core_wasm, host},
};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set_local(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn set_add(chunks: &mut [Chunk], current: usize, set_slot: u16, value_slot: u16, line: u32) {
    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], value_slot, line);
    host::emit(&mut chunks[current], "ecma:set", "add", 2, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn set_delete(chunks: &mut [Chunk], current: usize, set_slot: u16, value_slot: u16, line: u32) {
    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], value_slot, line);
    host::emit(&mut chunks[current], "ecma:set", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn set_has(chunks: &mut [Chunk], current: usize, set_slot: u16, value_slot: u16, line: u32) {
    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], value_slot, line);
    host::emit(&mut chunks[current], "ecma:set", "has", 2, line);
}

fn values_snapshot(chunks: &mut [Chunk], current: usize, set_slot: u16, out_slot: u16, line: u32) {
    get(&mut chunks[current], set_slot, line);
    host::emit(&mut chunks[current], "ecma:set", "values", 1, line);
    set_local(&mut chunks[current], out_slot, line);
}

pub fn emit_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    host::emit(&mut chunks[current], "ecma:set", "new", 0, line);
}

pub fn emit_value_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    set_local(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:set", "new", 0, line);
    set_local(&mut chunks[current], out, line);
    get(&mut chunks[current], out, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    host::emit(&mut chunks[current], "ecma:set", "add", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], value, line);
    core_wasm::i32_const(&mut chunks[current], line, 5);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], out, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    host::emit(&mut chunks[current], "ecma:set", "add", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], out, line);
}

pub fn emit_cardinality(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:set", "size", 1, line);
}

pub fn emit_is_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_cardinality(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_size(chunks: &mut [Chunk], current: usize, line: u32) {
    let bs = chunks[current].alloc_scratch(1);
    set_local(&mut chunks[current], bs, line);
    get(&mut chunks[current], bs, line);
    emit_length(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 8);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, 8);
    chunks[current].emit_else(line);
    get(&mut chunks[current], bs, line);
    emit_length(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_get(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 3 {
        let to = chunks[current].alloc_scratch(1);
        let from = chunks[current].alloc_scratch(1);
        let bs = chunks[current].alloc_scratch(1);
        let out = chunks[current].alloc_scratch(1);
        let i = chunks[current].alloc_scratch(1);
        set_local(&mut chunks[current], to, line);
        set_local(&mut chunks[current], from, line);
        set_local(&mut chunks[current], bs, line);
        host::emit(&mut chunks[current], "ecma:set", "new", 0, line);
        set_local(&mut chunks[current], out, line);
        get(&mut chunks[current], from, line);
        set_local(&mut chunks[current], i, line);
        range_loop(chunks, current, i, to, line, |chunks, current, i, line| {
            set_has(chunks, current, bs, i, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if_value(line);
            set_add(chunks, current, out, i, line);
            chunks[current].emit_end(line);
        });
        get(&mut chunks[current], out, line);
        return;
    }
    let index = chunks[current].alloc_scratch(1);
    let bs = chunks[current].alloc_scratch(1);
    set_local(&mut chunks[current], index, line);
    set_local(&mut chunks[current], bs, line);
    set_has(chunks, current, bs, index, line);
}

pub fn emit_set(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 2 {
        let index = chunks[current].alloc_scratch(1);
        let bs = chunks[current].alloc_scratch(1);
        set_local(&mut chunks[current], index, line);
        set_local(&mut chunks[current], bs, line);
        set_add(chunks, current, bs, index, line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }

    let second = chunks[current].alloc_scratch(1);
    let first = chunks[current].alloc_scratch(1);
    let bs = chunks[current].alloc_scratch(1);
    set_local(&mut chunks[current], second, line);
    set_local(&mut chunks[current], first, line);
    set_local(&mut chunks[current], bs, line);
    get(&mut chunks[current], second, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("boolean", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], second, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    set_add(chunks, current, bs, first, line);
    chunks[current].emit_else(line);
    set_delete(chunks, current, bs, first, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    let i = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], first, line);
    set_local(&mut chunks[current], i, line);
    range_loop(
        chunks,
        current,
        i,
        second,
        line,
        |chunks, current, i, line| {
            set_add(chunks, current, bs, i, line);
        },
    );
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_clear(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 1 {
        host::emit(&mut chunks[current], "ecma:set", "clear", 1, line);
        return;
    }
    let end = if argc == 3 {
        let slot = chunks[current].alloc_scratch(1);
        set_local(&mut chunks[current], slot, line);
        Some(slot)
    } else {
        None
    };
    let start = chunks[current].alloc_scratch(1);
    let bs = chunks[current].alloc_scratch(1);
    set_local(&mut chunks[current], start, line);
    set_local(&mut chunks[current], bs, line);
    if let Some(end) = end {
        let i = chunks[current].alloc_scratch(1);
        get(&mut chunks[current], start, line);
        set_local(&mut chunks[current], i, line);
        range_loop(chunks, current, i, end, line, |chunks, current, i, line| {
            set_delete(chunks, current, bs, i, line);
        });
    } else {
        set_delete(chunks, current, bs, start, line);
    }
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_flip(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let end = if argc == 3 {
        let slot = chunks[current].alloc_scratch(1);
        set_local(&mut chunks[current], slot, line);
        Some(slot)
    } else {
        None
    };
    let start = chunks[current].alloc_scratch(1);
    let bs = chunks[current].alloc_scratch(1);
    set_local(&mut chunks[current], start, line);
    set_local(&mut chunks[current], bs, line);
    if let Some(end) = end {
        let i = chunks[current].alloc_scratch(1);
        get(&mut chunks[current], start, line);
        set_local(&mut chunks[current], i, line);
        range_loop(chunks, current, i, end, line, |chunks, current, i, line| {
            flip_one(chunks, current, bs, i, line);
        });
    } else {
        flip_one(chunks, current, bs, start, line);
    }
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

fn flip_one(chunks: &mut [Chunk], current: usize, bs: u16, index: u16, line: u32) {
    set_has(chunks, current, bs, index, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    set_delete(chunks, current, bs, index, line);
    chunks[current].emit_else(line);
    set_add(chunks, current, bs, index, line);
    chunks[current].emit_end(line);
}

fn range_loop<F: Fn(&mut [Chunk], usize, u16, u32)>(
    chunks: &mut [Chunk],
    current: usize,
    i: u16,
    end: u16,
    line: u32,
    body: F,
) {
    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], end, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    body(chunks, current, i, line);
    get(&mut chunks[current], i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set_local(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
}

pub fn emit_length(chunks: &mut [Chunk], current: usize, line: u32) {
    let bs = chunks[current].alloc_scratch(1);
    let vals = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let max = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    set_local(&mut chunks[current], bs, line);
    values_snapshot(chunks, current, bs, vals, line);
    get(&mut chunks[current], vals, line);
    collections::emit_len(chunks, current, line);
    set_local(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    set_local(&mut chunks[current], max, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set_local(&mut chunks[current], i, line);
    range_loop(chunks, current, i, len, line, |chunks, current, i, line| {
        get(&mut chunks[current], vals, line);
        get(&mut chunks[current], i, line);
        collections::emit_get(chunks, current, line);
        set_local(&mut chunks[current], value, line);
        get(&mut chunks[current], value, line);
        get(&mut chunks[current], max, line);
        vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], value, line);
        set_local(&mut chunks[current], max, line);
        chunks[current].emit_end(line);
    });
    get(&mut chunks[current], max, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
}

pub fn emit_next_set_bit(chunks: &mut [Chunk], current: usize, line: u32) {
    scan_set_bit(chunks, current, true, line);
}

pub fn emit_previous_set_bit(chunks: &mut [Chunk], current: usize, line: u32) {
    scan_set_bit(chunks, current, false, line);
}

fn scan_set_bit(chunks: &mut [Chunk], current: usize, forward: bool, line: u32) {
    let start = chunks[current].alloc_scratch(1);
    let bs = chunks[current].alloc_scratch(1);
    let vals = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    set_local(&mut chunks[current], start, line);
    set_local(&mut chunks[current], bs, line);
    values_snapshot(chunks, current, bs, vals, line);
    get(&mut chunks[current], vals, line);
    collections::emit_len(chunks, current, line);
    set_local(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    set_local(&mut chunks[current], result, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set_local(&mut chunks[current], i, line);
    range_loop(chunks, current, i, len, line, |chunks, current, i, line| {
        get(&mut chunks[current], vals, line);
        get(&mut chunks[current], i, line);
        collections::emit_get(chunks, current, line);
        set_local(&mut chunks[current], value, line);
        get(&mut chunks[current], result, line);
        core_wasm::i32_const(&mut chunks[current], line, -1);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        if forward {
            get(&mut chunks[current], value, line);
            get(&mut chunks[current], start, line);
            vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
            vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
        } else {
            get(&mut chunks[current], value, line);
            get(&mut chunks[current], start, line);
            vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
            vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
        }
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], value, line);
        set_local(&mut chunks[current], result, line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);
        if forward {
            get(&mut chunks[current], value, line);
            get(&mut chunks[current], start, line);
            vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
            vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if_value(line);
            get(&mut chunks[current], value, line);
            get(&mut chunks[current], result, line);
            vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
        } else {
            get(&mut chunks[current], value, line);
            get(&mut chunks[current], start, line);
            vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
            vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if_value(line);
            get(&mut chunks[current], value, line);
            get(&mut chunks[current], result, line);
            vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
        }
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], value, line);
        set_local(&mut chunks[current], result, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    });
    get(&mut chunks[current], result, line);
}

pub fn emit_next_clear_bit(chunks: &mut [Chunk], current: usize, line: u32) {
    let start = chunks[current].alloc_scratch(1);
    let bs = chunks[current].alloc_scratch(1);
    set_local(&mut chunks[current], start, line);
    set_local(&mut chunks[current], bs, line);
    set_has(chunks, current, bs, start, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], start, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], start, line);
    chunks[current].emit_end(line);
}

pub fn emit_previous_clear_bit(chunks: &mut [Chunk], current: usize, line: u32) {
    let start = chunks[current].alloc_scratch(1);
    set_local(&mut chunks[current], start, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], start, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
}

pub fn emit_and(chunks: &mut [Chunk], current: usize, line: u32) {
    mutate_with_other(
        chunks,
        current,
        line,
        |chunks, current, target, other, value, line| {
            set_has(chunks, current, other, value, line);
            vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
            chunks[current].emit_if_value(line);
            set_delete(chunks, current, target, value, line);
            chunks[current].emit_end(line);
        },
    );
}

pub fn emit_or(chunks: &mut [Chunk], current: usize, line: u32) {
    mutate_from_other(
        chunks,
        current,
        line,
        |chunks, current, target, value, line| {
            set_add(chunks, current, target, value, line);
        },
    );
}

pub fn emit_xor(chunks: &mut [Chunk], current: usize, line: u32) {
    mutate_from_other(
        chunks,
        current,
        line,
        |chunks, current, target, value, line| {
            flip_one(chunks, current, target, value, line);
        },
    );
}

pub fn emit_and_not(chunks: &mut [Chunk], current: usize, line: u32) {
    mutate_from_other(
        chunks,
        current,
        line,
        |chunks, current, target, value, line| {
            set_delete(chunks, current, target, value, line);
        },
    );
}

fn mutate_with_other<F: Fn(&mut [Chunk], usize, u16, u16, u16, u32)>(
    chunks: &mut [Chunk],
    current: usize,
    line: u32,
    body: F,
) {
    let other = chunks[current].alloc_scratch(1);
    let target = chunks[current].alloc_scratch(1);
    let vals = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    set_local(&mut chunks[current], other, line);
    set_local(&mut chunks[current], target, line);
    values_snapshot(chunks, current, target, vals, line);
    get(&mut chunks[current], vals, line);
    collections::emit_len(chunks, current, line);
    set_local(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set_local(&mut chunks[current], i, line);
    range_loop(chunks, current, i, len, line, |chunks, current, i, line| {
        get(&mut chunks[current], vals, line);
        get(&mut chunks[current], i, line);
        collections::emit_get(chunks, current, line);
        set_local(&mut chunks[current], value, line);
        body(chunks, current, target, other, value, line);
    });
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

fn mutate_from_other<F: Fn(&mut [Chunk], usize, u16, u16, u32)>(
    chunks: &mut [Chunk],
    current: usize,
    line: u32,
    body: F,
) {
    let other = chunks[current].alloc_scratch(1);
    let target = chunks[current].alloc_scratch(1);
    let vals = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    set_local(&mut chunks[current], other, line);
    set_local(&mut chunks[current], target, line);
    values_snapshot(chunks, current, other, vals, line);
    get(&mut chunks[current], vals, line);
    collections::emit_len(chunks, current, line);
    set_local(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set_local(&mut chunks[current], i, line);
    range_loop(chunks, current, i, len, line, |chunks, current, i, line| {
        get(&mut chunks[current], vals, line);
        get(&mut chunks[current], i, line);
        collections::emit_get(chunks, current, line);
        set_local(&mut chunks[current], value, line);
        body(chunks, current, target, value, line);
    });
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_intersects(chunks: &mut [Chunk], current: usize, line: u32) {
    let other = chunks[current].alloc_scratch(1);
    let target = chunks[current].alloc_scratch(1);
    let vals = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    set_local(&mut chunks[current], other, line);
    set_local(&mut chunks[current], target, line);
    chunks[current].emit_bool_const(false, line);
    set_local(&mut chunks[current], result, line);
    values_snapshot(chunks, current, target, vals, line);
    get(&mut chunks[current], vals, line);
    collections::emit_len(chunks, current, line);
    set_local(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set_local(&mut chunks[current], i, line);
    range_loop(chunks, current, i, len, line, |chunks, current, i, line| {
        get(&mut chunks[current], vals, line);
        get(&mut chunks[current], i, line);
        collections::emit_get(chunks, current, line);
        set_local(&mut chunks[current], value, line);
        set_has(chunks, current, other, value, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_bool_const(true, line);
        set_local(&mut chunks[current], result, line);
        chunks[current].emit_end(line);
    });
    get(&mut chunks[current], result, line);
}

pub fn emit_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let other = chunks[current].alloc_scratch(1);
    let target = chunks[current].alloc_scratch(1);
    set_local(&mut chunks[current], other, line);
    set_local(&mut chunks[current], target, line);
    get(&mut chunks[current], target, line);
    get(&mut chunks[current], other, line);
    emit_intersects_like_equals(chunks, current, target, other, line);
}

fn emit_intersects_like_equals(
    chunks: &mut [Chunk],
    current: usize,
    target: u16,
    other: u16,
    line: u32,
) {
    get(&mut chunks[current], target, line);
    host::emit(&mut chunks[current], "ecma:set", "size", 1, line);
    get(&mut chunks[current], other, line);
    host::emit(&mut chunks[current], "ecma:set", "size", 1, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], target, line);
    host::emit(&mut chunks[current], "ecma:set", "size", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], target, line);
    get(&mut chunks[current], other, line);
    emit_intersects(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

pub fn emit_clone(chunks: &mut [Chunk], current: usize, line: u32) {
    let original = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    set_local(&mut chunks[current], original, line);
    host::emit(&mut chunks[current], "ecma:set", "new", 0, line);
    set_local(&mut chunks[current], out, line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], original, line);
    emit_or(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], out, line);
}

pub fn emit_stream(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:set", "values", 1, line);
}

pub fn emit_to_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let bs = chunks[current].alloc_scratch(1);
    set_local(&mut chunks[current], bs, line);
    get(&mut chunks[current], bs, line);
    host::emit(&mut chunks[current], "ecma:set", "size", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_array_new(chunks, current, 1, line);
    chunks[current].emit_else(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_end(line);
}

pub fn emit_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:set", "values", 1, line);
    chunks[current].emit_string_const(",", line);
    host::emit(&mut chunks[current], "ecma:array", "join", 2, line);
}

pub fn emit_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_cardinality(chunks, current, line);
}
