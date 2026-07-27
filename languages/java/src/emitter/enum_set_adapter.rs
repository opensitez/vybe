//! Java EnumSet backed by an array of enum display names.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use vybe_compiler::compiler::{
    collections,
    instructions::{core_wasm, host},
};

const NAMES_KEY: &str = "__java_enum_names";
const CLASS_KEY: &str = "__java_class_name";

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn attach_names(chunks: &mut [Chunk], current: usize, set_slot: u16, names_slot: u16, line: u32) {
    get(&mut chunks[current], set_slot, line);
    chunks[current].emit_string_const(NAMES_KEY, line);
    get(&mut chunks[current], names_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], set_slot, line);
    chunks[current].emit_string_const(CLASS_KEY, line);
    chunks[current].emit_string_const("EnumSet", line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_names(chunks: &mut [Chunk], current: usize, set_slot: u16, line: u32) {
    get(&mut chunks[current], set_slot, line);
    chunks[current].emit_string_const(NAMES_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
}

fn emit_name_for_value(
    chunks: &mut [Chunk],
    current: usize,
    set_slot: u16,
    value_slot: u16,
    line: u32,
) {
    emit_names(chunks, current, set_slot, line);
    get(&mut chunks[current], value_slot, line);
    collections::emit_get(chunks, current, line);
}

fn emit_contains_name(
    chunks: &mut [Chunk],
    current: usize,
    set_slot: u16,
    name_slot: u16,
    line: u32,
) {
    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], name_slot, line);
    collections::emit_contains(chunks, current, line);
}

fn push_name_if_absent(
    chunks: &mut [Chunk],
    current: usize,
    set_slot: u16,
    name_slot: u16,
    line: u32,
) {
    emit_contains_name(chunks, current, set_slot, name_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], name_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

pub fn emit_none_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let names = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], names, line);
    collections::emit_array_new(chunks, current, 0, line);
    set(&mut chunks[current], out, line);
    attach_names(chunks, current, out, names, line);
    get(&mut chunks[current], out, line);
}

pub fn emit_all_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let names = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], names, line);
    get(&mut chunks[current], names, line);
    collections::emit_clone(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    attach_names(chunks, current, out, names, line);
    get(&mut chunks[current], out, line);
}

pub fn emit_of(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let value_count = argc.saturating_sub(1);
    let values = chunks[current].alloc_scratch(value_count.max(1) as u16);
    for i in (0..value_count).rev() {
        set(&mut chunks[current], values + i as u16, line);
    }
    let names = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], names, line);
    collections::emit_array_new(chunks, current, 0, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    attach_names(chunks, current, out, names, line);
    for i in 0..value_count {
        get(&mut chunks[current], names, line);
        get(&mut chunks[current], values + i as u16, line);
        collections::emit_get(chunks, current, line);
        let name = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], name, line);
        push_name_if_absent(chunks, current, out, name, line);
    }
    get(&mut chunks[current], out, line);
}

pub fn emit_copy_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);
    get(&mut chunks[current], source, line);
    collections::emit_clone(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], source, line);
    chunks[current].emit_string_const(NAMES_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    let names = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], names, line);
    attach_names(chunks, current, out, names, line);
    get(&mut chunks[current], out, line);
}

pub fn emit_complement_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);
    emit_names(chunks, current, source, line);
    let names = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], names, line);
    collections::emit_array_new(chunks, current, 0, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    attach_names(chunks, current, out, names, line);

    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let name = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], i, line);
    get(&mut chunks[current], names, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    let _block = chunks[current].emit_block(line);
    let (_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], len, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], names, line);
    get(&mut chunks[current], i, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], name, line);
    emit_contains_name(chunks, current, source, name, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], name, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::compiler::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], out, line);
}

pub fn emit_range(chunks: &mut [Chunk], current: usize, line: u32) {
    let end = chunks[current].alloc_scratch(1);
    let start = chunks[current].alloc_scratch(1);
    let names = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], end, line);
    set(&mut chunks[current], start, line);
    set(&mut chunks[current], names, line);
    collections::emit_array_new(chunks, current, 0, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    attach_names(chunks, current, out, names, line);
    let i = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], start, line);
    set(&mut chunks[current], i, line);
    let _block = chunks[current].emit_block(line);
    let (_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], end, line);
    vybe_compiler::compiler::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], names, line);
    get(&mut chunks[current], i, line);
    collections::emit_get(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::compiler::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], out, line);
}

pub fn emit_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let set_slot = chunks[current].alloc_scratch(1);
    let name = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], set_slot, line);
    emit_name_for_value(chunks, current, set_slot, value, line);
    set(&mut chunks[current], name, line);
    emit_contains_name(chunks, current, set_slot, name, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], name, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_end(line);
}

pub fn emit_add_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    let target = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let name = chunks[current].alloc_scratch(1);
    let changed = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);
    set(&mut chunks[current], target, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], i, line);
    get(&mut chunks[current], source, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], changed, line);
    let _block = chunks[current].emit_block(line);
    let (_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], len, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], source, line);
    get(&mut chunks[current], i, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], name, line);
    emit_contains_name(chunks, current, target, name, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], target, line);
    get(&mut chunks[current], name, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], changed, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::compiler::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], changed, line);
}

pub fn emit_contains(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let set_slot = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], set_slot, line);
    emit_name_for_value(chunks, current, set_slot, value, line);
    let name = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], name, line);
    emit_contains_name(chunks, current, set_slot, name, line);
}

pub fn emit_contains_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    let target = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let name = chunks[current].alloc_scratch(1);
    let ok = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);
    set(&mut chunks[current], target, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], ok, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], i, line);
    get(&mut chunks[current], source, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    let _block = chunks[current].emit_block(line);
    let (_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], len, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], source, line);
    get(&mut chunks[current], i, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], name, line);
    emit_contains_name(chunks, current, target, name, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], ok, line);
    get(&mut chunks[current], len, line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::compiler::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], ok, line);
}

pub fn emit_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let other = chunks[current].alloc_scratch(1);
    let set_slot = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], other, line);
    set(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], set_slot, line);
    collections::emit_len(chunks, current, line);
    get(&mut chunks[current], other, line);
    collections::emit_len(chunks, current, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], other, line);
    emit_contains_all(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

pub fn emit_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
}

pub fn emit_iterator(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:array", "values", 1, line);
}

pub fn emit_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let set_slot = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], set_slot, line);
    emit_name_for_value(chunks, current, set_slot, value, line);
    let name = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], name, line);
    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], name, line);
    collections::emit_remove_value(chunks, current, line);
}

pub fn emit_get_class(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_string_const("EnumSet", line);
}
