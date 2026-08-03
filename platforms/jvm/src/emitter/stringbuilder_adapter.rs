//! JVM `java.lang.StringBuilder` adapter.

use vybe_compiler::primitives::{instructions::core_wasm, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const BUFFER_KEY: &str = "__buffer";

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

pub fn emit_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let initial = chunks[current].alloc_scratch(1);
    if argc > 0 {
        set(&mut chunks[current], initial, line);
        for _ in 1..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    } else {
        chunks[current].emit_string_const("", line);
        set(&mut chunks[current], initial, line);
    }

    chunks[current].emit_struct_new(0, 0, line);
    core_wasm::dup(&mut chunks[current], line);
    get(&mut chunks[current], initial, line);
    let key = chunks[current].add_constant(vybe_runtime::Value::String(BUFFER_KEY.into()));
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    chunks[current].emit_op(Op::DROP, line);
}

pub fn emit_append(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let sb = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    if argc > 2 {
        for _ in 2..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    }
    set(&mut chunks[current], sb, line);
    append_slot(chunks, current, sb, value, false, line);
}

pub fn emit_append_line(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let sb = chunks[current].alloc_scratch(1);
    if argc > 1 {
        set(&mut chunks[current], value, line);
        if argc > 2 {
            for _ in 2..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
        }
    } else {
        chunks[current].emit_string_const("", line);
        set(&mut chunks[current], value, line);
    }
    set(&mut chunks[current], sb, line);
    append_slot(chunks, current, sb, value, true, line);
}

fn append_slot(
    chunks: &mut [Chunk],
    current: usize,
    sb: u16,
    value: u16,
    newline: bool,
    line: u32,
) {
    let key = chunks[current].add_constant(vybe_runtime::Value::String(BUFFER_KEY.into()));
    get(&mut chunks[current], sb, line);
    get(&mut chunks[current], sb, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    get(&mut chunks[current], value, line);
    strings::emit_str_concat_coercing(&mut chunks[current], line);
    if newline {
        chunks[current].emit_string_const("\n", line);
        strings::emit_str_concat(&mut chunks[current], line);
    }
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], sb, line);
}

pub fn emit_to_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 1 {
        for _ in 1..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    }
    let key = chunks[current].add_constant(vybe_runtime::Value::String(BUFFER_KEY.into()));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
}
