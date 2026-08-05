//! JVM `java.util.StringTokenizer` adapter.

use vybe_compiler::primitives::{instructions::host, ops};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const TOKENS_KEY: &str = "__tokens";
const INDEX_KEY: &str = "__index";

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn prop_get(chunk: &mut Chunk, obj: u16, key_name: &str, line: u32) {
    let key = chunk.add_constant(vybe_runtime::Value::String(key_name.into()));
    get(chunk, obj, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
}

fn prop_set_from_slot(chunk: &mut Chunk, obj: u16, key_name: &str, value: u16, line: u32) {
    let key = chunk.add_constant(vybe_runtime::Value::String(key_name.into()));
    get(chunk, obj, line);
    get(chunk, value, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
}

pub fn emit_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let source = chunks[current].alloc_scratch(2);
    let delim = source + 1;
    if argc > 2 {
        for _ in 2..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    }
    if argc > 1 {
        set(&mut chunks[current], delim, line);
    } else {
        chunks[current].emit_string_const(" \t\n\r\u{000c}", line);
        set(&mut chunks[current], delim, line);
    }
    if argc > 0 {
        set(&mut chunks[current], source, line);
    } else {
        chunks[current].emit_string_const("", line);
        set(&mut chunks[current], source, line);
    }

    get(&mut chunks[current], source, line);
    get(&mut chunks[current], delim, line);
    host::emit(&mut chunks[current], "ecma:string", "split", 2, line);
    let tokens = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], tokens, line);

    chunks[current].emit_struct_new(0, 0, line);
    let tokenizer = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], tokenizer, line);
    prop_set_from_slot(&mut chunks[current], tokenizer, TOKENS_KEY, tokens, line);
    chunks[current].emit_i32_const(0, line);
    let index = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], index, line);
    prop_set_from_slot(&mut chunks[current], tokenizer, INDEX_KEY, index, line);
    get(&mut chunks[current], tokenizer, line);
}

pub fn emit_has_more(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let tokenizer = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], tokenizer, line);
    prop_get(&mut chunks[current], tokenizer, INDEX_KEY, line);
    prop_get(&mut chunks[current], tokenizer, TOKENS_KEY, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
}

pub fn emit_count(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let tokenizer = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], tokenizer, line);
    prop_get(&mut chunks[current], tokenizer, TOKENS_KEY, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    prop_get(&mut chunks[current], tokenizer, INDEX_KEY, line);
    chunks[current].emit_op(Op::I32_SUB, line);
}

pub fn emit_next(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let tokenizer = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], tokenizer, line);
    prop_get(&mut chunks[current], tokenizer, TOKENS_KEY, line);
    prop_get(&mut chunks[current], tokenizer, INDEX_KEY, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);

    prop_get(&mut chunks[current], tokenizer, INDEX_KEY, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    let index = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], index, line);
    prop_set_from_slot(&mut chunks[current], tokenizer, INDEX_KEY, index, line);
    get(&mut chunks[current], out, line);
}
