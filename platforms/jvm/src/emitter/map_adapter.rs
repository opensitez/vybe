//! JVM `java.util.Map` constructor adapters.

use vybe_compiler::primitives::{collections, instructions::host};
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

const CONCURRENT_MAP_KEY: &str = "__java_concurrent_map";
const IDENTITY_MAP_KEY: &str = "__java_identity_map";
const LINKED_MAP_KEY: &str = "__java_linked_map";

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn drop_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
}

fn mark_bool(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_bool_const(true, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], map, line);
}

pub fn emit_hash_map_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    drop_args(chunks, current, argc, line);
    collections::emit_map_new(chunks, current, line);
}

pub fn emit_concurrent_hash_map_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_hash_map_new(chunks, current, argc, line);
    mark_bool(chunks, current, CONCURRENT_MAP_KEY, line);
}

pub fn emit_identity_hash_map_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_hash_map_new(chunks, current, argc, line);
    mark_bool(chunks, current, IDENTITY_MAP_KEY, line);
}

pub fn emit_linked_hash_map_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_hash_map_new(chunks, current, argc, line);
    mark_bool(chunks, current, LINKED_MAP_KEY, line);
}
