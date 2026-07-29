//! Shared .NET JSON adapters.

use vybe_bytecode::{opcode::Op, Chunk};

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(argc, line);
}

pub fn emit_json_serialize(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    call_import(chunks, current, "ecma:json", "stringify", argc, line);
}

pub fn emit_json_deserialize(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    call_import(chunks, current, "ecma:json", "parse", argc, line);
}
