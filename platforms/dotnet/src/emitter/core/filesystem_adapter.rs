use vybe_emitter::instructions::host;
use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use vybe_emitter::{collections, loops};

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

pub fn emit_file_read_all_lines(chunks: &mut [Chunk], current: usize, line: u32) {
    let read_idx = chunks[0].add_import("node:fs", "readFileSync");
    let chunk = &mut chunks[current];
    let path_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    push_const(chunk, Value::String(Arc::from("utf8")), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, read_idx, line);
    chunk.emit(2, line);
    push_const(chunk, Value::String(Arc::from("\n")), line);
    host::emit(chunk, "ecma:string", "split", 2, line);
}

fn emit_directory_entries(chunks: &mut [Chunk], current: usize, line: u32, want_directories: bool) {
    let list_idx = chunks[0].add_import("wasi:filesystem", "listDir");
    let is_dir_idx = chunks[0].add_import("wasi:filesystem", "isDir");
    let resolve_idx = chunks[0].add_import("node:path", "resolve");

    let chunk = &mut chunks[current];
    let root_slot = reserve_slot(chunk);
    let entries_slot = reserve_slot(chunk);
    let idx_slot = reserve_slot(chunk);
    let entry_slot = reserve_slot(chunk);
    let full_path_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, root_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, root_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, list_idx, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, entries_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = loops::emit_for_in_start(chunks, current, entries_slot, idx_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, entry_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, root_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, entry_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, resolve_idx, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, full_path_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, full_path_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, is_dir_idx, line);
    chunk.emit(1, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    if !want_directories {
        vybe_emitter::ops::emit_dyn_not(chunk, line);
    }

    let skip_push = chunk.emit_block(line);
    vybe_emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, full_path_slot, line);
    collections::emit_push(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
    chunk.patch_block(skip_push);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_directory_get_files(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_directory_entries(chunks, current, line, false);
}

pub fn emit_directory_get_directories(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_directory_entries(chunks, current, line, true);
}
