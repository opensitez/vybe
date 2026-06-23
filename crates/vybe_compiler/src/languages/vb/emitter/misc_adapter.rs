use crate::emitter::instructions::core_wasm;
use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

const VB_FILE_PATH_BY_HANDLE: &str = "__vb_file_path_by_handle";
const VB_FILE_EOF_BY_HANDLE: &str = "__vb_file_eof_by_handle";
const VB_RECORD_ROWS_BY_HANDLE: &str = "__vb_record_rows_by_handle";
const VB_RECORD_NEXT_INDEX_BY_HANDLE: &str = "__vb_record_next_index_by_handle";

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let slot = chunk.local_count;
    chunk.local_count = slot + 1;
    slot
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op(Op::DROP, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn gget(chunk: &mut Chunk, name: &str, line: u32) {
    let key = chunk.add_constant(Value::String(Arc::from(name)));
    chunk.emit_op_u16(Op::GLOBAL_GET, key, line);
}

fn gset(chunk: &mut Chunk, name: &str, line: u32) {
    let key = chunk.add_constant(Value::String(Arc::from(name)));
    chunk.emit_op_u16(Op::GLOBAL_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

fn ensure_global_map(chunks: &mut [Chunk], current: usize, name: &str, line: u32) -> u16 {
    let slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };

    {
        let chunk = &mut chunks[current];
        gget(chunk, name, line);
        core_wasm::dup(chunk, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }

    chunks[current].emit_if(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
    }
    crate::emitter::collections::emit_map_new(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        core_wasm::dup(chunk, line);
        gset(chunk, name, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);

    {
        let chunk = &mut chunks[current];
        lset(chunk, slot, line);
    }
    slot
}

pub fn emit_vb_dir(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        push_const(&mut chunks[current], Value::String(Arc::from("")), line);
        return;
    }

    let exists = chunks[0].add_import("wasi:filesystem", "exists");
    let file_name = chunks[0].add_import("wasi:filesystem", "pathGetFileName");
    let path_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };

    {
        let chunk = &mut chunks[current];
        lset(chunk, path_slot, line);
        lget(chunk, path_slot, line);
        chunk.emit_op_u16(Op::CALL_IMPORT, exists, line);
        chunk.emit(1, line);
        crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    }

    chunks[current].emit_if(line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, path_slot, line);
        chunk.emit_op_u16(Op::CALL_IMPORT, file_name, line);
        chunk.emit(1, line);
    }
    chunks[current].emit_else(line);
    push_const(&mut chunks[current], Value::String(Arc::from("")), line);
    chunks[current].emit_end(line);
}

pub fn emit_vb_filedatetime(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let stat = chunks[0].add_import("wasi:filesystem", "stat");
    let to_iso = chunks[0].add_import("ecma:date", "toISOString");
    let modified_key = chunks[current].add_constant(Value::String(Arc::from("modified")));
    let path_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };

    {
        let chunk = &mut chunks[current];
        lset(chunk, path_slot, line);
        lget(chunk, path_slot, line);
        chunk.emit_op_u16(Op::CALL_IMPORT, stat, line);
        chunk.emit(1, line);
        core_wasm::dup(chunk, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }

    chunks[current].emit_if(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
    }
    push_const(&mut chunks[current], Value::String(Arc::from("")), line);

    chunks[current].emit_else(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::STRUCT_GET, modified_key, line);
        chunk.emit_op_u16(Op::CALL_IMPORT, to_iso, line);
        chunk.emit(1, line);
    }
    chunks[current].emit_end(line);
}

pub fn emit_vb_lof(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let stat = chunks[0].add_import("wasi:filesystem", "stat");
    let size_key = chunks[current].add_constant(Value::String(Arc::from("size")));
    let handle_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let path_map_slot = ensure_global_map(chunks, current, VB_FILE_PATH_BY_HANDLE, line);

    {
        let chunk = &mut chunks[current];
        lset(chunk, handle_slot, line);
        lget(chunk, path_map_slot, line);
        lget(chunk, handle_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        core_wasm::dup(chunk, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }

    chunks[current].emit_if(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
    }
    push_const(&mut chunks[current], Value::I32(0), line);

    chunks[current].emit_else(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::CALL_IMPORT, stat, line);
        chunk.emit(1, line);
        core_wasm::dup(chunk, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }

    chunks[current].emit_if(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
    }
    push_const(&mut chunks[current], Value::I32(0), line);

    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, size_key, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_vb_eof(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let handle_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let next_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let rows_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let len_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let eof_map_slot = ensure_global_map(chunks, current, VB_FILE_EOF_BY_HANDLE, line);
    let next_map_slot = ensure_global_map(chunks, current, VB_RECORD_NEXT_INDEX_BY_HANDLE, line);
    let rows_map_slot = ensure_global_map(chunks, current, VB_RECORD_ROWS_BY_HANDLE, line);

    {
        let chunk = &mut chunks[current];
        lset(chunk, handle_slot, line);
        lget(chunk, next_map_slot, line);
        lget(chunk, handle_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, next_slot, line);

        lget(chunk, rows_map_slot, line);
        lget(chunk, handle_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, rows_slot, line);

        lget(chunk, next_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }

    chunks[current].emit_if(line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, eof_map_slot, line);
        lget(chunk, handle_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        core_wasm::dup(chunk, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }
    chunks[current].emit_if(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
    }
    push_const(&mut chunks[current], Value::Bool(false), line);
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, rows_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }
    chunks[current].emit_if(line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, eof_map_slot, line);
        lget(chunk, handle_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        core_wasm::dup(chunk, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }
    chunks[current].emit_if(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
    }
    push_const(&mut chunks[current], Value::Bool(false), line);
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, rows_slot, line);
    }
    crate::emitter::collections::emit_len(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, len_slot, line);
        lget(chunk, next_slot, line);
        lget(chunk, len_slot, line);
        crate::emitter::ops::emit_dyn_ge(chunk, line);
    }
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_vb_shell_pid(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let spawn_sync = chunks[0].add_import("node:child_process", "spawnSync");
    let pid_key = chunks[current].add_constant(Value::String(Arc::from("pid")));
    let command_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };

    {
        let chunk = &mut chunks[current];
        if argc >= 2 {
            chunk.emit_op(Op::DROP, line);
        }
        lset(chunk, command_slot, line);

        push_const(chunk, Value::String(Arc::from("/bin/sh")), line);
        push_const(chunk, Value::String(Arc::from("-c")), line);
        lget(chunk, command_slot, line);
        chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 2, line);
        chunk.emit_op_u16(Op::CALL_IMPORT, spawn_sync, line);
        chunk.emit(2, line);
        chunk.emit_op_u16(Op::STRUCT_GET, pid_key, line);
        core_wasm::dup(chunk, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }

    chunks[current].emit_if(line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
        push_const(chunk, Value::I32(0), line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);
}
