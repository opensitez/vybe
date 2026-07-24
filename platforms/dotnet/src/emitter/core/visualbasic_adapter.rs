//! Shared Microsoft.VisualBasic runtime helpers for .NET languages.

use std::sync::Arc;
use vybe_emitter::instructions::core_wasm;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

const VB_FILE_PATH_BY_HANDLE: &str = "__vb_file_path_by_handle";
const VB_FILE_EOF_BY_HANDLE: &str = "__vb_file_eof_by_handle";
const VB_RECORD_ROWS_BY_HANDLE: &str = "__vb_record_rows_by_handle";
const VB_RECORD_NEXT_INDEX_BY_HANDLE: &str = "__vb_record_next_index_by_handle";

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
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
}

fn emit_host_call(
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
    vybe_emitter::collections::emit_map_new(chunks, current, line);
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

pub fn emit_vb_filecopy(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "wasi:filesystem", "copy", argc, line);
}

pub fn emit_vb_kill(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "wasi:filesystem", "remove", argc, line);
}

pub fn emit_vb_fileexists(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "wasi:filesystem", "exists", argc, line);
}

pub fn emit_vb_filelen(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "wasi:filesystem", "fileSize", argc, line);
}

pub fn emit_vb_curdir(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "node:process", "cwd", argc, line);
}

pub fn emit_vb_chdir(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "node:process", "chdir", argc, line);
}

pub fn emit_vb_mkdir(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "wasi:filesystem", "mkdir", argc, line);
}

pub fn emit_vb_rmdir(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "wasi:filesystem", "remove", argc, line);
}

pub fn emit_vb_name(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "wasi:filesystem", "rename", argc, line);
}

pub fn emit_vb_get(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "wasi:filesystem", "readFile", argc, line);
}

pub fn emit_vb_put(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "wasi:filesystem", "writeFile", argc, line);
}

pub fn emit_vb_app_path(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "node:process", "cwd", argc, line);
}

pub fn emit_vb_app_title(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "node:process", "argv0", argc, line);
}

pub fn emit_vb_to_number(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "ecma:number", "Number", argc, line);
}

pub fn emit_vb_to_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 1 {
        emit_host_call(chunks, current, "ecma:string", "String", argc, line);
        return;
    }

    let value_slot = alloc_local(&mut chunks[current]);
    let result_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, value_slot, line);
    super::console_adapter::emit_dotnet_stringify(chunk, value_slot, result_slot, line);
    lget(chunk, result_slot, line);
}

pub fn emit_vb_random(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "ecma:math", "random", argc, line);
}

pub fn emit_vb_lset(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "ecma:string", "padEnd", argc, line);
}

pub fn emit_vb_rset(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "ecma:string", "padStart", argc, line);
}

pub fn emit_vb_array(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "ecma:array", "from", argc, line);
}

pub fn emit_vb_debug_print(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "wasi:cli", "log", argc, line);
}

pub fn emit_vb_print(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "wasi:cli", "log", argc, line);
}

pub fn emit_vb_input(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "wasi:cli", "readLine", argc, line);
}

pub fn emit_vb_app(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "wasi:cli", "args", argc, line);
}

pub fn emit_vb_open(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_host_call(chunks, current, "wasi:filesystem", "readFile", argc, line);
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
        vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
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
    vybe_emitter::collections::emit_len(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, len_slot, line);
        lget(chunk, next_slot, line);
        lget(chunk, len_slot, line);
        vybe_emitter::ops::emit_dyn_ge(chunk, line);
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
