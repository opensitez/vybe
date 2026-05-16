use std::sync::Arc;

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

const VB_FILE_PATH_BY_HANDLE: &str = "__vb_file_path_by_handle";
const VB_FILE_EOF_BY_HANDLE: &str = "__vb_file_eof_by_handle";

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let slot = chunk.local_count;
    chunk.local_count = slot + 1;
    slot
}

fn push_const(chunk: &mut Chunk, value: Value, line: u32) {
    let idx = chunk.add_constant(value);
    chunk.emit_op_u16(Op::CONST, idx, line);
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
        chunk.emit_op(Op::DUP, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }

    let has_map = chunks[current].emit_jump(Op::BR_IF_FALSE, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
    }
    crate::emitter::collections::emit_map_new(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DUP, line);
        gset(chunk, name, line);
    }
    chunks[current].patch_jump(has_map);

    {
        let chunk = &mut chunks[current];
        lset(chunk, slot, line);
    }
    slot
}

fn emit_array_get_const_index(chunk: &mut Chunk, array_slot: u16, index: f64, line: u32) {
    lget(chunk, array_slot, line);
    push_const(chunk, Value::F64(index), line);
    chunk.emit_op(Op::ARRAY_GET, line);
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
    }

    let missing = chunks[current].emit_jump(Op::BR_IF_FALSE, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, path_slot, line);
        chunk.emit_op_u16(Op::CALL_IMPORT, file_name, line);
        chunk.emit(1, line);
    }
    let done = chunks[current].emit_jump(Op::BR, line);

    chunks[current].patch_jump(missing);
    push_const(&mut chunks[current], Value::String(Arc::from("")), line);
    chunks[current].patch_jump(done);
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
        chunk.emit_op(Op::DUP, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }

    let has_stat = chunks[current].emit_jump(Op::BR_IF_FALSE, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
    }
    push_const(&mut chunks[current], Value::String(Arc::from("")), line);
    let done = chunks[current].emit_jump(Op::BR, line);

    chunks[current].patch_jump(has_stat);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::STRUCT_GET, modified_key, line);
        chunk.emit_op_u16(Op::CALL_IMPORT, to_iso, line);
        chunk.emit(1, line);
    }
    chunks[current].patch_jump(done);
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
        chunk.emit_op(Op::DUP, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }

    let has_path = chunks[current].emit_jump(Op::BR_IF_FALSE, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
    }
    push_const(&mut chunks[current], Value::I32(0), line);
    let done = chunks[current].emit_jump(Op::BR, line);

    chunks[current].patch_jump(has_path);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::CALL_IMPORT, stat, line);
        chunk.emit(1, line);
        chunk.emit_op(Op::DUP, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }

    let has_stat = chunks[current].emit_jump(Op::BR_IF_FALSE, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
    }
    push_const(&mut chunks[current], Value::I32(0), line);
    let stat_done = chunks[current].emit_jump(Op::BR, line);

    chunks[current].patch_jump(has_stat);
    chunks[current].emit_op_u16(Op::STRUCT_GET, size_key, line);
    chunks[current].patch_jump(stat_done);
    chunks[current].patch_jump(done);
}

pub fn emit_vb_eof(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let handle_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let eof_map_slot = ensure_global_map(chunks, current, VB_FILE_EOF_BY_HANDLE, line);

    {
        let chunk = &mut chunks[current];
        lset(chunk, handle_slot, line);
        emit_array_get_const_index(chunk, eof_map_slot, 0.0, line);
        chunk.emit_op(Op::DROP, line);
        lget(chunk, eof_map_slot, line);
        lget(chunk, handle_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op(Op::DUP, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }

    let has_value = chunks[current].emit_jump(Op::BR_IF_FALSE, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
    }
    push_const(&mut chunks[current], Value::Bool(false), line);
    let done = chunks[current].emit_jump(Op::BR, line);

    chunks[current].patch_jump(has_value);
    chunks[current].patch_jump(done);
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
        chunk.emit_op(Op::DUP, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
    }

    let has_pid = chunks[current].emit_jump(Op::BR_IF_FALSE, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
        push_const(chunk, Value::I32(0), line);
    }
    chunks[current].patch_jump(has_pid);
}