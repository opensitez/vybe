use std::sync::Arc;
use vybe_compiler::primitives::instructions::{core_wasm, host};

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use vybe_compiler::primitives::collections;

const ENV_OVERRIDES_GLOBAL: &str = "__dotnet_environment_overrides";
const ENV_EXIT_CODE_GLOBAL: &str = "__dotnet_environment_exit_code";

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

fn load_or_create_env_overrides(chunks: &mut [Chunk], current: usize, line: u32) -> u16 {
    let chunk = &mut chunks[current];
    let global = chunk.add_constant(Value::String(Arc::from(ENV_OVERRIDES_GLOBAL)));
    let object_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::GLOBAL_GET, global, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let create_block = chunk.emit_block(line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::DROP, line);

    let object_new = chunks[current].add_import("ecma:object", "new");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, object_new, line);
    chunk.emit(0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::GLOBAL_SET, global, line);
    chunk.emit_end(line);
    chunk.patch_block(create_block);

    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, object_slot, line);
    object_slot
}

pub fn emit_environment_username(chunks: &mut [Chunk], current: usize, line: u32) {
    let user_info = chunks[current].add_import("node:os", "userInfo");
    let chunk = &mut chunks[current];
    let username_key = chunk.add_constant(Value::String(Arc::from("username")));
    chunk.emit_op_u16(Op::CALL_IMPORT, user_info, line);
    chunk.emit(0, line);
    chunk.emit_op_u16(Op::STRUCT_GET, username_key, line);
}

pub fn emit_environment_version(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    for value in [8.0, 0.0, 0.0, 0.0] {
        push_const(chunk, Value::F64(value), line);
    }
    crate::emitter::core::version_adapter::emit_version_new(chunks, current, 4, line);
}

pub fn emit_environment_exit_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let global = chunk.add_constant(Value::String(Arc::from(ENV_EXIT_CODE_GLOBAL)));
    let value_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::GLOBAL_GET, global, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    push_const(chunk, Value::I32(0), line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_end(line);
}

pub fn emit_environment_set_exit_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let global = chunk.add_constant(Value::String(Arc::from(ENV_EXIT_CODE_GLOBAL)));
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::GLOBAL_SET, global, line);
}

pub fn emit_environment_system_directory(chunks: &mut [Chunk], current: usize, line: u32) {
    let cwd = chunks[current].add_import("node:process", "cwd");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, cwd, line);
    chunk.emit(0, line);
}

pub fn emit_environment_get_command_line_args(chunks: &mut [Chunk], current: usize, line: u32) {
    let get_args = chunks[current].add_import("wasi:cli/environment", "get-arguments");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, get_args, line);
    chunk.emit(0, line);
}

pub fn emit_environment_get_folder_path(chunks: &mut [Chunk], current: usize, line: u32) {
    let cwd = chunks[current].add_import("node:process", "cwd");
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, cwd, line);
    chunk.emit(0, line);
}

pub fn emit_environment_special_folder(
    name: &str,
    chunks: &mut [Chunk],
    current: usize,
    line: u32,
) {
    let chunk = &mut chunks[current];
    push_const(chunk, Value::String(Arc::from(name)), line);
}

pub fn emit_environment_processor_count(chunks: &mut [Chunk], current: usize, line: u32) {
    let cpus = chunks[current].add_import("node:os", "cpus");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, cpus, line);
    chunk.emit(0, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
}

pub fn emit_environment_tick_count(chunks: &mut [Chunk], current: usize, line: u32) {
    let uptime = chunks[current].add_import("node:process", "uptime");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, uptime, line);
    chunk.emit(0, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
}

pub fn emit_environment_get(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let name_slot = reserve_slot(chunk);
    let object_slot = reserve_slot(chunk);
    let value_slot = reserve_slot(chunk);

    if argc > 1 {
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);

    let global = chunk.add_constant(Value::String(Arc::from(ENV_OVERRIDES_GLOBAL)));
    chunk.emit_op_u16(Op::GLOBAL_GET, global, line);
    chunk.emit_op_u16(Op::LOCAL_SET, object_slot, line);

    chunk.emit_op(Op::NULL, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let done = chunk.emit_block(line);
    let fallback_block = chunk.emit_block(line);
    chunk.emit_br_if(0, line);

    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
    collections::emit_get(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_br(1, line);

    let get_env = chunks[current].add_import("wasi:cli/environment", "get-environment");
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    chunk.patch_block(fallback_block);
    // `get-environment` takes NO argument and answers the whole
    // `list<tuple<string, string>>` (wasi-cli `wit/environment.wit`). Keying it
    // by name is this adapter's job, not the interface's, so the pairs become a
    // map and the lookup happens here.
    chunk.emit_op_u16(Op::CALL_IMPORT, get_env, line);
    chunk.emit(0, line);
    host::emit(chunk, "ecma:map", "fromEntries", 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
    collections::emit_get(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    chunk.patch_block(done);
}

pub fn emit_environment_set(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = reserve_slot(chunk);
    let name_slot = reserve_slot(chunk);

    if argc > 2 {
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);

    let object_slot = load_or_create_env_overrides(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_set(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::NULL, line);
}

pub fn emit_environment_get_all(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let get_env = chunks[current].add_import("wasi:cli/environment", "get-environment");
    let chunk = &mut chunks[current];
    let map_slot = reserve_slot(chunk);
    let object_slot = reserve_slot(chunk);
    let entries_slot = reserve_slot(chunk);
    let idx_slot = reserve_slot(chunk);
    let pair_slot = reserve_slot(chunk);
    let key_slot = reserve_slot(chunk);
    let value_slot = reserve_slot(chunk);

    if argc > 0 {
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, get_env, line);
    chunk.emit(0, line);
    host::emit(chunk, "ecma:map", "fromEntries", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, map_slot, line);

    let global = chunk.add_constant(Value::String(Arc::from(ENV_OVERRIDES_GLOBAL)));
    chunk.emit_op_u16(Op::GLOBAL_GET, global, line);
    chunk.emit_op_u16(Op::LOCAL_SET, object_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let done = chunk.emit_block(line);
    chunk.emit_br_if(0, line);

    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
    host::emit(chunk, "ecma:object", "entries", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, entries_slot, line);

    let state = vybe_compiler::primitives::loops::emit_for_in_start(
        chunks,
        current,
        entries_slot,
        idx_slot,
        line,
    );
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, pair_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    collections::emit_get(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, key_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    collections::emit_get(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(chunk, "ecma:map", "set", 3, line);
    chunk.emit_op(Op::DROP, line);

    vybe_compiler::primitives::loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    chunk.patch_block(done);
    chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
}

pub fn emit_environment_expand(chunks: &mut [Chunk], current: usize, line: u32) {
    let get_env = chunks[current].add_import("wasi:cli/environment", "get-environment");
    let chunk = &mut chunks[current];
    let input_slot = reserve_slot(chunk);
    let env_slot = reserve_slot(chunk);
    let idx_slot = reserve_slot(chunk);
    let pair_slot = reserve_slot(chunk);
    let key_slot = reserve_slot(chunk);
    let value_slot = reserve_slot(chunk);
    let token_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, get_env, line);
    chunk.emit(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, env_slot, line);

    let state = vybe_compiler::primitives::loops::emit_for_in_start(
        chunks, current, env_slot, idx_slot, line,
    );
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, pair_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    collections::emit_get(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, key_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    collections::emit_get(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);

    push_const(chunk, Value::String(Arc::from("%")), line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
    host::emit(chunk, "ecma:string", "concat", 2, line);
    push_const(chunk, Value::String(Arc::from("%")), line);
    host::emit(chunk, "ecma:string", "concat", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, token_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, token_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(chunk, "ecma:string", "replaceAll", 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);

    vybe_compiler::primitives::loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    let chunk = &mut chunks[current];
    let global = chunk.add_constant(Value::String(Arc::from(ENV_OVERRIDES_GLOBAL)));
    let overrides_slot = reserve_slot(chunk);
    let override_entries_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::GLOBAL_GET, global, line);
    chunk.emit_op_u16(Op::LOCAL_SET, overrides_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, overrides_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let done = chunk.emit_block(line);
    chunk.emit_br_if(0, line);

    chunk.emit_op_u16(Op::LOCAL_GET, overrides_slot, line);
    host::emit(chunk, "ecma:object", "entries", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, override_entries_slot, line);

    let state = vybe_compiler::primitives::loops::emit_for_in_start(
        chunks,
        current,
        override_entries_slot,
        idx_slot,
        line,
    );
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, pair_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    collections::emit_get(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, key_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    collections::emit_get(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);

    push_const(chunk, Value::String(Arc::from("%")), line);
    chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
    host::emit(chunk, "ecma:string", "concat", 2, line);
    push_const(chunk, Value::String(Arc::from("%")), line);
    host::emit(chunk, "ecma:string", "concat", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, token_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, token_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(chunk, "ecma:string", "replaceAll", 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, input_slot, line);

    vybe_compiler::primitives::loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    chunk.patch_block(done);
    chunk.emit_op_u16(Op::LOCAL_GET, input_slot, line);
}

pub fn emit_environment_target(name: &str, chunks: &mut [Chunk], current: usize, line: u32) {
    push_const(&mut chunks[current], Value::String(Arc::from(name)), line);
}
