use crate::emitter::instructions::core_wasm;
use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use crate::emitter::collections;

const ENV_OVERRIDES_GLOBAL: &str = "__dotnet_environment_overrides";

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

    let object_new = chunks[0].add_import("ecma:object", "new");
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
    let user_info = chunks[0].add_import("node:os", "userInfo");
    let chunk = &mut chunks[current];
    let username_key = chunk.add_constant(Value::String(Arc::from("username")));
    chunk.emit_op_u16(Op::CALL_IMPORT, user_info, line);
    chunk.emit(0, line);
    chunk.emit_op_u16(Op::STRUCT_GET, username_key, line);
}

pub fn emit_environment_processor_count(chunks: &mut [Chunk], current: usize, line: u32) {
    let cpus = chunks[0].add_import("node:os", "cpus");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, cpus, line);
    chunk.emit(0, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
}

pub fn emit_environment_tick_count(chunks: &mut [Chunk], current: usize, line: u32) {
    let uptime = chunks[0].add_import("node:process", "uptime");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, uptime, line);
    chunk.emit(0, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
}

pub fn emit_environment_get(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let name_slot = reserve_slot(chunk);
    let object_slot = reserve_slot(chunk);
    let value_slot = reserve_slot(chunk);

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

    let get_env = chunks[0].add_import("wasi:cli/environment", "get-environment");
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    chunk.patch_block(fallback_block);
    chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, get_env, line);
    chunk.emit(1, line);
    chunk.emit_end(line);
    chunk.patch_block(done);
}

pub fn emit_environment_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = reserve_slot(chunk);
    let name_slot = reserve_slot(chunk);

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
