//! .NET `System.Guid` adapter — bytecode-only.
//!
//! `Guid` values are represented as plain Objects carrying the normalized
//! lowercase text form under `__value` plus a `__type="Guid"` tag so the
//! shared .NET dispatch layer can preserve value-type semantics without host
//! changes.

use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};
use vybe_compiler::primitives::functions::create_function_chunk;
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::object::emit_bind_method_with_slot;

const TYPE_KEY: &str = "__type";
const VALUE_KEY: &str = "__value";
const BYTES_KEY: &str = "__bytes";
const EMPTY_GUID: &str = "00000000-0000-0000-0000-000000000000";
const GUID_PATTERN: &str =
    "^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$";
const GUID_N_PATTERN: &str = "^[0-9a-fA-F]{32}$";
const FORMAT_EXCEPTION_MSG: &str =
    "Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).";

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

fn emit_throw_guid_format_exception(chunk: &mut Chunk, line: u32) {
    chunk.emit_string_const(FORMAT_EXCEPTION_MSG, line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(chunk, "FormatException", line);
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
}

fn bind_guid_to_string(chunks: &mut Vec<Chunk>, current: usize, this_slot: u16, line: u32) {
    let mut method = create_function_chunk("__guid_tostring", 1);
    let value_key = method.add_constant(Value::String(Arc::from(VALUE_KEY)));
    method.emit_op_u16(Op::LOCAL_GET, 0, line);
    method.emit_op_u16(Op::STRUCT_GET, value_key, line);
    method.emit_op(Op::RETURN, line);
    method.local_count = 1;
    chunks.push(method);
    let method_idx = chunks.len() - 1;
    // `.NET`'s own spelling of the member, plus the lowercased vtable key a
    // case-insensitive caller (VB) lands on, plus the ToString ROLE for a
    // caller in any other language. The first two are this type's real API
    // surface — not a guess at what some other language might call it, which
    // is what the deleted synonym table was doing.
    for name in ["tostring", "ToString"] {
        emit_bind_method_with_slot(
            &mut chunks[current],
            this_slot,
            name,
            Some(vybe_ast::ProtocolSlot::ToString),
            method_idx,
            None,
            line,
        );
    }
}

fn emit_wrap_guid_from_slot(chunks: &mut Vec<Chunk>, current: usize, text_slot: u16, line: u32) {
    let chunk = &mut chunks[current];
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let value_key = chunk.add_constant(Value::String(Arc::from(VALUE_KEY)));
    let obj_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Guid")), line);
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, value_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::DROP, line);

    bind_guid_to_string(chunks, current, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

fn emit_wrap_guid_with_bytes_from_slots(
    chunks: &mut Vec<Chunk>,
    current: usize,
    text_slot: u16,
    bytes_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let value_key = chunk.add_constant(Value::String(Arc::from(VALUE_KEY)));
    let bytes_key = chunk.add_constant(Value::String(Arc::from(BYTES_KEY)));
    let obj_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Guid")), line);
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, value_key, line);
    chunk.emit_op(Op::DROP, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, bytes_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::DROP, line);

    bind_guid_to_string(chunks, current, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

fn emit_validate_guid_text(chunk: &mut Chunk, test_idx: u16, text_slot: u16, line: u32) {
    let ok_block = chunk.emit_block(line);
    push_const(chunk, Value::String(Arc::from(GUID_PATTERN)), line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, test_idx, line);
    chunk.emit(2, line);
    chunk.emit_br_if(0, line);
    emit_throw_guid_format_exception(chunk, line);
    chunk.emit_end(line);
    chunk.patch_block(ok_block);
}

fn emit_build_guid_from_stack(
    chunks: &mut Vec<Chunk>,
    current: usize,
    normalize: bool,
    validate: bool,
    line: u32,
) {
    let to_str_idx = chunks[current].add_import("ecma:string", "String");
    let lower_idx = chunks[current].add_import("ecma:string", "toLowerCase");
    let test_idx = chunks[current].add_import("ecma:regexp", "test");
    let replace_all_idx = chunks[current].add_import("ecma:string", "replaceAll");
    let substr_idx = chunks[current].add_import("ecma:string", "substr");
    let concat_idx = chunks[current].add_import("ecma:string", "concat");

    let chunk = &mut chunks[current];
    let text_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::CALL_IMPORT, to_str_idx, line);
    chunk.emit(1, line);
    push_const(chunk, Value::String(Arc::from("{")), line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, replace_all_idx, line);
    chunk.emit(3, line);
    push_const(chunk, Value::String(Arc::from("}")), line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, replace_all_idx, line);
    chunk.emit(3, line);
    push_const(chunk, Value::String(Arc::from("(")), line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, replace_all_idx, line);
    chunk.emit(3, line);
    push_const(chunk, Value::String(Arc::from(")")), line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, replace_all_idx, line);
    chunk.emit(3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);

    push_const(chunk, Value::String(Arc::from(GUID_N_PATTERN)), line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, test_idx, line);
    chunk.emit(2, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    push_const(chunk, Value::F64(8.0), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, substr_idx, line);
    chunk.emit(3, line);
    push_const(chunk, Value::String(Arc::from("-")), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, concat_idx, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    push_const(chunk, Value::F64(8.0), line);
    push_const(chunk, Value::F64(4.0), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, substr_idx, line);
    chunk.emit(3, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, concat_idx, line);
    chunk.emit(2, line);
    push_const(chunk, Value::String(Arc::from("-")), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, concat_idx, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    push_const(chunk, Value::F64(12.0), line);
    push_const(chunk, Value::F64(4.0), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, substr_idx, line);
    chunk.emit(3, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, concat_idx, line);
    chunk.emit(2, line);
    push_const(chunk, Value::String(Arc::from("-")), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, concat_idx, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    push_const(chunk, Value::F64(16.0), line);
    push_const(chunk, Value::F64(4.0), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, substr_idx, line);
    chunk.emit(3, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, concat_idx, line);
    chunk.emit(2, line);
    push_const(chunk, Value::String(Arc::from("-")), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, concat_idx, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    push_const(chunk, Value::F64(20.0), line);
    push_const(chunk, Value::F64(12.0), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, substr_idx, line);
    chunk.emit(3, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, concat_idx, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    chunk.emit_end(line);

    if validate {
        emit_validate_guid_text(chunk, test_idx, text_slot, line);
    }

    if normalize {
        chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
        chunk.emit_op_u16(Op::CALL_IMPORT, lower_idx, line);
        chunk.emit(1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    }

    emit_wrap_guid_from_slot(chunks, current, text_slot, line);
}

pub fn emit_guid_empty(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let text_slot = reserve_slot(chunk);
    push_const(chunk, Value::String(Arc::from(EMPTY_GUID)), line);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    emit_wrap_guid_from_slot(chunks, current, text_slot, line);
}

pub fn emit_guid_new_guid(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let random_uuid_idx = chunks[current].add_import("web:crypto", "randomUUID");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, random_uuid_idx, line);
    chunk.emit(0, line);
    emit_build_guid_from_stack(chunks, current, true, true, line);
}

pub fn emit_guid_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    emit_build_guid_from_stack(chunks, current, true, true, line);
}

pub fn emit_guid_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    match argc {
        0 => emit_guid_empty(chunks, current, line),
        1 => {
            let is_array_idx = chunks[current].add_import("ecma:array", "isArray");
            let join_idx = chunks[current].add_import("ecma:array", "join");
            let chunk = &mut chunks[current];
            let value_slot = reserve_slot(chunk);
            let text_slot = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
            chunk.emit_op_u16(Op::CALL_IMPORT, is_array_idx, line);
            chunk.emit(1, line);
            chunk.emit_if_value(line);
            chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
            push_const(chunk, Value::String(Arc::from(",")), line);
            chunk.emit_op_u16(Op::CALL_IMPORT, join_idx, line);
            chunk.emit(2, line);
            chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
            emit_wrap_guid_with_bytes_from_slots(chunks, current, text_slot, value_slot, line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
            emit_build_guid_from_stack(chunks, current, true, true, line);
            chunks[current].emit_end(line);
        }
        _ => {
            let chunk = &mut chunks[current];
            for _ in 1..argc {
                chunk.emit_op(Op::DROP, line);
            }
            emit_guid_parse(chunks, current, line);
        }
    }
}

pub fn emit_guid_to_byte_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let bytes_key = chunk.add_constant(Value::String(Arc::from(BYTES_KEY)));
    chunk.emit_op_u16(Op::STRUCT_GET, bytes_key, line);
}

pub fn emit_guid_get_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value_key = chunk.add_constant(Value::String(Arc::from(VALUE_KEY)));
    chunk.emit_op_u16(Op::STRUCT_GET, value_key, line);
    push_const(chunk, Value::F64(0.0), line);
    let char_code_idx = chunk.add_import("ecma:string", "charCodeAt");
    chunk.emit_op_u16(Op::CALL_IMPORT, char_code_idx, line);
    chunk.emit(2, line);
}

pub fn emit_guid_to_string(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let replace_all_idx = chunks[current].add_import("ecma:string", "replaceAll");
    let chunk = &mut chunks[current];
    let obj_slot = reserve_slot(chunk);
    let fmt_slot = reserve_slot(chunk);
    let value_key = chunk.add_constant(Value::String(Arc::from(VALUE_KEY)));
    if argc > 0 {
        chunk.emit_op_u16(Op::LOCAL_SET, fmt_slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, value_key, line);
    if argc > 0 {
        let value_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
        push_const(chunk, Value::String(Arc::from("N")), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if_value(line);
        chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
        push_const(chunk, Value::String(Arc::from("-")), line);
        push_const(chunk, Value::String(Arc::from("")), line);
        chunk.emit_op_u16(Op::CALL_IMPORT, replace_all_idx, line);
        chunk.emit(3, line);
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunk.emit_end(line);
    }
}

pub fn emit_guid_try_parse(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let to_str_idx = chunks[current].add_import("ecma:string", "String");
    let lower_idx = chunks[current].add_import("ecma:string", "toLowerCase");
    let test_idx = chunks[current].add_import("ecma:regexp", "test");

    let text_slot;
    {
        let chunk = &mut chunks[current];
        for _ in 1..argc {
            chunk.emit_op(Op::DROP, line);
        }

        text_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::CALL_IMPORT, to_str_idx, line);
        chunk.emit(1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);

        push_const(chunk, Value::String(Arc::from(GUID_PATTERN)), line);
        chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
        chunk.emit_op_u16(Op::CALL_IMPORT, test_idx, line);
        chunk.emit(2, line);
        chunk.emit_if(line);

        chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
        chunk.emit_op_u16(Op::CALL_IMPORT, lower_idx, line);
        chunk.emit(1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    }
    emit_wrap_guid_from_slot(chunks, current, text_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
}
