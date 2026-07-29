//! .NET `System.Version` adapter — bytecode-only.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};
use vybe_compiler::compiler::functions::create_function_chunk;
use vybe_compiler::compiler::instructions::{core_wasm, host};
use vybe_compiler::compiler::object::emit_bind_method_with_slot;

const TYPE_KEY: &str = "__type";
const TYPES_KEY: &str = "__types";
const MAJOR_KEY: &str = "Major";
const MINOR_KEY: &str = "Minor";
const BUILD_KEY: &str = "Build";
const REVISION_KEY: &str = "Revision";
const MAJOR_REVISION_KEY: &str = "MajorRevision";
const MINOR_REVISION_KEY: &str = "MinorRevision";

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

fn emit_throw_dotnet_exception(chunk: &mut Chunk, exception_name: &str, message: &str, line: u32) {
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_string_const(message, line);
    vybe_compiler::compiler::errors::emit_exception_new_finalize(chunk, exception_name, line);
    vybe_compiler::compiler::errors::emit_throw(chunk, line);
}

fn emit_array_get_const_index(chunk: &mut Chunk, array_slot: u16, index: f64, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
    push_const(chunk, Value::F64(index), line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

fn emit_parse_number_from_slot(chunks: &mut [Chunk], current: usize, text_slot: u16, line: u32) {
    let parse_int_idx = chunks[current].add_import("ecma:number", "parseInt");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, parse_int_idx, line);
    chunk.emit(1, line);
    chunk.emit_op(Op::F64_FLOOR, line);
}

fn emit_store_optional_array_part_as_number(
    chunks: &mut [Chunk],
    current: usize,
    array_slot: u16,
    index: f64,
    out_slot: u16,
    default_value: f64,
    line: u32,
) {
    let chunk = &mut chunks[current];
    emit_array_get_const_index(chunk, array_slot, index, line);
    let text_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    push_const(chunk, Value::F64(default_value), line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    host::emit(chunk, "wasm:js-undefined", "test", 1, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::F64(default_value), line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunk.emit_else(line);
    let _ = chunk;
    emit_parse_number_from_slot(chunks, current, text_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_validate_number_slot(
    chunk: &mut Chunk,
    slot: u16,
    allow_negative: bool,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    emit_throw_dotnet_exception(
        chunk,
        "ArgumentException",
        "Version string portion was not valid.",
        line,
    );
    chunk.emit_end(line);

    if !allow_negative {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        push_const(chunk, Value::F64(0.0), line);
        chunk.emit_op(Op::F64_LT, line);
        chunk.emit_if(line);
        emit_throw_dotnet_exception(
            chunk,
            "ArgumentOutOfRangeException",
            "Version component must be non-negative.",
            line,
        );
        chunk.emit_end(line);
    }
}

fn emit_validate_optional_number_slot(
    chunk: &mut Chunk,
    len_slot: u16,
    min_len: i32,
    slot: u16,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_i32_const(min_len, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_if(line);
    emit_validate_number_slot(chunk, slot, false, line);
    chunk.emit_end(line);
}

fn emit_validate_version_parts(
    chunk: &mut Chunk,
    len_slot: u16,
    major_slot: u16,
    minor_slot: u16,
    build_slot: u16,
    revision_slot: u16,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_i32_const(2, line);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_if(line);
    emit_throw_dotnet_exception(
        chunk,
        "ArgumentException",
        "Version string must contain between two and four components.",
        line,
    );
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_i32_const(4, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);
    emit_throw_dotnet_exception(
        chunk,
        "ArgumentException",
        "Version string must contain between two and four components.",
        line,
    );
    chunk.emit_end(line);

    emit_validate_number_slot(chunk, major_slot, false, line);
    emit_validate_number_slot(chunk, minor_slot, false, line);
    emit_validate_optional_number_slot(chunk, len_slot, 3, build_slot, line);
    emit_validate_optional_number_slot(chunk, len_slot, 4, revision_slot, line);
}

fn emit_version_parts_from_string(
    chunks: &mut [Chunk],
    current: usize,
    text_slot: u16,
    parts_slot: u16,
    len_slot: u16,
    major_slot: u16,
    minor_slot: u16,
    build_slot: u16,
    revision_slot: u16,
    validate: bool,
    line: u32,
) {
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    push_const(chunk, Value::String(Arc::from(".")), line);
    host::emit(chunk, "ecma:string", "split", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, parts_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, parts_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);

    emit_store_optional_array_part_as_number(
        chunks, current, parts_slot, 0.0, major_slot, 0.0, line,
    );
    emit_store_optional_array_part_as_number(
        chunks, current, parts_slot, 1.0, minor_slot, 0.0, line,
    );
    emit_store_optional_array_part_as_number(
        chunks, current, parts_slot, 2.0, build_slot, -1.0, line,
    );
    emit_store_optional_array_part_as_number(
        chunks,
        current,
        parts_slot,
        3.0,
        revision_slot,
        -1.0,
        line,
    );

    if validate {
        let chunk = &mut chunks[current];
        emit_validate_version_parts(
            chunk,
            len_slot,
            major_slot,
            minor_slot,
            build_slot,
            revision_slot,
            line,
        );
    }
}

fn emit_build_version_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    tostring_method_idx: usize,
    major_slot: u16,
    minor_slot: u16,
    build_slot: u16,
    revision_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let types_key = chunk.add_constant(Value::String(Arc::from(TYPES_KEY)));
    let major_key = chunk.add_constant(Value::String(Arc::from(MAJOR_KEY)));
    let minor_key = chunk.add_constant(Value::String(Arc::from(MINOR_KEY)));
    let build_key = chunk.add_constant(Value::String(Arc::from(BUILD_KEY)));
    let revision_key = chunk.add_constant(Value::String(Arc::from(REVISION_KEY)));
    let major_revision_key = chunk.add_constant(Value::String(Arc::from(MAJOR_REVISION_KEY)));
    let minor_revision_key = chunk.add_constant(Value::String(Arc::from(MINOR_REVISION_KEY)));

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Version")), line);
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);

    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Version")), line);
    push_const(chunk, Value::String(Arc::from("Object")), line);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 2, line);
    chunk.emit_op_u16(Op::STRUCT_SET, types_key, line);
    chunk.emit_op(Op::DROP, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, major_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, major_key, line);
    chunk.emit_op(Op::DROP, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, minor_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, minor_key, line);
    chunk.emit_op(Op::DROP, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, build_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, build_key, line);
    chunk.emit_op(Op::DROP, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, revision_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, revision_key, line);
    chunk.emit_op(Op::DROP, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, revision_slot, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
    chunk.emit_i32_const(16, line);
    chunk.emit_op(Op::I32_SHR_U, line);
    chunk.emit_op(Op::F64_CONVERT_I32_U, line);
    chunk.emit_op_u16(Op::STRUCT_SET, major_revision_key, line);
    chunk.emit_op(Op::DROP, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, revision_slot, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
    chunk.emit_i32_const(0xffff, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::F64_CONVERT_I32_U, line);
    chunk.emit_op_u16(Op::STRUCT_SET, minor_revision_key, line);
    chunk.emit_op(Op::DROP, line);

    let obj_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    bind_version_to_string(chunk, obj_slot, tostring_method_idx, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

fn emit_version_part(chunk: &mut Chunk, obj_slot: u16, key: &str, line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, key_idx, line);
}

fn emit_append_version_part(
    chunk: &mut Chunk,
    obj_slot: u16,
    out_slot: u16,
    key: &str,
    line: u32,
) {
    let to_str_idx = chunk.add_import("ecma:string", "String");
    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
    push_const(chunk, Value::String(Arc::from(".")), line);
    vybe_compiler::compiler::ops::emit_dyn_add(chunk, line);
    emit_version_part(chunk, obj_slot, key, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, to_str_idx, line);
    chunk.emit(1, line);
    vybe_compiler::compiler::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
}

fn push_version_tostring_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut method = create_function_chunk("__dotnet_version_tostring", 1);
    method.local_count = 2;
    let to_str_idx = method.add_import("ecma:string", "String");
    let obj_slot = 0;
    let out_slot = 1;

    emit_version_part(&mut method, obj_slot, MAJOR_KEY, line);
    method.emit_op_u16(Op::CALL_IMPORT, to_str_idx, line);
    method.emit(1, line);
    method.emit_op_u16(Op::LOCAL_SET, out_slot, line);

    emit_append_version_part(&mut method, obj_slot, out_slot, MINOR_KEY, line);
    for key in [BUILD_KEY, REVISION_KEY] {
        emit_version_part(&mut method, obj_slot, key, line);
        push_const(&mut method, Value::F64(-1.0), line);
        vybe_compiler::compiler::ops::emit_dyn_gt(&mut method, line);
        vybe_compiler::compiler::ops::emit_dyn_to_bool(&mut method, line);
        method.emit_if(line);
        emit_append_version_part(&mut method, obj_slot, out_slot, key, line);
        method.emit_end(line);
    }

    method.emit_op_u16(Op::LOCAL_GET, out_slot, line);
    method.emit_op(Op::RETURN, line);
    chunks.push(method);
    chunks.len() - 1
}

fn bind_version_to_string(chunk: &mut Chunk, obj_slot: u16, method_idx: usize, line: u32) {
    for name in ["tostring", "ToString"] {
        emit_bind_method_with_slot(
            chunk,
            obj_slot,
            name,
            Some(vybe_ast::ProtocolSlot::ToString),
            method_idx,
            None,
            line,
        );
    }
}

fn emit_version_compare_internal(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let right_slot = reserve_slot(chunk);
    let left_slot = reserve_slot(chunk);
    let left_part_slot = reserve_slot(chunk);
    let right_part_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);

    let done = chunk.emit_block(line);
    for key in [MAJOR_KEY, MINOR_KEY, BUILD_KEY, REVISION_KEY] {
        emit_version_part(chunk, left_slot, key, line);
        chunk.emit_op_u16(Op::LOCAL_SET, left_part_slot, line);
        emit_version_part(chunk, right_slot, key, line);
        chunk.emit_op_u16(Op::LOCAL_SET, right_part_slot, line);

        chunk.emit_op_u16(Op::LOCAL_GET, left_part_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, right_part_slot, line);
        vybe_compiler::compiler::ops::emit_dyn_lt(chunk, line);
        vybe_compiler::compiler::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        push_const(chunk, Value::F64(-1.0), line);
        chunk.emit_br(1, line);
        chunk.emit_end(line);

        chunk.emit_op_u16(Op::LOCAL_GET, left_part_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, right_part_slot, line);
        vybe_compiler::compiler::ops::emit_dyn_gt(chunk, line);
        vybe_compiler::compiler::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_br(1, line);
        chunk.emit_end(line);
    }

    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_end(line);
    chunk.patch_block(done);
}

pub fn emit_version_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let tostring_method_idx = push_version_tostring_chunk(chunks, line);
    let chunk = &mut chunks[current];
    let revision_slot = reserve_slot(chunk);
    let build_slot = reserve_slot(chunk);
    let minor_slot = reserve_slot(chunk);
    let major_slot = reserve_slot(chunk);

    match argc {
        4 => {
            chunk.emit_op_u16(Op::LOCAL_SET, revision_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, build_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, minor_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, major_slot, line);
        }
        3 => {
            push_const(chunk, Value::F64(-1.0), line);
            chunk.emit_op_u16(Op::LOCAL_SET, revision_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, build_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, minor_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, major_slot, line);
        }
        2 => {
            push_const(chunk, Value::F64(-1.0), line);
            chunk.emit_op_u16(Op::LOCAL_SET, revision_slot, line);
            push_const(chunk, Value::F64(-1.0), line);
            chunk.emit_op_u16(Op::LOCAL_SET, build_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, minor_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, major_slot, line);
        }
        _ => {
            push_const(chunk, Value::F64(-1.0), line);
            chunk.emit_op_u16(Op::LOCAL_SET, revision_slot, line);
            push_const(chunk, Value::F64(-1.0), line);
            chunk.emit_op_u16(Op::LOCAL_SET, build_slot, line);
            push_const(chunk, Value::F64(0.0), line);
            chunk.emit_op_u16(Op::LOCAL_SET, minor_slot, line);
            push_const(chunk, Value::F64(0.0), line);
            chunk.emit_op_u16(Op::LOCAL_SET, major_slot, line);
        }
    }

    emit_build_version_from_slots(
        chunks,
        current,
        tostring_method_idx,
        major_slot,
        minor_slot,
        build_slot,
        revision_slot,
        line,
    );
}

pub fn emit_version_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let tostring_method_idx = push_version_tostring_chunk(chunks, line);
    let to_str_idx = chunks[current].add_import("ecma:string", "String");
    let chunk = &mut chunks[current];
    let text_slot = reserve_slot(chunk);
    let parts_slot = reserve_slot(chunk);
    let len_slot = reserve_slot(chunk);
    let major_slot = reserve_slot(chunk);
    let minor_slot = reserve_slot(chunk);
    let build_slot = reserve_slot(chunk);
    let revision_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::CALL_IMPORT, to_str_idx, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);

    emit_version_parts_from_string(
        chunks,
        current,
        text_slot,
        parts_slot,
        len_slot,
        major_slot,
        minor_slot,
        build_slot,
        revision_slot,
        true,
        line,
    );

    emit_build_version_from_slots(
        chunks,
        current,
        tostring_method_idx,
        major_slot,
        minor_slot,
        build_slot,
        revision_slot,
        line,
    );
}

pub fn emit_version_try_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let tostring_method_idx = push_version_tostring_chunk(chunks, line);
    let to_str_idx = chunks[current].add_import("ecma:string", "String");
    let chunk = &mut chunks[current];
    let text_slot = reserve_slot(chunk);
    let parts_slot = reserve_slot(chunk);
    let len_slot = reserve_slot(chunk);
    let major_slot = reserve_slot(chunk);
    let minor_slot = reserve_slot(chunk);
    let build_slot = reserve_slot(chunk);
    let revision_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::CALL_IMPORT, to_str_idx, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);

    emit_version_parts_from_string(
        chunks,
        current,
        text_slot,
        parts_slot,
        len_slot,
        major_slot,
        minor_slot,
        build_slot,
        revision_slot,
        false,
        line,
    );

    let chunk = &mut chunks[current];
    let invalid_slot = reserve_slot(chunk);
    core_wasm::bool_const(chunk, line, false);
    chunk.emit_op_u16(Op::LOCAL_SET, invalid_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_i32_const(2, line);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_if(line);
    core_wasm::bool_const(chunk, line, true);
    chunk.emit_op_u16(Op::LOCAL_SET, invalid_slot, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_i32_const(4, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);
    core_wasm::bool_const(chunk, line, true);
    chunk.emit_op_u16(Op::LOCAL_SET, invalid_slot, line);
    chunk.emit_end(line);

    for (slot, min_len) in [
        (major_slot, 2),
        (minor_slot, 2),
        (build_slot, 3),
        (revision_slot, 4),
    ] {
        if min_len > 2 {
            chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
            chunk.emit_i32_const(min_len, line);
            chunk.emit_op(Op::I32_GE_S, line);
            chunk.emit_if(line);
        }
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::compiler::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        core_wasm::bool_const(chunk, line, true);
        chunk.emit_op_u16(Op::LOCAL_SET, invalid_slot, line);
        chunk.emit_end(line);

        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        push_const(chunk, Value::F64(0.0), line);
        chunk.emit_op(Op::F64_LT, line);
        chunk.emit_if(line);
        core_wasm::bool_const(chunk, line, true);
        chunk.emit_op_u16(Op::LOCAL_SET, invalid_slot, line);
        chunk.emit_end(line);
        if min_len > 2 {
            chunk.emit_end(line);
        }
    }

    chunk.emit_op_u16(Op::LOCAL_GET, invalid_slot, line);
    chunk.emit_if_value(line);
    chunk.emit_op(Op::NULL, line);
    chunk.emit_else(line);
    let _ = chunk;
    emit_build_version_from_slots(
        chunks,
        current,
        tostring_method_idx,
        major_slot,
        minor_slot,
        build_slot,
        revision_slot,
        line,
    );
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
}

pub fn emit_version_to_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let to_str_idx = chunks[current].add_import("ecma:string", "String");
    let chunk = &mut chunks[current];
    let field_count_slot = reserve_slot(chunk);
    let obj_slot = reserve_slot(chunk);
    let out_slot = reserve_slot(chunk);
    let defined_count_slot = reserve_slot(chunk);

    if argc > 0 {
        chunk.emit_op_u16(Op::LOCAL_SET, field_count_slot, line);
    } else {
        push_const(chunk, Value::F64(4.0), line);
        chunk.emit_op_u16(Op::LOCAL_SET, field_count_slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op_u16(Op::LOCAL_SET, defined_count_slot, line);
    emit_version_part(chunk, obj_slot, BUILD_KEY, line);
    push_const(chunk, Value::F64(-1.0), line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_if(line);
    push_const(chunk, Value::F64(3.0), line);
    chunk.emit_op_u16(Op::LOCAL_SET, defined_count_slot, line);
    chunk.emit_end(line);
    emit_version_part(chunk, obj_slot, REVISION_KEY, line);
    push_const(chunk, Value::F64(-1.0), line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_if(line);
    push_const(chunk, Value::F64(4.0), line);
    chunk.emit_op_u16(Op::LOCAL_SET, defined_count_slot, line);
    chunk.emit_end(line);

    if argc > 0 {
        chunk.emit_op_u16(Op::LOCAL_GET, field_count_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, defined_count_slot, line);
        chunk.emit_op(Op::F64_GT, line);
        chunk.emit_if(line);
        emit_throw_dotnet_exception(
            chunk,
            "ArgumentException",
            "Field count exceeds the number of defined Version components.",
            line,
        );
        chunk.emit_end(line);
    }

    emit_version_part(chunk, obj_slot, MAJOR_KEY, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, to_str_idx, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
    push_const(chunk, Value::String(Arc::from(".")), line);
    vybe_compiler::compiler::ops::emit_dyn_add(chunk, line);
    emit_version_part(chunk, obj_slot, MINOR_KEY, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, to_str_idx, line);
    chunk.emit(1, line);
    vybe_compiler::compiler::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);

    for (idx, key) in [(3.0, BUILD_KEY), (4.0, REVISION_KEY)] {
        if argc > 0 {
            chunk.emit_op_u16(Op::LOCAL_GET, field_count_slot, line);
            push_const(chunk, Value::F64(idx), line);
            chunk.emit_op(Op::F64_GE, line);
            chunk.emit_if(line);
        }
        emit_version_part(chunk, obj_slot, key, line);
        push_const(chunk, Value::F64(-1.0), line);
        chunk.emit_op(Op::F64_GT, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
        push_const(chunk, Value::String(Arc::from(".")), line);
        vybe_compiler::compiler::ops::emit_dyn_add(chunk, line);
        emit_version_part(chunk, obj_slot, key, line);
        chunk.emit_op_u16(Op::CALL_IMPORT, to_str_idx, line);
        chunk.emit(1, line);
        vybe_compiler::compiler::ops::emit_dyn_add(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
        chunk.emit_end(line);
        if argc > 0 {
            chunk.emit_end(line);
        }
    }

    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

pub fn emit_version_clone(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let tostring_method_idx = push_version_tostring_chunk(chunks, line);
    let chunk = &mut chunks[current];
    let obj_slot = reserve_slot(chunk);
    let major_slot = reserve_slot(chunk);
    let minor_slot = reserve_slot(chunk);
    let build_slot = reserve_slot(chunk);
    let revision_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    emit_version_part(chunk, obj_slot, MAJOR_KEY, line);
    chunk.emit_op_u16(Op::LOCAL_SET, major_slot, line);
    emit_version_part(chunk, obj_slot, MINOR_KEY, line);
    chunk.emit_op_u16(Op::LOCAL_SET, minor_slot, line);
    emit_version_part(chunk, obj_slot, BUILD_KEY, line);
    chunk.emit_op_u16(Op::LOCAL_SET, build_slot, line);
    emit_version_part(chunk, obj_slot, REVISION_KEY, line);
    chunk.emit_op_u16(Op::LOCAL_SET, revision_slot, line);
    emit_build_version_from_slots(
        chunks,
        current,
        tostring_method_idx,
        major_slot,
        minor_slot,
        build_slot,
        revision_slot,
        line,
    );
}

pub fn emit_version_compare(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_version_compare_internal(chunks, current, line);
}

pub fn emit_version_compare_instance(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_version_compare_internal(chunks, current, line);
    chunks[current].emit_op(Op::F64_NEG, line);
}

pub fn emit_version_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_version_compare_internal(chunks, current, line);
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(chunk, line);
    vybe_compiler::compiler::ops::emit_i32_to_bool(chunk, line);
}

pub fn emit_version_lt(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_version_compare_internal(chunks, current, line);
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::compiler::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(chunk, line);
    vybe_compiler::compiler::ops::emit_i32_to_bool(chunk, line);
}

pub fn emit_version_gt(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_version_compare_internal(chunks, current, line);
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::compiler::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(chunk, line);
    vybe_compiler::compiler::ops::emit_i32_to_bool(chunk, line);
}

pub fn emit_version_eq(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_version_equals(chunks, current, line);
}

pub fn emit_version_ne(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_version_equals(chunks, current, line);
    let chunk = &mut chunks[current];
    vybe_compiler::compiler::ops::emit_dyn_not(chunk, line);
}
