//! .NET `System.Version` adapter — bytecode-only.

use std::sync::Arc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

const TYPE_KEY: &str = "__type";
const MAJOR_KEY: &str = "Major";
const MINOR_KEY: &str = "Minor";
const BUILD_KEY: &str = "Build";
const REVISION_KEY: &str = "Revision";

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    let idx = chunk.add_constant(val);
    chunk.emit_op_u16(Op::CONST, idx, line);
}

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    let slot = chunk.local_count;
    chunk.local_count = slot + 1;
    slot
}

fn emit_array_get_const_index(chunk: &mut Chunk, array_slot: u16, index: f64, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, array_slot, line);
    push_const(chunk, Value::F64(index), line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

fn emit_parse_number_from_slot(chunks: &mut [Chunk], current: usize, text_slot: u16, line: u32) {
    let parse_int_idx = chunks[0].add_import("ecma:number", "parseInt");
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
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let use_default = chunk.emit_jump(Op::BR_IF_TRUE, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunk.emit_op(Op::REF_IS_UNDEFINED, line);
    let use_default_undefined = chunk.emit_jump(Op::BR_IF_TRUE, line);

    emit_parse_number_from_slot(chunks, current, text_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunk.emit_op(Op::DROP, line);
    let done = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(use_default);
    chunk.patch_jump(use_default_undefined);
    push_const(chunk, Value::F64(default_value), line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.patch_jump(done);
}

fn emit_build_version_from_slots(
    chunk: &mut Chunk,
    major_slot: u16,
    minor_slot: u16,
    build_slot: u16,
    revision_slot: u16,
    line: u32,
) {
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let major_key = chunk.add_constant(Value::String(Arc::from(MAJOR_KEY)));
    let minor_key = chunk.add_constant(Value::String(Arc::from(MINOR_KEY)));
    let build_key = chunk.add_constant(Value::String(Arc::from(BUILD_KEY)));
    let revision_key = chunk.add_constant(Value::String(Arc::from(REVISION_KEY)));

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op(Op::DUP, line);
    push_const(chunk, Value::String(Arc::from("Version")), line);
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, major_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, major_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, minor_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, minor_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, build_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, build_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, revision_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, revision_key, line);
    chunk.emit_op(Op::DROP, line);
}

fn emit_version_part(chunk: &mut Chunk, obj_slot: u16, key: &str, line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, key_idx, line);
}

fn emit_version_compare_internal(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let right_slot = reserve_slot(chunk);
    let left_slot = reserve_slot(chunk);
    let left_part_slot = reserve_slot(chunk);
    let right_part_slot = reserve_slot(chunk);
    let mut done_jumps = Vec::new();

    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
    chunk.emit_op(Op::DROP, line);

    for key in [MAJOR_KEY, MINOR_KEY, BUILD_KEY, REVISION_KEY] {
        emit_version_part(chunk, left_slot, key, line);
        chunk.emit_op_u16(Op::LOCAL_SET, left_part_slot, line);
        chunk.emit_op(Op::DROP, line);
        emit_version_part(chunk, right_slot, key, line);
        chunk.emit_op_u16(Op::LOCAL_SET, right_part_slot, line);
        chunk.emit_op(Op::DROP, line);

        chunk.emit_op_u16(Op::LOCAL_GET, left_part_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, right_part_slot, line);
        chunk.emit_op(Op::DYN_LT, line);
        let not_lt = chunk.emit_jump(Op::BR_IF_FALSE, line);
        push_const(chunk, Value::F64(-1.0), line);
        done_jumps.push(chunk.emit_jump(Op::BR, line));
        chunk.patch_jump(not_lt);

        chunk.emit_op_u16(Op::LOCAL_GET, left_part_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, right_part_slot, line);
        chunk.emit_op(Op::DYN_GT, line);
        let not_gt = chunk.emit_jump(Op::BR_IF_FALSE, line);
        push_const(chunk, Value::F64(1.0), line);
        done_jumps.push(chunk.emit_jump(Op::BR, line));
        chunk.patch_jump(not_gt);
    }

    push_const(chunk, Value::F64(0.0), line);
    for jump in done_jumps {
        chunk.patch_jump(jump);
    }
}

pub fn emit_version_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let revision_slot = reserve_slot(chunk);
    let build_slot = reserve_slot(chunk);
    let minor_slot = reserve_slot(chunk);
    let major_slot = reserve_slot(chunk);

    match argc {
        4 => {
            chunk.emit_op_u16(Op::LOCAL_SET, revision_slot, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op_u16(Op::LOCAL_SET, build_slot, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op_u16(Op::LOCAL_SET, minor_slot, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op_u16(Op::LOCAL_SET, major_slot, line);
            chunk.emit_op(Op::DROP, line);
        }
        3 => {
            push_const(chunk, Value::F64(-1.0), line);
            chunk.emit_op_u16(Op::LOCAL_SET, revision_slot, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op_u16(Op::LOCAL_SET, build_slot, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op_u16(Op::LOCAL_SET, minor_slot, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op_u16(Op::LOCAL_SET, major_slot, line);
            chunk.emit_op(Op::DROP, line);
        }
        2 => {
            push_const(chunk, Value::F64(-1.0), line);
            chunk.emit_op_u16(Op::LOCAL_SET, revision_slot, line);
            chunk.emit_op(Op::DROP, line);
            push_const(chunk, Value::F64(-1.0), line);
            chunk.emit_op_u16(Op::LOCAL_SET, build_slot, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op_u16(Op::LOCAL_SET, minor_slot, line);
            chunk.emit_op(Op::DROP, line);
            chunk.emit_op_u16(Op::LOCAL_SET, major_slot, line);
            chunk.emit_op(Op::DROP, line);
        }
        _ => {
            push_const(chunk, Value::F64(-1.0), line);
            chunk.emit_op_u16(Op::LOCAL_SET, revision_slot, line);
            chunk.emit_op(Op::DROP, line);
            push_const(chunk, Value::F64(-1.0), line);
            chunk.emit_op_u16(Op::LOCAL_SET, build_slot, line);
            chunk.emit_op(Op::DROP, line);
            push_const(chunk, Value::F64(0.0), line);
            chunk.emit_op_u16(Op::LOCAL_SET, minor_slot, line);
            chunk.emit_op(Op::DROP, line);
            push_const(chunk, Value::F64(0.0), line);
            chunk.emit_op_u16(Op::LOCAL_SET, major_slot, line);
            chunk.emit_op(Op::DROP, line);
        }
    }

    emit_build_version_from_slots(chunk, major_slot, minor_slot, build_slot, revision_slot, line);
}

pub fn emit_version_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    let to_str_idx = chunks[0].add_import("ecma:string", "String");
    let chunk = &mut chunks[current];
    let text_slot = reserve_slot(chunk);
    let parts_slot = reserve_slot(chunk);
    let major_slot = reserve_slot(chunk);
    let minor_slot = reserve_slot(chunk);
    let build_slot = reserve_slot(chunk);
    let revision_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::CALL_IMPORT, to_str_idx, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    push_const(chunk, Value::String(Arc::from(".")), line);
    chunk.emit_op(Op::STR_SPLIT, line);
    chunk.emit_op_u16(Op::LOCAL_SET, parts_slot, line);
    chunk.emit_op(Op::DROP, line);

    emit_store_optional_array_part_as_number(chunks, current, parts_slot, 0.0, major_slot, 0.0, line);
    emit_store_optional_array_part_as_number(chunks, current, parts_slot, 1.0, minor_slot, 0.0, line);
    emit_store_optional_array_part_as_number(chunks, current, parts_slot, 2.0, build_slot, -1.0, line);
    emit_store_optional_array_part_as_number(chunks, current, parts_slot, 3.0, revision_slot, -1.0, line);

    let chunk = &mut chunks[current];
    emit_build_version_from_slots(chunk, major_slot, minor_slot, build_slot, revision_slot, line);
}

pub fn emit_version_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let to_str_idx = chunks[0].add_import("ecma:string", "String");
    let chunk = &mut chunks[current];
    let obj_slot = reserve_slot(chunk);
    let out_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunk.emit_op(Op::DROP, line);

    emit_version_part(chunk, obj_slot, MAJOR_KEY, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, to_str_idx, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
    push_const(chunk, Value::String(Arc::from(".")), line);
    chunk.emit_op(Op::DYN_ADD, line);
    emit_version_part(chunk, obj_slot, MINOR_KEY, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, to_str_idx, line);
    chunk.emit(1, line);
    chunk.emit_op(Op::DYN_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunk.emit_op(Op::DROP, line);

    for key in [BUILD_KEY, REVISION_KEY] {
        emit_version_part(chunk, obj_slot, key, line);
        push_const(chunk, Value::F64(-1.0), line);
        chunk.emit_op(Op::DYN_GT, line);
        let skip = chunk.emit_jump(Op::BR_IF_FALSE, line);
        chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
        push_const(chunk, Value::String(Arc::from(".")), line);
        chunk.emit_op(Op::DYN_ADD, line);
        emit_version_part(chunk, obj_slot, key, line);
        chunk.emit_op_u16(Op::CALL_IMPORT, to_str_idx, line);
        chunk.emit(1, line);
        chunk.emit_op(Op::DYN_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
        chunk.emit_op(Op::DROP, line);
        chunk.patch_jump(skip);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

pub fn emit_version_compare(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_version_compare_internal(chunks, current, line);
}

pub fn emit_version_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_version_compare_internal(chunks, current, line);
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_EQ, line);
}

pub fn emit_version_lt(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_version_compare_internal(chunks, current, line);
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_LT, line);
}

pub fn emit_version_gt(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_version_compare_internal(chunks, current, line);
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_GT, line);
}

pub fn emit_version_eq(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_version_equals(chunks, current, line);
}

pub fn emit_version_ne(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_version_equals(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DYN_NOT, line);
}