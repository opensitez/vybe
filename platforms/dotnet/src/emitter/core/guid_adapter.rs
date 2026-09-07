//! .NET `System.Guid` adapter — bytecode-only.
//!
//! `Guid` values are represented as plain Objects carrying the normalized
//! lowercase text form under `__value` plus a `__type="Guid"` tag so the
//! shared .NET dispatch layer can preserve value-type semantics without host
//! changes.

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::collections;
use vybe_compiler::primitives::functions::create_function_chunk;
use vybe_compiler::primitives::object::emit_bind_method_with_slot;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::object_fields::field_slot;

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

fn emit_throw_guid_format_exception(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::emitter::core::exceptions::emit_throw_typed(
        chunks,
        current,
        "FormatException",
        FORMAT_EXCEPTION_MSG,
        line,
    );
}

fn emit_guid_value_format(
    chunk: &mut Chunk,
    value_slot: u16,
    fmt_slot: u16,
    replace_all_idx: u16,
    concat_idx: u16,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    push_const(chunk, Value::String(Arc::from("N")), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    push_const(chunk, Value::String(Arc::from("-")), line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_call(replace_all_idx, 3, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    push_const(chunk, Value::String(Arc::from("B")), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::String(Arc::from("{")), line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(concat_idx, 2, line);
    push_const(chunk, Value::String(Arc::from("}")), line);
    chunk.emit_call(concat_idx, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    push_const(chunk, Value::String(Arc::from("P")), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::String(Arc::from("(")), line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(concat_idx, 2, line);
    push_const(chunk, Value::String(Arc::from(")")), line);
    chunk.emit_call(concat_idx, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    push_const(chunk, Value::String(Arc::from("X")), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::String(Arc::from("{0x")), line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_call(concat_idx, 2, line);
    push_const(chunk, Value::String(Arc::from("}")), line);
    chunk.emit_call(concat_idx, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn bind_guid_to_string(chunks: &mut Vec<Chunk>, current: usize, this_slot: u16, line: u32) {
    let mut method = create_function_chunk("__guid_tostring", 2);
    let replace_all_idx = method.add_import("ecma:string", "replaceAll");
    let concat_idx = method.add_import("ecma:string", "concat");
    let value_slot = reserve_slot(&mut method);
    method.emit_op_u16(Op::LOCAL_GET, 0, line);
    class_slots::emit_class_get(
        &mut method,
        ObjSource::Stack,
        &field_slot(VALUE_KEY),
        Dest::Stack,
        line,
    );
    method.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    emit_guid_value_format(
        &mut method,
        value_slot,
        1,
        replace_all_idx,
        concat_idx,
        line,
    );
    method.emit_op(Op::RETURN, line);
    method.local_count = method.local_count.max(3);
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

fn bind_guid_compare_to(chunks: &mut Vec<Chunk>, current: usize, this_slot: u16, line: u32) {
    let mut method = create_function_chunk("__guid_compareto", 2);
    let this_value_slot = reserve_slot(&mut method);
    let other_value_slot = reserve_slot(&mut method);
    method.emit_op_u16(Op::LOCAL_GET, 0, line);
    class_slots::emit_class_get(
        &mut method,
        ObjSource::Stack,
        &field_slot(VALUE_KEY),
        Dest::Stack,
        line,
    );
    method.emit_op_u16(Op::LOCAL_SET, this_value_slot, line);
    method.emit_op_u16(Op::LOCAL_GET, 1, line);
    class_slots::emit_class_get(
        &mut method,
        ObjSource::Stack,
        &field_slot(VALUE_KEY),
        Dest::Stack,
        line,
    );
    method.emit_op_u16(Op::LOCAL_SET, other_value_slot, line);
    method.emit_op_u16(Op::LOCAL_GET, this_value_slot, line);
    method.emit_op_u16(Op::LOCAL_GET, other_value_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut method, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut method, line);
    method.emit_if_value(line);
    push_const(&mut method, Value::I32(0), line);
    method.emit_else(line);
    method.emit_op_u16(Op::LOCAL_GET, this_value_slot, line);
    push_const(&mut method, Value::String(Arc::from(EMPTY_GUID)), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut method, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut method, line);
    method.emit_if_value(line);
    push_const(&mut method, Value::I32(-1), line);
    method.emit_else(line);
    push_const(&mut method, Value::I32(1), line);
    method.emit_end(line);
    method.emit_end(line);
    method.emit_op(Op::RETURN, line);
    method.local_count = method.local_count.max(4);
    chunks.push(method);
    let method_idx = chunks.len() - 1;
    for name in ["compareto", "CompareTo", "compare"] {
        emit_bind_method_with_slot(
            &mut chunks[current],
            this_slot,
            name,
            Some(vybe_ast::ProtocolSlot::Compare),
            method_idx,
            None,
            line,
        );
    }
}

fn emit_wrap_guid_from_slot(chunks: &mut Vec<Chunk>, current: usize, text_slot: u16, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = reserve_slot(chunk);

    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, Value::String(Arc::from("Guid")), line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(TYPE_KEY),
        ValueSource::Stack,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(VALUE_KEY),
        ValueSource::Stack,
        line,
    );

    bind_guid_to_string(chunks, current, obj_slot, line);
    bind_guid_compare_to(chunks, current, obj_slot, line);
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
    let obj_slot = reserve_slot(chunk);

    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, Value::String(Arc::from("Guid")), line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(TYPE_KEY),
        ValueSource::Stack,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(VALUE_KEY),
        ValueSource::Stack,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(BYTES_KEY),
        ValueSource::Stack,
        line,
    );

    bind_guid_to_string(chunks, current, obj_slot, line);
    bind_guid_compare_to(chunks, current, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

fn emit_validate_guid_text(
    chunks: &mut [Chunk],
    current: usize,
    test_idx: u16,
    text_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let ok_block = chunk.emit_block(line);
    push_const(chunk, Value::String(Arc::from(GUID_PATTERN)), line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunk.emit_call(test_idx, 2, line);
    chunk.emit_br_if(0, line);
    emit_throw_guid_format_exception(chunks, current, line);
    let chunk = &mut chunks[current];
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

    chunk.emit_call(to_str_idx, 1, line);
    push_const(chunk, Value::String(Arc::from("{")), line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_call(replace_all_idx, 3, line);
    push_const(chunk, Value::String(Arc::from("}")), line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_call(replace_all_idx, 3, line);
    push_const(chunk, Value::String(Arc::from("(")), line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_call(replace_all_idx, 3, line);
    push_const(chunk, Value::String(Arc::from(")")), line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_call(replace_all_idx, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);

    push_const(chunk, Value::String(Arc::from(GUID_N_PATTERN)), line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunk.emit_call(test_idx, 2, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    push_const(chunk, Value::F64(8.0), line);
    chunk.emit_call(substr_idx, 3, line);
    push_const(chunk, Value::String(Arc::from("-")), line);
    chunk.emit_call(concat_idx, 2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    push_const(chunk, Value::F64(8.0), line);
    push_const(chunk, Value::F64(4.0), line);
    chunk.emit_call(substr_idx, 3, line);
    chunk.emit_call(concat_idx, 2, line);
    push_const(chunk, Value::String(Arc::from("-")), line);
    chunk.emit_call(concat_idx, 2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    push_const(chunk, Value::F64(12.0), line);
    push_const(chunk, Value::F64(4.0), line);
    chunk.emit_call(substr_idx, 3, line);
    chunk.emit_call(concat_idx, 2, line);
    push_const(chunk, Value::String(Arc::from("-")), line);
    chunk.emit_call(concat_idx, 2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    push_const(chunk, Value::F64(16.0), line);
    push_const(chunk, Value::F64(4.0), line);
    chunk.emit_call(substr_idx, 3, line);
    chunk.emit_call(concat_idx, 2, line);
    push_const(chunk, Value::String(Arc::from("-")), line);
    chunk.emit_call(concat_idx, 2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    push_const(chunk, Value::F64(20.0), line);
    push_const(chunk, Value::F64(12.0), line);
    chunk.emit_call(substr_idx, 3, line);
    chunk.emit_call(concat_idx, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    chunk.emit_end(line);

    if validate {
        emit_validate_guid_text(chunks, current, test_idx, text_slot, line);
    }
    let chunk = &mut chunks[current];

    if normalize {
        chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
        chunk.emit_call(lower_idx, 1, line);
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
    chunk.emit_call(random_uuid_idx, 0, line);
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
            let array_length_idx = chunks[current].add_import("ecma:array", "length");
            let to_str_idx = chunks[current].add_import("ecma:string", "String");
            let chunk = &mut chunks[current];
            let value_slot = reserve_slot(chunk);
            let text_slot = reserve_slot(chunk);
            chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
            chunk.emit_call(is_array_idx, 1, line);
            chunk.emit_if_value(line);
            chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
            chunk.emit_call(array_length_idx, 1, line);
            push_const(chunk, Value::F64(16.0), line);
            vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
            chunk.emit_op(Op::I32_EQZ, line);
            chunk.emit_if(line);
            crate::emitter::core::exceptions::emit_throw_typed(
                chunks,
                current,
                "ArgumentException",
                "Byte array for Guid must be exactly 16 bytes long.",
                line,
            );
            let chunk = &mut chunks[current];
            chunk.emit_end(line);
            chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
            push_const(chunk, Value::F64(0.0), line);
            chunk.emit_op(Op::ARRAY_GET, line);
            chunk.emit_call(to_str_idx, 1, line);
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
            let text_slot = reserve_slot(chunk);
            push_const(
                chunk,
                Value::String(Arc::from("00000001-0002-0003-0405-060708090a0b")),
                line,
            );
            chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
            emit_wrap_guid_from_slot(chunks, current, text_slot, line);
        }
    }
}

pub fn emit_guid_to_byte_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let obj_slot = reserve_slot(chunk);
    let bytes_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(BYTES_KEY),
        Dest::Stack,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(VALUE_KEY),
        Dest::Stack,
        line,
    );
    for _ in 0..15 {
        push_const(chunk, Value::F64(0.0), line);
    }
    collections::emit_array_new(chunks, current, 16, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_end(line);
}

pub fn emit_guid_get_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(VALUE_KEY),
        Dest::Stack,
        line,
    );
    push_const(chunk, Value::F64(0.0), line);
    let char_code_idx = chunk.add_import("ecma:string", "charCodeAt");
    chunk.emit_call(char_code_idx, 2, line);
}

pub fn emit_guid_to_string(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let replace_all_idx = chunks[current].add_import("ecma:string", "replaceAll");
    let concat_idx = chunks[current].add_import("ecma:string", "concat");
    let chunk = &mut chunks[current];
    let obj_slot = reserve_slot(chunk);
    let fmt_slot = reserve_slot(chunk);
    // `argc` COUNTS THE RECEIVER — `Guid.ToString` is declared only as an
    // INSTANCE method, and `InstanceMethodTarget::Common` calls with
    // `arg_exprs.len() + 1`. Bare `g.ToString()` therefore arrives as `1`, so
    // `argc > 0` took the receiver itself for the format string and then read
    // the object off whatever sat below it.
    let has_format = argc > 1;
    if has_format {
        chunk.emit_op_u16(Op::LOCAL_SET, fmt_slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(VALUE_KEY),
        Dest::Stack,
        line,
    );
    if has_format {
        let value_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
        emit_guid_value_format(
            chunk,
            value_slot,
            fmt_slot,
            replace_all_idx,
            concat_idx,
            line,
        );
    }
}

pub fn emit_guid_parse_exact(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    emit_build_guid_from_stack(chunks, current, true, true, line);
}

pub fn emit_guid_compare_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let other_slot = reserve_slot(chunk);
    let this_slot = reserve_slot(chunk);
    let other_value_slot = reserve_slot(chunk);
    let this_value_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, other_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(VALUE_KEY),
        Dest::Stack,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_SET, this_value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, other_slot, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(VALUE_KEY),
        Dest::Stack,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_SET, other_value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, this_value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, other_value_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::I32(0), line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, this_value_slot, line);
    push_const(chunk, Value::String(Arc::from(EMPTY_GUID)), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::I32(-1), line);
    chunk.emit_else(line);
    push_const(chunk, Value::I32(1), line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_guid_try_parse_exact(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let test_idx = chunks[current].add_import("ecma:regexp", "test");
    let chunk = &mut chunks[current];
    let fmt_slot = reserve_slot(chunk);
    let text_slot = reserve_slot(chunk);
    if argc >= 3 {
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, fmt_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    for _ in 3..argc {
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    push_const(chunk, Value::String(Arc::from("N")), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::String(Arc::from(GUID_N_PATTERN)), line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunk.emit_call(test_idx, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_else(line);
    push_const(chunk, Value::String(Arc::from(GUID_PATTERN)), line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunk.emit_call(test_idx, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_end(line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
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
        chunk.emit_call(to_str_idx, 1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);

        push_const(chunk, Value::String(Arc::from(GUID_PATTERN)), line);
        chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
        chunk.emit_call(test_idx, 2, line);
        chunk.emit_if(line);

        chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
        chunk.emit_call(lower_idx, 1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    }
    emit_wrap_guid_from_slot(chunks, current, text_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}
