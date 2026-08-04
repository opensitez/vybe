//! Dart string helpers — Rust inline opcode emitters.
//!
//! Composes `compiler_common` emitters (collections + strings) so the
//! provider can be swapped in one place. No raw opcode emits where a
//! `common::*` helper exists.

use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};
use vybe_compiler::primitives::instructions::{core_wasm, host};
use vybe_compiler::primitives::{
    collections, errors, functions, generators, loops, reflection, strings };

const SB_BUFFER_KEY: &str = "__dart_string_buffer";
const SB_MARKER_KEY: &str = "__dart_string_buffer_marker";
const URI_HREF_KEY: &str = "href";
const URI_MARKER_KEY: &str = "__dart_uri_marker";
const SET_MARKER_KEY: &str = "__dart_set_marker";
const STOPWATCH_MARKER_KEY: &str = "__dart_stopwatch_marker";
const STOPWATCH_RUNNING_KEY: &str = "isrunning";
const MAP_ORDER_KEY: &str = "__dart_map_order";
const SORTED_MAP_MARKER_KEY: &str = "__dart_sorted_map_marker";

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn string_key(chunk: &mut Chunk, key: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(key)))
}

fn emit_sb_buffer_get(chunk: &mut Chunk, line: u32) {
    let key = string_key(chunk, SB_BUFFER_KEY);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
}

fn emit_sb_marker_test(chunk: &mut Chunk, line: u32) {
    let key = string_key(chunk, SB_MARKER_KEY);
    core_wasm::dup(chunk, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

fn emit_dart_value_to_string(chunk: &mut Chunk, line: u32) {
    let to_str = chunk.add_import("ecma:string", "String");
    chunk.emit_op_u16(Op::CALL_IMPORT, to_str, line);
    chunk.emit(1, line);
}

fn emit_string_field(chunk: &mut Chunk, key: &str, value: &str, line: u32) {
    core_wasm::dup(chunk, line);
    chunk.emit_string_const(value, line);
    let key = string_key(chunk, key);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    chunk.emit_op(Op::DROP, line);
}

fn stamp_runtime_type(
    chunk: &mut Chunk,
    dart_name: &str,
    kind: reflection::ReflectKind,
    line: u32,
) {
    emit_string_field(chunk, reflection::FIELD_TYPE, dart_name, line);
    emit_string_field(chunk, reflection::FIELD_TYPE_NAME, dart_name, line);
    emit_string_field(chunk, reflection::FIELD_KIND, kind.as_str(), line);
}

/// Shared by every Dart exception emit, including `io_adapter`'s — one place
/// builds the object, stamps `__exception_type`, and wires the instanceof
/// chain, so `on FileSystemException` catches what the io adapter throws.
pub(crate) fn emit_dart_exception_new(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    dart_name: &str,
    types: &[&str],
    line: u32,
) {
    let chunk = &mut chunks[current];
    let msg_slot = reserve_slot(chunk);
    if argc > 0 {
        chunk.emit_op_u16(Op::LOCAL_SET, msg_slot, line);
    } else {
        chunk.emit_string_const("", line);
        chunk.emit_op_u16(Op::LOCAL_SET, msg_slot, line);
    }
    chunk.emit_struct_new(0, 0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, msg_slot, line);
    errors::emit_exception_new_finalize(chunk, dart_name, line);
    stamp_runtime_type(chunk, dart_name, reflection::ReflectKind::Exception, line);
    emit_string_field(chunk, "__exception_type", dart_name, line);
    emit_string_field(chunk, "name", dart_name, line);
    let obj_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    for ty in types {
        vybe_compiler::primitives::reflection::emit_instanceof_chain(
            chunks, current, obj_slot, ty, line,
        );
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

fn emit_is_generator(chunk: &mut Chunk, line: u32) {
    let is_gen = chunk.add_import("ecma:value", "isGenerator");
    chunk.emit_call(is_gen, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

fn emit_slot_is_bigint(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    host::emit(chunk, "wasm:js-bigint", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

fn emit_bigint_i32(chunks: &mut [Chunk], current: usize, value: i32, line: u32) {
    chunks[current].emit_i32_const(value, line);
    host::emit(&mut chunks[current], "ecma:bigint", "BigInt", 1, line);
}

fn emit_dart_sb_append_value(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    let chunk = &mut chunks[current];
    let sb_slot = reserve_slot(chunk);
    let buffer_key = string_key(chunk, SB_BUFFER_KEY);
    chunk.emit_op_u16(Op::LOCAL_SET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, buffer_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_dart_value_to_string(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
}

/// Dart `s.isEmpty` — true iff length == 0. Stack: [s] → [bool].
pub fn emit_dart_is_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_dart_length(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Dart `s.isNotEmpty` — true iff length != 0. Stack: [s] → [bool].
pub fn emit_dart_is_not_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_dart_length(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_ne(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Dart `n.isEven` — true iff `n % 2 == 0`. Stack: [n] → [bool].
pub fn emit_dart_is_even(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    emit_slot_is_bigint(&mut chunks[current], value_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_bigint_i32(chunks, current, 2, line);
    host::emit(&mut chunks[current], "ecma:bigint", "rem", 2, line);
    emit_bigint_i32(chunks, current, 0, line);
    host::emit(&mut chunks[current], "ecma:bigint", "eq", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_REM_S, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

/// Dart `n.isOdd` — true iff `n % 2 != 0`. Stack: [n] → [bool].
pub fn emit_dart_is_odd(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    emit_slot_is_bigint(&mut chunks[current], value_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_bigint_i32(chunks, current, 2, line);
    host::emit(&mut chunks[current], "ecma:bigint", "rem", 2, line);
    emit_bigint_i32(chunks, current, 0, line);
    host::emit(&mut chunks[current], "ecma:bigint", "ne", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_REM_S, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_ne(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

/// Dart `StringBuffer()` — plain object with one mutable string field.
pub fn emit_dart_sb_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = string_key(chunk, SB_BUFFER_KEY);
    let marker_key = string_key(chunk, SB_MARKER_KEY);
    chunk.emit_struct_new(0, 0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_string_const("", line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    core_wasm::dup(chunk, line);
    chunk.emit_bool_const(true, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, marker_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Dart `buf.write(value)` — append stringified value, return receiver.
pub fn emit_dart_sb_write(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    emit_dart_sb_append_value(chunks, current, value_slot, line);
}

/// Dart `buf.writeln([value])` — append value then newline, return receiver.
pub fn emit_dart_sb_writeln(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    if argc > 1 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    } else {
        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    }
    emit_dart_sb_append_value(chunks, current, value_slot, line);
    let newline_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_string_const("\n", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, newline_slot, line);
    emit_dart_sb_append_value(chunks, current, newline_slot, line);
}

/// Dart `buf.writeAll(iterable, [separator])`.
pub fn emit_dart_sb_write_all(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let sep_slot = reserve_slot(&mut chunks[current]);
    let iterable_slot = reserve_slot(&mut chunks[current]);
    if argc > 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, sep_slot, line);
    } else {
        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, sep_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, iterable_slot, line);
    let joined_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, iterable_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep_slot, line);
    collections::emit_join(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, joined_slot, line);
    emit_dart_sb_append_value(chunks, current, joined_slot, line);
}

/// Dart `buf.writeCharCode(code)`.
pub fn emit_dart_sb_write_char_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "fromCharCode",
        1,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    emit_dart_sb_append_value(chunks, current, value_slot, line);
}

/// Dart `buf.clear()` — empty mutable content, return receiver.
pub fn emit_dart_sb_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let sb_slot = reserve_slot(chunk);
    let buffer_key = string_key(chunk, SB_BUFFER_KEY);
    chunk.emit_op_u16(Op::LOCAL_SET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_string_const("", line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
}

/// Dart `RegExp(pattern, ...)` after walker flag-normalisation.
/// Stack: [pattern] or [pattern, flags] -> [regexp].
pub fn emit_dart_regexp_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        chunks[current].emit_string_const("", line);
    }
    host::emit(&mut chunks[current], "ecma:regexp", "new", 2, line);
}

/// Dart `re.hasMatch(input)`.
pub fn emit_dart_regexp_has_match(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:regexp", "test", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Dart `re.firstMatch(input)` -> JS match array/null.
pub fn emit_dart_regexp_first_match(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:regexp", "exec", 2, line);
}

/// Dart `re.allMatches(input)` -> iterable match array.
pub fn emit_dart_regexp_all_matches(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:regexp", "matchAll", 2, line);
}

/// Dart `match.group(index)`; JS match arrays store group 0..N by index.
pub fn emit_dart_regexp_group(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_get(chunks, current, line);
}

pub fn emit_dart_exception(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_dart_exception_new(chunks, current, argc, "Exception", &["Exception"], line);
}

pub fn emit_dart_format_exception(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_dart_exception_new(
        chunks,
        current,
        argc,
        "FormatException",
        &["FormatException", "Exception"],
        line,
    );
}

pub fn emit_dart_range_error(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_dart_exception_new(
        chunks,
        current,
        argc,
        "RangeError",
        &["RangeError", "Error"],
        line,
    );
}

pub fn emit_dart_state_error(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_dart_exception_new(
        chunks,
        current,
        argc,
        "StateError",
        &["StateError", "Error"],
        line,
    );
}

pub fn emit_dart_argument_error(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_dart_exception_new(
        chunks,
        current,
        argc,
        "ArgumentError",
        &["ArgumentError", "Error"],
        line,
    );
}

pub fn emit_dart_unimplemented_error(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_dart_exception_new(
        chunks,
        current,
        argc,
        "UnimplementedError",
        &["UnimplementedError", "Error"],
        line,
    );
}

pub fn emit_dart_stack_trace(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let _error_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, _error_slot, line);
    chunk.emit_struct_new(0, 0, line);
    stamp_runtime_type(chunk, "StackTrace", reflection::ReflectKind::Object, line);
    emit_string_field(chunk, "__exception_type", "StackTrace", line);
    emit_string_field(chunk, "name", "StackTrace", line);
    emit_string_field(chunk, "message", "StackTrace", line);
    let obj_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    vybe_compiler::primitives::reflection::emit_instanceof_chain(
        chunks,
        current,
        obj_slot,
        "StackTrace",
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

fn emit_dart_format_exception_throw(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_struct_new(0, 0, line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_string_const("Invalid number", line);
    errors::emit_exception_new_finalize(&mut chunks[current], "FormatException", line);
    stamp_runtime_type(
        &mut chunks[current],
        "FormatException",
        reflection::ReflectKind::Exception,
        line,
    );
    emit_string_field(
        &mut chunks[current],
        "__exception_type",
        "FormatException",
        line,
    );
    emit_string_field(&mut chunks[current], "name", "FormatException", line);
    let exc_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, exc_slot, line);
    vybe_compiler::primitives::reflection::emit_instanceof_chain(
        chunks,
        current,
        exc_slot,
        "FormatException",
        line,
    );
    vybe_compiler::primitives::reflection::emit_instanceof_chain(
        chunks,
        current,
        exc_slot,
        "Exception",
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, exc_slot, line);
    errors::emit_throw(&mut chunks[current], line);
}

fn emit_trimmed_string_empty(chunk: &mut Chunk, text_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    strings::emit_trim(chunk, line);
    strings::emit_length(chunk, line);
    chunk.emit_i32_const(0, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

fn emit_number_is_nan(chunk: &mut Chunk, number_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, number_slot, line);
    host::emit(chunk, "ecma:number", "isNaN", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

fn emit_number_is_finite(chunk: &mut Chunk, number_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, number_slot, line);
    host::emit(chunk, "ecma:number", "isFinite", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

fn emit_parse_int_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    text_slot: u16,
    radix_slot: Option<u16>,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    if let Some(radix_slot) = radix_slot {
        chunks[current].emit_op_u16(Op::LOCAL_GET, radix_slot, line);
        host::emit(&mut chunks[current], "ecma:number", "parseInt", 2, line);
    } else {
        host::emit(&mut chunks[current], "ecma:number", "Number", 1, line);
    }
}

fn emit_regexp_test_slot(chunk: &mut Chunk, text_slot: u16, pattern: &str, line: u32) {
    chunk.emit_string_const(pattern, line);
    chunk.emit_string_const("", line);
    host::emit(chunk, "ecma:regexp", "new", 2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    host::emit(chunk, "ecma:regexp", "test", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

fn emit_radix_eq(chunk: &mut Chunk, radix_slot: u16, radix: i32, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, radix_slot, line);
    core_wasm::i32_const(chunk, line, radix);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

fn emit_int_radix_text_valid(chunk: &mut Chunk, text_slot: u16, radix_slot: u16, line: u32) {
    emit_radix_eq(chunk, radix_slot, 2, line);
    chunk.emit_if(line);
    emit_regexp_test_slot(chunk, text_slot, "^[+-]?[01]+$", line);
    chunk.emit_else(line);
    emit_radix_eq(chunk, radix_slot, 8, line);
    chunk.emit_if(line);
    emit_regexp_test_slot(chunk, text_slot, "^[+-]?[0-7]+$", line);
    chunk.emit_else(line);
    emit_radix_eq(chunk, radix_slot, 10, line);
    chunk.emit_if(line);
    emit_regexp_test_slot(chunk, text_slot, "^[+-]?[0-9]+$", line);
    chunk.emit_else(line);
    emit_radix_eq(chunk, radix_slot, 16, line);
    chunk.emit_if(line);
    emit_regexp_test_slot(chunk, text_slot, "^[+-]?[0-9a-fA-F]+$", line);
    chunk.emit_else(line);
    emit_radix_eq(chunk, radix_slot, 36, line);
    chunk.emit_if(line);
    emit_regexp_test_slot(chunk, text_slot, "^[+-]?[0-9a-zA-Z]+$", line);
    chunk.emit_else(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_int_parse_result_or_null(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    throw_on_error: bool,
    line: u32,
) {
    let radix_slot = if argc >= 2 {
        let slot = reserve_slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
        Some(slot)
    } else {
        None
    };
    let text_slot = reserve_slot(&mut chunks[current]);
    let value_slot = reserve_slot(&mut chunks[current]);
    let int_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);

    emit_trimmed_string_empty(&mut chunks[current], text_slot, line);
    chunks[current].emit_if(line);
    if throw_on_error {
        emit_dart_format_exception_throw(chunks, current, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    chunks[current].emit_else(line);

    if let Some(radix_slot) = radix_slot {
        emit_int_radix_text_valid(&mut chunks[current], text_slot, radix_slot, line);
        chunks[current].emit_if(line);
    }
    emit_parse_int_from_slots(chunks, current, text_slot, radix_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    emit_number_is_nan(&mut chunks[current], value_slot, line);
    if radix_slot.is_none() {
        emit_number_is_finite(&mut chunks[current], value_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
        chunks[current].emit_op(Op::I32_OR, line);
    }
    chunks[current].emit_if(line);
    if throw_on_error {
        emit_dart_format_exception_throw(chunks, current, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op(Op::F64_TRUNC, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, int_slot, line);
    if radix_slot.is_none() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, int_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, int_slot, line);
    chunks[current].emit_op(Op::I32_FROM_F64, line);
    if radix_slot.is_none() {
        chunks[current].emit_else(line);
        if throw_on_error {
            emit_dart_format_exception_throw(chunks, current, line);
        } else {
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        }
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);
    if radix_slot.is_some() {
        chunks[current].emit_else(line);
        if throw_on_error {
            emit_dart_format_exception_throw(chunks, current, line);
        } else {
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        }
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);
}

pub fn emit_dart_int_parse(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_int_parse_result_or_null(chunks, current, argc, true, line);
}

pub fn emit_dart_int_try_parse(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_int_parse_result_or_null(chunks, current, argc, false, line);
}

pub fn emit_dart_double_try_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    let text_slot = reserve_slot(&mut chunks[current]);
    let value_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);
    emit_trimmed_string_empty(&mut chunks[current], text_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    host::emit(&mut chunks[current], "ecma:number", "Number", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    emit_number_is_nan(&mut chunks[current], value_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    strings::emit_trim(&mut chunks[current], line);
    chunks[current].emit_string_const("NaN", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_dart_double_parse(chunks: &mut [Chunk], current: usize, line: u32) {
    let text_slot = reserve_slot(&mut chunks[current]);
    let value_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);
    emit_trimmed_string_empty(&mut chunks[current], text_slot, line);
    chunks[current].emit_if(line);
    emit_dart_format_exception_throw(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    host::emit(&mut chunks[current], "ecma:number", "Number", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    emit_number_is_nan(&mut chunks[current], value_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    strings::emit_trim(&mut chunks[current], line);
    chunks[current].emit_string_const("NaN", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_else(line);
    emit_dart_format_exception_throw(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_dart_bigint_from(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:bigint", "BigInt", 1, line);
}

pub fn emit_dart_stream_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    let out_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

pub fn emit_dart_stream_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_array_new(chunks, current, 0, line);
}

pub fn emit_dart_stream_error(chunks: &mut [Chunk], current: usize, line: u32) {
    let error_slot = reserve_slot(&mut chunks[current]);
    let stream_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, error_slot, line);
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stream_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunks[current].emit_bool_const(true, line);
    let marker_key = string_key(&mut chunks[current], "__dart_stream_error");
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, marker_key, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_set_string_field_from_slot(&mut chunks[current], stream_slot, "error", error_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stream_slot, line);
}

fn emit_get_string_field_to_slot(
    chunk: &mut Chunk,
    obj_slot: u16,
    key: &str,
    dst_slot: u16,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let key = string_key(chunk, key);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, dst_slot, line);
}

fn emit_set_bool_field(chunk: &mut Chunk, obj_slot: u16, key: &str, value: bool, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_bool_const(value, line);
    let key = string_key(chunk, key);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    chunk.emit_op(Op::DROP, line);
}

pub fn emit_dart_stream_listen(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let done_slot = reserve_slot(&mut chunks[current]);
    let error_or_done_slot = reserve_slot(&mut chunks[current]);
    let on_data_slot = reserve_slot(&mut chunks[current]);
    let stream_slot = reserve_slot(&mut chunks[current]);
    if argc > 3 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, done_slot, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, done_slot, line);
    }
    if argc > 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, error_or_done_slot, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, error_or_done_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, on_data_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stream_slot, line);
    chunks[current].emit_struct_new(0, 0, line);
    let sub_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sub_slot, line);
    emit_set_string_field_from_slot(&mut chunks[current], sub_slot, "stream", stream_slot, line);
    emit_set_string_field_from_slot(&mut chunks[current], sub_slot, "onData", on_data_slot, line);
    emit_set_string_field_from_slot(
        &mut chunks[current],
        sub_slot,
        "errorOrDone",
        error_or_done_slot,
        line,
    );
    emit_set_string_field_from_slot(&mut chunks[current], sub_slot, "onDone", done_slot, line);
    emit_set_bool_field(&mut chunks[current], sub_slot, "cancelled", false, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sub_slot, line);
}

pub fn emit_dart_stream_cancel(chunks: &mut [Chunk], current: usize, line: u32) {
    let sub_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sub_slot, line);
    emit_set_bool_field(&mut chunks[current], sub_slot, "cancelled", true, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_dart_stream_as_future(chunks: &mut [Chunk], current: usize, line: u32) {
    let sub_slot = reserve_slot(&mut chunks[current]);
    let stream_slot = reserve_slot(&mut chunks[current]);
    let on_data_slot = reserve_slot(&mut chunks[current]);
    let error_or_done_slot = reserve_slot(&mut chunks[current]);
    let done_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sub_slot, line);
    emit_get_string_field_to_slot(&mut chunks[current], sub_slot, "stream", stream_slot, line);
    emit_get_string_field_to_slot(&mut chunks[current], sub_slot, "onData", on_data_slot, line);
    emit_get_string_field_to_slot(
        &mut chunks[current],
        sub_slot,
        "errorOrDone",
        error_or_done_slot,
        line,
    );
    emit_get_string_field_to_slot(&mut chunks[current], sub_slot, "onDone", done_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    let marker_key = string_key(&mut chunks[current], "__dart_stream_error");
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, marker_key, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    let err_slot = reserve_slot(&mut chunks[current]);
    emit_get_string_field_to_slot(&mut chunks[current], stream_slot, "error", err_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, error_or_done_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, error_or_done_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, err_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, err_slot, line);
    errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_else(line);
    let arr_slot = materialize_slot(chunks, current, stream_slot, line);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sub_slot, line);
    let cancelled_key = string_key(&mut chunks[current], "cancelled");
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, cancelled_key, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, on_data_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, done_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, error_or_done_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, done_slot, line);
    chunks[current].emit_end(line);
    let callback_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, callback_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, callback_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, callback_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

pub fn emit_dart_queue_remove_first(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_shift(chunks, current, line);
}

pub fn emit_dart_bigint_parse(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 1 {
        let radix_slot = reserve_slot(&mut chunks[current]);
        let text_slot = reserve_slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, radix_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);
        emit_parse_int_from_slots(chunks, current, text_slot, Some(radix_slot), line);
        host::emit(&mut chunks[current], "ecma:bigint", "BigInt", 1, line);
    } else {
        host::emit(&mut chunks[current], "ecma:bigint", "BigInt", 1, line);
    }
}

pub fn emit_dart_bigint_abs(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    let result_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_bigint_i32(chunks, current, 0, line);
    host::emit(&mut chunks[current], "ecma:bigint", "lt", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(&mut chunks[current], "ecma:bigint", "neg", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_dart_bigint_gcd(chunks: &mut [Chunk], current: usize, line: u32) {
    let b_slot = reserve_slot(&mut chunks[current]);
    let a_slot = reserve_slot(&mut chunks[current]);
    let tmp_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a_slot, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
    emit_bigint_i32(chunks, current, 0, line);
    host::emit(&mut chunks[current], "ecma:bigint", "eq", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
    host::emit(&mut chunks[current], "ecma:bigint", "rem", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, tmp_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tmp_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
    emit_dart_bigint_abs(chunks, current, line);
}

pub fn emit_dart_stopwatch_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_struct_new(0, 0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_bool_const(true, line);
    let marker_key = string_key(chunk, STOPWATCH_MARKER_KEY);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, marker_key, line);
    chunk.emit_op(Op::DROP, line);
    core_wasm::dup(chunk, line);
    chunk.emit_bool_const(false, line);
    let running_key = string_key(chunk, STOPWATCH_RUNNING_KEY);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, running_key, line);
    chunk.emit_op(Op::DROP, line);
}

fn emit_dart_stopwatch_set_running(chunks: &mut [Chunk], current: usize, running: bool, line: u32) {
    let sw_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sw_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sw_slot, line);
    chunks[current].emit_bool_const(running, line);
    let running_key = string_key(&mut chunks[current], STOPWATCH_RUNNING_KEY);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, running_key, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_dart_stopwatch_start(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_dart_stopwatch_set_running(chunks, current, true, line);
}

pub fn emit_dart_stopwatch_stop(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_dart_stopwatch_set_running(chunks, current, false, line);
}

pub fn emit_dart_stopwatch_reset(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_dart_stopwatch_set_running(chunks, current, false, line);
}

pub fn emit_dart_stopwatch_is_running(chunks: &mut [Chunk], current: usize, line: u32) {
    let running_key = string_key(&mut chunks[current], STOPWATCH_RUNNING_KEY);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, running_key, line);
}

pub fn emit_dart_stopwatch_elapsed(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    crate::emitter::core_adapter::emit_duration_zero(chunks, current, line);
}

pub fn emit_dart_stopwatch_elapsed_milliseconds(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_i32_const(0, line);
}

pub fn emit_dart_stopwatch_elapsed_microseconds(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_i32_const(0, line);
}

pub fn emit_dart_index_get(chunks: &mut [Chunk], current: usize, line: u32) {
    let index_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    // A STRING receiver must not go through `ecma:array.get` — it returned
    // undefined for every index, so `"café"[3]` and even `"abc"[1]` were null.
    // Dart indexes a string by UTF-16 code unit and yields a one-character
    // String, which is `ecma:string.charAt`.
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    host::emit(&mut chunks[current], "ecma:number", "Number", 1, line);
    host::emit(&mut chunks[current], "ecma:string", "charAt", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    host::emit(&mut chunks[current], "ecma:number", "Number", 1, line);
    host::emit(&mut chunks[current], "ecma:array", "get", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    emit_undefined_to_null(&mut chunks[current], line);
}

/// Dart `print(value)` — calls value.toString() before logging, per
/// dart:core. Stack: [value] → []. Composes ecma:string.String (which
/// invokes the object's toString method) and wasi:cli.log. Import
/// tables are PER CHUNK: register on chunks[current], the chunk whose
/// CALL_IMPORT indexes them (registering on chunks[0] made the index
/// misresolve whenever the tables diverged).
pub fn emit_dart_print(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    use vybe_runtime::Op as VOp;
    emit_dart_to_string(chunks, current, line);
    let log_idx = chunks[current].add_import("wasi:logging/logging", "log");
    chunks[current].emit_op_u16(VOp::CALL_IMPORT, log_idx, line);
    chunks[current].emit(1, line);
}

/// Dart `value.toString()` — route through ECMA string coercion.
/// Stack: [value] → [string].
pub fn emit_dart_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let uri_key = string_key(&mut chunks[current], URI_HREF_KEY);
    let uri_marker = string_key(&mut chunks[current], URI_MARKER_KEY);
    core_wasm::dup(&mut chunks[current], line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    let tuple_slot = reserve_slot(&mut chunks[current]);
    core_wasm::dup(&mut chunks[current], line);
    vybe_compiler::primitives::tuples::emit_is_tuple(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, tuple_slot, line);
    chunks[current].emit_string_const(", ", line);
    collections::emit_join(chunks, current, line);
    let joined_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, joined_slot, line);
    chunks[current].emit_string_const("[", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, joined_slot, line);
    chunks[current].emit_string_const("]", line);
    strings::emit_concat(&mut chunks[current], 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tuple_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    vybe_compiler::primitives::tuples::emit_list_string_to_tuple(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    core_wasm::dup(&mut chunks[current], line);
    host::emit(&mut chunks[current], "wasm:js-bigint", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    host::emit(&mut chunks[current], "ecma:bigint", "toString", 1, line);
    chunks[current].emit_else(line);
    core_wasm::dup(&mut chunks[current], line);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, uri_marker, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, uri_key, line);
    chunks[current].emit_else(line);
    emit_sb_marker_test(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_sb_buffer_get(&mut chunks[current], line);
    chunks[current].emit_else(line);
    core_wasm::dup(&mut chunks[current], line);
    emit_dart_enum_like_to_string(chunks, current, line);
    chunks[current].emit_if(line);
    emit_dart_enum_to_string(chunks, current, line);
    chunks[current].emit_else(line);
    core_wasm::dup(&mut chunks[current], line);
    emit_dart_plain_map_like(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_dart_map_to_string(chunks, current, line);
    chunks[current].emit_else(line);
    emit_dart_value_to_string(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_dart_enum_like_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    let type_key = string_key(&mut chunks[current], reflection::FIELD_TYPE);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, type_key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    let name_key = string_key(&mut chunks[current], "name");
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, name_key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_AND, line);
}

fn emit_dart_enum_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    let type_key = string_key(&mut chunks[current], reflection::FIELD_TYPE);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, type_key, line);
    chunks[current].emit_string_const(".", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    let name_key = string_key(&mut chunks[current], "name");
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, name_key, line);
    strings::emit_concat(&mut chunks[current], 3, line);
}

fn emit_dart_plain_map_like(chunk: &mut Chunk, line: u32) {
    let value_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    reflection::emit_typeof_in_chunk(chunk, line);
    chunk.emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    let type_key = string_key(chunk, reflection::FIELD_TYPE);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, type_key, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_end(line);
}

fn emit_dart_map_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let map_slot = reserve_slot(&mut chunks[current]);
    let entries_slot = reserve_slot(&mut chunks[current]);
    let parts_slot = reserve_slot(&mut chunks[current]);
    let pair_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parts_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "entries", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries_slot, line);
    let state = loops::emit_for_in_start(chunks, current, entries_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pair_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parts_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    emit_dart_value_to_string(&mut chunks[current], line);
    chunks[current].emit_string_const(": ", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    emit_dart_value_to_string(&mut chunks[current], line);
    strings::emit_concat(&mut chunks[current], 3, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parts_slot, line);
    chunks[current].emit_string_const(", ", line);
    collections::emit_join(chunks, current, line);
    let joined_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, joined_slot, line);
    chunks[current].emit_string_const("{", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, joined_slot, line);
    chunks[current].emit_string_const("}", line);
    strings::emit_concat(&mut chunks[current], 3, line);
}

fn emit_dart_double_to_string_normal(
    chunks: &mut [Chunk],
    current: usize,
    value_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(&mut chunks[current], "ecma:number", "isInteger", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_i32_const(1, line);
    host::emit(&mut chunks[current], "ecma:number", "toFixed", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(&mut chunks[current], "ecma:number", "toString", 1, line);
    chunks[current].emit_end(line);
}

pub fn emit_dart_double_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    vybe_compiler::primitives::math::emit_signbit(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("-0.0", line);
    chunks[current].emit_else(line);
    emit_dart_double_to_string_normal(chunks, current, value_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    emit_dart_double_to_string_normal(chunks, current, value_slot, line);
    chunks[current].emit_end(line);
}

/// Dart `replaceFirst(pattern, replacement)` — first match only.
/// Routed at the profile level to `host:ecma:string.replace`
/// (ECMA-262 §22.1.3.18 replaces only the first occurrence).
/// This adapter is unused; kept as a stub so the dispatch arm
/// stays uniform with the other Dart helpers.
pub fn emit_dart_replace_first(chunks: &mut [Chunk], current: usize, line: u32) {
    let _ = (chunks, current, line);
}

/// Dart `list.first` — `list[0]`. Polymorphic so user classes whose
/// fields happen to be named `first`/`last`/`length` keep working
/// through plain STRUCT_GET when the receiver isn't a list.
pub fn emit_dart_list_first(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let receiver_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    core_wasm::dup(chunk, line);
    host::emit(chunk, "ecma:array", "isArray", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_is_generator(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    generators::emit_drain_into_array(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    let iter_slot = reserve_slot(&mut chunks[current]);
    emit_getter_or_field_to_slot(chunks, current, receiver_slot, "iterator", iter_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, iter_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_dart_materialize_iterable(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    let key = string_key(&mut chunks[current], "first");
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// Dart `list.last` — `list[length - 1]`. Polymorphic; non-list
/// receivers fall through to STRUCT_GET("last").
pub fn emit_dart_list_last(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let receiver_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    core_wasm::dup(chunk, line);
    host::emit(chunk, "ecma:array", "isArray", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    core_wasm::dup(&mut chunks[current], line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_is_generator(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    generators::emit_drain_into_array(chunks, current, line);
    core_wasm::dup(&mut chunks[current], line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    let iter_slot = reserve_slot(&mut chunks[current]);
    emit_getter_or_field_to_slot(chunks, current, receiver_slot, "iterator", iter_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, iter_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_dart_materialize_iterable(chunks, current, line);
    core_wasm::dup(&mut chunks[current], line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    let key = string_key(&mut chunks[current], "last");
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// Dart polymorphic `.length` — string | list | map. Stack: [coll] → [i32].
///
/// String/array routes through `compiler_common::{strings,collections}`
/// emitters; Map fall-through goes to `ecma:object.length` which returns
/// the property count (Dart `Map.length` semantics).
pub fn emit_dart_length(chunks: &mut [Chunk], current: usize, line: u32) {
    use vybe_runtime::Op as VOp;
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    // Dart's `String.length` is a count of UTF-16 CODE UNITS, so it goes
    // through `ecma:string.length` (`encode_utf16().count()`) rather than
    // `strings::emit_length`, which is `wasm:js-string.length` and counts
    // UTF-8 bytes — that made `'café'.length` 5 and every non-ASCII string
    // test wrong. PHP wants the byte count and keeps the shared helper.
    host::emit(&mut chunks[current], "ecma:string", "length", 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_is_generator(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    generators::emit_drain_into_array(chunks, current, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_sb_marker_test(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_sb_buffer_get(&mut chunks[current], line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_else(line);
    let length_slot = reserve_slot(&mut chunks[current]);
    emit_getter_or_field_to_slot(chunks, current, receiver_slot, "length", length_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, length_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, length_slot, line);
    chunks[current].emit_else(line);
    let iter_slot = reserve_slot(&mut chunks[current]);
    emit_getter_or_field_to_slot(chunks, current, receiver_slot, "iterator", iter_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, iter_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_dart_materialize_iterable(chunks, current, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    // Object/Map fall-through — count own enumerable properties via
    // `ecma:object.length`. Import tables are per chunk: register on
    // the chunk whose CALL_IMPORT indexes them.
    let idx = chunks[current].add_import("ecma:object", "length");
    chunks[current].emit_op_u16(VOp::CALL_IMPORT, idx, line);
    chunks[current].emit(1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// Dart `iterable.toList()` — materialize generators/custom iterables.
/// Stack: [iterable] -> [array].
pub fn emit_dart_iter_to_list(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_dart_materialize_iterable(chunks, current, line);
}

/// Dart `for-in` preparation. Keep generator continuations lazy so the common
/// compiler generator loop can honor `break`/manual take, but materialize
/// Dart custom `IterableBase` objects before the shared array loop sees them.
pub fn emit_dart_for_in_iterable(chunks: &mut [Chunk], current: usize, line: u32) {
    let slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    emit_is_generator(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    emit_dart_materialize_iterable(chunks, current, line);
    chunks[current].emit_end(line);
}

/// Dart `iterable.join([separator])`.
/// Stack: [iterable] or [iterable, sep] -> [string].
pub fn emit_dart_iter_join(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let sep_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    if argc > 1 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, sep_slot, line);
    } else {
        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, sep_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_dart_materialize_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep_slot, line);
    collections::emit_join(chunks, current, line);
}

/// Dart `Future.sync(fn)` / `Future.microtask(fn)`.
pub fn emit_dart_future_call0(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
    host::emit(&mut chunks[current], "ecma:promise", "resolve", 1, line);
}

/// Dart `Future.delayed(duration, fn)`; duration scheduling is outside this
/// test-level lowering, but the result is still a real resolved Promise.
pub fn emit_dart_future_delayed(chunks: &mut [Chunk], current: usize, line: u32) {
    let callback_slot = reserve_slot(&mut chunks[current]);
    let duration_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, callback_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, duration_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, callback_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
    host::emit(&mut chunks[current], "ecma:promise", "resolve", 1, line);
}

/// Dart polymorphic `contains`: String.contains keeps string semantics;
/// Iterable.contains materializes generators/custom iterables first.
pub fn emit_dart_contains(chunks: &mut [Chunk], current: usize, line: u32) {
    let needle_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, needle_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    core_wasm::dup(&mut chunks[current], line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle_slot, line);
    host::emit(&mut chunks[current], "ecma:string", "includes", 2, line);
    chunks[current].emit_else(line);
    emit_dart_materialize_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle_slot, line);
    collections::emit_contains(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_undefined_to_null(chunk: &mut Chunk, line: u32) {
    let slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    host::emit(chunk, "wasm:js-undefined", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_end(line);
}

fn emit_get_field_or_null_to_slot(
    chunk: &mut Chunk,
    obj_slot: u16,
    key: &str,
    dst_slot: u16,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let key = string_key(chunk, key);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    emit_undefined_to_null(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, dst_slot, line);
}

fn emit_call_ref_on_receiver(
    chunks: &mut [Chunk],
    current: usize,
    receiver_slot: u16,
    fn_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    vybe_compiler::primitives::globals::emit_write(&mut chunks[current], "__js_this", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
}

fn emit_getter_or_field_to_slot(
    chunks: &mut [Chunk],
    current: usize,
    obj_slot: u16,
    field: &str,
    dst_slot: u16,
    line: u32,
) {
    let getter_slot = reserve_slot(&mut chunks[current]);
    let getter_name = format!("__get_{}", field);
    emit_get_field_or_null_to_slot(
        &mut chunks[current],
        obj_slot,
        &getter_name,
        getter_slot,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, getter_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_call_ref_on_receiver(chunks, current, obj_slot, getter_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dst_slot, line);
    chunks[current].emit_else(line);
    emit_get_field_or_null_to_slot(&mut chunks[current], obj_slot, field, dst_slot, line);
    chunks[current].emit_end(line);
}

fn emit_set_string_field_from_slot(
    chunk: &mut Chunk,
    obj_slot: u16,
    key: &str,
    val_slot: u16,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, val_slot, line);
    let key = string_key(chunk, key);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    chunk.emit_op(Op::DROP, line);
}

fn emit_mark_set_top(chunk: &mut Chunk, line: u32) {
    core_wasm::dup(chunk, line);
    chunk.emit_bool_const(true, line);
    let key = string_key(chunk, SET_MARKER_KEY);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    chunk.emit_op(Op::DROP, line);
}

fn emit_slot_is_set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let key = string_key(chunk, SET_MARKER_KEY);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

fn emit_slot_is_array(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    host::emit(chunk, "ecma:array", "isArray", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

fn emit_pair_to_map_entry(chunks: &mut [Chunk], current: usize, pair_slot: u16, line: u32) {
    let key_slot = reserve_slot(&mut chunks[current]);
    let value_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    let chunk = &mut chunks[current];
    emit_set_string_field_from_slot(chunk, pair_slot, "key", key_slot, line);
    emit_set_string_field_from_slot(chunk, pair_slot, "value", value_slot, line);
}

pub fn emit_dart_identity(_chunks: &mut [Chunk], _current: usize, _line: u32) {}

pub fn emit_dart_map_new(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_struct_new(0, 0, line);
}

/// `SplayTreeMap` — a plain map tagged so its key/value/entry reads sort
/// ascending (see `emit_dart_map_keys`). Ordering falls out of the shared
/// natural sort at read time; storage/lookup stay identical to a plain map.
pub fn emit_dart_sorted_map_new(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_dart_map_new(chunks, current, line);
    let chunk = &mut chunks[current];
    core_wasm::dup(chunk, line);
    chunk.emit_bool_const(true, line);
    let key = string_key(chunk, SORTED_MAP_MARKER_KEY);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
    chunk.emit_op(Op::DROP, line);
}

pub fn emit_dart_set_new(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_array_new(chunks, current, 0, line);
    emit_mark_set_top(&mut chunks[current], line);
}

pub fn emit_dart_set_from(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_dart_materialize_iterable(chunks, current, line);
    let arr_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    emit_dart_iter_distinct(chunks, current, 0, line);
    emit_mark_set_top(&mut chunks[current], line);
}

fn emit_freeze_top(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:object", "freeze", 1, line);
}

fn emit_throw_if_frozen(chunks: &mut [Chunk], current: usize, receiver_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "isFrozen", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("Cannot modify an unmodifiable collection", line);
    errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

pub fn emit_dart_list_unmodifiable(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_dart_list_from(chunks, current, 1, line);
    emit_freeze_top(chunks, current, line);
}

pub fn emit_dart_map_unmodifiable(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_dart_map_from(chunks, current, line);
    emit_freeze_top(chunks, current, line);
}

pub fn emit_dart_map_unmodifiable_entries(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_dart_map_from_entries(chunks, current, line);
    emit_freeze_top(chunks, current, line);
}

pub fn emit_dart_set_unmodifiable(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_dart_set_from(chunks, current, line);
    emit_freeze_top(chunks, current, line);
}

pub fn emit_dart_map_entry(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    let key_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    let pair_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pair_slot, line);
    let chunk = &mut chunks[current];
    emit_set_string_field_from_slot(chunk, pair_slot, "key", key_slot, line);
    emit_set_string_field_from_slot(chunk, pair_slot, "value", value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, pair_slot, line);
}

pub fn emit_dart_map_keys(chunks: &mut [Chunk], current: usize, line: u32) {
    let receiver_slot = reserve_slot(&mut chunks[current]);
    let order_slot = reserve_slot(&mut chunks[current]);
    let keys_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    let order_key = string_key(&mut chunks[current], MAP_ORDER_KEY);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, order_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, order_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, order_slot, line);
    host::emit(&mut chunks[current], "wasm:js-undefined", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "keys", 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, order_slot, line);
    chunks[current].emit_end(line);
    // SplayTreeMap: keys enumerate in ascending order. When the map carries the
    // sorted marker, sort the key array once here — every values/entries/forEach
    // read funnels through this, so ordering follows for all of them.
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    let sorted_key = string_key(&mut chunks[current], SORTED_MAP_MARKER_KEY);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, sorted_key, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    // Map keys are stored as JS object-property strings, so a `Map<int,_>`
    // would sort "10" before "2" lexicographically. When the keys are numeric
    // (homogeneous key type, tested on the first key), coerce the whole array
    // to numbers first so the ascending order — and the enumerated keys — are
    // numeric. Non-numeric keys keep the default (lexicographic) order.
    let len_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    // numeric = !isNaN(Number(keys[0]))
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    host::emit(&mut chunks[current], "ecma:number", "parseFloat", 1, line);
    host::emit(&mut chunks[current], "ecma:number", "isNaN", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    // coerce every key to a number in place
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    host::emit(&mut chunks[current], "ecma:number", "parseFloat", 1, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    collections::emit_sort(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
}

pub fn emit_dart_map_values(chunks: &mut [Chunk], current: usize, line: u32) {
    let receiver_slot = reserve_slot(&mut chunks[current]);
    let keys_slot = reserve_slot(&mut chunks[current]);
    let values_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let key_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_dart_map_keys(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, values_slot, line);
    let state = loops::emit_for_in_start(chunks, current, keys_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, values_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, values_slot, line);
}

pub fn emit_dart_map_entries(chunks: &mut [Chunk], current: usize, line: u32) {
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_dart_map_keys(chunks, current, line);
    let entries_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let key_slot = reserve_slot(&mut chunks[current]);
    let value_slot = reserve_slot(&mut chunks[current]);
    let pair_slot = reserve_slot(&mut chunks[current]);
    let keys_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries_slot, line);
    let state = loops::emit_for_in_start(chunks, current, keys_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pair_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, entries_slot, line);
    emit_pair_to_map_entry(chunks, current, pair_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, entries_slot, line);
}

pub fn emit_dart_map_from(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:object", "entries", 1, line);
    emit_dart_map_from_entries(chunks, current, line);
}

pub fn emit_dart_map_from_entries(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_dart_materialize_iterable(chunks, current, line);
    let entries_slot = reserve_slot(&mut chunks[current]);
    let map_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let pair_slot = reserve_slot(&mut chunks[current]);
    let key_slot = reserve_slot(&mut chunks[current]);
    let value_slot = reserve_slot(&mut chunks[current]);
    let order_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries_slot, line);
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, order_slot, line);
    let state = loops::emit_for_in_start(chunks, current, entries_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pair_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, order_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, order_slot, line);
    let order_key = string_key(&mut chunks[current], MAP_ORDER_KEY);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, order_key, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
}

pub fn emit_dart_map_from_iterables(chunks: &mut [Chunk], current: usize, line: u32) {
    let values_slot = reserve_slot(&mut chunks[current]);
    let keys_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, values_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    let keys_arr = materialize_slot(chunks, current, keys_slot, line);
    let values_arr = materialize_slot(chunks, current, values_slot, line);
    let map_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map_slot, line);
    let state = loops::emit_for_in_start(chunks, current, keys_arr, idx_slot, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, values_arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
}

pub fn emit_dart_map_contains_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "values", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_contains(chunks, current, line);
}

pub fn emit_dart_index_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    let key_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    emit_throw_if_frozen(chunks, current, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
}

pub fn emit_dart_add_general(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 2 {
        let value_slot = reserve_slot(&mut chunks[current]);
        let key_slot = reserve_slot(&mut chunks[current]);
        let receiver_slot = reserve_slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
        emit_throw_if_frozen(chunks, current, receiver_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
        host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
        emit_undefined_to_null(&mut chunks[current], line);
        let prev_slot = reserve_slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, prev_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, prev_slot, line);
        return;
    }
    let value_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    emit_throw_if_frozen(chunks, current, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    let type_key = string_key(&mut chunks[current], reflection::FIELD_TYPE);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, type_key, line);
    chunks[current].emit_string_const("DateTime", line);
    chunks[current].emit_op(Op::EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    crate::emitter::core_adapter::emit_dart_add(chunks, current, line);
    chunks[current].emit_else(line);
    emit_slot_is_set(&mut chunks[current], receiver_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_contains(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    let cached_hash_slot = reserve_slot(&mut chunks[current]);
    emit_get_field_or_null_to_slot(
        &mut chunks[current],
        receiver_slot,
        "__dart_identity_hash",
        cached_hash_slot,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_restore_cached_hash_if_any(chunks, current, receiver_slot, cached_hash_slot, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_restore_cached_hash_if_any(
    chunks: &mut [Chunk],
    current: usize,
    receiver_slot: u16,
    cached_hash_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, cached_hash_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cached_hash_slot, line);
    let field_key = string_key(&mut chunks[current], "__dart_identity_hash");
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, field_key, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

fn emit_array_is_array(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
}

pub fn emit_dart_add_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let other_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, other_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    emit_throw_if_frozen(chunks, current, receiver_slot, line);
    emit_array_is_array(chunks, current, receiver_slot, line);
    chunks[current].emit_if(line);
    emit_slot_is_set(&mut chunks[current], receiver_slot, line);
    chunks[current].emit_if(line);
    let other_arr = materialize_slot(chunks, current, other_slot, line);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    let state = loops::emit_for_in_start(chunks, current, other_arr, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_contains(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    let other_arr = materialize_slot(chunks, current, other_slot, line);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    let state = loops::emit_for_in_start(chunks, current, other_arr, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, other_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "entries", 1, line);
    let entries_slot = reserve_slot(&mut chunks[current]);
    let pair_slot = reserve_slot(&mut chunks[current]);
    let key_slot = reserve_slot(&mut chunks[current]);
    let val_slot = reserve_slot(&mut chunks[current]);
    let idx2_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries_slot, line);
    let state2 = loops::emit_for_in_start(chunks, current, entries_slot, idx2_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pair_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, val_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, val_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx2_slot, state2, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

pub fn emit_dart_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let key_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    emit_throw_if_frozen(chunks, current, receiver_slot, line);
    emit_array_is_array(chunks, current, receiver_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    collections::emit_index_of(chunks, current, line);
    let idx_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    emit_undefined_to_null(&mut chunks[current], line);
    let old_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, old_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, old_slot, line);
    chunks[current].emit_end(line);
}

pub fn emit_dart_lookup(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_contains(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

pub fn emit_dart_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    emit_throw_if_frozen(chunks, current, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_sb_marker_test(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_dart_sb_clear(chunks, current, line);
    chunks[current].emit_else(line);
    emit_array_is_array(chunks, current, receiver_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    collections::emit_clear(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "keys", 1, line);
    let keys_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let key_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    let state = loops::emit_for_in_start(chunks, current, keys_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_dart_map_update(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let absent_slot = reserve_slot(&mut chunks[current]);
    let fn_slot = reserve_slot(&mut chunks[current]);
    let key_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    if argc > 3 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, absent_slot, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, absent_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    emit_throw_if_frozen(chunks, current, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "has", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, absent_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
    chunks[current].emit_end(line);
    let new_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, new_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, new_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, new_slot, line);
}

pub fn emit_dart_map_update_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let fn_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    emit_throw_if_frozen(chunks, current, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "entries", 1, line);
    let entries_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let pair_slot = reserve_slot(&mut chunks[current]);
    let key_slot = reserve_slot(&mut chunks[current]);
    let value_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries_slot, line);
    let state = loops::emit_for_in_start(chunks, current, entries_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pair_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_dart_for_each_general(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, fn_slot) = stash_iterable_and_arg(chunks, current, line);
    emit_array_is_array(chunks, current, receiver_slot, line);
    chunks[current].emit_if(line);
    let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
    let idx_slot = reserve_slot(&mut chunks[current]);
    loops::emit_foreach(chunks, current, fn_slot, arr_slot, idx_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_dart_map_keys(chunks, current, line);
    let keys_slot = reserve_slot(&mut chunks[current]);
    let idx2_slot = reserve_slot(&mut chunks[current]);
    let key_slot = reserve_slot(&mut chunks[current]);
    let value_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    let state = loops::emit_for_in_start(chunks, current, keys_slot, idx2_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx2_slot, state, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

pub fn emit_dart_map_general(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, fn_slot) = stash_iterable_and_arg(chunks, current, line);
    emit_array_is_array(chunks, current, receiver_slot, line);
    chunks[current].emit_if(line);
    let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
    let result_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    loops::emit_map(
        chunks,
        current,
        fn_slot,
        arr_slot,
        result_slot,
        idx_slot,
        line,
    );
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_is_generator(&mut chunks[current], line);
    chunks[current].emit_if(line);
    let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
    let result_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    loops::emit_map(
        chunks,
        current,
        fn_slot,
        arr_slot,
        result_slot,
        idx_slot,
        line,
    );
    chunks[current].emit_else(line);
    let iter_slot = reserve_slot(&mut chunks[current]);
    emit_getter_or_field_to_slot(chunks, current, receiver_slot, "iterator", iter_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, iter_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
    let result_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    loops::emit_map(
        chunks,
        current,
        fn_slot,
        arr_slot,
        result_slot,
        idx_slot,
        line,
    );
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "entries", 1, line);
    let entries_slot = reserve_slot(&mut chunks[current]);
    let pairs_slot = reserve_slot(&mut chunks[current]);
    let idx2_slot = reserve_slot(&mut chunks[current]);
    let pair_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pairs_slot, line);
    let state = loops::emit_for_in_start(chunks, current, entries_slot, idx2_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pair_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pairs_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx2_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pairs_slot, line);
    host::emit(&mut chunks[current], "ecma:map", "fromEntries", 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_dart_remove_where(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, fn_slot) = stash_iterable_and_arg(chunks, current, line);
    emit_throw_if_frozen(chunks, current, receiver_slot, line);
    emit_array_is_array(chunks, current, receiver_slot, line);
    chunks[current].emit_if(line);
    emit_dart_set_retain_or_remove_where(chunks, current, receiver_slot, fn_slot, false, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "entries", 1, line);
    let entries_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let pair_slot = reserve_slot(&mut chunks[current]);
    let key_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries_slot, line);
    let state = loops::emit_for_in_start(chunks, current, entries_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pair_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pair_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

fn emit_set_push_unique(
    chunks: &mut [Chunk],
    current: usize,
    set_slot: u16,
    value_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_contains(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

fn emit_replace_array_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    receiver_slot: u16,
    values_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    collections::emit_clear(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    let state = loops::emit_for_in_start(chunks, current, values_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

fn emit_set_filter_against_slot(
    chunks: &mut [Chunk],
    current: usize,
    receiver_slot: u16,
    other_slot: u16,
    keep_present: bool,
    mutate_receiver: bool,
    line: u32,
) {
    let other_arr = materialize_slot(chunks, current, other_slot, line);
    let result_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    collections::emit_array_new(chunks, current, 0, line);
    emit_mark_set_top(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    let state = loops::emit_for_in_start(chunks, current, receiver_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, other_arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_contains(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    if !keep_present {
        chunks[current].emit_op(Op::I32_EQZ, line);
    }
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    if mutate_receiver {
        emit_replace_array_from_slot(chunks, current, receiver_slot, result_slot, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    }
}

fn emit_dart_set_retain_or_remove_where(
    chunks: &mut [Chunk],
    current: usize,
    receiver_slot: u16,
    fn_slot: u16,
    retain_matches: bool,
    line: u32,
) {
    let result_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    collections::emit_array_new(chunks, current, 0, line);
    emit_mark_set_top(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    let state = loops::emit_for_in_start(chunks, current, receiver_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    if !retain_matches {
        chunks[current].emit_op(Op::I32_EQZ, line);
    }
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    emit_replace_array_from_slot(chunks, current, receiver_slot, result_slot, line);
}

pub fn emit_dart_set_remove_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, other_slot) = stash_iterable_and_arg(chunks, current, line);
    emit_set_filter_against_slot(
        chunks,
        current,
        receiver_slot,
        other_slot,
        false,
        true,
        line,
    );
}

pub fn emit_dart_set_retain_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, other_slot) = stash_iterable_and_arg(chunks, current, line);
    emit_set_filter_against_slot(chunks, current, receiver_slot, other_slot, true, true, line);
}

pub fn emit_dart_set_retain_where(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, fn_slot) = stash_iterable_and_arg(chunks, current, line);
    emit_throw_if_frozen(chunks, current, receiver_slot, line);
    emit_dart_set_retain_or_remove_where(chunks, current, receiver_slot, fn_slot, true, line);
}

pub fn emit_dart_set_union(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, other_slot) = stash_iterable_and_arg(chunks, current, line);
    let other_arr = materialize_slot(chunks, current, other_slot, line);
    let result_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    collections::emit_array_new(chunks, current, 0, line);
    emit_mark_set_top(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    let state = loops::emit_for_in_start(chunks, current, receiver_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    emit_set_push_unique(chunks, current, result_slot, elem_slot, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    let idx2_slot = reserve_slot(&mut chunks[current]);
    let state2 = loops::emit_for_in_start(chunks, current, other_arr, idx2_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    emit_set_push_unique(chunks, current, result_slot, elem_slot, line);
    loops::emit_for_in_end(chunks, current, idx2_slot, state2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_dart_set_intersection(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, other_slot) = stash_iterable_and_arg(chunks, current, line);
    emit_set_filter_against_slot(
        chunks,
        current,
        receiver_slot,
        other_slot,
        true,
        false,
        line,
    );
}

pub fn emit_dart_difference(chunks: &mut [Chunk], current: usize, line: u32) {
    let other_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, other_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    let type_key = string_key(&mut chunks[current], reflection::FIELD_TYPE);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, type_key, line);
    chunks[current].emit_string_const("DateTime", line);
    chunks[current].emit_op(Op::EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, other_slot, line);
    crate::emitter::core_adapter::emit_datetime_difference(chunks, current, line);
    chunks[current].emit_else(line);
    emit_set_filter_against_slot(
        chunks,
        current,
        receiver_slot,
        other_slot,
        false,
        false,
        line,
    );
    chunks[current].emit_end(line);
}

pub fn emit_dart_set_contains_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, other_slot) = stash_iterable_and_arg(chunks, current, line);
    let other_arr = materialize_slot(chunks, current, other_slot, line);
    let result_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    let state = loops::emit_for_in_start(chunks, current, other_arr, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_contains(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `x == null` / `null == x` fast path: a single reference null-test instead of
/// the full `emit_dart_eq` deep-equality routine. Stack: `[value]` → `[bool]`.
/// Comparing against the `null` literal never needs value/collection equality,
/// so routing it here keeps every null guard a couple of instructions rather
/// than inlining the ~thousand-instruction equality cascade at each site.
pub fn emit_dart_is_null(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    // Normalize the raw i32 test to a Dart bool so the result is usable as a
    // value (`!`, interpolation, storage), not only as an `if` condition.
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

/// Structural equality for two objects: same number of own keys, and every
/// key of the left equal (by `==`) on the right. Stack: `[] → [bool]`.
///
/// Deliberately key-WISE rather than comparing `JSON.stringify` of each side:
/// two objects built from the same constructor do not necessarily serialise
/// their keys in the same order, so the stringify form reported identical
/// values as unequal — and did so intermittently, since the order depends on
/// the object rather than the program.
fn emit_dart_fields_equal(
    chunks: &mut [Chunk],
    current: usize,
    left_slot: u16,
    right_slot: u16,
    line: u32,
) {
    let left_keys = reserve_slot(&mut chunks[current]);
    let right_keys = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let key_slot = reserve_slot(&mut chunks[current]);
    let result_slot = reserve_slot(&mut chunks[current]);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "keys", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "keys", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_keys, line);

    // Start out equal, then let any differing key falsify it.
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    // A differing key count is already a mismatch.
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_keys, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_keys, line);
    collections::emit_len(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    let state = loops::emit_for_in_start(chunks, current, left_keys, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    // `left[key] == right[key]`. Compared with the PRIMITIVE equality, not a
    // nested `emit_dart_eq`: this emitter inlines its body, so calling the
    // full comparison here would expand forever at compile time. A value
    // type's fields are scalars (numbers, strings, bools, enum spellings),
    // which is exactly what the primitive form handles.
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_dart_eq(chunks: &mut [Chunk], current: usize, line: u32) {
    let right_slot = reserve_slot(&mut chunks[current]);
    let left_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_slot, line);
    emit_slot_is_set(&mut chunks[current], left_slot, line);
    emit_slot_is_set(&mut chunks[current], right_slot, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    collections::emit_len(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    let same_len_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, same_len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, same_len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    emit_dart_set_contains_all(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    emit_slot_is_array(&mut chunks[current], left_slot, line);
    emit_slot_is_array(&mut chunks[current], right_slot, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    collections::emit_sequence_equal(chunks, current, line);
    chunks[current].emit_else(line);
    // Structural (JSON) equality when BOTH sides are plain map-like objects OR
    // both are stamped value-equality types (Flutter `ValueKey`/`Color`/… whose
    // `operator ==` is by value). Otherwise fall through to identity.
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    emit_dart_plain_map_like(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    emit_dart_plain_map_like(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_AND, line);
    let veq_key = string_key(&mut chunks[current], "__value_eq");
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, veq_key, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, veq_key, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    emit_dart_fields_equal(chunks, current, left_slot, right_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_dart_identical(chunks: &mut [Chunk], current: usize, line: u32) {
    let right_slot = reserve_slot(&mut chunks[current]);
    let left_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("number", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("boolean", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunks[current].emit_op(Op::REF_EQ, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_dart_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    let getter_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    emit_slot_is_object_or_function(&mut chunks[current], value_slot, line);
    chunks[current].emit_if(line);
    emit_get_field_or_null_to_slot(
        &mut chunks[current],
        value_slot,
        "__get_hash",
        getter_slot,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, getter_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_call_ref_on_receiver(chunks, current, value_slot, getter_slot, line);
    chunks[current].emit_else(line);
    emit_slot_is_array(&mut chunks[current], value_slot, line);
    chunks[current].emit_if(line);
    emit_dart_cached_array_hash(chunks, current, value_slot, line);
    chunks[current].emit_else(line);
    emit_dart_identity_hash(chunks, current, value_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    vybe_compiler::primitives::object::emit_hash_code(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

pub fn emit_dart_object_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    let getter_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    emit_get_field_or_null_to_slot(
        &mut chunks[current],
        value_slot,
        "__get_hash",
        getter_slot,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, getter_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_call_ref_on_receiver(chunks, current, value_slot, getter_slot, line);
    chunks[current].emit_else(line);
    emit_slot_is_array(&mut chunks[current], value_slot, line);
    chunks[current].emit_if(line);
    emit_dart_cached_array_hash(chunks, current, value_slot, line);
    chunks[current].emit_else(line);
    emit_dart_identity_hash(chunks, current, value_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_dart_cached_array_hash(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    let cached_slot = reserve_slot(&mut chunks[current]);
    let result_slot = reserve_slot(&mut chunks[current]);
    emit_get_field_or_null_to_slot(
        &mut chunks[current],
        value_slot,
        "__dart_identity_hash",
        cached_slot,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, cached_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cached_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    vybe_compiler::primitives::object::emit_hash_code(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    let field_key = string_key(&mut chunks[current], "__dart_identity_hash");
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, field_key, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_end(line);
}

fn emit_slot_is_object_or_function(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_string_const("function", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn emit_dart_identity_hash(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    let existing_slot = reserve_slot(&mut chunks[current]);
    let next_slot = reserve_slot(&mut chunks[current]);
    emit_get_field_or_null_to_slot(
        &mut chunks[current],
        value_slot,
        "__dart_identity_hash",
        existing_slot,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, existing_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, existing_slot, line);
    chunks[current].emit_else(line);

    vybe_compiler::primitives::globals::emit_read(&mut chunks[current], "__dart_identity_hash_next", line);
    emit_undefined_to_null(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, next_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, next_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, next_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_TEE, next_slot, line);
    vybe_compiler::primitives::globals::emit_write(&mut chunks[current], "__dart_identity_hash_next", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, next_slot, line);
    let field_key = string_key(&mut chunks[current], "__dart_identity_hash");
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, field_key, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, next_slot, line);
    chunks[current].emit_end(line);
}

pub fn emit_dart_map_put_if_absent(chunks: &mut [Chunk], current: usize, line: u32) {
    let fn_slot = reserve_slot(&mut chunks[current]);
    let key_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    emit_throw_if_frozen(chunks, current, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "has", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    emit_undefined_to_null(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
    let value_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_end(line);
}

pub fn emit_dart_list_from(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 1 {
        chunks[current].emit_op(Op::DROP, line);
    }
    emit_dart_materialize_iterable(chunks, current, line);
}

pub fn emit_dart_string_from_char_codes(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = reserve_slot(&mut chunks[current]);
    emit_dart_materialize_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "fromCharCodeArray",
        3,
        line,
    );
}

pub fn emit_dart_string_code_units(chunks: &mut [Chunk], current: usize, line: u32) {
    let str_slot = reserve_slot(&mut chunks[current]);
    let out_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let len_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, str_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_slot, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    // The length is a DYNAMIC (f64-boxed) value, not a raw i32, so `I32_GE_S`
    // compared an i32 counter against a boxed number and the guard fired on
    // the first iteration — `s.runes` and `s.codeUnits` returned an empty
    // array for every non-literal receiver. Compare dynamically, the way
    // `emit_dart_list_generate` below already does.
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    host::emit(&mut chunks[current], "ecma:string", "charCodeAt", 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

pub fn emit_dart_string_runes(chunks: &mut [Chunk], current: usize, line: u32) {
    let chars_slot = reserve_slot(&mut chunks[current]);
    let out_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let len_slot = reserve_slot(&mut chunks[current]);
    host::emit(&mut chunks[current], "ecma:array", "from", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, chars_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, chars_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    // The length is a DYNAMIC (f64-boxed) value, not a raw i32, so `I32_GE_S`
    // compared an i32 counter against a boxed number and the guard fired on
    // the first iteration — `s.runes` and `s.codeUnits` returned an empty
    // array for every non-literal receiver. Compare dynamically, the way
    // `emit_dart_list_generate` below already does.
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, chars_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    host::emit(&mut chunks[current], "ecma:string", "codePointAt", 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

pub fn emit_dart_list_generate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let fn_slot = reserve_slot(&mut chunks[current]);
    let len_slot = reserve_slot(&mut chunks[current]);
    let result_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    if argc > 2 {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    let state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_dart_index_of(chunks: &mut [Chunk], current: usize, argc: u8, last: bool, line: u32) {
    let from_slot = reserve_slot(&mut chunks[current]);
    let needle_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    if argc > 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, from_slot, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, from_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, needle_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle_slot, line);
    if argc > 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, from_slot, line);
        host::emit(
            &mut chunks[current],
            "ecma:string",
            if last { "lastIndexOf" } else { "indexOf" },
            3,
            line,
        );
    } else {
        host::emit(
            &mut chunks[current],
            "ecma:string",
            if last { "lastIndexOf" } else { "indexOf" },
            2,
            line,
        );
    }
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle_slot, line);
    if argc > 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, from_slot, line);
        if last {
            collections::emit_last_index_of_from(chunks, current, line);
        } else {
            collections::emit_index_of_from(chunks, current, line);
        }
    } else if last {
        collections::emit_last_index_of(chunks, current, line);
    } else {
        collections::emit_index_of(chunks, current, line);
    }
    chunks[current].emit_end(line);
}

pub fn emit_dart_list_insert(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    let index_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    emit_throw_if_frozen(chunks, current, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_insert(chunks, current, line);
}

pub fn emit_dart_list_remove_at(chunks: &mut [Chunk], current: usize, line: u32) {
    let index_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    emit_throw_if_frozen(chunks, current, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_remove_at(chunks, current, line);
}

pub fn emit_dart_list_remove_last(chunks: &mut [Chunk], current: usize, line: u32) {
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    emit_throw_if_frozen(chunks, current, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    collections::emit_pop(chunks, current, line);
}

pub fn emit_dart_list_remove_range(chunks: &mut [Chunk], current: usize, line: u32) {
    let end_slot = reserve_slot(&mut chunks[current]);
    let start_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_remove_range(chunks, current, line);
}

pub fn emit_dart_list_get_range(chunks: &mut [Chunk], current: usize, line: u32) {
    let end_slot = reserve_slot(&mut chunks[current]);
    let start_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_slot, line);
    collections::emit_slice(chunks, current, line);
}

pub fn emit_dart_list_fill_range(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    let end_slot = reserve_slot(&mut chunks[current]);
    let start_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_slot, line);
    collections::emit_fill(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_dart_list_set_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let source_slot = reserve_slot(&mut chunks[current]);
    let start_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    let source_arr = materialize_slot(chunks, current, source_slot, line);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    let state = loops::emit_for_in_start(chunks, current, source_arr, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_dart_list_set_range(chunks: &mut [Chunk], current: usize, line: u32) {
    let source_slot = reserve_slot(&mut chunks[current]);
    let end_slot = reserve_slot(&mut chunks[current]);
    let start_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    emit_throw_if_frozen(chunks, current, receiver_slot, line);
    let source_arr = materialize_slot(chunks, current, source_slot, line);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    let state = loops::emit_for_in_start(chunks, current, source_arr, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_dart_list_as_map(chunks: &mut [Chunk], current: usize, line: u32) {
    let receiver_slot = reserve_slot(&mut chunks[current]);
    let map_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map_slot, line);
    let state = loops::emit_for_in_start(chunks, current, receiver_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
}

pub fn emit_dart_list_reversed(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_reversed(chunks, current, line);
}

fn emit_dart_list_sort_by_compare_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let list_slot = reserve_slot(&mut chunks[current]);
    let len_slot = reserve_slot(&mut chunks[current]);
    let i_slot = reserve_slot(&mut chunks[current]);
    let j_slot = reserve_slot(&mut chunks[current]);
    let a_slot = reserve_slot(&mut chunks[current]);
    let b_slot = reserve_slot(&mut chunks[current]);

    chunks[current].emit_op_u16(Op::LOCAL_SET, list_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let outer = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    loops::emit_loop_cond(chunks, current, line);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j_slot, line);
    let inner = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, j_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j_slot, line);
    loops::emit_loop_end(chunks, current, inner, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    loops::emit_loop_end(chunks, current, outer, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
}

pub fn emit_dart_list_sort(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 1 {
        let cmp_slot = reserve_slot(&mut chunks[current]);
        let list_slot = reserve_slot(&mut chunks[current]);
        let len_slot = reserve_slot(&mut chunks[current]);
        let i_slot = reserve_slot(&mut chunks[current]);
        let j_slot = reserve_slot(&mut chunks[current]);
        let a_slot = reserve_slot(&mut chunks[current]);
        let b_slot = reserve_slot(&mut chunks[current]);

        chunks[current].emit_op_u16(Op::LOCAL_SET, cmp_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, list_slot, line);
        emit_throw_if_frozen(chunks, current, list_slot, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
        collections::emit_len(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
        let outer = loops::emit_loop_start(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
        loops::emit_loop_cond(chunks, current, line);

        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, j_slot, line);
        let inner = loops::emit_loop_start(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, j_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_SUB, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
        loops::emit_loop_cond(chunks, current, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, j_slot, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, a_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, j_slot, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, b_slot, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, cmp_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
        chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, j_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
        collections::emit_set(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, j_slot, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
        collections::emit_set(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, j_slot, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, j_slot, line);
        loops::emit_loop_end(chunks, current, inner, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
        loops::emit_loop_end(chunks, current, outer, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
    } else {
        let list_slot = reserve_slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, list_slot, line);
        emit_throw_if_frozen(chunks, current, list_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
        emit_dart_list_sort_by_compare_to(chunks, current, line);
    }
}

pub fn emit_dart_list_single(
    chunks: &mut [Chunk],
    current: usize,
    null_when_not_single: bool,
    line: u32,
) {
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    if null_when_not_single {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    } else {
        core_wasm::undefined(&mut chunks[current], line);
    }
    chunks[current].emit_end(line);
}

pub fn emit_dart_list_where_search(chunks: &mut [Chunk], current: usize, mode: u8, line: u32) {
    let (receiver_slot, fn_slot) = stash_iterable_and_arg(chunks, current, line);
    let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
    let result_slot = reserve_slot(&mut chunks[current]);
    if mode <= 1 {
        core_wasm::i32_const(&mut chunks[current], line, -1);
    } else {
        core_wasm::undefined(&mut chunks[current], line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    let len_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    if mode == 1 || mode == 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_SUB, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    let state = loops::emit_loop_start(chunks, current, line);
    if mode == 1 || mode == 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    }
    loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    if mode <= 1 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
        collections::emit_get(chunks, current, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    if mode == 1 || mode == 3 {
        chunks[current].emit_op(Op::I32_SUB, line);
    } else {
        chunks[current].emit_op(Op::I32_ADD, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

fn stash_iterable_and_arg(chunks: &mut [Chunk], current: usize, line: u32) -> (u16, u16) {
    let arg_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arg_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    (receiver_slot, arg_slot)
}

fn emit_dart_drain_iterator_to_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let iter_slot = reserve_slot(&mut chunks[current]);
    let result_slot = reserve_slot(&mut chunks[current]);
    let move_next_slot = reserve_slot(&mut chunks[current]);
    let current_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, iter_slot, line);
    emit_get_field_or_null_to_slot(
        &mut chunks[current],
        iter_slot,
        "moveNext",
        move_next_slot,
        line,
    );
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    let state = loops::emit_loop_start(chunks, current, line);
    emit_call_ref_on_receiver(chunks, current, iter_slot, move_next_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    loops::emit_loop_cond(chunks, current, line);
    emit_getter_or_field_to_slot(chunks, current, iter_slot, "current", current_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, current_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

fn emit_dart_drain_iterator_to_array_precurrent(chunks: &mut [Chunk], current: usize, line: u32) {
    let iter_slot = reserve_slot(&mut chunks[current]);
    let result_slot = reserve_slot(&mut chunks[current]);
    let move_next_slot = reserve_slot(&mut chunks[current]);
    let current_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, iter_slot, line);
    emit_get_field_or_null_to_slot(
        &mut chunks[current],
        iter_slot,
        "moveNext",
        move_next_slot,
        line,
    );
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    let state = loops::emit_loop_start(chunks, current, line);
    emit_getter_or_field_to_slot(chunks, current, iter_slot, "current", current_slot, line);
    emit_call_ref_on_receiver(chunks, current, iter_slot, move_next_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, current_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

fn emit_dart_materialize_iterable(chunks: &mut [Chunk], current: usize, line: u32) {
    let slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    emit_is_generator(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    generators::emit_drain_into_array(chunks, current, line);
    chunks[current].emit_else(line);
    let move_next_slot = reserve_slot(&mut chunks[current]);
    emit_get_field_or_null_to_slot(&mut chunks[current], slot, "moveNext", move_next_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, move_next_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    emit_dart_drain_iterator_to_array(chunks, current, line);
    chunks[current].emit_else(line);
    let iterator_slot = reserve_slot(&mut chunks[current]);
    emit_getter_or_field_to_slot(chunks, current, slot, "iterator", iterator_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, iterator_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    let iterator_move_next_slot = reserve_slot(&mut chunks[current]);
    emit_get_field_or_null_to_slot(
        &mut chunks[current],
        iterator_slot,
        "moveNext",
        iterator_move_next_slot,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, iterator_move_next_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, iterator_slot, line);
    emit_dart_drain_iterator_to_array(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_dart_materialize_iterable_precurrent(chunks: &mut [Chunk], current: usize, line: u32) {
    let slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    emit_is_generator(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    generators::emit_drain_into_array(chunks, current, line);
    chunks[current].emit_else(line);
    let move_next_slot = reserve_slot(&mut chunks[current]);
    emit_get_field_or_null_to_slot(&mut chunks[current], slot, "moveNext", move_next_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, move_next_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    emit_dart_drain_iterator_to_array_precurrent(chunks, current, line);
    chunks[current].emit_else(line);
    let iterator_slot = reserve_slot(&mut chunks[current]);
    emit_getter_or_field_to_slot(chunks, current, slot, "iterator", iterator_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, iterator_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    let iterator_move_next_slot = reserve_slot(&mut chunks[current]);
    emit_get_field_or_null_to_slot(
        &mut chunks[current],
        iterator_slot,
        "moveNext",
        iterator_move_next_slot,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, iterator_move_next_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, iterator_slot, line);
    emit_dart_drain_iterator_to_array_precurrent(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn materialize_slot(chunks: &mut [Chunk], current: usize, receiver_slot: u16, line: u32) -> u16 {
    let arr_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_dart_materialize_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    arr_slot
}

/// Dart `iterable.map(fn)` over a generator-aware materialized iterable.
pub fn emit_dart_iter_map(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, fn_slot) = stash_iterable_and_arg(chunks, current, line);
    let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
    let result_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    loops::emit_map(
        chunks,
        current,
        fn_slot,
        arr_slot,
        result_slot,
        idx_slot,
        line,
    );
}

/// Dart `stream.asyncMap(fn)`: map with promise assimilation using the same
/// JSPI await primitive used by shared promise-chain emitters.
pub fn emit_dart_iter_async_map(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, fn_slot) = stash_iterable_and_arg(chunks, current, line);
    let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
    let result_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    functions::emit_await(&mut chunks[current], line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Dart `iterable.where(fn)`.
pub fn emit_dart_iter_where(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, fn_slot) = stash_iterable_and_arg(chunks, current, line);
    let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
    let result_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    loops::emit_filter(
        chunks,
        current,
        fn_slot,
        arr_slot,
        result_slot,
        idx_slot,
        elem_slot,
        line,
    );
}

/// Dart `iterable.any(fn)`.
pub fn emit_dart_iter_any(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, fn_slot) = stash_iterable_and_arg(chunks, current, line);
    let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let _result_slot = reserve_slot(&mut chunks[current]);
    loops::emit_any_every(chunks, current, fn_slot, arr_slot, idx_slot, true, line);
}

/// Dart `iterable.every(fn)`.
pub fn emit_dart_iter_every(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, fn_slot) = stash_iterable_and_arg(chunks, current, line);
    let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let _result_slot = reserve_slot(&mut chunks[current]);
    loops::emit_any_every(chunks, current, fn_slot, arr_slot, idx_slot, false, line);
}

/// Dart `iterable.reduce(fn)` plus walker-normalized
/// `iterable.fold(initial, fn)` as `reduce(fn, initial)`.
pub fn emit_dart_iter_reduce(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 2 {
        let initial_slot = reserve_slot(&mut chunks[current]);
        let fn_slot = reserve_slot(&mut chunks[current]);
        let receiver_slot = reserve_slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, initial_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
        let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
        let idx_slot = reserve_slot(&mut chunks[current]);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
        let state = loops::emit_loop_start(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
        collections::emit_len(chunks, current, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
        loops::emit_loop_cond(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, initial_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, initial_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
        loops::emit_loop_end(chunks, current, state, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, initial_slot, line);
        return;
    }
    let (receiver_slot, fn_slot) = stash_iterable_and_arg(chunks, current, line);
    let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
    let acc_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    loops::emit_reduce(chunks, current, fn_slot, arr_slot, acc_slot, idx_slot, line);
}

/// Dart `iterable.elementAt(index)`.
pub fn emit_dart_iter_element_at(chunks: &mut [Chunk], current: usize, line: u32) {
    let index_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_dart_materialize_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    collections::emit_get(chunks, current, line);
}

/// Dart `iterable.take(n)`. Generator receivers use bounded continuation
/// draining so infinite generators stay safe.
pub fn emit_dart_iter_take(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, limit_slot) = stash_iterable_and_arg(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_is_generator(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, limit_slot, line);
    generators::emit_take_into_array(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_dart_materialize_iterable(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, limit_slot, line);
    host::emit(&mut chunks[current], "ecma:array", "slice", 3, line);
    chunks[current].emit_end(line);
}

/// Dart `iterable.skip(n)`.
pub fn emit_dart_iter_skip(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, count_slot) = stash_iterable_and_arg(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_dart_materialize_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
    host::emit(&mut chunks[current], "ecma:array", "slice", 2, line);
}

/// Dart `iterable.expand(fn)`.
pub fn emit_dart_iter_expand(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, fn_slot) = stash_iterable_and_arg(chunks, current, line);
    let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
    let result_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    let inner_slot = reserve_slot(&mut chunks[current]);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    emit_dart_materialize_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, inner_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, inner_slot, line);
    collections::emit_concat(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_dart_iter_expand_precurrent(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, fn_slot) = stash_iterable_and_arg(chunks, current, line);
    let arr_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_dart_materialize_iterable_precurrent(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    let result_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    let inner_slot = reserve_slot(&mut chunks[current]);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    emit_dart_materialize_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, inner_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, inner_slot, line);
    collections::emit_concat(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Dart `iterable.followedBy(other)`.
pub fn emit_dart_iter_followed_by(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, other_slot) = stash_iterable_and_arg(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
    emit_dart_materialize_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, other_slot, line);
    emit_dart_materialize_iterable(chunks, current, line);
    collections::emit_concat(chunks, current, line);
}

/// Dart `iterable.takeWhile(fn)`.
pub fn emit_dart_iter_take_while(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, fn_slot) = stash_iterable_and_arg(chunks, current, line);
    let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
    let result_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Dart `iterable.skipWhile(fn)`.
pub fn emit_dart_iter_skip_while(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, fn_slot) = stash_iterable_and_arg(chunks, current, line);
    let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
    let result_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    let skipping_slot = reserve_slot(&mut chunks[current]);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, skipping_slot, line);
    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, skipping_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, skipping_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Dart `iterable.distinct()` for the adjacent-dupe stream/list cases.
pub fn emit_dart_iter_distinct(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let eq_slot = reserve_slot(&mut chunks[current]);
    let receiver_slot = reserve_slot(&mut chunks[current]);
    if argc > 1 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, eq_slot, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, eq_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_slot, line);
    let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
    let result_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    let prev_slot = reserve_slot(&mut chunks[current]);
    let first_slot = reserve_slot(&mut chunks[current]);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, first_slot, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, prev_slot, line);
    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, first_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, first_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, prev_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, eq_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, prev_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_ne(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, eq_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, prev_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_end(line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, prev_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Dart `iterable.forEach(fn)`.
pub fn emit_dart_iter_for_each(chunks: &mut [Chunk], current: usize, line: u32) {
    let (receiver_slot, fn_slot) = stash_iterable_and_arg(chunks, current, line);
    let arr_slot = materialize_slot(chunks, current, receiver_slot, line);
    let idx_slot = reserve_slot(&mut chunks[current]);
    loops::emit_foreach(chunks, current, fn_slot, arr_slot, idx_slot, line);
}
