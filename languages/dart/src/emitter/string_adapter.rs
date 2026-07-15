//! Dart string helpers — Rust inline opcode emitters.
//!
//! Composes `compiler_common` emitters (collections + strings) so the
//! provider can be swapped in one place. No raw opcode emits where a
//! `common::*` helper exists.

use vybe_emitter::instructions::{core_wasm, host};
use vybe_emitter::{collections, strings};
use std::sync::Arc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

const SB_BUFFER_KEY: &str = "__dart_string_buffer";
const SB_MARKER_KEY: &str = "__dart_string_buffer_marker";
const URI_HREF_KEY: &str = "href";
const URI_MARKER_KEY: &str = "__dart_uri_marker";

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn string_key(chunk: &mut Chunk, key: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(key)))
}

fn emit_sb_buffer_get(chunk: &mut Chunk, line: u32) {
    let key = string_key(chunk, SB_BUFFER_KEY);
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
}

fn emit_sb_marker_test(chunk: &mut Chunk, line: u32) {
    let key = string_key(chunk, SB_MARKER_KEY);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
}

fn emit_dart_value_to_string(chunk: &mut Chunk, line: u32) {
    let to_str = chunk.add_import("ecma:string", "String");
    chunk.emit_op_u16(Op::CALL_IMPORT, to_str, line);
    chunk.emit(1, line);
}

fn emit_dart_sb_append_value(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    let chunk = &mut chunks[current];
    let sb_slot = reserve_slot(chunk);
    let buffer_key = string_key(chunk, SB_BUFFER_KEY);
    chunk.emit_op_u16(Op::LOCAL_SET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buffer_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_dart_value_to_string(chunk, line);
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, sb_slot, line);
}

/// Dart `s.isEmpty` — true iff length == 0. Stack: [s] → [bool].
pub fn emit_dart_is_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_dart_length(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Dart `s.isNotEmpty` — true iff length != 0. Stack: [s] → [bool].
pub fn emit_dart_is_not_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_dart_length(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_ne(&mut chunks[current], line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Dart `n.isEven` — true iff `n % 2 == 0`. Stack: [n] → [bool].
pub fn emit_dart_is_even(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_REM_S, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Dart `n.isOdd` — true iff `n % 2 != 0`. Stack: [n] → [bool].
pub fn emit_dart_is_odd(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_REM_S, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_ne(&mut chunks[current], line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Dart `StringBuffer()` — plain object with one mutable string field.
pub fn emit_dart_sb_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buffer_key = string_key(chunk, SB_BUFFER_KEY);
    let marker_key = string_key(chunk, SB_MARKER_KEY);
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
    chunk.emit_op(Op::DROP, line);
    core_wasm::dup(chunk, line);
    chunk.emit_bool_const(true, line);
    chunk.emit_op_u16(Op::STRUCT_SET, marker_key, line);
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
    host::emit(&mut chunks[current], "wasm:js-string", "fromCharCode", 1, line);
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
    chunk.emit_op_u16(Op::STRUCT_SET, buffer_key, line);
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
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
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

/// Dart `print(value)` — calls value.toString() before logging, per
/// dart:core. Stack: [value] → []. Composes ecma:string.String (which
/// invokes the object's toString method) and wasi:cli.log. Import
/// tables are PER CHUNK: register on chunks[current], the chunk whose
/// CALL_IMPORT indexes them (registering on chunks[0] made the index
/// misresolve whenever the tables diverged).
pub fn emit_dart_print(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    use vybe_bytecode::Op as VOp;
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
    chunks[current].emit_op_u16(Op::STRUCT_GET, uri_marker, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, uri_key, line);
    chunks[current].emit_else(line);
    emit_sb_marker_test(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_sb_buffer_get(&mut chunks[current], line);
    chunks[current].emit_else(line);
    emit_dart_value_to_string(&mut chunks[current], line);
    chunks[current].emit_end(line);
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
    use std::sync::Arc;
    use vybe_bytecode::Value;
    let chunk = &mut chunks[current];
    core_wasm::dup(chunk, line);
    host::emit(chunk, "ecma:array", "isArray", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    // Not an array — read the `first` field via STRUCT_GET.
    let key = chunks[current].add_constant(Value::String(Arc::from("first")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, key, line);
    chunks[current].emit_end(line);
}

/// Dart `list.last` — `list[length - 1]`. Polymorphic; non-list
/// receivers fall through to STRUCT_GET("last").
pub fn emit_dart_list_last(chunks: &mut [Chunk], current: usize, line: u32) {
    use std::sync::Arc;
    use vybe_bytecode::Value;
    let chunk = &mut chunks[current];
    core_wasm::dup(chunk, line);
    host::emit(chunk, "ecma:array", "isArray", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    core_wasm::dup(&mut chunks[current], line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    let key = chunks[current].add_constant(Value::String(Arc::from("last")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, key, line);
    chunks[current].emit_end(line);
}

/// Dart polymorphic `.length` — string | list | map. Stack: [coll] → [i32].
///
/// String/array routes through `compiler_common::{strings,collections}`
/// emitters; Map fall-through goes to `ecma:object.length` which returns
/// the property count (Dart `Map.length` semantics).
pub fn emit_dart_length(chunks: &mut [Chunk], current: usize, line: u32) {
    use vybe_bytecode::Op as VOp;
    core_wasm::dup(&mut chunks[current], line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_else(line);
    core_wasm::dup(&mut chunks[current], line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_else(line);
    emit_sb_marker_test(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_sb_buffer_get(&mut chunks[current], line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_else(line);
    // Object/Map fall-through — count own enumerable properties via
    // `ecma:object.length`. Import tables are per chunk: register on
    // the chunk whose CALL_IMPORT indexes them.
    let idx = chunks[current].add_import("ecma:object", "length");
    chunks[current].emit_op_u16(VOp::CALL_IMPORT, idx, line);
    chunks[current].emit(1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}
