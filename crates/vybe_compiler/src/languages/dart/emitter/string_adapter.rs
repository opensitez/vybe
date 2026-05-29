//! Dart string helpers — Rust inline opcode emitters.
//!
//! Composes `compiler_common` emitters (collections + strings) so the
//! provider can be swapped in one place. No raw opcode emits where a
//! `common::*` helper exists.

use crate::emitter::{collections, strings};
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

/// Dart `s.isEmpty` — true iff length == 0. Stack: [s] → [bool].
pub fn emit_dart_is_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
}

/// Dart `s.isNotEmpty` — true iff length != 0. Stack: [s] → [bool].
pub fn emit_dart_is_not_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    crate::emitter::ops::emit_dyn_ne(&mut chunks[current], line);
}

/// Dart `print(value)` — calls value.toString() before logging, per
/// dart:core. Stack: [value] → []. Composes ecma:string.String (which
/// invokes the object's toString method) and wasi:cli.log. Imports
/// register on chunks[0] (the module-level imports chunk).
pub fn emit_dart_print(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    use vybe_bytecode::Op as VOp;
    let to_str = chunks[0].add_import("ecma:string", "String");
    chunks[current].emit_op_u16(VOp::CALL_IMPORT, to_str, line);
    chunks[current].emit(1, line);
    let log_idx = chunks[0].add_import("wasi:cli", "log");
    chunks[current].emit_op_u16(VOp::CALL_IMPORT, log_idx, line);
    chunks[current].emit(1, line);
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
    use vybe_bytecode::Value;
    use std::sync::Arc;
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op(Op::REF_IS_ARRAY, line);
    let to_arr = chunk.emit_jump(Op::BR_IF_TRUE, line);
    // Not an array — read the `first` field via STRUCT_GET.
    let key = chunks[current].add_constant(Value::String(Arc::from("first")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, key, line);
    let end = chunks[current].emit_jump(Op::BR, line);
    chunks[current].patch_jump(to_arr);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    collections::emit_get(chunks, current, line);
    chunks[current].patch_jump(end);
}

/// Dart `list.last` — `list[length - 1]`. Polymorphic; non-list
/// receivers fall through to STRUCT_GET("last").
pub fn emit_dart_list_last(chunks: &mut [Chunk], current: usize, line: u32) {
    use vybe_bytecode::Value;
    use std::sync::Arc;
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op(Op::REF_IS_ARRAY, line);
    let to_arr = chunk.emit_jump(Op::BR_IF_TRUE, line);
    let key = chunks[current].add_constant(Value::String(Arc::from("last")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, key, line);
    let end = chunks[current].emit_jump(Op::BR, line);
    chunks[current].patch_jump(to_arr);
    chunks[current].emit_op(Op::DUP, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_CONST_1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
    chunks[current].patch_jump(end);
}

/// Dart polymorphic `.length` — string | list | map. Stack: [coll] → [i32].
///
/// String/array routes through `compiler_common::{strings,collections}`
/// emitters; Map fall-through goes to `ecma:object.length` which returns
/// the property count (Dart `Map.length` semantics).
pub fn emit_dart_length(chunks: &mut [Chunk], current: usize, line: u32) {
    use vybe_bytecode::Op as VOp;
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op(Op::REF_IS_STRING, line);
    let to_str = chunk.emit_jump(Op::BR_IF_TRUE, line);
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op(Op::REF_IS_ARRAY, line);
    let to_arr = chunk.emit_jump(Op::BR_IF_TRUE, line);
    // Object/Map fall-through — count own enumerable properties via
    // `ecma:object.length`. Imports must register on chunks[0] (the
    // module-level imports chunk).
    let idx = chunks[0].add_import("ecma:object", "length");
    chunks[current].emit_op_u16(VOp::CALL_IMPORT, idx, line);
    chunks[current].emit(1, line);
    let end_all = chunks[current].emit_jump(Op::BR, line);
    chunks[current].patch_jump(to_arr);
    collections::emit_len(chunks, current, line);
    let end_arr = chunks[current].emit_jump(Op::BR, line);
    chunks[current].patch_jump(to_str);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].patch_jump(end_arr);
    chunks[current].patch_jump(end_all);
}
