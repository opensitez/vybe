//! Common object helpers shared by language adapters.
//!
//! These are ECMA-shaped object/value operations. Language frontends should
//! normalize their own API names (`java.util.Objects.equals`, PHP object
//! helpers, etc.) into profile builtins that route here when the semantics are
//! genuinely shared.

use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

/// Dynamic value equality. Stack: [left, right] -> [Bool]
pub fn emit_equals(chunk: &mut Chunk, line: u32) {
    crate::ops::emit_dyn_eq(chunk, line);
    crate::ops::emit_i32_to_bool(chunk, line);
}

/// Null test. Stack: [value] -> [Bool]
pub fn emit_is_null(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::REF_IS_NULL, line);
    crate::ops::emit_i32_to_bool(chunk, line);
}

/// Non-null test. Stack: [value] -> [Bool]
pub fn emit_non_null(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    crate::ops::emit_i32_to_bool(chunk, line);
}

/// Set an object monitor's notified marker.
///
/// This is the object side of monitor notify/notifyAll. Languages still own
/// their scheduling/catch semantics, but the object-state mutation is common.
/// Stack: [object] -> [null]
pub fn emit_monitor_notify(chunk: &mut Chunk, line: u32) {
    chunk.emit_bool_const(true, line);
    let key = chunk.add_constant(Value::String(Arc::from("__j_notified")));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::NULL, line);
}

/// Deterministic object hash based on ECMA string conversion.
///
/// This mirrors Java's common `31*h + codeUnit` polynomial and is useful for
/// language APIs that need stable object/value hashing without identity hash
/// support. Null hashes to 0.
/// Stack: [value] -> [i32]
pub fn emit_hash_code(chunk: &mut Chunk, line: u32) {
    let value = chunk.alloc_scratch(4);
    let text = value + 1;
    let hash = value + 2;
    let index = value + 3;
    let to_string = chunk.add_import("ecma:string", "String");
    let length = chunk.add_import("wasm:js-string", "length");
    let char_code_at = chunk.add_import("wasm:js-string", "charCodeAt");

    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(to_string, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, text, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, hash, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);

    let outer = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);

    chunk.emit_i32_const(31, line);
    chunk.emit_op_u16(Op::LOCAL_GET, hash, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_call(char_code_at, 2, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, hash, line);

    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(outer);

    chunk.emit_op_u16(Op::LOCAL_GET, hash, line);
    chunk.emit_end(line);
}

/// Hash an array of values using the Java/List-style accumulator.
/// Stack: [array] -> [i32]
pub fn emit_hash_array(chunk: &mut Chunk, line: u32) {
    let items = chunk.alloc_scratch(3);
    let hash = items + 1;
    let index = items + 2;
    let length = chunk.add_import("ecma:array", "length");
    let get = chunk.add_import("ecma:array", "get");

    chunk.emit_op_u16(Op::LOCAL_SET, items, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, hash, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);

    let outer = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_op_u16(Op::LOCAL_GET, items, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);

    chunk.emit_i32_const(31, line);
    chunk.emit_op_u16(Op::LOCAL_GET, hash, line);
    chunk.emit_op(Op::I32_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, items, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_call(get, 2, line);
    emit_hash_code(chunk, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, hash, line);

    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(outer);

    chunk.emit_op_u16(Op::LOCAL_GET, hash, line);
}

/// Null-aware comparator dispatch. Stack: [a, b, comparator] -> [value]
pub fn emit_compare(chunk: &mut Chunk, line: u32) {
    let cmp = chunk.alloc_scratch(3);
    let b = cmp + 1;
    let a = cmp + 2;

    chunk.emit_op_u16(Op::LOCAL_SET, cmp, line);
    chunk.emit_op_u16(Op::LOCAL_SET, b, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a, line);

    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    crate::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(-1, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, cmp, line);
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    chunk.emit_op(Op::CALL_REF, line);
    chunk.emit(2, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// String conversion with a null fallback. Stack: [value, fallback] -> [string]
pub fn emit_to_string_or(chunk: &mut Chunk, line: u32) {
    let fallback = chunk.alloc_scratch(2);
    let value = fallback + 1;
    let to_string = chunk.add_import("ecma:string", "String");

    chunk.emit_op_u16(Op::LOCAL_SET, fallback, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, fallback, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_string_const("null", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, fallback, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(to_string, 1, line);
    chunk.emit_end(line);
}
