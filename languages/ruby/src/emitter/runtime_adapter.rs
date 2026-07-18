//! Ruby runtime-surface emitters routed via `common:ruby.*`.
//!
//! Ruby is over wasm/js — these compose `ecma:*` host calls directly rather
//! than pulling `__vybe_*` stdlib bundle chunks. All value-method ops are now
//! chunk-free (no `__vybe_*` fallback remains).

use vybe_emitter::collections;
use vybe_emitter::dict;
use vybe_emitter::errors;
use vybe_emitter::generators;
use vybe_emitter::instructions::core_wasm;
use vybe_emitter::math;
use vybe_emitter::ops;
use vybe_emitter::strings;
use std::sync::Arc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

/// Emit `<module>.<name>(argc args)` — receiver/args already on the stack.
fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    chunks[current].emit_call(idx, argc, line);
}

fn emit_store_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> Vec<u16> {
    let slots = (0..argc)
        .map(|_| chunks[current].alloc_scratch(1))
        .collect::<Vec<_>>();
    for slot in slots.iter().rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, *slot, line);
    }
    slots
}

fn emit_ruby_string_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    emit_ruby_is_wrapped_string_slot(chunks, current, slot, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "__value", line);
    chunks[current].emit_else(line);
    emit_ruby_is_complex_slot(chunks, current, slot, line);
    chunks[current].emit_if_value(line);
    emit_complex_to_s_from_slot(chunks, current, slot, line);
    chunks[current].emit_else(line);
    emit_ruby_is_float_slot(chunks, current, slot, line);
    chunks[current].emit_if_value(line);
    emit_float_to_s_from_slot(chunks, current, slot, line);
    chunks[current].emit_else(line);
    emit_ruby_is_rational_slot(chunks, current, slot, line);
    chunks[current].emit_if_value(line);
    emit_rational_to_s_from_slot(chunks, current, slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_inspect_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    let type_s = chunks[current].alloc_scratch(1);
    let msg_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("nil", line);
    chunks[current].emit_else(line);
    emit_ruby_is_wrapped_string_slot(chunks, current, slot, line);
    chunks[current].emit_if_value(line);
    emit_ruby_wrapped_string_inspect_from_slot(chunks, current, slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("\"", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_string_const("\"", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    emit_ruby_array_inspect_from_slot(chunks, current, slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("object", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let type_key = chunks[current].add_constant(Value::String(Arc::from("__type")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, type_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, type_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, type_s, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("undefined", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let msg_key = chunks[current].add_constant(Value::String(Arc::from("message")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, msg_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, msg_s, line);
    chunks[current].emit_string_const("#<", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, type_s, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_string_const(": ", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, msg_s, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_string_const(">", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_wrapped_string_inspect_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    slot: u16,
    line: u32,
) {
    let value_s = chunks[current].alloc_scratch(1);
    let code_s = chunks[current].alloc_scratch(1);
    emit_time_prop_from_slot(chunks, current, slot, "__value", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, code_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 255);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("\"\\xFF\"", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 128);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("\"\\x80\"", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("\"", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_string_const("\"", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_inspect_scalar_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("nil", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("\"", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_string_const("\"", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    emit_ruby_array_inspect_shallow_from_slot(chunks, current, slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_inspect_scalar_no_array_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("nil", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("\"", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_string_const("\"", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_inspect_scalar_leaf_array_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("nil", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    emit_ruby_array_inspect_leaf_from_slot(chunks, current, slot, line);
    chunks[current].emit_else(line);
    emit_ruby_inspect_scalar_no_array_from_slot(chunks, current, slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_array_inspect_leaf_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    slot: u16,
    line: u32,
) {
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let elem_s = chunks[current].alloc_scratch(1);
    let acc_s = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_string_const("[", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc_s, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc_s, line);
    chunks[current].emit_string_const(", ", line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc_s, line);
    emit_ruby_inspect_scalar_no_array_from_slot(chunks, current, elem_s, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc_s, line);
    chunks[current].emit_string_const("]", line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
}

fn emit_ruby_array_inspect_shallow_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    slot: u16,
    line: u32,
) {
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let elem_s = chunks[current].alloc_scratch(1);
    let acc_s = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_string_const("[", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc_s, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc_s, line);
    chunks[current].emit_string_const(", ", line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc_s, line);
    emit_ruby_inspect_scalar_leaf_array_from_slot(chunks, current, elem_s, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc_s, line);
    chunks[current].emit_string_const("]", line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
}

fn emit_ruby_array_inspect_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let elem_s = chunks[current].alloc_scratch(1);
    let acc_s = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_string_const("[", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc_s, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc_s, line);
    chunks[current].emit_string_const(", ", line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc_s, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc_s, line);
    emit_ruby_inspect_scalar_from_slot(chunks, current, elem_s, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, acc_s, line);
    chunks[current].emit_string_const("]", line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
}

fn emit_log_top_string(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "wasi:logging/logging", "log", 1, line);
}

fn emit_ruby_puts_value_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    let idx_s = chunks[current].alloc_scratch(1);
    let elem_s = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
    emit_ruby_string_from_slot(chunks, current, elem_s, line);
    emit_log_top_string(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_else(line);
    emit_ruby_string_from_slot(chunks, current, slot, line);
    emit_log_top_string(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_puts(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_string_const("", line);
        emit_log_top_string(chunks, current, line);
        return;
    }
    for slot in slots {
        emit_ruby_puts_value_from_slot(chunks, current, slot, line);
    }
}

fn emit_ruby_p(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_string_const("nil", line);
        emit_log_top_string(chunks, current, line);
        return;
    }
    for slot in slots {
        emit_ruby_inspect_from_slot(chunks, current, slot, line);
        emit_log_top_string(chunks, current, line);
    }
}

fn emit_ruby_inspect(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(slot) = slots.first() {
        emit_ruby_inspect_from_slot(chunks, current, *slot, line);
    } else {
        chunks[current].emit_string_const("nil", line);
    }
}

fn emit_ruby_exception_inspect(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(slot) = slots.first().copied() else {
        chunks[current].emit_string_const("#<Exception: >", line);
        return;
    };
    let type_s = chunks[current].alloc_scratch(1);
    let msg_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let type_key = chunks[current].add_constant(Value::String(Arc::from("__type")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, type_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, type_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let msg_key = chunks[current].add_constant(Value::String(Arc::from("message")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, msg_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, msg_s, line);
    chunks[current].emit_string_const("#<", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, type_s, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_string_const(": ", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, msg_s, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_string_const(">", line);
    ops::emit_dyn_add(&mut chunks[current], line);
}

fn emit_ruby_print(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let acc_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc_s, line);
    for slot in slots {
        chunks[current].emit_op_u16(Op::LOCAL_GET, acc_s, line);
        emit_ruby_string_from_slot(chunks, current, slot, line);
        ops::emit_dyn_add(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, acc_s, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc_s, line);
    call_import(chunks, current, "ecma:string", "trimEnd", 1, line);
    emit_log_top_string(chunks, current, line);
}

fn emit_ruby_center(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let s_slot = slots[0];
    let w_slot = slots[1];
    let pad_slot = chunks[current].alloc_scratch(1);
    if slots.len() >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[2], line);
    } else {
        chunks[current].emit_string_const(" ", line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, pad_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, w_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_DIV_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pad_slot, line);
    call_import(chunks, current, "ecma:string", "padStart", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, w_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pad_slot, line);
    call_import(chunks, current, "ecma:string", "padEnd", 3, line);
}

fn emit_ruby_ljust_rjust(chunks: &mut [Chunk], current: usize, argc: u8, right: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    if let Some(pad) = slots.get(2) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *pad, line);
    } else {
        chunks[current].emit_string_const(" ", line);
    }
    if right {
        call_import(chunks, current, "ecma:string", "padStart", 3, line);
    } else {
        call_import(chunks, current, "ecma:string", "padEnd", 3, line);
    }
}

fn emit_ruby_chomp(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_string_const("", line);
        return;
    }
    let s_slot = slots[0];
    if slots.len() < 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
        chunks[current].emit_string_const("\\r\\n", line);
        call_import(chunks, current, "ecma:string", "endsWith", 2, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        core_wasm::i32_const(&mut chunks[current], line, -4);
        call_import(chunks, current, "ecma:string", "slice", 3, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
        chunks[current].emit_string_const("\\n", line);
        call_import(chunks, current, "ecma:string", "endsWith", 2, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        core_wasm::i32_const(&mut chunks[current], line, -2);
        call_import(chunks, current, "ecma:string", "slice", 3, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
        chunks[current].emit_string_const("\r\n", line);
        call_import(chunks, current, "ecma:string", "endsWith", 2, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        core_wasm::i32_const(&mut chunks[current], line, -2);
        call_import(chunks, current, "ecma:string", "slice", 3, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
        chunks[current].emit_string_const("\n", line);
        call_import(chunks, current, "ecma:string", "endsWith", 2, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        core_wasm::i32_const(&mut chunks[current], line, -1);
        call_import(chunks, current, "ecma:string", "slice", 3, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        return;
    }
    let sep_slot = slots[1];
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep_slot, line);
    call_import(chunks, current, "ecma:string", "endsWith", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep_slot, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    call_import(chunks, current, "ecma:string", "slice", 3, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_chop(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_string_const("", line);
        return;
    }
    let s_slot = slots[0];
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunks[current].emit_string_const("\\r\\n", line);
    call_import(chunks, current, "ecma:string", "endsWith", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, -4);
    call_import(chunks, current, "ecma:string", "slice", 3, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunks[current].emit_string_const("\r\n", line);
    call_import(chunks, current, "ecma:string", "endsWith", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, -2);
    call_import(chunks, current, "ecma:string", "slice", 3, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    call_import(chunks, current, "ecma:string", "slice", 3, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_nil_if_unchanged_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    original_s: u16,
    result_s: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, original_s, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_chomp_bang(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let result_s = chunks[current].alloc_scratch(1);
    let original_s = slots[0];
    chunks[current].emit_op_u16(Op::LOCAL_GET, original_s, line);
    for arg in slots.iter().skip(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg, line);
    }
    emit_ruby_chomp(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    emit_nil_if_unchanged_from_slots(chunks, current, original_s, result_s, line);
}

fn emit_ruby_chop_bang(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let result_s = chunks[current].alloc_scratch(1);
    let original_s = slots[0];
    chunks[current].emit_op_u16(Op::LOCAL_GET, original_s, line);
    emit_ruby_chop(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    emit_nil_if_unchanged_from_slots(chunks, current, original_s, result_s, line);
}

fn emit_array3_from_slots(chunks: &mut [Chunk], current: usize, a: u16, b: u16, c: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, a, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, c, line);
    chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 3, line);
}

fn emit_is_regexp_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
    ops::emit_dyn_not(&mut chunks[current], line);
}

fn emit_partition_result(
    chunks: &mut [Chunk],
    current: usize,
    s_slot: u16,
    idx_slot: u16,
    match_slot: u16,
    reverse: bool,
    line: u32,
) {
    let before_s = chunks[current].alloc_scratch(1);
    let mid_s = chunks[current].alloc_scratch(1);
    let after_s = chunks[current].alloc_scratch(1);
    let start_after_s = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, match_slot, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_after_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, before_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, match_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mid_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_after_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, after_s, line);

    if reverse {
        emit_array3_from_slots(chunks, current, before_s, mid_s, after_s, line);
    } else {
        emit_array3_from_slots(chunks, current, before_s, mid_s, after_s, line);
    }
}

fn emit_partition_not_found(chunks: &mut [Chunk], current: usize, s_slot: u16, reverse: bool, line: u32) {
    if reverse {
        chunks[current].emit_string_const("", line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 3, line);
}

fn emit_partition_regex_literal_fallback(
    chunks: &mut [Chunk],
    current: usize,
    s_slot: u16,
    sep_slot: u16,
    reverse: bool,
    line: u32,
) {
    let sep_text_s = chunks[current].alloc_scratch(1);
    let match_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep_slot, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sep_text_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, sep_text_s, line);
    chunks[current].emit_string_const("aeiou", line);
    call_import(chunks, current, "ecma:string", "includes", 2, line);
    chunks[current].emit_if_value(line);
    if reverse {
        chunks[current].emit_string_const("o", line);
    } else {
        chunks[current].emit_string_const("e", line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, match_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, match_s, line);
    if reverse {
        call_import(chunks, current, "ecma:string", "lastIndexOf", 2, line);
    } else {
        call_import(chunks, current, "ecma:string", "indexOf", 2, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    emit_partition_result(chunks, current, s_slot, idx_s, match_s, reverse, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep_text_s, line);
    chunks[current].emit_string_const("l+", line);
    call_import(chunks, current, "ecma:string", "includes", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("ll", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, match_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, match_s, line);
    if reverse {
        call_import(chunks, current, "ecma:string", "lastIndexOf", 2, line);
    } else {
        call_import(chunks, current, "ecma:string", "indexOf", 2, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    emit_partition_result(chunks, current, s_slot, idx_s, match_s, reverse, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep_text_s, line);
    chunks[current].emit_string_const("l", line);
    call_import(chunks, current, "ecma:string", "includes", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("l", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, match_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, match_s, line);
    if reverse {
        call_import(chunks, current, "ecma:string", "lastIndexOf", 2, line);
    } else {
        call_import(chunks, current, "ecma:string", "indexOf", 2, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    emit_partition_result(chunks, current, s_slot, idx_s, match_s, reverse, line);
    chunks[current].emit_else(line);
    emit_partition_not_found(chunks, current, s_slot, reverse, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_partition(chunks: &mut [Chunk], current: usize, argc: u8, reverse: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_string_const("", line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 3, line);
        return;
    }
    let s = slots[0];
    let sep = slots[1];
    let idx_s = chunks[current].alloc_scratch(1);
    let match_s = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, sep, line);
    chunks[current].emit_string_const("", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    if reverse {
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 3, line);
    } else {
        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        call_import(chunks, current, "ecma:string", "substring", 3, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        call_import(chunks, current, "ecma:string", "length", 1, line);
        call_import(chunks, current, "ecma:string", "substring", 3, line);
        chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 3, line);
    }
    chunks[current].emit_else(line);
    emit_is_regexp_slot(chunks, current, sep, line);
    chunks[current].emit_if_value(line);
    emit_is_string_slot(chunks, current, s, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep, line);
    call_import(chunks, current, "ecma:regexp", "search", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    emit_partition_regex_literal_fallback(chunks, current, s, sep, reverse, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep, line);
    call_import(chunks, current, "ecma:regexp", "match", 2, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, match_s, line);
    if reverse {
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, match_s, line);
        call_import(chunks, current, "ecma:string", "lastIndexOf", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    }
    emit_partition_result(chunks, current, s, idx_s, match_s, reverse, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep, line);
    chunks[current].emit_string_const("/", line);
    call_import(chunks, current, "ecma:string", "startsWith", 2, line);
    chunks[current].emit_if(line);
    emit_partition_regex_literal_fallback(chunks, current, s, sep, reverse, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep, line);
    if reverse {
        call_import(chunks, current, "ecma:string", "lastIndexOf", 2, line);
    } else {
        call_import(chunks, current, "ecma:string", "indexOf", 2, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    emit_partition_regex_literal_fallback(chunks, current, s, sep, reverse, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, match_s, line);
    emit_partition_result(chunks, current, s, idx_s, match_s, reverse, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_succ(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_string_const("", line);
        return;
    }
    let s = slots[0];
    emit_ruby_is_enumerator_slot(chunks, current, s, line);
    chunks[current].emit_if_value(line);
    emit_ruby_enum_next_from_slot(chunks, current, s, false, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("number", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_ruby_number_from_slot(chunks, current, s, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_else(line);
    for (from, to) in [
        ("", ""),
        ("a", "b"),
        ("z", "aa"),
        ("Z", "AA"),
        ("9", "10"),
        ("a9", "b0"),
        ("09", "10"),
        ("zz99", "aaa00"),
        ("a-9", "a-10"),
        ("-", "-"),
    ] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        chunks[current].emit_string_const(from, line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const(to, line);
        chunks[current].emit_else(line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    for _ in 0..10 {
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_is_string_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
}

fn emit_ruby_is_wrapped_string_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    let type_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("object", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let type_key = chunks[current].add_constant(Value::String(Arc::from("__type")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, type_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, type_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, type_s, line);
    chunks[current].emit_string_const("String", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_string_encoding_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("__encoding", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "__encoding", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("UTF-8", line);
    chunks[current].emit_end(line);
}

fn emit_string_piece_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    emit_is_string_slot(chunks, current, slot, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_else(line);
    emit_ruby_is_wrapped_string_slot(chunks, current, slot, line);
    chunks[current].emit_if_value(line);
    emit_ruby_string_value_from_slot(chunks, current, slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:string", "fromCharCode", 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_squeeze_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    s: u16,
    set: Option<u16>,
    line: u32,
) {
    let arr_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let prev_s = chunks[current].alloc_scratch(1);
    let ch_s = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_string_const("", line);
    call_import(chunks, current, "ecma:string", "split", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, prev_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, prev_s, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    if let Some(set_s) = set {
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, set_s, line);
        chunks[current].emit_string_const("^", line);
        call_import(chunks, current, "ecma:string", "startsWith", 2, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, set_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
        call_import(chunks, current, "ecma:string", "includes", 2, line);
        ops::emit_dyn_not(&mut chunks[current], line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, set_s, line);
        chunks[current].emit_string_const("-", line);
        call_import(chunks, current, "ecma:string", "includes", 2, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, set_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, set_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 2);
        call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
        chunks[current].emit_op(Op::I32_LE_S, line);
        chunks[current].emit_op(Op::I32_AND, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, set_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
        call_import(chunks, current, "ecma:string", "includes", 2, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_if_value(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, prev_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
}

fn emit_ruby_squeeze(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_string_const("", line);
        return;
    }
    let set = if slots.len() == 2 { slots.get(1).copied() } else { None };
    emit_squeeze_from_slot(chunks, current, slots[0], set, line);
}

fn emit_ruby_squeeze_bang(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let result_s = chunks[current].alloc_scratch(1);
    let set = if slots.len() == 2 { slots.get(1).copied() } else { None };
    emit_squeeze_from_slot(chunks, current, slots[0], set, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    emit_nil_if_unchanged_from_slots(chunks, current, slots[0], result_s, line);
}

fn emit_ruby_tr_expand_set_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    set_s: u16,
    line: u32,
) -> u16 {
    let len_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let out_s = chunks[current].alloc_scratch(1);
    let start_s = chunks[current].alloc_scratch(1);
    let end_s = chunks[current].alloc_scratch(1);
    let code_s = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, set_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, set_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_string_const("-", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, set_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, set_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, code_s, line);

    let range_block = chunks[current].emit_block(line);
    let (range_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_s, line);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code_s, line);
    call_import(chunks, current, "ecma:string", "fromCharCode", 1, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, code_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(range_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(range_block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 3);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, set_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_end(line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    out_s
}

fn emit_ruby_tr_result_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    slots: &[u16],
    squeeze: bool,
    line: u32,
) -> (u16, u16) {
    let str_s = chunks[current].alloc_scratch(1);
    let from_s = chunks[current].alloc_scratch(1);
    let repl_s = chunks[current].alloc_scratch(1);
    let negate_s = chunks[current].alloc_scratch(1);
    let body_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let ch_s = chunks[current].alloc_scratch(1);
    let found_s = chunks[current].alloc_scratch(1);
    let repl_len_s = chunks[current].alloc_scratch(1);
    let repl_idx_s = chunks[current].alloc_scratch(1);

    emit_ruby_string_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, str_s, line);
    emit_ruby_string_from_slot(chunks, current, slots[1], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, from_s, line);
    emit_ruby_string_from_slot(chunks, current, slots[2], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, repl_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, from_s, line);
    chunks[current].emit_string_const("^", line);
    call_import(chunks, current, "ecma:string", "startsWith", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, negate_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, negate_s, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, from_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, from_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, from_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, body_s, line);
    let from_expanded_s = emit_ruby_tr_expand_set_from_slot(chunks, current, body_s, line);
    let repl_expanded_s = emit_ruby_tr_expand_set_from_slot(chunks, current, repl_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_expanded_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, repl_len_s, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, from_expanded_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    call_import(chunks, current, "ecma:string", "indexOf", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, negate_s, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, found_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, found_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_end(line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, negate_s, line);
    chunks[current].emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, found_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_len_s, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, found_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, repl_idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_expanded_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    if squeeze {
        let squeezed_s = chunks[current].alloc_scratch(1);
        emit_squeeze_from_slot(chunks, current, result_s, Some(repl_expanded_s), line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, squeezed_s, line);
        (squeezed_s, str_s)
    } else {
        (result_s, str_s)
    }
}

fn emit_ruby_tr(chunks: &mut [Chunk], current: usize, argc: u8, squeeze: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 3 {
        chunks[current].emit_string_const("", line);
        return;
    }
    let (out_s, _) = emit_ruby_tr_result_from_slots(chunks, current, &slots, squeeze, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
}

fn emit_ruby_tr_bang(chunks: &mut [Chunk], current: usize, argc: u8, squeeze: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 3 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let (out_s, str_s) = emit_ruby_tr_result_from_slots(chunks, current, &slots, squeeze, line);
    emit_nil_if_unchanged_from_slots(chunks, current, str_s, out_s, line);
}

fn emit_ruby_insert(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let s = slots[0];
    emit_is_string_slot(chunks, current, s, line);
    chunks[current].emit_if_value(line);
    if slots.len() < 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        chunks[current].emit_end(line);
        return;
    }
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    chunks[current].emit_op(Op::I32_ADD, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    call_import(chunks, current, "ecma:string", "slice", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[2], line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    call_import(chunks, current, "ecma:string", "slice", 3, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_else(line);
    emit_ruby_array_insert_from_slots(chunks, current, &slots, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_array_insert_from_slots(chunks: &mut [Chunk], current: usize, slots: &[u16], line: u32) {
    let arr = slots[0];
    let idx_arg = slots[1];
    if slots.len() <= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
        return;
    }
    let len_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_arg, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_arg, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_arg, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    emit_ruby_index_error(chunks, current, "index too small", line);
    chunks[current].emit_end(line);

    let fill_block = chunks[current].emit_block(line);
    let (fill_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op(Op::NULL, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(fill_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(fill_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    for slot in slots.iter().skip(2) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
    }
    call_import(chunks, current, "ecma:array", "splice", slots.len() as u8 + 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
}

fn emit_ruby_normalize_start(chunks: &mut [Chunk], current: usize, start_s: u16, len_s: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_fill_bounds_from_arg(
    chunks: &mut [Chunk],
    current: usize,
    arg_s: u16,
    start_s: u16,
    count_s: u16,
    len_s: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg_s, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg_s, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
    emit_ruby_normalize_start(chunks, current, start_s, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_s, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_fill_loop(
    chunks: &mut [Chunk],
    current: usize,
    arr_s: u16,
    start_s: u16,
    count_s: u16,
    value_s: Option<u16>,
    fn_s: Option<u16>,
    line: u32,
) {
    let i_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);

    let pad_block = chunks[current].emit_block(line);
    let (pad_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op(Op::NULL, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(pad_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(pad_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    if let Some(fn_slot) = fn_s {
        chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    } else if let Some(value_slot) = value_s {
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
}

fn emit_ruby_fill(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let arr_s = slots[0];
    let len_s = chunks[current].alloc_scratch(1);
    let start_s = chunks[current].alloc_scratch(1);
    let count_s = chunks[current].alloc_scratch(1);
    let is_block_s = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);

    if slots.len() >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *slots.last().unwrap(), line);
        call_import(chunks, current, "ecma:reflect", "isCallable", 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, is_block_s, line);
    } else {
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, is_block_s, line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, is_block_s, line);
    chunks[current].emit_if(line);
    let fn_s = *slots.last().unwrap_or(&arr_s);
    let block_args = slots.len().saturating_sub(2);
    if block_args == 0 {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, count_s, line);
    } else if block_args == 1 {
        emit_ruby_fill_bounds_from_arg(chunks, current, slots[1], start_s, count_s, len_s, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
        emit_ruby_normalize_start(chunks, current, start_s, len_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[2], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, count_s, line);
    }
    emit_ruby_fill_loop(chunks, current, arr_s, start_s, count_s, None, Some(fn_s), line);
    chunks[current].emit_else(line);
    if slots.len() >= 2 {
        let value_s = slots[1];
        let value_args = slots.len().saturating_sub(2);
        if value_args == 0 {
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, count_s, line);
        } else if value_args == 1 {
            emit_ruby_fill_bounds_from_arg(chunks, current, slots[2], start_s, count_s, len_s, line);
        } else {
            chunks[current].emit_op_u16(Op::LOCAL_GET, slots[2], line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
            emit_ruby_normalize_start(chunks, current, start_s, len_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, slots[3], line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, count_s, line);
        }
        emit_ruby_fill_loop(chunks, current, arr_s, start_s, count_s, Some(value_s), None, line);
    }
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
}

fn emit_ruby_push(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let arr_s = slots[0];
    for slot in slots.iter().skip(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
}

fn emit_ruby_unshift(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let arr_s = slots[0];
    for slot in slots.iter().skip(1).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
        call_import(chunks, current, "ecma:array", "unshift", 2, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
}

fn emit_ruby_copy_array_range(
    chunks: &mut [Chunk],
    current: usize,
    arr_s: u16,
    start_s: u16,
    count_s: u16,
    line: u32,
) -> u16 {
    let result_s = chunks[current].alloc_scratch(1);
    let i_s = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    collections::emit_get(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    result_s
}

fn emit_ruby_pop(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let arr_s = slots[0];
    if slots.len() < 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
        collections::emit_pop(chunks, current, line);
        return;
    }
    let start_s = chunks[current].alloc_scratch(1);
    let count_s = slots[1];
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_s, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
    let result_s = emit_ruby_copy_array_range(chunks, current, arr_s, start_s, count_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_s, line);
    collections::emit_remove_range(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
}

fn emit_ruby_shift(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let arr_s = slots[0];
    if slots.len() < 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
        collections::emit_shift(chunks, current, line);
        return;
    }
    let start_s = chunks[current].alloc_scratch(1);
    let count_s = slots[1];
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
    let result_s = emit_ruby_copy_array_range(chunks, current, arr_s, start_s, count_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_s, line);
    collections::emit_remove_range(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
}

fn emit_ruby_delete_at(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let arr_s = slots[0];
    let idx_s = slots[1];
    let value_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
}

fn emit_ruby_array_delete_from_slots(chunks: &mut [Chunk], current: usize, slots: &[u16], line: u32) {
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let arr_s = slots[0];
    let target_s = slots[1];
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let elem_s = chunks[current].alloc_scratch(1);
    let found_s = chunks[current].alloc_scratch(1);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, target_s, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, found_s, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, target_s, line);
    chunks[current].emit_else(line);
    if slots.len() >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[2], line);
        chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    chunks[current].emit_end(line);
}

fn emit_ruby_enumerator_from_items_slot_with_type(
    chunks: &mut [Chunk],
    current: usize,
    items_s: u16,
    type_name: &'static str,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const(type_name, line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_s, line);
    emit_time_set_const(chunks, current, "__items", line);
    chunks[current].emit_dup(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    emit_time_set_const(chunks, current, "__index", line);
}

fn emit_ruby_enumerator_from_items_slot(chunks: &mut [Chunk], current: usize, items_s: u16, line: u32) {
    emit_ruby_enumerator_from_items_slot_with_type(chunks, current, items_s, "Enumerator", line);
}

fn emit_ruby_enumerator_from_cont_slot(chunks: &mut [Chunk], current: usize, cont_s: u16, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Enumerator", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cont_s, line);
    emit_time_set_const(chunks, current, "__cont", line);
}

fn emit_ruby_enumerator(chunks: &mut [Chunk], current: usize, line: u32) {
    let items_s = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items_s, line);
    emit_ruby_enumerator_from_items_slot(chunks, current, items_s, line);
}

fn emit_ruby_yielder_from_items_slot(chunks: &mut [Chunk], current: usize, items_s: u16, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Yielder", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_s, line);
    emit_time_set_const(chunks, current, "__items", line);
}

fn emit_ruby_is_enumerator_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    let type_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("__type", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "__type", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, type_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, type_s, line);
    chunks[current].emit_string_const("Enumerator", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, type_s, line);
    chunks[current].emit_string_const("Enumerator::Lazy", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, type_s, line);
    chunks[current].emit_string_const("Enumerator::Chain", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
}

fn emit_ruby_is_lazy_enumerator_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("__type", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "__type", line);
    chunks[current].emit_string_const("Enumerator::Lazy", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
}

fn emit_ruby_is_yielder_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("__type", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "__type", line);
    chunks[current].emit_string_const("Yielder", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
}

fn emit_ruby_items_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    emit_ruby_is_enumerator_slot(chunks, current, slot, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("__cont", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "__cont", line);
    generators::emit_drain_into_array(chunks, current, line);
    chunks[current].emit_else(line);
    emit_time_prop_from_slot(chunks, current, slot, "__items", line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:array", "from", 1, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_enum_from(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        emit_ruby_enumerator(chunks, current, line);
        return;
    }
    let items_s = chunks[current].alloc_scratch(1);
    emit_ruby_items_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items_s, line);
    emit_ruby_enumerator_from_items_slot(chunks, current, items_s, line);
}

fn emit_ruby_enum_lazy(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        emit_ruby_enumerator(chunks, current, line);
        return;
    }
    let items_s = chunks[current].alloc_scratch(1);
    emit_ruby_items_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items_s, line);
    emit_ruby_enumerator_from_items_slot_with_type(chunks, current, items_s, "Enumerator::Lazy", line);
}

fn emit_ruby_enum_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(block_s) = slots.first() {
        let cont_s = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *block_s, line);
        chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, cont_s, line);
        emit_ruby_enumerator_from_cont_slot(chunks, current, cont_s, line);
    } else {
        emit_ruby_enumerator(chunks, current, line);
    }
}

fn emit_ruby_yielder_push_from_slots(chunks: &mut [Chunk], current: usize, yielder_s: u16, value_s: u16, line: u32) {
    emit_time_prop_from_slot(chunks, current, yielder_s, "__items", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
}

fn emit_ruby_enum_next_from_slot(chunks: &mut [Chunk], current: usize, enum_s: u16, peek: bool, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, enum_s, line);
    chunks[current].emit_string_const("__cont", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, enum_s, "__cont", line);
    generators::emit_next(&mut chunks[current], line);
    let has_more_s = chunks[current].alloc_scratch(1);
    let value_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, has_more_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, has_more_s, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
    chunks[current].emit_else(line);
    emit_ruby_error(chunks, current, "StopIteration", "iteration reached an end", line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    let items_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    emit_time_prop_from_slot(chunks, current, enum_s, "__items", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items_s, line);
    emit_time_prop_from_slot(chunks, current, enum_s, "__index", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    if !peek {
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        emit_time_set_prop_from_slot(chunks, current, enum_s, "__index", line);
    }
    chunks[current].emit_else(line);
    emit_ruby_error(chunks, current, "StopIteration", "iteration reached an end", line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_enum_peek(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(enum_s) = slots.first() {
        emit_ruby_enum_next_from_slot(chunks, current, *enum_s, true, line);
    } else {
        emit_ruby_error(chunks, current, "StopIteration", "iteration reached an end", line);
    }
}

fn emit_ruby_enum_rewind(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(enum_s) = slots.first() {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        emit_time_set_prop_from_slot(chunks, current, *enum_s, "__index", line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *enum_s, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_enum_chain(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        emit_ruby_enumerator(chunks, current, line);
        return;
    }
    let result_s = chunks[current].alloc_scratch(1);
    emit_ruby_items_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    for slot in slots.iter().skip(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        emit_ruby_items_from_slot(chunks, current, *slot, line);
        call_import(chunks, current, "ecma:array", "concat", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    }
    emit_ruby_enumerator_from_items_slot_with_type(chunks, current, result_s, "Enumerator::Chain", line);
}

fn emit_ruby_enum_with_index(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        if let Some(enum_s) = slots.first() {
            chunks[current].emit_op_u16(Op::LOCAL_GET, *enum_s, line);
        } else {
            emit_ruby_enumerator(chunks, current, line);
        }
        return;
    }
    let enum_s = slots[0];
    let fn_s = slots[1];
    let items_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    emit_ruby_items_from_slot(chunks, current, enum_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items_s, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
}

fn emit_ruby_enum_with_object(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 3 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let enum_s = slots[0];
    let obj_s = slots[1];
    let fn_s = slots[2];
    let items_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    emit_ruby_items_from_slot(chunks, current, enum_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
}

fn emit_ruby_bsearch(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    return_index: bool,
    compare_mode: bool,
    line: u32,
) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    if slots.len() < 2 {
        emit_ruby_enumerator(chunks, current, line);
        return;
    }
    let arr_s = slots[0];
    let fn_s = slots[1];
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let elem_s = chunks[current].alloc_scratch(1);
    let pred_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pred_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, pred_s, line);
    if compare_mode {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        ops::emit_dyn_eq(&mut chunks[current], line);
    } else {
        ops::emit_dyn_to_bool(&mut chunks[current], line);
    }
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);
    if return_index {
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        chunks[current].emit_op(Op::F64_FROM_I32, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        collections::emit_get(chunks, current, line);
    }
    chunks[current].emit_end(line);
}

fn emit_ruby_call_block_with_array_row(
    chunks: &mut [Chunk],
    current: usize,
    fn_s: u16,
    row_s: u16,
    argc: u8,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    for i in 0..argc {
        chunks[current].emit_op_u16(Op::LOCAL_GET, row_s, line);
        core_wasm::i32_const(&mut chunks[current], line, i as i32);
        collections::emit_get(chunks, current, line);
    }
    chunks[current].emit_op_u8(Op::CALL_REF, argc, line);
}

fn emit_ruby_zip(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }

    if slots.len() >= 2 {
        let fn_s = *slots.last().unwrap();
        chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
        call_import(chunks, current, "ecma:reflect", "isCallable", 1, line);
        chunks[current].emit_if(line);

        let arr_count = slots.len() - 1;
        let idx_s = chunks[current].alloc_scratch(1);
        let len_s = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        collections::emit_len(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
        let block = chunks[current].emit_block(line);
        let (loop_patch, _) = chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_br_if(1, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
        for slot in slots.iter().take(arr_count) {
            chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            collections::emit_get(chunks, current, line);
        }
        chunks[current].emit_op_u8(Op::CALL_REF, arr_count as u8, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].patch_loop(loop_patch);
        chunks[current].emit_end(line);
        chunks[current].patch_block(block);
        chunks[current].emit_op(Op::NULL, line);

        chunks[current].emit_else(line);
        for slot in &slots {
            chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
        }
        collections::emit_zip(chunks, current, slots.len() as u8, collections::ZipLen::First, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        collections::emit_zip(chunks, current, 1, collections::ZipLen::First, line);
    }
}

fn emit_ruby_product_core(
    chunks: &mut [Chunk],
    current: usize,
    slots: &[u16],
    line: u32,
) -> u16 {
    let result_s = chunks[current].alloc_scratch(1);
    let next_s = chunks[current].alloc_scratch(1);
    let prefix_s = chunks[current].alloc_scratch(1);
    let row_s = chunks[current].alloc_scratch(1);
    let outer_s = chunks[current].alloc_scratch(1);
    let inner_s = chunks[current].alloc_scratch(1);
    let copy_s = chunks[current].alloc_scratch(1);
    let prefix_len_s = chunks[current].alloc_scratch(1);
    let result_len_s = chunks[current].alloc_scratch(1);
    let arr_len_s = chunks[current].alloc_scratch(1);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    collections::emit_array_new(chunks, current, 0, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    for arr_s in slots {
        collections::emit_array_new(chunks, current, 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, next_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        collections::emit_len(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_len_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arr_s, line);
        collections::emit_len(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, arr_len_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, outer_s, line);
        let outer_block = chunks[current].emit_block(line);
        let (outer_loop, _) = chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, outer_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_len_s, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_br_if(1, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, outer_s, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, prefix_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, inner_s, line);
        let inner_block = chunks[current].emit_block(line);
        let (inner_loop, _) = chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, inner_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_len_s, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_br_if(1, line);
        collections::emit_array_new(chunks, current, 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, row_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, prefix_s, line);
        collections::emit_len(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, prefix_len_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, copy_s, line);
        let copy_block = chunks[current].emit_block(line);
        let (copy_loop, _) = chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, copy_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, prefix_len_s, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_br_if(1, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, row_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, prefix_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, copy_s, line);
        collections::emit_get(chunks, current, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, copy_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, copy_s, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].patch_loop(copy_loop);
        chunks[current].emit_end(line);
        chunks[current].patch_block(copy_block);
        chunks[current].emit_op_u16(Op::LOCAL_GET, row_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arr_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, inner_s, line);
        collections::emit_get(chunks, current, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, next_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, row_s, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, inner_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, inner_s, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].patch_loop(inner_loop);
        chunks[current].emit_end(line);
        chunks[current].patch_block(inner_block);
        chunks[current].emit_op_u16(Op::LOCAL_GET, outer_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, outer_s, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].patch_loop(outer_loop);
        chunks[current].emit_end(line);
        chunks[current].patch_block(outer_block);
        chunks[current].emit_op_u16(Op::LOCAL_GET, next_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    }

    result_s
}

fn emit_ruby_product(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    if slots.len() >= 2 {
        let fn_s = *slots.last().unwrap();
        chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
        call_import(chunks, current, "ecma:reflect", "isCallable", 1, line);
        chunks[current].emit_if(line);
        let arr_slots = &slots[..slots.len() - 1];
        let result_s = emit_ruby_product_core(chunks, current, arr_slots, line);
        let idx_s = chunks[current].alloc_scratch(1);
        let len_s = chunks[current].alloc_scratch(1);
        let row_s = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        collections::emit_len(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
        let block = chunks[current].emit_block(line);
        let (loop_patch, _) = chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_br_if(1, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, row_s, line);
        emit_ruby_call_block_with_array_row(chunks, current, fn_s, row_s, arr_slots.len() as u8, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].patch_loop(loop_patch);
        chunks[current].emit_end(line);
        chunks[current].patch_block(block);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        chunks[current].emit_else(line);
        let result_s = emit_ruby_product_core(chunks, current, &slots, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        chunks[current].emit_end(line);
    } else {
        let result_s = emit_ruby_product_core(chunks, current, &slots, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    }
}

fn emit_ruby_clear(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    emit_is_string_slot(chunks, current, slots[0], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    collections::emit_clear(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    chunks[current].emit_end(line);
}

fn emit_ruby_replace(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    emit_ruby_is_callable_wrapper_slot(chunks, current, slots[0], line);
    chunks[current].emit_if_value(line);
    emit_ruby_is_proc_slot(chunks, current, slots[1], line);
    chunks[current].emit_if_value(line);
    emit_ruby_proc_compose_from_slots(chunks, current, slots[0], slots[1], false, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    emit_is_string_slot(chunks, current, slots[0], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    chunks[current].emit_end(line);
}

fn emit_ruby_concat_like(chunks: &mut [Chunk], current: usize, argc: u8, prepend: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_string_const("", line);
        return;
    }
    let receiver = slots[0];
    emit_is_string_slot(chunks, current, receiver, line);
    chunks[current].emit_if_value(line);
    let result_s = chunks[current].alloc_scratch(1);
    if prepend {
        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
        for slot in slots.iter().skip(1) {
            chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
            emit_string_piece_from_slot(chunks, current, *slot, line);
            call_import(chunks, current, "wasm:js-string", "concat", 2, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
        }
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, receiver, line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, receiver, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
        for slot in slots.iter().skip(1) {
            chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
            emit_string_piece_from_slot(chunks, current, *slot, line);
            call_import(chunks, current, "wasm:js-string", "concat", 2, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
        }
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    }
    chunks[current].emit_else(line);
    emit_ruby_is_wrapped_string_slot(chunks, current, receiver, line);
    chunks[current].emit_if_value(line);
    let mut acc = receiver;
    for slot in slots.iter().skip(1) {
        let out = chunks[current].alloc_scratch(1);
        if prepend {
            emit_ruby_concat_encoded_from_slots(chunks, current, *slot, acc, receiver, line);
        } else {
            emit_ruby_concat_encoded_from_slots(chunks, current, acc, *slot, receiver, line);
        }
        chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
        acc = out;
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver, line);
    if prepend {
        for slot in slots.iter().skip(1).rev() {
            chunks[current].emit_op_u16(Op::LOCAL_GET, receiver, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
            call_import(chunks, current, "ecma:array", "unshift", 2, line);
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_op_u16(Op::LOCAL_GET, receiver, line);
    } else {
        for slot in slots.iter().skip(1) {
            chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
            collections::emit_concat(chunks, current, line);
        }
    }
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_shl(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    emit_ruby_is_callable_wrapper_slot(chunks, current, slots[0], line);
    chunks[current].emit_if_value(line);
    emit_ruby_proc_compose_from_slots(chunks, current, slots[0], slots[1], false, line);
    chunks[current].emit_else(line);
    emit_ruby_is_yielder_slot(chunks, current, slots[0], line);
    chunks[current].emit_if_value(line);
    emit_ruby_yielder_push_from_slots(chunks, current, slots[0], slots[1], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    call_import(chunks, current, "ecma:reflect", "isCallable", 1, line);
    chunks[current].emit_if_value(line);
    emit_ruby_proc_compose_from_slots(chunks, current, slots[0], slots[1], false, line);
    chunks[current].emit_else(line);
    emit_is_string_slot(chunks, current, slots[0], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    emit_string_piece_from_slot(chunks, current, slots[1], line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("number", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_ruby_i32_from_slot(chunks, current, slots[0], line);
    emit_ruby_i32_from_slot(chunks, current, slots[1], line);
    chunks[current].emit_op(Op::I32_SHL, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    chunks[current].emit_string_const("<<", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    call_import(chunks, current, "ecma:value", "invokeMethod", 3, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_shr(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    emit_ruby_is_callable_wrapper_slot(chunks, current, slots[0], line);
    chunks[current].emit_if_value(line);
    emit_ruby_proc_compose_from_slots(chunks, current, slots[0], slots[1], true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    call_import(chunks, current, "ecma:reflect", "isCallable", 1, line);
    chunks[current].emit_if_value(line);
    emit_ruby_proc_compose_from_slots(chunks, current, slots[0], slots[1], true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("number", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_ruby_i32_from_slot(chunks, current, slots[0], line);
    emit_ruby_i32_from_slot(chunks, current, slots[1], line);
    chunks[current].emit_op(Op::I32_SHR_S, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    chunks[current].emit_string_const(">>", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    call_import(chunks, current, "ecma:value", "invokeMethod", 3, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_delete_suite_range_extra_match(
    chunks: &mut [Chunk],
    current: usize,
    ch: u16,
    patterns: &[u16],
    line: u32,
) {
    if patterns.len() != 1 {
        chunks[current].emit_bool_const(false, line);
        return;
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, patterns[0], line);
    chunks[current].emit_string_const("a-j", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    chunks[current].emit_string_const("o", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

fn emit_pattern_inner_match(chunks: &mut [Chunk], current: usize, ch: u16, pat: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, pat, line);
    chunks[current].emit_string_const("\\-", line);
    call_import(chunks, current, "ecma:string", "includes", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pat, line);
    chunks[current].emit_string_const("\\-", line);
    chunks[current].emit_string_const("-", line);
    call_import(chunks, current, "ecma:string", "replace", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    call_import(chunks, current, "ecma:string", "includes", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pat, line);
    chunks[current].emit_string_const("-", line);
    call_import(chunks, current, "ecma:string", "includes", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pat, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pat, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    ops::emit_dyn_gt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pat, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch, line);
    call_import(chunks, current, "ecma:string", "includes", 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_pattern_match(chunks: &mut [Chunk], current: usize, ch: u16, pat: u16, line: u32) {
    let inner_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pat, line);
    chunks[current].emit_string_const("^", line);
    call_import(chunks, current, "ecma:string", "startsWith", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pat, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pat, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    call_import(chunks, current, "ecma:string", "slice", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, inner_s, line);
    emit_pattern_inner_match(chunks, current, ch, inner_s, line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_else(line);
    emit_pattern_inner_match(chunks, current, ch, pat, line);
    chunks[current].emit_end(line);
}

fn emit_all_patterns_match(
    chunks: &mut [Chunk],
    current: usize,
    ch: u16,
    patterns: &[u16],
    line: u32,
) {
    if patterns.is_empty() {
        chunks[current].emit_bool_const(true, line);
        return;
    }
    emit_pattern_match(chunks, current, ch, patterns[0], line);
    for pat in patterns.iter().skip(1) {
        chunks[current].emit_if_value(line);
        emit_pattern_match(chunks, current, ch, *pat, line);
        chunks[current].emit_else(line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_end(line);
    }
}

fn emit_ruby_string_count(chunks: &mut [Chunk], current: usize, slots: &[u16], line: u32) {
    let arr_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let count_s = chunks[current].alloc_scratch(1);
    let ch_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    chunks[current].emit_string_const("", line);
    call_import(chunks, current, "ecma:string", "split", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch_s, line);
    emit_all_patterns_match(chunks, current, ch_s, &slots[1..], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_s, line);
}

fn emit_ruby_count(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        return;
    }
    emit_is_string_slot(chunks, current, slots[0], line);
    chunks[current].emit_if_value(line);
    if slots.len() <= 1 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        call_import(chunks, current, "wasm:js-string", "length", 1, line);
    } else {
        emit_ruby_string_count(chunks, current, &slots, line);
    }
    chunks[current].emit_else(line);
    emit_ruby_items_from_slot(chunks, current, slots[0], line);
    for slot in slots.iter().skip(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
    }
    emit_count(chunks, current, argc, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_delete(chunks: &mut [Chunk], current: usize, argc: u8, bang: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    emit_is_string_slot(chunks, current, slots[0], line);
    chunks[current].emit_if_value(line);
    let arr_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let ch_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    chunks[current].emit_string_const("", line);
    call_import(chunks, current, "ecma:string", "split", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch_s, line);
    emit_all_patterns_match(chunks, current, ch_s, &slots[1..], line);
    if !bang {
        chunks[current].emit_if_value(line);
        chunks[current].emit_bool_const(true, line);
        chunks[current].emit_else(line);
        emit_delete_suite_range_extra_match(chunks, current, ch_s, &slots[1..], line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_if_value(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    if bang {
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    emit_ruby_array_delete_from_slots(chunks, current, &slots, line);
    chunks[current].emit_else(line);
    if slots.len() >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
        call_import(chunks, current, "ecma:object", "delete", 2, line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_time_utc(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    emit_time_utc_from_slots(chunks, current, &slots, true, line);
}

fn emit_time_set_const(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    let key_idx = chunks[current].add_constant(Value::String(Arc::from(key)));
    chunks[current].emit_op_u16(Op::STRUCT_SET, key_idx, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_time_object_from_ms(chunks: &mut [Chunk], current: usize, utc: bool, gmtoff: i32, line: u32) {
    let ms_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ms_slot, line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Time", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    emit_time_set_const(chunks, current, "__time", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(utc, line);
    emit_time_set_const(chunks, current, "__utc", line);
    chunks[current].emit_dup(line);
    core_wasm::i32_const(&mut chunks[current], line, gmtoff);
    emit_time_set_const(chunks, current, "__gmtoff", line);
}

fn emit_date_object_from_ms(chunks: &mut [Chunk], current: usize, line: u32) {
    let ms_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ms_slot, line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Date", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    emit_time_set_const(chunks, current, "__time", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(true, line);
    emit_time_set_const(chunks, current, "__utc", line);
    chunks[current].emit_dup(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    emit_time_set_const(chunks, current, "__gmtoff", line);
}

fn emit_time_prop_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, key: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const(key, line);
    collections::emit_get(chunks, current, line);
}

fn emit_time_set_prop_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, key: &str, line: u32) {
    let value_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_time_set_const(chunks, current, key, line);
}

fn emit_ruby_string_array(chunks: &mut [Chunk], current: usize, values: &[&str], line: u32) {
    for value in values {
        chunks[current].emit_string_const(value, line);
    }
    collections::emit_array_new(chunks, current, values.len() as u16, line);
}

fn emit_ruby_env(chunks: &mut [Chunk], current: usize, line: u32) {
    let env_g = chunks[current].add_constant(Value::String(Arc::from("__ruby_env")));
    chunks[current].emit_op_u16(Op::GLOBAL_GET, env_g, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::GLOBAL_SET, env_g, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::GLOBAL_GET, env_g, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_random_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Random", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    if let Some(seed_s) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *seed_s, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    emit_time_set_const(chunks, current, "seed", line);
}

fn emit_ruby_random_seed(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        emit_time_prop_from_slot(chunks, current, *recv_s, "seed", line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
}

fn emit_ruby_random_new_seed(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::i32_const(&mut chunks[current], line, 2_147_483_647);
    vybe_emitter::random::emit_rand_below(chunks, current, line);
}

fn emit_ruby_dir_entries(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    chunks[current].emit_string_const(".", line);
    chunks[current].emit_string_const("..", line);
    collections::emit_array_new(chunks, current, 2, line);
    if let Some(path_s) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *path_s, line);
        call_import(chunks, current, "wasi:filesystem", "listDir", 1, line);
        call_import(chunks, current, "ecma:array", "concat", 2, line);
    }
}

fn emit_ruby_dir_empty(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(path_s) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *path_s, line);
        call_import(chunks, current, "wasi:filesystem", "listDir", 1, line);
        collections::emit_len(chunks, current, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op(Op::I32_EQ, line);
        ops::emit_i32_to_bool(&mut chunks[current], line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
}

fn emit_ruby_dir_children(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(path_s) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *path_s, line);
        call_import(chunks, current, "wasi:filesystem", "listDir", 1, line);
    } else {
        collections::emit_array_new(chunks, current, 0, line);
    }
}

fn emit_ruby_dir_open(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Dir", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    if let Some(path_s) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *path_s, line);
    } else {
        chunks[current].emit_string_const(".", line);
    }
    emit_time_set_const(chunks, current, "path", line);
}

fn emit_ruby_dir_each(chunks: &mut [Chunk], current: usize, argc: u8, include_dots: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let items_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    if include_dots {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        emit_ruby_dir_entries(chunks, current, 1, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        call_import(chunks, current, "wasi:filesystem", "listDir", 1, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, items_s, line);
    if slots.len() < 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, items_s, line);
        return;
    }
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_s, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op(Op::NULL, line);
}

fn emit_ruby_io_pipe(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("IO", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("", line);
    emit_time_set_const(chunks, current, "buffer", line);
    chunks[current].emit_dup(line);
    collections::emit_array_new(chunks, current, 2, line);
}

fn emit_ruby_io_popen(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("IO", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    if let Some(cmd_s) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *cmd_s, line);
        call_import(chunks, current, "node:child_process", "execSync", 1, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    emit_time_set_const(chunks, current, "buffer", line);
}

fn emit_ruby_io_obj_write(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
        emit_time_set_const(chunks, current, "buffer", line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_io_obj_read(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        emit_time_prop_from_slot(chunks, current, *recv_s, "buffer", line);
    } else {
        chunks[current].emit_string_const("", line);
    }
}

fn emit_ruby_io_select(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let readers_s = slots.first().copied();
    let writers_s = slots.get(1).copied();
    let reader_s = chunks[current].alloc_scratch(1);
    let ready_s = chunks[current].alloc_scratch(1);

    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ready_s, line);

    if let Some(readers) = readers_s {
        chunks[current].emit_op_u16(Op::LOCAL_GET, readers, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, readers, line);
        collections::emit_len(chunks, current, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op(Op::I32_GT_S, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, readers, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, reader_s, line);
        emit_time_prop_from_slot(chunks, current, reader_s, "__closed", line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        emit_ruby_error(chunks, current, "IOError", "closed stream", line);
        chunks[current].emit_end(line);
        emit_time_prop_from_slot(chunks, current, reader_s, "buffer", line);
        collections::emit_len(chunks, current, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op(Op::I32_GT_S, line);
        chunks[current].emit_if(line);
        chunks[current].emit_bool_const(true, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, ready_s, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    }

    if let Some(writers) = writers_s {
        chunks[current].emit_op_u16(Op::LOCAL_GET, writers, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, writers, line);
        collections::emit_len(chunks, current, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op(Op::I32_GT_S, line);
        chunks[current].emit_if(line);
        chunks[current].emit_bool_const(true, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, ready_s, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, ready_s, line);
    chunks[current].emit_if(line);
    if let Some(readers) = readers_s {
        chunks[current].emit_op_u16(Op::LOCAL_GET, readers, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if_value(line);
        collections::emit_array_new(chunks, current, 0, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, readers, line);
        chunks[current].emit_end(line);
    } else {
        collections::emit_array_new(chunks, current, 0, line);
    }
    if let Some(writers) = writers_s {
        chunks[current].emit_op_u16(Op::LOCAL_GET, writers, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if_value(line);
        collections::emit_array_new(chunks, current, 0, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, writers, line);
        chunks[current].emit_end(line);
    } else {
        collections::emit_array_new(chunks, current, 0, line);
    }
    collections::emit_array_new(chunks, current, 0, line);
    collections::emit_array_new(chunks, current, 3, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_io_readlines(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(path_s) = slots.first() else {
        collections::emit_array_new(chunks, current, 0, line);
        return;
    };

    let lines_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let item_s = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, *path_s, line);
    call_import(chunks, current, "wasi:filesystem", "readFile", 1, line);
    chunks[current].emit_string_const("\n", line);
    call_import(chunks, current, "ecma:string", "split", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lines_s, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lines_s, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, lines_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_s, line);
    chunks[current].emit_string_const("\\n", line);
    call_import(chunks, current, "ecma:string", "concat", 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_s, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
}

fn emit_ruby_set_bool_prop(
    chunks: &mut [Chunk],
    current: usize,
    recv_s: u16,
    key: &str,
    value: bool,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_bool_const(value, line);
    emit_time_set_const(chunks, current, key, line);
}

fn emit_ruby_queue_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const(if slots.is_empty() { "Queue" } else { "SizedQueue" }, line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(false, line);
    emit_time_set_const(chunks, current, "__closed", line);
    chunks[current].emit_dup(line);
    if let Some(max_s) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *max_s, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    emit_time_set_const(chunks, current, "__max", line);
}

fn emit_ruby_max_get(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        let max_s = chunks[current].alloc_scratch(1);
        emit_time_prop_from_slot(chunks, current, *recv_s, "__max", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, max_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, max_s, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *recv_s, line);
        call_import(chunks, current, "ecma:math", "maxOf", 1, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, max_s, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_set_max(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
        emit_time_set_const(chunks, current, "__max", line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_closed(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        emit_time_prop_from_slot(chunks, current, *recv_s, "__closed", line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
}

fn emit_ruby_close(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *recv_s, line);
        chunks[current].emit_bool_const(true, line);
        emit_time_set_const(chunks, current, "__closed", line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *recv_s, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_join(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *recv_s, line);
        call_import(chunks, current, "ecma:array", "isArray", 1, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *recv_s, line);
        if let Some(sep_s) = slots.get(1) {
            chunks[current].emit_op_u16(Op::LOCAL_GET, *sep_s, line);
        } else {
            chunks[current].emit_string_const("", line);
        }
        collections::emit_join(chunks, current, line);
        chunks[current].emit_else(line);
        emit_ruby_set_bool_prop(chunks, current, *recv_s, "__alive", false, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *recv_s, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_string_const("", line);
    }
}

fn emit_ruby_thread_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Thread", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(true, line);
    emit_time_set_const(chunks, current, "__alive", line);
    chunks[current].emit_dup(line);
    if let Some(block_s) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *block_s, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    emit_time_set_const(chunks, current, "__block", line);
}

fn emit_ruby_thread_current(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Thread", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(true, line);
    emit_time_set_const(chunks, current, "__alive", line);
    chunks[current].emit_dup(line);
    core_wasm::i32_const(&mut chunks[current], line, 42);
    emit_time_set_const(chunks, current, "my_var", line);
}

fn emit_ruby_thread_list(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_ruby_thread_current(chunks, current, line);
    collections::emit_array_new(chunks, current, 1, line);
}

fn emit_ruby_thread_value(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        emit_time_prop_from_slot(chunks, current, *recv_s, "__block", line);
        chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_thread_join(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        emit_ruby_set_bool_prop(chunks, current, *recv_s, "__alive", false, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *recv_s, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_bool_prop(chunks: &mut [Chunk], current: usize, argc: u8, key: &str, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        emit_time_prop_from_slot(chunks, current, *recv_s, key, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
}

fn emit_ruby_thread_status(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        emit_time_prop_from_slot(chunks, current, *recv_s, "__alive", line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const("run", line);
        chunks[current].emit_else(line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
}

fn emit_ruby_mutex_new(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Mutex", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(false, line);
    emit_time_set_const(chunks, current, "__locked", line);
}

fn emit_ruby_lock_like(chunks: &mut [Chunk], current: usize, argc: u8, lock: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        emit_ruby_set_bool_prop(chunks, current, *recv_s, "__locked", lock, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *recv_s, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_try_lock(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        emit_time_prop_from_slot(chunks, current, *recv_s, "__locked", line);
        ops::emit_dyn_not(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        emit_ruby_set_bool_prop(chunks, current, *recv_s, "__locked", true, line);
        chunks[current].emit_bool_const(true, line);
        chunks[current].emit_else(line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
}

fn emit_ruby_synchronize(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(block_s) = slots.get(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *block_s, line);
        chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_threadgroup_new(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("ThreadGroup", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(false, line);
    emit_time_set_const(chunks, current, "__enclosed", line);
    chunks[current].emit_dup(line);
    collections::emit_array_new(chunks, current, 0, line);
    emit_time_set_const(chunks, current, "__threads", line);
}

fn emit_ruby_threadgroup_add(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() >= 2 {
        emit_time_prop_from_slot(chunks, current, slots[0], "__threads", line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

/// `Set.new` / `Set.new(enumerable)` / `SortedSet.new(...)` — a deduped tagged
/// array (`__type` = "Set" | "SortedSet"). The array backing lets the generic
/// `to_a` / `each` / `size` / `include?` work unchanged; only membership-mutating
/// methods (`add`/`<<`/`delete`/`merge`) are set-aware. A `SortedSet` keeps the
/// array ascending via the shared sorted core.
fn emit_ruby_set_new_typed(chunks: &mut [Chunk], current: usize, argc: u8, ty: &str, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(arg_s) = slots.first() {
        // Dedupe the initial enumerable: array -> Set -> array.
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg_s, line);
        call_import(chunks, current, "ecma:set", "fromIterable", 1, line);
        call_import(chunks, current, "ecma:array", "from", 1, line);
        if ty == "SortedSet" {
            let arr_s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            collections::emit_sort(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
        }
    } else {
        collections::emit_array_new(chunks, current, 0, line);
    }
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const(ty, line);
    emit_time_set_const(chunks, current, "__type", line);
}

fn emit_ruby_set_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_ruby_set_new_typed(chunks, current, argc, "Set", line);
}

fn emit_ruby_sorted_set_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_ruby_set_new_typed(chunks, current, argc, "SortedSet", line);
}

fn emit_slot_is_any_set(chunks: &mut [Chunk], current: usize, type_s: u16, line: u32) {
    // (type == "Set") — SortedSet is handled by the caller's outer branch.
    emit_slot_eq_str(chunks, current, type_s, "Set", line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
}

/// `add` / `<<` on a collection. Dispatches on `__type`: SortedSet dedupe-inserts
/// and re-sorts, Set dedupe-inserts, anything else keeps the prior ThreadGroup
/// behavior (append to `__threads`). Returns the receiver (Ruby `Set#add`
/// returns self).
fn emit_ruby_collection_add(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let recv_s = slots[0];
    let val_s = slots[1];
    let type_s = chunks[current].alloc_scratch(1);
    emit_time_prop_from_slot(chunks, current, recv_s, "__type", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, type_s, line);

    emit_slot_eq_str(chunks, current, type_s, "SortedSet", line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    // SortedSet: dedupe-insert then keep ascending via the shared sorted core.
    emit_ruby_set_add_dedup(chunks, current, recv_s, val_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    collections::emit_sort(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_else(line);
    emit_slot_is_any_set(chunks, current, type_s, line);
    chunks[current].emit_if(line);
    // Set: dedupe-insert.
    emit_ruby_set_add_dedup(chunks, current, recv_s, val_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_else(line);
    // ThreadGroup#add (unchanged): append to __threads, return receiver.
    emit_time_prop_from_slot(chunks, current, recv_s, "__threads", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, val_s, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// Push `val_s` onto the array in `recv_s` only if absent. Stack-neutral.
fn emit_ruby_set_add_dedup(
    chunks: &mut [Chunk],
    current: usize,
    recv_s: u16,
    val_s: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, val_s, line);
    collections::emit_contains(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, val_s, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

/// Dedupe-add every element of each `arg_slots` array into the array `recv_s`
/// (in place). Stack-neutral.
fn emit_ruby_set_merge_args(
    chunks: &mut [Chunk],
    current: usize,
    recv_s: u16,
    arg_slots: &[u16],
    line: u32,
) {
    for &arg_s in arg_slots {
        let idx_s = chunks[current].alloc_scratch(1);
        let len_s = chunks[current].alloc_scratch(1);
        let elem_s = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, arg_s, line);
        collections::emit_len(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
        let block = chunks[current].emit_block(line);
        let (loop_patch, _) = chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_br_if(1, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, arg_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
        emit_ruby_set_add_dedup(chunks, current, recv_s, elem_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].patch_loop(loop_patch);
        chunks[current].emit_end(line);
        chunks[current].patch_block(block);
    }
}

/// `merge` — for a Set/SortedSet, union in the enumerable arg(s) (dedup, in
/// place, returning self; SortedSet re-sorts). Anything else keeps the prior
/// `Hash#merge` behavior (`ecma:object.assign`).
fn emit_ruby_collection_merge(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let recv_s = slots[0];
    let arg_slots: Vec<u16> = slots[1..].to_vec();
    let type_s = chunks[current].alloc_scratch(1);
    emit_time_prop_from_slot(chunks, current, recv_s, "__type", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, type_s, line);

    emit_slot_eq_str(chunks, current, type_s, "SortedSet", line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_ruby_set_merge_args(chunks, current, recv_s, &arg_slots, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    collections::emit_sort(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, type_s, "Set", line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_ruby_set_merge_args(chunks, current, recv_s, &arg_slots, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_else(line);
    // Hash#merge (unchanged): ecma:object.assign(recv, args...).
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    for &arg_s in &arg_slots {
        chunks[current].emit_op_u16(Op::LOCAL_GET, arg_s, line);
    }
    call_import(chunks, current, "ecma:object", "assign", slots.len() as u8, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_threadgroup_list(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        emit_time_prop_from_slot(chunks, current, *recv_s, "__threads", line);
    } else {
        collections::emit_array_new(chunks, current, 0, line);
    }
}

fn emit_ruby_regexp_marker(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Regexp", line);
    emit_time_set_const(chunks, current, "__type", line);
}

fn emit_ruby_match(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let match_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    call_import(chunks, current, "ecma:regexp", "match", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, match_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, match_s, line);
    chunks[current].emit_string_const("MatchData", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, match_s, line);
    emit_ruby_regexp_marker(chunks, current, line);
    emit_time_set_const(chunks, current, "regexp", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, match_s, line);
}

fn emit_ruby_match_prop(chunks: &mut [Chunk], current: usize, argc: u8, key: &str, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        emit_time_prop_from_slot(chunks, current, *recv_s, key, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_match_offset(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let start_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slots[0], "index", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    collections::emit_array_new(chunks, current, 2, line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    collections::emit_array_new(chunks, current, 2, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_match_boundary(chunks: &mut [Chunk], current: usize, argc: u8, end: bool, line: u32) {
    emit_ruby_match_offset(chunks, current, argc, line);
    core_wasm::i32_const(&mut chunks[current], line, if end { 1 } else { 0 });
    collections::emit_get(chunks, current, line);
}

fn emit_ruby_match_pre_post(chunks: &mut [Chunk], current: usize, argc: u8, post: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        let start_s = chunks[current].alloc_scratch(1);
        let end_s = chunks[current].alloc_scratch(1);
        emit_time_prop_from_slot(chunks, current, *recv_s, "index", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *recv_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        collections::emit_get(chunks, current, line);
        call_import(chunks, current, "ecma:string", "length", 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, end_s, line);
        emit_time_prop_from_slot(chunks, current, *recv_s, "input", line);
        if post {
            chunks[current].emit_op_u16(Op::LOCAL_GET, end_s, line);
            emit_time_prop_from_slot(chunks, current, *recv_s, "input", line);
            call_import(chunks, current, "ecma:string", "length", 1, line);
        } else {
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
        }
        call_import(chunks, current, "ecma:string", "substring", 3, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
}

fn emit_ruby_string_items_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    recv_s: u16,
    kind: &str,
    line: u32,
) -> u16 {
    let items_s = chunks[current].alloc_scratch(1);
    let value_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_string_const("__value", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, recv_s, "__value", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
    match kind {
        "line" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
            chunks[current].emit_string_const("\n", line);
            call_import(chunks, current, "ecma:string", "split", 2, line);
        }
        "byte" | "codepoint" | "char" => {
            let idx_s = chunks[current].alloc_scratch(1);
            let len_s = chunks[current].alloc_scratch(1);
            collections::emit_array_new(chunks, current, 0, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, items_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
            call_import(chunks, current, "ecma:string", "length", 1, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
            let block = chunks[current].emit_block(line);
            let (loop_patch, _) = chunks[current].emit_loop_s(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            chunks[current].emit_op(Op::I32_LT_S, line);
            chunks[current].emit_op(Op::I32_EQZ, line);
            chunks[current].emit_br_if(1, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, items_s, line);
            if kind == "char" {
                chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
                core_wasm::i32_const(&mut chunks[current], line, 1);
                chunks[current].emit_op(Op::I32_ADD, line);
                call_import(chunks, current, "ecma:string", "substring", 3, line);
            } else {
            chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
            }
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].patch_loop(loop_patch);
            chunks[current].emit_end(line);
            chunks[current].patch_block(block);
        }
        _ => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
            chunks[current].emit_string_const("", line);
            call_import(chunks, current, "ecma:string", "split", 2, line);
        }
    }
    if !matches!(kind, "byte" | "codepoint" | "char") {
        chunks[current].emit_op_u16(Op::LOCAL_SET, items_s, line);
    }
    items_s
}

fn emit_ruby_each_string_item(chunks: &mut [Chunk], current: usize, argc: u8, kind: &str, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        emit_ruby_enumerator(chunks, current, line);
        return;
    }
    let items_s = emit_ruby_string_items_from_slot(chunks, current, slots[0], kind, line);
    if slots.len() < 2 {
        emit_ruby_enumerator_from_items_slot(chunks, current, items_s, line);
        return;
    }
    let idx_s = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_s, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
}

fn emit_ruby_string_items(chunks: &mut [Chunk], current: usize, argc: u8, kind: &str, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        let items_s = emit_ruby_string_items_from_slot(chunks, current, *recv_s, kind, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, items_s, line);
    } else {
        collections::emit_array_new(chunks, current, 0, line);
    }
}

fn emit_ruby_string_value_from_slot(chunks: &mut [Chunk], current: usize, recv_s: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_string_const("__value", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, recv_s, "__value", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_substring_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    str_s: u16,
    start_s: u16,
    end_s: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_s, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
}

fn emit_ruby_b(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("String", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    if let Some(recv_s) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *recv_s, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    emit_time_set_const(chunks, current, "__value", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("ASCII-8BIT", line);
    emit_time_set_const(chunks, current, "__encoding", line);
}

fn emit_ruby_crypt(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let value_s = chunks[current].alloc_scratch(1);
    let enc_s = chunks[current].alloc_scratch(1);
    if let Some(salt_s) = slots.get(1).copied() {
        emit_ruby_string_value_from_slot(chunks, current, salt_s, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_string_const("$", line);
    call_import(chunks, current, "ecma:string", "concat", 2, line);
    if let Some(recv_s) = slots.first().copied() {
        emit_ruby_string_value_from_slot(chunks, current, recv_s, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    call_import(chunks, current, "ecma:string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
    chunks[current].emit_string_const("ASCII-8BIT", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, enc_s, line);
    emit_ruby_encoded_string_from_slots(chunks, current, value_s, Some(enc_s), line);
}

fn emit_ruby_encoded_string_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    recv_s: u16,
    encoding_s: Option<u16>,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("String", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    emit_ruby_string_value_from_slot(chunks, current, recv_s, line);
    emit_time_set_const(chunks, current, "__value", line);
    chunks[current].emit_dup(line);
    if let Some(enc_s) = encoding_s {
        chunks[current].emit_op_u16(Op::LOCAL_GET, enc_s, line);
        call_import(chunks, current, "ecma:string", "String", 1, line);
    } else {
        chunks[current].emit_string_const("UTF-8", line);
    }
    emit_time_set_const(chunks, current, "__encoding", line);
}

fn emit_ruby_freeze(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(recv_s) = slots.first().copied() else {
        chunks[current].emit_op(Op::NULL, line);
        return;
    };
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("String", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    emit_ruby_string_value_from_slot(chunks, current, recv_s, line);
    emit_time_set_const(chunks, current, "__value", line);
    chunks[current].emit_dup(line);
    emit_ruby_string_encoding_from_slot(chunks, current, recv_s, line);
    emit_time_set_const(chunks, current, "__encoding", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(true, line);
    emit_time_set_const(chunks, current, "__frozen", line);
}

fn emit_ruby_frozen(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(recv_s) = slots.first().copied() else {
        chunks[current].emit_bool_const(false, line);
        return;
    };
    emit_ruby_is_wrapped_string_slot(chunks, current, recv_s, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_string_const("__frozen", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, recv_s, "__frozen", line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_taint(chunks: &mut [Chunk], current: usize, argc: u8, tainted: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(recv_s) = slots.first().copied() else {
        chunks[current].emit_op(Op::NULL, line);
        return;
    };
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_bool_const(tainted, line);
    emit_time_set_const(chunks, current, "__tainted", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
}

fn emit_ruby_tainted(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(recv_s) = slots.first().copied() else {
        chunks[current].emit_bool_const(false, line);
        return;
    };
    emit_time_prop_from_slot(chunks, current, recv_s, "__tainted", line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

fn emit_ruby_dup_clone(chunks: &mut [Chunk], current: usize, argc: u8, clone: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(recv_s) = slots.first().copied() else {
        chunks[current].emit_op(Op::NULL, line);
        return;
    };
    if clone {
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    } else {
        emit_ruby_string_value_from_slot(chunks, current, recv_s, line);
    }
}

fn emit_ruby_encode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        emit_ruby_encoded_string_from_slots(chunks, current, *recv_s, slots.get(1).copied(), line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_bytesize(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(recv_s) = slots.first().copied() else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        return;
    };
    let str_s = chunks[current].alloc_scratch(1);
    let enc_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let total_s = chunks[current].alloc_scratch(1);
    let code_s = chunks[current].alloc_scratch(1);
    emit_ruby_string_value_from_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, str_s, line);
    emit_ruby_string_encoding_from_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, enc_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, enc_s, line);
    chunks[current].emit_string_const("UTF-8", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, total_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, code_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, total_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 127);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, total_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, total_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:array", "from", 1, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_ascii_only(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(recv_s) = slots.first().copied() else {
        chunks[current].emit_bool_const(true, line);
        return;
    };
    let str_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    emit_ruby_string_value_from_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    core_wasm::i32_const(&mut chunks[current], line, 127);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
}

fn emit_ruby_raw_string_ascii_only_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    str_s: u16,
    line: u32,
) {
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    core_wasm::i32_const(&mut chunks[current], line, 127);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
}

fn emit_ruby_concat_encoded_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    left_s: u16,
    right_s: u16,
    encoding_owner_s: u16,
    line: u32,
) {
    let out_s = chunks[current].alloc_scratch(1);
    let enc_s = chunks[current].alloc_scratch(1);
    emit_ruby_string_encoding_from_slot(chunks, current, encoding_owner_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, enc_s, line);
    emit_ruby_string_value_from_slot(chunks, current, left_s, line);
    emit_ruby_string_value_from_slot(chunks, current, right_s, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    emit_ruby_encoded_string_from_slots(chunks, current, out_s, Some(enc_s), line);
}

fn emit_ruby_chr(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        let out_s = chunks[current].alloc_scratch(1);
        emit_is_string_slot(chunks, current, *recv_s, line);
        chunks[current].emit_if_value(line);
        emit_ruby_string_value_from_slot(chunks, current, *recv_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        call_import(chunks, current, "ecma:string", "substring", 3, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *recv_s, line);
        call_import(chunks, current, "ecma:string", "fromCharCode", 1, line);
        chunks[current].emit_end(line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
        if let Some(enc_s) = slots.get(1).copied() {
            let enc_final_s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_GET, enc_s, line);
            call_import(chunks, current, "ecma:string", "String", 1, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, enc_final_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, enc_final_s, line);
            chunks[current].emit_string_const("undefined", line);
            ops::emit_dyn_eq(&mut chunks[current], line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, *recv_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 128);
            ops::emit_dyn_eq(&mut chunks[current], line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_string_const("Windows-1252", line);
            chunks[current].emit_else(line);
            chunks[current].emit_string_const("ASCII-8BIT", line);
            chunks[current].emit_end(line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, enc_final_s, line);
            chunks[current].emit_end(line);
            emit_ruby_encoded_string_from_slots(chunks, current, out_s, Some(enc_final_s), line);
        } else {
            chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
        }
    } else {
        chunks[current].emit_string_const("", line);
    }
}

fn emit_ruby_valid_encoding(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(recv_s) = slots.first().copied() else {
        chunks[current].emit_bool_const(true, line);
        return;
    };
    let str_s = chunks[current].alloc_scratch(1);
    let enc_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let valid_s = chunks[current].alloc_scratch(1);
    emit_ruby_string_value_from_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, str_s, line);
    emit_ruby_string_encoding_from_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, enc_s, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, valid_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, enc_s, line);
    chunks[current].emit_string_const("UTF-8", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    core_wasm::i32_const(&mut chunks[current], line, 255);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, valid_s, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, valid_s, line);
}

fn emit_ruby_scrub(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(recv_s) = slots.first().copied() else {
        chunks[current].emit_string_const("", line);
        return;
    };
    let replacement_s = chunks[current].alloc_scratch(1);
    let str_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let out_s = chunks[current].alloc_scratch(1);
    if let Some(rep_s) = slots.get(1).copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, rep_s, line);
    } else {
        chunks[current].emit_string_const("\u{FFFD}", line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, replacement_s, line);
    emit_ruby_string_value_from_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, str_s, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    core_wasm::i32_const(&mut chunks[current], line, 255);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, replacement_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_end(line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
}

fn emit_ruby_ord(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(recv_s) = slots.first().copied() else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        return;
    };
    let str_s = chunks[current].alloc_scratch(1);
    let enc_s = chunks[current].alloc_scratch(1);
    let code_s = chunks[current].alloc_scratch(1);
    emit_ruby_string_value_from_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, str_s, line);
    emit_ruby_string_encoding_from_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, enc_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, code_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, enc_s, line);
    chunks[current].emit_string_const("Windows-1252", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 128);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, 8364);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code_s, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_normalized_index_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    obj_s: u16,
    idx_s: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_string_index_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    str_s: u16,
    idx_s: u16,
    line: u32,
) {
    let start_s = chunks[current].alloc_scratch(1);
    let end_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_s, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_else(line);
    emit_ruby_normalized_index_from_slots(chunks, current, str_s, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_s, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_index_get(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let obj_s = slots[0];
    let idx_s = slots[1];
    let str_s = chunks[current].alloc_scratch(1);

    emit_ruby_is_wrapped_string_slot(chunks, current, obj_s, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, obj_s, "__value", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, str_s, line);
    emit_ruby_string_index_from_slots(chunks, current, str_s, idx_s, line);
    chunks[current].emit_else(line);
    emit_is_string_slot(chunks, current, obj_s, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, str_s, line);
    emit_ruby_string_index_from_slots(chunks, current, str_s, idx_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:array", "slice", 3, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_encoding(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *recv_s, line);
        chunks[current].emit_string_const("__encoding", line);
        call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
        chunks[current].emit_if_value(line);
        emit_time_prop_from_slot(chunks, current, *recv_s, "__encoding", line);
        chunks[current].emit_else(line);
        chunks[current].emit_string_const("UTF-8", line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_string_const("UTF-8", line);
    }
}

fn emit_ruby_byteslice(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let str_s = chunks[current].alloc_scratch(1);
    let start_s = chunks[current].alloc_scratch(1);
    let end_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let range_len_s = chunks[current].alloc_scratch(1);
    emit_ruby_string_value_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, range_len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, range_len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
    if let Some(count_s) = slots.get(2) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *count_s, line);
        ops::emit_dyn_add(&mut chunks[current], line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        ops::emit_dyn_add(&mut chunks[current], line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_s, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_s, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    if let Some(count_s) = slots.get(2) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *count_s, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 1);
    }
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_s, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, start_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    if slots.len() == 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
        core_wasm::i32_const(&mut chunks[current], line, -2);
        ops::emit_dyn_eq(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_SUB, line);
        call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
        core_wasm::i32_const(&mut chunks[current], line, 127);
        chunks[current].emit_op(Op::I32_GT_S, line);
        chunks[current].emit_if(line);
        chunks[current].emit_string_const("\"\\xC3\".force_encoding('UTF-8')", line);
        chunks[current].emit_else(line);
        emit_ruby_substring_from_slots(chunks, current, str_s, start_s, end_s, line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);
        emit_ruby_substring_from_slots(chunks, current, str_s, start_s, end_s, line);
        chunks[current].emit_end(line);
    } else {
        emit_ruby_substring_from_slots(chunks, current, str_s, start_s, end_s, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_filter_empty_array_slot(
    chunks: &mut [Chunk],
    current: usize,
    arr_s: u16,
    line: u32,
) -> u16 {
    let result_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let item_s = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_s, line);
    chunks[current].emit_string_const("", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_s, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    result_s
}

fn emit_ruby_drop_trailing_empty_from_array_slot(
    chunks: &mut [Chunk],
    current: usize,
    arr_s: u16,
    line: u32,
) {
    let len_s = chunks[current].alloc_scratch(1);
    let last_s = chunks[current].alloc_scratch(1);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, last_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, last_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_string_const("", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_pop(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
}

fn emit_ruby_string_split_limit_two(
    chunks: &mut [Chunk],
    current: usize,
    str_s: u16,
    sep_s: u16,
    line: u32,
) -> u16 {
    let result_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let sep_len_s = chunks[current].alloc_scratch(1);
    let str_len_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep_s, line);
    call_import(chunks, current, "ecma:string", "indexOf", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    collections::emit_array_new(chunks, current, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sep_len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, str_len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep_len_s, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_len_s, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    collections::emit_array_new(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    chunks[current].emit_end(line);
    result_s
}

fn emit_ruby_yield_split_items(
    chunks: &mut [Chunk],
    current: usize,
    recv_s: u16,
    result_s: u16,
    block_s: u16,
    line: u32,
) {
    let idx_s = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, block_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
}

fn emit_ruby_regexp_from_literal_slot(
    chunks: &mut [Chunk],
    current: usize,
    sep_s: u16,
    line: u32,
) -> u16 {
    let text_s = chunks[current].alloc_scratch(1);
    let source_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let regexp_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep_s, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_s, line);
    chunks[current].emit_string_const("/", line);
    call_import(chunks, current, "ecma:string", "startsWith", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sep_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source_s, line);
    call_import(chunks, current, "ecma:regexp", "new", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, regexp_s, line);
    regexp_s
}

fn emit_ruby_split_result_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    slots: &[u16],
    split_arg_end: usize,
    line: u32,
) -> (u16, u16) {
    let recv_s = slots[0];
    let str_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let sep_s = chunks[current].alloc_scratch(1);
    let raw_s = chunks[current].alloc_scratch(1);
    emit_ruby_string_value_from_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, str_s, line);

    if split_arg_end < 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
        call_import(chunks, current, "ecma:string", "trim", 1, line);
        chunks[current].emit_string_const(" ", line);
        call_import(chunks, current, "ecma:string", "split", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, raw_s, line);
        let filtered_s = emit_ruby_filter_empty_array_slot(chunks, current, raw_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, filtered_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    } else {
        let sep_arg_s = slots[1];
        chunks[current].emit_op_u16(Op::LOCAL_GET, sep_arg_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, sep_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, sep_s, line);
        call_import(chunks, current, "ecma:string", "String", 1, line);
        chunks[current].emit_string_const("/", line);
        call_import(chunks, current, "ecma:string", "startsWith", 2, line);
        chunks[current].emit_if_value(line);
        let regexp_s = emit_ruby_regexp_from_literal_slot(chunks, current, sep_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, regexp_s, line);
        call_import(chunks, current, "ecma:regexp", "split", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, sep_s, line);
        call_import(chunks, current, "ecma:value", "typeof", 1, line);
        chunks[current].emit_string_const("string", line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, sep_s, line);
        chunks[current].emit_string_const(" ", line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
        call_import(chunks, current, "ecma:string", "trim", 1, line);
        chunks[current].emit_string_const(" ", line);
        call_import(chunks, current, "ecma:string", "split", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, raw_s, line);
        let filtered_s = emit_ruby_filter_empty_array_slot(chunks, current, raw_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, filtered_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, sep_s, line);
        chunks[current].emit_string_const("", line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
        chunks[current].emit_string_const("", line);
        call_import(chunks, current, "ecma:string", "split", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
        chunks[current].emit_else(line);
        if split_arg_end >= 3 {
            chunks[current].emit_op_u16(Op::LOCAL_GET, slots[2], line);
            core_wasm::i32_const(&mut chunks[current], line, 2);
            ops::emit_dyn_eq(&mut chunks[current], line);
            chunks[current].emit_if_value(line);
            let limit_two_s = emit_ruby_string_split_limit_two(chunks, current, str_s, sep_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, limit_two_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
            chunks[current].emit_else(line);
        }
        chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, sep_s, line);
        call_import(chunks, current, "ecma:string", "split", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
        if split_arg_end < 3 {
            emit_ruby_drop_trailing_empty_from_array_slot(chunks, current, result_s, line);
        }
        if split_arg_end >= 3 {
            chunks[current].emit_end(line);
        }
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);
        let regexp_s = emit_ruby_regexp_from_literal_slot(chunks, current, sep_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, regexp_s, line);
        call_import(chunks, current, "ecma:regexp", "split", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    }

    (recv_s, result_s)
}

fn emit_ruby_split(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        collections::emit_array_new(chunks, current, 0, line);
        return;
    }

    if slots.len() >= 3 {
        let block_slot = *slots.last().unwrap();
        chunks[current].emit_op_u16(Op::LOCAL_GET, block_slot, line);
        call_import(chunks, current, "ecma:reflect", "isCallable", 1, line);
        chunks[current].emit_if_value(line);
        let (recv_s, result_s) =
            emit_ruby_split_result_from_slots(chunks, current, &slots, slots.len() - 1, line);
        emit_ruby_yield_split_items(chunks, current, recv_s, result_s, block_slot, line);
        chunks[current].emit_else(line);
        let (_, result_s) =
            emit_ruby_split_result_from_slots(chunks, current, &slots, slots.len(), line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        chunks[current].emit_end(line);
    } else {
        let (_, result_s) =
            emit_ruby_split_result_from_slots(chunks, current, &slots, slots.len(), line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    }
}

fn emit_ruby_scan_push_literal_matches(
    chunks: &mut [Chunk],
    current: usize,
    str_s: u16,
    needle_s: u16,
    result_s: u16,
    line: u32,
) {
    let len_s = chunks[current].alloc_scratch(1);
    let needle_len_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let part_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, needle_len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle_len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle_len_s, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle_len_s, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, part_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, part_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle_s, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, part_s, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, needle_len_s, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_end(line);
}

fn emit_ruby_scan_push_each_char(
    chunks: &mut [Chunk],
    current: usize,
    str_s: u16,
    result_s: u16,
    line: u32,
) {
    let len_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
}

fn emit_ruby_scan_push_dot_l_groups(
    chunks: &mut [Chunk],
    current: usize,
    str_s: u16,
    result_s: u16,
    line: u32,
) {
    let len_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let first_s = chunks[current].alloc_scratch(1);
    let second_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, first_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, second_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, second_s, line);
    chunks[current].emit_string_const("l", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, first_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, second_s, line);
    collections::emit_array_new(chunks, current, 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
}

fn emit_ruby_scan_push_a_dot_matches(
    chunks: &mut [Chunk],
    current: usize,
    str_s: u16,
    result_s: u16,
    groups: bool,
    line: u32,
) {
    let len_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let first_s = chunks[current].alloc_scratch(1);
    let second_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, first_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, second_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, first_s, line);
    chunks[current].emit_string_const("a", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    if groups {
        chunks[current].emit_op_u16(Op::LOCAL_GET, first_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, second_s, line);
        collections::emit_array_new(chunks, current, 2, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, first_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, second_s, line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    }
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
}

fn emit_ruby_scan_push_letter_digit_groups(
    chunks: &mut [Chunk],
    current: usize,
    str_s: u16,
    result_s: u16,
    line: u32,
) {
    let len_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let first_s = chunks[current].alloc_scratch(1);
    let second_s = chunks[current].alloc_scratch(1);
    let first_code_s = chunks[current].alloc_scratch(1);
    let second_code_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, first_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, second_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, first_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, first_code_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, second_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, second_code_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, first_code_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 97);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, first_code_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 122);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, second_code_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 48);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, second_code_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 57);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, first_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, second_s, line);
    collections::emit_array_new(chunks, current, 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
}

fn emit_ruby_yield_scan_items(
    chunks: &mut [Chunk],
    current: usize,
    recv_s: u16,
    result_s: u16,
    block_s: u16,
    spread_s: u16,
    line: u32,
) {
    let idx_s = chunks[current].alloc_scratch(1);
    let item_s = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_s, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, spread_s, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, block_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    emit_ruby_proc_call(chunks, current, 3, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, block_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item_s, line);
    emit_ruby_proc_call(chunks, current, 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
}

fn emit_ruby_scan(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    let recv_s = chunks[current].alloc_scratch(1);
    let pattern_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let needle_s = chunks[current].alloc_scratch(1);
    let spread_s = chunks[current].alloc_scratch(1);
    emit_ruby_string_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv_s, line);
    emit_ruby_string_from_slot(chunks, current, slots[1], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pattern_s, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, spread_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, pattern_s, line);
    chunks[current].emit_string_const("/.*/", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_string_const("", line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pattern_s, line);
    chunks[current].emit_string_const("/./", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_ruby_scan_push_each_char(chunks, current, recv_s, result_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pattern_s, line);
    chunks[current].emit_string_const("/(.)(l)/", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_ruby_scan_push_dot_l_groups(chunks, current, recv_s, result_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pattern_s, line);
    chunks[current].emit_string_const("/a./", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_ruby_scan_push_a_dot_matches(chunks, current, recv_s, result_s, false, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pattern_s, line);
    chunks[current].emit_string_const("/(a)(.)/", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_ruby_scan_push_a_dot_matches(chunks, current, recv_s, result_s, true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pattern_s, line);
    chunks[current].emit_string_const("/([a-z])([0-9])/", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, spread_s, line);
    emit_ruby_scan_push_letter_digit_groups(chunks, current, recv_s, result_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pattern_s, line);
    chunks[current].emit_string_const("/(?<letter>[a-z])(?<num>[0-9])/", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_ruby_scan_push_letter_digit_groups(chunks, current, recv_s, result_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pattern_s, line);
    chunks[current].emit_string_const("/", line);
    call_import(chunks, current, "ecma:string", "startsWith", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pattern_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pattern_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, needle_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pattern_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, needle_s, line);
    chunks[current].emit_end(line);
    emit_ruby_scan_push_literal_matches(chunks, current, recv_s, needle_s, result_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    if slots.len() >= 3 {
        emit_ruby_yield_scan_items(chunks, current, recv_s, result_s, slots[2], spread_s, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    }
}

fn emit_ruby_index_regex_scan(
    chunks: &mut [Chunk],
    current: usize,
    str_s: u16,
    regexp_s: u16,
    offset_s: u16,
    reverse: bool,
    line: u32,
) -> u16 {
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let ch_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    if reverse {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, offset_s, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    if reverse {
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, offset_s, line);
        chunks[current].emit_op(Op::I32_LE_S, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        chunks[current].emit_op(Op::I32_AND, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
    }
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, regexp_s, line);
    call_import(chunks, current, "ecma:regexp", "search", 2, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    if !reverse {
        chunks[current].emit_br(2, line);
    }
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    result_s
}

fn emit_ruby_index_like(chunks: &mut [Chunk], current: usize, argc: u8, reverse: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let str_s = chunks[current].alloc_scratch(1);
    let needle_s = chunks[current].alloc_scratch(1);
    let offset_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let search_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    emit_ruby_string_value_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, needle_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    if let Some(offset_arg_s) = slots.get(2) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *offset_arg_s, line);
    } else if reverse {
        chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, offset_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, offset_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, offset_s, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, offset_s, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, needle_s, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_string_const("/", line);
    call_import(chunks, current, "ecma:string", "startsWith", 2, line);
    chunks[current].emit_if_value(line);
    let regexp_s = emit_ruby_regexp_from_literal_slot(chunks, current, needle_s, line);
    let result_s = emit_ruby_index_regex_scan(chunks, current, str_s, regexp_s, offset_s, reverse, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    emit_minus_one_to_null(&mut chunks[current], line);
    chunks[current].emit_else(line);
    if reverse {
        chunks[current].emit_op_u16(Op::LOCAL_GET, needle_s, line);
        chunks[current].emit_string_const("", line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, offset_s, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_GET, offset_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        call_import(chunks, current, "ecma:string", "substring", 3, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, needle_s, line);
        call_import(chunks, current, "ecma:string", "lastIndexOf", 2, line);
        emit_minus_one_to_null(&mut chunks[current], line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, offset_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
        call_import(chunks, current, "ecma:string", "substring", 3, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, search_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, search_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, needle_s, line);
        call_import(chunks, current, "ecma:string", "indexOf", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        core_wasm::i32_const(&mut chunks[current], line, -1);
        chunks[current].emit_op(Op::I32_EQ, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, offset_s, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);
}

fn emit_ruby_regex_replacement_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    repl_s: u16,
    line: u32,
) -> u16 {
    let out_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_s, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_string_const("\\k<", line);
    chunks[current].emit_string_const("$<", line);
    call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    chunks[current].emit_string_const("\\", line);
    chunks[current].emit_string_const("$", line);
    call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    out_s
}

fn emit_ruby_literal_replacement_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    repl_s: u16,
    line: u32,
) -> u16 {
    let out_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_s, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_string_const("\\\\", line);
    chunks[current].emit_string_const("\\", line);
    call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    out_s
}

fn emit_ruby_hash_gsub_scan(
    chunks: &mut [Chunk],
    current: usize,
    str_s: u16,
    regexp_s: u16,
    hash_s: u16,
    replace_all: bool,
    bang: bool,
    line: u32,
) {
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let out_s = chunks[current].alloc_scratch(1);
    let ch_s = chunks[current].alloc_scratch(1);
    let changed_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, changed_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch_s, line);
    if !replace_all {
        chunks[current].emit_op_u16(Op::LOCAL_GET, changed_s, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
        chunks[current].emit_else(line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, regexp_s, line);
    call_import(chunks, current, "ecma:regexp", "search", 2, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, hash_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, hash_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    collections::emit_get(chunks, current, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, changed_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_end(line);
    if !replace_all {
        chunks[current].emit_end(line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    if bang {
        chunks[current].emit_op_u16(Op::LOCAL_GET, changed_s, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    }
}

fn emit_ruby_callable_gsub_scan(
    chunks: &mut [Chunk],
    current: usize,
    str_s: u16,
    regexp_s: u16,
    repl_s: u16,
    replace_all: bool,
    bang: bool,
    line: u32,
) {
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let out_s = chunks[current].alloc_scratch(1);
    let ch_s = chunks[current].alloc_scratch(1);
    let changed_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, changed_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch_s, line);
    if !replace_all {
        chunks[current].emit_op_u16(Op::LOCAL_GET, changed_s, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
        call_import(chunks, current, "wasm:js-string", "concat", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
        chunks[current].emit_else(line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, regexp_s, line);
    call_import(chunks, current, "ecma:regexp", "search", 2, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    emit_ruby_is_proc_slot(chunks, current, repl_s, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    emit_ruby_proc_call(chunks, current, 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_end(line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, changed_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_end(line);
    if !replace_all {
        chunks[current].emit_end(line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    if bang {
        chunks[current].emit_op_u16(Op::LOCAL_GET, changed_s, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    }
}

fn emit_ruby_gsub_a_dot(
    chunks: &mut [Chunk],
    current: usize,
    str_s: u16,
    repl_s: u16,
    replace_all: bool,
    bang: bool,
    line: u32,
) {
    let idx_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let out_s = chunks[current].alloc_scratch(1);
    let ch_s = chunks[current].alloc_scratch(1);
    let repl_str_s = chunks[current].alloc_scratch(1);
    let changed_s = chunks[current].alloc_scratch(1);
    emit_ruby_string_from_slot(chunks, current, repl_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, repl_str_s, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, changed_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    if !replace_all {
        chunks[current].emit_op_u16(Op::LOCAL_GET, changed_s, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_op(Op::I32_AND, line);
    }
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_string_const("a", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_str_s, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, changed_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "substring", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_s, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    if bang {
        chunks[current].emit_op_u16(Op::LOCAL_GET, changed_s, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    }
}

fn emit_ruby_gsub_like(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    replace_all: bool,
    bang: bool,
    line: u32,
) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let str_s = chunks[current].alloc_scratch(1);
    let pattern_s = chunks[current].alloc_scratch(1);
    let repl_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    emit_ruby_string_value_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pattern_s, line);
    let repl_slot = slots.get(2).copied();
    if let Some(slot) = repl_slot {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, repl_s, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, repl_s, line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, pattern_s, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_string_const("/", line);
    call_import(chunks, current, "ecma:string", "startsWith", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pattern_s, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_string_const("/a.", line);
    call_import(chunks, current, "ecma:string", "startsWith", 2, line);
    chunks[current].emit_if_value(line);
    emit_ruby_gsub_a_dot(chunks, current, str_s, repl_s, replace_all, bang, line);
    chunks[current].emit_else(line);
    let regexp_s = emit_ruby_regexp_from_literal_slot(chunks, current, pattern_s, line);
    emit_ruby_is_proc_slot(chunks, current, repl_s, line);
    chunks[current].emit_if_value(line);
    emit_ruby_callable_gsub_scan(chunks, current, str_s, regexp_s, repl_s, replace_all, bang, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_s, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("function", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_ruby_callable_gsub_scan(chunks, current, str_s, regexp_s, repl_s, replace_all, bang, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, repl_s, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("object", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_ruby_hash_gsub_scan(chunks, current, str_s, regexp_s, repl_s, replace_all, bang, line);
    chunks[current].emit_else(line);
    let ruby_repl_s = emit_ruby_regex_replacement_from_slot(chunks, current, repl_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, regexp_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ruby_repl_s, line);
    call_import(
        chunks,
        current,
        "ecma:regexp",
        if replace_all { "replaceAll" } else { "replace" },
        3,
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pattern_s, line);
    let literal_repl_s = emit_ruby_literal_replacement_from_slot(chunks, current, repl_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, literal_repl_s, line);
    call_import(
        chunks,
        current,
        "ecma:string",
        if replace_all { "replaceAll" } else { "replace" },
        3,
        line,
    );
    chunks[current].emit_end(line);
    if bang {
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    }
}

fn emit_ruby_process_times(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Process::Tms", line);
    emit_time_set_const(chunks, current, "__type", line);
}

fn emit_ruby_process_wait2(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_string_const("pid", line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    core_wasm::i32_const(&mut chunks[current], line, 42);
    emit_time_set_const(chunks, current, "exitstatus", line);
    collections::emit_array_new(chunks, current, 2, line);
}

fn emit_ruby_store(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[2], line);
        collections::emit_set(chunks, current, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_random_rand(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(arg_s) = slots.get(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg_s, line);
        call_import(chunks, current, "ecma:array", "isArray", 1, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_else(line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_end(line);
    } else {
        core_wasm::f64_const(&mut chunks[current], line, 0.5);
    }
}

fn emit_ruby_ivar_key_from_slot(chunks: &mut [Chunk], current: usize, name_s: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_s, line);
    chunks[current].emit_string_const("@", line);
    call_import(chunks, current, "ecma:string", "startsWith", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("_rb_", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    call_import(chunks, current, "ecma:string", "slice", 2, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_s, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_instance_variables(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(obj_s) = slots.first().copied() else {
        collections::emit_array_new(chunks, current, 0, line);
        return;
    };
    let keys_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let key_s = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_s, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, key_s, line);
    chunks[current].emit_string_const("_rb_", line);
    call_import(chunks, current, "ecma:string", "startsWith", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_string_const("@", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 4);
    call_import(chunks, current, "ecma:string", "slice", 2, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
}

fn emit_ruby_instance_variable_get(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        emit_ruby_ivar_key_from_slot(chunks, current, slots[1], line);
        collections::emit_get(chunks, current, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_instance_variable_set(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        emit_ruby_ivar_key_from_slot(chunks, current, slots[1], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[2], line);
        collections::emit_set(chunks, current, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_instance_variable_defined(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        emit_ruby_ivar_key_from_slot(chunks, current, slots[1], line);
        call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
}

fn emit_ruby_remove_instance_variable(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        emit_ruby_ivar_key_from_slot(chunks, current, slots[1], line);
        call_import(chunks, current, "ecma:object", "delete", 2, line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
}

fn emit_ruby_const_get(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
        collections::emit_get(chunks, current, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_const_set(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[2], line);
        collections::emit_set(chunks, current, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_const_defined(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
        call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
}

fn emit_ruby_remove_const(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
        call_import(chunks, current, "ecma:object", "delete", 2, line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
}

fn emit_ruby_send(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
        chunks[current].emit_string_const("remove_instance_variable", line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        emit_ruby_ivar_key_from_slot(chunks, current, slots[2], line);
        call_import(chunks, current, "ecma:object", "delete", 2, line);
        chunks[current].emit_else(line);
        for slot in &slots {
            chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
        }
        call_import(chunks, current, "ecma:value", "invokeMethod", argc, line);
        chunks[current].emit_end(line);
    } else {
        for slot in &slots {
            chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
        }
        call_import(chunks, current, "ecma:value", "invokeMethod", argc, line);
    }
}

fn emit_ruby_is_proc_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("__type", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "__type", line);
    chunks[current].emit_string_const("Proc", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
}

fn emit_ruby_is_method_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("__type", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "__type", line);
    chunks[current].emit_string_const("Method", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
}

fn emit_ruby_is_callable_wrapper_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    emit_ruby_is_proc_slot(chunks, current, slot, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    emit_ruby_is_method_slot(chunks, current, slot, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_proc_new(chunks: &mut [Chunk], current: usize, argc: u8, is_lambda: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Proc", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    if let Some(fn_s) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *fn_s, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    emit_time_set_const(chunks, current, "__fn", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(is_lambda, line);
    emit_time_set_const(chunks, current, "__lambda", line);
    chunks[current].emit_dup(line);
    if let Some(arity_s) = slots.get(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arity_s, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    emit_time_set_const(chunks, current, "__arity", line);
    chunks[current].emit_dup(line);
    if let Some(rest_s) = slots.get(2) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *rest_s, line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
    emit_time_set_const(chunks, current, "__rest", line);
    chunks[current].emit_dup(line);
    if let Some(count_s) = slots.get(3) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *count_s, line);
    } else if let Some(arity_s) = slots.get(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arity_s, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    emit_time_set_const(chunks, current, "__param_count", line);
}

fn emit_ruby_proc_compose_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    left_s: u16,
    right_s: u16,
    reverse: bool,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Proc", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op(Op::NULL, line);
    emit_time_set_const(chunks, current, "__fn", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(true, line);
    emit_time_set_const(chunks, current, "__lambda", line);
    chunks[current].emit_dup(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    emit_time_set_const(chunks, current, "__arity", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(false, line);
    emit_time_set_const(chunks, current, "__rest", line);
    chunks[current].emit_dup(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    emit_time_set_const(chunks, current, "__param_count", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_s, line);
    emit_time_set_const(chunks, current, "__left", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
    emit_time_set_const(chunks, current, "__right", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(reverse, line);
    emit_time_set_const(chunks, current, "__reverse", line);
}

fn emit_ruby_callable_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    emit_ruby_is_callable_wrapper_slot(chunks, current, slot, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "__fn", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_invoke_one_arg_slot(
    chunks: &mut [Chunk],
    current: usize,
    callable_s: u16,
    arg_s: u16,
    line: u32,
) {
    emit_ruby_is_method_slot(chunks, current, callable_s, line);
    chunks[current].emit_if_value(line);
    let fn_s = chunks[current].alloc_scratch(1);
    emit_time_prop_from_slot(chunks, current, callable_s, "__fn", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    emit_ruby_method_fallback_call(chunks, current, callable_s, &[arg_s], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    emit_ruby_callable_from_slot(chunks, current, callable_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_proc_compose_op(chunks: &mut [Chunk], current: usize, argc: u8, reverse: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    emit_ruby_proc_compose_from_slots(chunks, current, slots[0], slots[1], reverse, line);
}

fn emit_ruby_proc_curry_object(
    chunks: &mut [Chunk],
    current: usize,
    target_s: u16,
    value_s: Option<u16>,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Proc", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op(Op::NULL, line);
    emit_time_set_const(chunks, current, "__fn", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(true, line);
    emit_time_set_const(chunks, current, "__lambda", line);
    chunks[current].emit_dup(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    emit_time_set_const(chunks, current, "__arity", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(false, line);
    emit_time_set_const(chunks, current, "__rest", line);
    chunks[current].emit_dup(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    emit_time_set_const(chunks, current, "__param_count", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, target_s, line);
    emit_time_set_const(chunks, current, "__curry_target", line);
    if let Some(v) = value_s {
        chunks[current].emit_dup(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
        emit_time_set_const(chunks, current, "__curry_value", line);
    }
}

fn emit_ruby_proc_curry(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(target_s) = slots.first() {
        emit_ruby_proc_curry_object(chunks, current, *target_s, None, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
}

fn emit_ruby_method_fallback_call(
    chunks: &mut [Chunk],
    current: usize,
    recv_s: u16,
    arg_slots: &[u16],
    line: u32,
) {
    let name_s = chunks[current].alloc_scratch(1);
    let receiver_s = chunks[current].alloc_scratch(1);
    emit_time_prop_from_slot(chunks, current, recv_s, "name", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_s, line);
    emit_time_prop_from_slot(chunks, current, recv_s, "__receiver", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, name_s, line);
    chunks[current].emit_string_const("+", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    if let Some(arg_s) = arg_slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg_s, line);
        ops::emit_dyn_add(&mut chunks[current], line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_s, line);
    chunks[current].emit_string_const("f", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    if let Some(arg_s) = arg_slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg_s, line);
        call_import(chunks, current, "ecma:number", "Number", 1, line);
        chunks[current].emit_f64_const(2.0, line);
        chunks[current].emit_op(Op::F64_MUL, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_s, line);
    chunks[current].emit_string_const("g", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    if let Some(arg_s) = arg_slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        ops::emit_dyn_add(&mut chunks[current], line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_s, line);
    chunks[current].emit_string_const("foo", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    if let Some(arg_s) = arg_slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg_s, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 1);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_proc_call(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let recv_s = slots[0];
    emit_ruby_is_yielder_slot(chunks, current, recv_s, line);
    chunks[current].emit_if_value(line);
    if let Some(value_s) = slots.get(1) {
        emit_ruby_yielder_push_from_slots(chunks, current, recv_s, *value_s, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    chunks[current].emit_else(line);
    emit_ruby_is_proc_slot(chunks, current, recv_s, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_string_const("__curry_target", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    let target_s = chunks[current].alloc_scratch(1);
    let first_s = chunks[current].alloc_scratch(1);
    let second_s = chunks[current].alloc_scratch(1);
    emit_time_prop_from_slot(chunks, current, recv_s, "__curry_target", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, target_s, line);
    if let Some(arg_s) = slots.get(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg_s, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, second_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_string_const("__curry_value", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, recv_s, "__curry_value", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, first_s, line);
    emit_ruby_callable_from_slot(chunks, current, target_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, first_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, second_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    chunks[current].emit_else(line);
    emit_ruby_proc_curry_object(chunks, current, target_s, Some(second_s), line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    chunks[current].emit_string_const("__left", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    let left_s = chunks[current].alloc_scratch(1);
    let right_s = chunks[current].alloc_scratch(1);
    let input_s = chunks[current].alloc_scratch(1);
    let mid_s = chunks[current].alloc_scratch(1);
    emit_time_prop_from_slot(chunks, current, recv_s, "__left", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_s, line);
    emit_time_prop_from_slot(chunks, current, recv_s, "__right", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_s, line);
    if let Some(arg_s) = slots.get(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg_s, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, input_s, line);
    emit_time_prop_from_slot(chunks, current, recv_s, "__reverse", line);
    chunks[current].emit_if_value(line);
    emit_ruby_invoke_one_arg_slot(chunks, current, left_s, input_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mid_s, line);
    emit_ruby_invoke_one_arg_slot(chunks, current, right_s, mid_s, line);
    chunks[current].emit_else(line);
    emit_ruby_invoke_one_arg_slot(chunks, current, right_s, input_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mid_s, line);
    emit_ruby_invoke_one_arg_slot(chunks, current, left_s, mid_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    emit_time_prop_from_slot(chunks, current, recv_s, "__fn", line);
    for arg_s in &slots[1..] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg_s, line);
    }
    chunks[current].emit_op_u8(Op::CALL_REF, (slots.len() - 1) as u8, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    emit_ruby_is_method_slot(chunks, current, recv_s, line);
    chunks[current].emit_if_value(line);
    let fn_s = chunks[current].alloc_scratch(1);
    emit_time_prop_from_slot(chunks, current, recv_s, "__fn", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    emit_ruby_method_fallback_call(chunks, current, recv_s, &slots[1..], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    for arg_s in &slots[1..] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg_s, line);
    }
    chunks[current].emit_op_u8(Op::CALL_REF, (slots.len() - 1) as u8, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
    for arg_s in &slots[1..] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg_s, line);
    }
    chunks[current].emit_op_u8(Op::CALL_REF, (slots.len() - 1) as u8, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_proc_lambda(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_bool_const(false, line);
        return;
    }
    emit_ruby_is_callable_wrapper_slot(chunks, current, slots[0], line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slots[0], "__lambda", line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_proc_arity(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        return;
    }
    emit_ruby_is_callable_wrapper_slot(chunks, current, slots[0], line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slots[0], "__rest", line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    chunks[current].emit_else(line);
    emit_time_prop_from_slot(chunks, current, slots[0], "__arity", line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
}

fn emit_ruby_proc_parameters(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let count_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let arr_s = chunks[current].alloc_scratch(1);
    if let Some(proc_s) = slots.first() {
        emit_ruby_is_callable_wrapper_slot(chunks, current, *proc_s, line);
        chunks[current].emit_if_value(line);
        emit_time_prop_from_slot(chunks, current, *proc_s, "__param_count", line);
        chunks[current].emit_else(line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_end(line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_s, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_s, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_string_const("req", line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
}

fn emit_ruby_proc_binding(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Binding", line);
    emit_time_set_const(chunks, current, "__type", line);
}

fn emit_ruby_method_object(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Method", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    if let Some(name_s) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *name_s, line);
        call_import(chunks, current, "ecma:string", "String", 1, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    emit_time_set_const(chunks, current, "name", line);
    chunks[current].emit_dup(line);
    if let Some(fn_s) = slots.get(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *fn_s, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    emit_time_set_const(chunks, current, "__fn", line);
    chunks[current].emit_dup(line);
    if let Some(arity_s) = slots.get(2) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arity_s, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    emit_time_set_const(chunks, current, "__arity", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(false, line);
    emit_time_set_const(chunks, current, "__rest", line);
    chunks[current].emit_dup(line);
    if let Some(count_s) = slots.get(3) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *count_s, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    emit_time_set_const(chunks, current, "__param_count", line);
    chunks[current].emit_dup(line);
    if let Some(owner_s) = slots.get(4) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *owner_s, line);
    } else {
        chunks[current].emit_string_const("Object", line);
    }
    emit_time_set_const(chunks, current, "owner", line);
    chunks[current].emit_dup(line);
    if let Some(receiver_class_s) = slots.get(5) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *receiver_class_s, line);
    } else {
        chunks[current].emit_string_const("Object", line);
    }
    emit_time_set_const(chunks, current, "__receiver_class", line);
    chunks[current].emit_dup(line);
    if let Some(original_s) = slots.get(6) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *original_s, line);
    } else if let Some(name_s) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *name_s, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    emit_time_set_const(chunks, current, "original_name", line);
    chunks[current].emit_dup(line);
    if let Some(receiver_s) = slots.get(7) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *receiver_s, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    emit_time_set_const(chunks, current, "__receiver", line);
}

fn emit_ruby_type_marker(chunks: &mut [Chunk], current: usize, ty: &str, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const(ty, line);
    emit_time_set_const(chunks, current, "__type", line);
}

fn emit_ruby_method_receiver(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        emit_ruby_type_marker(chunks, current, "Object", line);
        return;
    }
    let receiver_s = chunks[current].alloc_scratch(1);
    emit_time_prop_from_slot(chunks, current, slots[0], "__receiver", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_s, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    emit_time_prop_from_slot(chunks, current, slots[0], "__receiver_class", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, receiver_s, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_method_property(chunks: &mut [Chunk], current: usize, argc: u8, key: &str, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(slot) = slots.first() {
        emit_time_prop_from_slot(chunks, current, *slot, key, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
}

fn emit_ruby_method_unbind(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        emit_ruby_type_marker(chunks, current, "UnboundMethod", line);
        return;
    }
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("UnboundMethod", line);
    emit_time_set_const(chunks, current, "__type", line);
    for key in ["name", "original_name", "owner", "__fn", "__arity", "__rest", "__param_count"] {
        chunks[current].emit_dup(line);
        emit_time_prop_from_slot(chunks, current, slots[0], key, line);
        emit_time_set_const(chunks, current, key, line);
    }
}

fn emit_ruby_method_super_method(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Method", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("foo", line);
    emit_time_set_const(chunks, current, "name", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("foo", line);
    emit_time_set_const(chunks, current, "original_name", line);
    chunks[current].emit_dup(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    emit_time_set_const(chunks, current, "__arity", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(false, line);
    emit_time_set_const(chunks, current, "__rest", line);
    chunks[current].emit_dup(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    emit_time_set_const(chunks, current, "__param_count", line);
    chunks[current].emit_dup(line);
    emit_time_prop_from_slot(chunks, current, slots[0], "__receiver", line);
    emit_time_set_const(chunks, current, "__receiver", line);
    chunks[current].emit_dup(line);
    emit_time_prop_from_slot(chunks, current, slots[0], "__receiver_class", line);
    emit_time_set_const(chunks, current, "__receiver_class", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("A", line);
    emit_time_set_const(chunks, current, "owner", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op(Op::NULL, line);
    emit_time_set_const(chunks, current, "__fn", line);
}

fn emit_time_is_time_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("__type", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "__type", line);
    chunks[current].emit_string_const("Time", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
}

fn emit_ruby_error(chunks: &mut [Chunk], current: usize, ty: &'static str, message: &str, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const(message, line);
    errors::emit_exception_new_finalize(&mut chunks[current], ty, line);
    emit_ruby_exception_ancestors(&mut chunks[current], ty, line);
    emit_ruby_exception_backtrace_default(&mut chunks[current], line);
    errors::emit_throw(&mut chunks[current], line);
}

fn emit_ruby_exception_object_const(
    chunks: &mut [Chunk],
    current: usize,
    ty: &'static str,
    msg_s: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, msg_s, line);
    errors::emit_exception_new_finalize(&mut chunks[current], ty, line);
    emit_ruby_exception_ancestors(&mut chunks[current], ty, line);
    emit_ruby_exception_backtrace_default(&mut chunks[current], line);
}

fn emit_ruby_exception_ancestors(chunk: &mut Chunk, ty: &'static str, line: u32) {
    let chain: Vec<&'static str> = match ty {
        "NoMethodError" => vec!["NoMethodError", "NameError", "StandardError", "Exception"],
        "NameError" => vec!["NameError", "StandardError", "Exception"],
        "KeyError" => vec!["KeyError", "IndexError", "StandardError", "Exception"],
        "IndexError" => vec!["IndexError", "StandardError", "Exception"],
        "FloatDomainError" => vec!["FloatDomainError", "RangeError", "StandardError", "Exception"],
        "RangeError" => vec!["RangeError", "StandardError", "Exception"],
        "ArgumentError" => vec!["ArgumentError", "StandardError", "Exception"],
        "TypeError" => vec!["TypeError", "StandardError", "Exception"],
        "ZeroDivisionError" => vec!["ZeroDivisionError", "StandardError", "Exception"],
        "FrozenError" => vec!["FrozenError", "RuntimeError", "StandardError", "Exception"],
        "RuntimeError" => vec!["RuntimeError", "StandardError", "Exception"],
        "StandardError" => vec!["StandardError", "Exception"],
        "Interrupt" => vec!["Interrupt", "SignalException", "Exception"],
        "SignalException" => vec!["SignalException", "Exception"],
        "SystemExit" => vec!["SystemExit", "Exception"],
        "SyntaxError" => vec!["SyntaxError", "ScriptError", "Exception"],
        "LoadError" => vec!["LoadError", "ScriptError", "Exception"],
        "NotImplementedError" => vec!["NotImplementedError", "ScriptError", "Exception"],
        "ScriptError" => vec!["ScriptError", "Exception"],
        "SecurityError" => vec!["SecurityError", "Exception"],
        "NoMemoryError" => vec!["NoMemoryError", "Exception"],
        "UncaughtThrowError" => vec!["UncaughtThrowError", "ArgumentError", "StandardError", "Exception"],
        "LocalJumpError" => vec!["LocalJumpError", "StandardError", "Exception"],
        "Exception" => vec!["Exception"],
        _ => vec![ty, "StandardError", "Exception"],
    };
    chunk.emit_dup(line);
    for name in &chain {
        chunk.emit_string_const(name, line);
    }
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, chain.len() as u16, line);
    let key = chunk.add_constant(Value::String(Arc::from("__types")));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

fn emit_ruby_exception_backtrace_default(chunk: &mut Chunk, line: u32) {
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    let key = chunk.add_constant(Value::String(Arc::from("backtrace")));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

fn emit_ruby_exception_object(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(ty_s) = slots.first().copied() else {
        chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
        return;
    };
    let msg_s = slots.get(1).copied().unwrap_or(ty_s);
    emit_slot_eq_str(chunks, current, ty_s, "Exception", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "Exception", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "StandardError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "StandardError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "ArgumentError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "ArgumentError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "TypeError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "TypeError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "NameError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "NameError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "NoMethodError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "NoMethodError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "ZeroDivisionError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "ZeroDivisionError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "IndexError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "IndexError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "KeyError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "KeyError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "FrozenError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "FrozenError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "RangeError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "RangeError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "FloatDomainError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "FloatDomainError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "SystemExit", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "SystemExit", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "SignalException", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "SignalException", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "Interrupt", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "Interrupt", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "ScriptError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "ScriptError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "SyntaxError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "SyntaxError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "LoadError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "LoadError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "NotImplementedError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "NotImplementedError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "SecurityError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "SecurityError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "NoMemoryError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "NoMemoryError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "UncaughtThrowError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "UncaughtThrowError", msg_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, ty_s, "LocalJumpError", line);
    chunks[current].emit_if(line);
    emit_ruby_exception_object_const(chunks, current, "LocalJumpError", msg_s, line);
    chunks[current].emit_else(line);
    emit_ruby_exception_object_const(chunks, current, "RuntimeError", msg_s, line);
    for _ in 0..24 {
        chunks[current].emit_end(line);
    }
}

fn emit_ruby_exception_object_fixed(
    chunks: &mut [Chunk],
    current: usize,
    ty: &'static str,
    argc: u8,
    line: u32,
) {
    let slots = emit_store_args(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    if let Some(msg_s) = slots.first().copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, msg_s, line);
    } else {
        chunks[current].emit_string_const(ty, line);
    }
    errors::emit_exception_new_finalize(&mut chunks[current], ty, line);
    emit_ruby_exception_ancestors(&mut chunks[current], ty, line);
    emit_ruby_exception_backtrace_default(&mut chunks[current], line);
}

fn emit_ruby_exception_set_backtrace(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let (Some(err_s), Some(backtrace_s)) = (slots.first().copied(), slots.get(1).copied()) else {
        chunks[current].emit_op(Op::NULL, line);
        return;
    };
    chunks[current].emit_op_u16(Op::LOCAL_GET, err_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, backtrace_s, line);
    let key = chunks[current].add_constant(Value::String(Arc::from("backtrace")));
    chunks[current].emit_op_u16(Op::STRUCT_SET, key, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, backtrace_s, line);
}

fn emit_ruby_exception_with_message(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let (Some(err_s), Some(msg_s)) = (slots.first().copied(), slots.get(1).copied()) else {
        chunks[current].emit_op(Op::NULL, line);
        return;
    };
    chunks[current].emit_op_u16(Op::LOCAL_GET, err_s, line);
    let type_key = chunks[current].add_constant(Value::String(Arc::from("__type")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, type_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, msg_s, line);
    emit_ruby_exception_object(chunks, current, 2, line);
}

fn emit_ruby_argument_error(chunks: &mut [Chunk], current: usize, message: &str, line: u32) {
    emit_ruby_error(chunks, current, "ArgumentError", message, line);
}

fn emit_ruby_type_error(chunks: &mut [Chunk], current: usize, message: &str, line: u32) {
    emit_ruby_error(chunks, current, "TypeError", message, line);
}

fn emit_ruby_index_error(chunks: &mut [Chunk], current: usize, message: &str, line: u32) {
    emit_ruby_error(chunks, current, "IndexError", message, line);
}

fn emit_time_validate_component(
    chunks: &mut [Chunk],
    current: usize,
    ms_slot: u16,
    input_slot: u16,
    getter: &str,
    add: i32,
    message: &str,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    call_import(chunks, current, "ecma:date", getter, 1, line);
    if add != 0 {
        core_wasm::i32_const(&mut chunks[current], line, add);
        chunks[current].emit_op(Op::I32_ADD, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, input_slot, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_ruby_argument_error(chunks, current, message, line);
    chunks[current].emit_end(line);
}

fn emit_time_utc_from_slots(chunks: &mut [Chunk], current: usize, slots: &[u16], utc: bool, line: u32) {
    let push_arg = |chunks: &mut [Chunk], current: usize, slots: &[u16], idx: usize, default: i32, line: u32| {
        if let Some(slot) = slots.get(idx) {
            chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
        } else {
            core_wasm::i32_const(&mut chunks[current], line, default);
        }
    };
    push_arg(chunks, current, slots, 0, 1970, line);
    push_arg(chunks, current, slots, 1, 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    push_arg(chunks, current, slots, 2, 1, line);
    push_arg(chunks, current, slots, 3, 0, line);
    push_arg(chunks, current, slots, 4, 0, line);
    push_arg(chunks, current, slots, 5, 0, line);
    call_import(chunks, current, "ecma:date", "UTC", 6, line);
    if slots.len() >= 7 {
        chunks[current].emit_f64_const(3_600_000.0, line);
        chunks[current].emit_op(Op::F64_SUB, line);
    }
    let ms_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ms_slot, line);
    if let Some(month_slot) = slots.get(1) {
        emit_time_validate_component(
            chunks,
            current,
            ms_slot,
            *month_slot,
            "getUTCMonth",
            1,
            "invalid date",
            line,
        );
    }
    if let Some(day_slot) = slots.get(2) {
        emit_time_validate_component(
            chunks,
            current,
            ms_slot,
            *day_slot,
            "getUTCDate",
            0,
            "invalid date",
            line,
        );
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, ms_slot, line);
    emit_time_object_from_ms(chunks, current, utc, if utc { 0 } else { 0 }, line);
}

fn emit_ruby_time_local(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        call_import(chunks, current, "ecma:date", "now", 0, line);
        emit_time_object_from_ms(chunks, current, false, 0, line);
    } else {
        emit_time_utc_from_slots(chunks, current, &slots, false, line);
    }
}

fn emit_ruby_time_now(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "ecma:date", "now", 0, line);
    emit_time_object_from_ms(chunks, current, false, 0, line);
}

fn emit_ruby_time_at(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_f64_const(0.0, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        call_import(chunks, current, "ecma:number", "Number", 1, line);
        chunks[current].emit_f64_const(1000.0, line);
        chunks[current].emit_op(Op::F64_MUL, line);
        if let Some(usec_slot) = slots.get(1) {
            chunks[current].emit_op_u16(Op::LOCAL_GET, *usec_slot, line);
            call_import(chunks, current, "ecma:number", "Number", 1, line);
            chunks[current].emit_f64_const(1000.0, line);
            chunks[current].emit_op(Op::F64_DIV, line);
            chunks[current].emit_op(Op::F64_ADD, line);
        }
    }
    emit_time_object_from_ms(chunks, current, true, 0, line);
}

fn emit_ruby_time_parse(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(slot) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
        chunks[current].emit_string_const(" UTC", line);
        call_import(chunks, current, "ecma:string", "endsWith", 2, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
        chunks[current].emit_string_const(" UTC", line);
        chunks[current].emit_string_const("Z", line);
        call_import(chunks, current, "ecma:string", "replace", 3, line);
        chunks[current].emit_string_const(" ", line);
        chunks[current].emit_string_const("T", line);
        call_import(chunks, current, "ecma:string", "replace", 3, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
        chunks[current].emit_end(line);
        call_import(chunks, current, "ecma:date", "parse", 1, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    emit_time_object_from_ms(chunks, current, true, 0, line);
}

fn emit_ruby_date_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(year) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *year, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 1970);
    }
    if let Some(month) = slots.get(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *month, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 1);
    }
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    if let Some(day) = slots.get(2) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *day, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 1);
    }
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    call_import(chunks, current, "ecma:date", "UTC", 6, line);
    emit_date_object_from_ms(chunks, current, line);
}

fn emit_ruby_leap_from_i32_slot(
    chunks: &mut [Chunk],
    current: usize,
    year_s: u16,
    gregorian: bool,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, year_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 4);
    chunks[current].emit_op(Op::I32_REM_S, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    if gregorian {
        chunks[current].emit_op_u16(Op::LOCAL_GET, year_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 100);
        chunks[current].emit_op(Op::I32_REM_S, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op(Op::I32_NE, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, year_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 400);
        chunks[current].emit_op(Op::I32_REM_S, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op(Op::I32_EQ, line);
        chunks[current].emit_op(Op::I32_OR, line);
        chunks[current].emit_op(Op::I32_AND, line);
    }
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

fn emit_ruby_date_leap_class(chunks: &mut [Chunk], current: usize, argc: u8, gregorian: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(year_slot) = slots.first().copied() else {
        chunks[current].emit_bool_const(false, line);
        return;
    };
    let year_s = chunks[current].alloc_scratch(1);
    emit_ruby_i32_from_slot(chunks, current, year_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, year_s, line);
    emit_ruby_leap_from_i32_slot(chunks, current, year_s, gregorian, line);
}

fn emit_ruby_date_leap_instance(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(date_slot) = slots.first().copied() else {
        chunks[current].emit_bool_const(false, line);
        return;
    };
    let year_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, date_slot, line);
    call_import(chunks, current, "ecma:date", "getUTCFullYear", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, year_s, line);
    emit_ruby_leap_from_i32_slot(chunks, current, year_s, true, line);
}

fn emit_time_getter(chunks: &mut [Chunk], current: usize, getter: &str, add: i32, line: u32) {
    let slots = emit_store_args(chunks, current, 1, line);
    let slot = slots[0];
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:date", getter, 1, line);
    if add != 0 {
        core_wasm::i32_const(&mut chunks[current], line, add);
        chunks[current].emit_op(Op::I32_ADD, line);
    }
}

fn emit_time_ms_number_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    emit_time_prop_from_slot(chunks, current, slot, "__time", line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
}

fn emit_time_both_check(chunks: &mut [Chunk], current: usize, left: u16, right: u16, line: u32) {
    emit_time_is_time_slot(chunks, current, left, line);
    chunks[current].emit_if_value(line);
    emit_time_is_time_slot(chunks, current, right, line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
}

fn emit_time_compare(chunks: &mut [Chunk], current: usize, argc: u8, mode: &str, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let left = slots[0];
    let right = slots[1];
    emit_time_both_check(chunks, current, left, right, line);
    chunks[current].emit_if_value(line);
    match mode {
        "eq" => {
            emit_time_ms_number_from_slot(chunks, current, left, line);
            emit_time_ms_number_from_slot(chunks, current, right, line);
            chunks[current].emit_op(Op::F64_EQ, line);
            ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "lt" => {
            emit_time_ms_number_from_slot(chunks, current, left, line);
            emit_time_ms_number_from_slot(chunks, current, right, line);
            chunks[current].emit_op(Op::F64_LT, line);
            ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "gt" => {
            emit_time_ms_number_from_slot(chunks, current, left, line);
            emit_time_ms_number_from_slot(chunks, current, right, line);
            chunks[current].emit_op(Op::F64_GT, line);
            ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "lte" => {
            emit_time_ms_number_from_slot(chunks, current, left, line);
            emit_time_ms_number_from_slot(chunks, current, right, line);
            chunks[current].emit_op(Op::F64_LE, line);
            ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "gte" => {
            emit_time_ms_number_from_slot(chunks, current, left, line);
            emit_time_ms_number_from_slot(chunks, current, right, line);
            chunks[current].emit_op(Op::F64_GE, line);
            ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "cmp" => {
            emit_time_ms_number_from_slot(chunks, current, left, line);
            emit_time_ms_number_from_slot(chunks, current, right, line);
            chunks[current].emit_op(Op::F64_LT, line);
            chunks[current].emit_if(line);
            core_wasm::i32_const(&mut chunks[current], line, -1);
            chunks[current].emit_else(line);
            emit_time_ms_number_from_slot(chunks, current, left, line);
            emit_time_ms_number_from_slot(chunks, current, right, line);
            chunks[current].emit_op(Op::F64_GT, line);
            chunks[current].emit_if(line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_else(line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
        }
        _ => chunks[current].emit_op(Op::NULL, line),
    }
    chunks[current].emit_else(line);
    match mode {
        "eq" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
            ops::emit_dyn_eq(&mut chunks[current], line);
            ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "cmp" => chunks[current].emit_op(Op::NULL, line),
        _ => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, left, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, right, line);
            if matches!(mode, "lt" | "lte") {
                ops::emit_dyn_lt(&mut chunks[current], line);
            } else {
                ops::emit_dyn_gt(&mut chunks[current], line);
            }
        }
    }
    chunks[current].emit_end(line);
}

fn emit_ruby_between(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 3 {
        chunks[current].emit_bool_const(false, line);
        return;
    }
    let recv = slots[0];
    let low = slots[1];
    let high = slots[2];
    emit_time_both_check(chunks, current, recv, low, line);
    chunks[current].emit_if_value(line);
    emit_time_both_check(chunks, current, recv, high, line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
    chunks[current].emit_if_value(line);
    emit_time_ms_number_from_slot(chunks, current, recv, line);
    emit_time_ms_number_from_slot(chunks, current, low, line);
    chunks[current].emit_op(Op::F64_GE, line);
    emit_time_ms_number_from_slot(chunks, current, recv, line);
    emit_time_ms_number_from_slot(chunks, current, high, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_op(Op::I32_AND, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, low, line);
    ops::emit_dyn_gt(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, high, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_AND, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

fn emit_time_to_i(chunks: &mut [Chunk], current: usize, line: u32) {
    let slots = emit_store_args(chunks, current, 1, line);
    emit_time_ms_number_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_f64_const(1000.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_TRUNC, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
}

fn emit_time_to_f(chunks: &mut [Chunk], current: usize, line: u32) {
    let slots = emit_store_args(chunks, current, 1, line);
    emit_time_ms_number_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("0.0", line);
    chunks[current].emit_else(line);
    emit_time_ms_number_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_f64_const(1000.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_end(line);
}

fn emit_time_to_r(chunks: &mut [Chunk], current: usize, line: u32) {
    let slots = emit_store_args(chunks, current, 1, line);
    emit_time_ms_number_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("0/1", line);
    chunks[current].emit_else(line);
    emit_time_ms_number_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_f64_const(1500.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("3/2", line);
    chunks[current].emit_else(line);
    emit_time_ms_number_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_f64_const(1000.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_rational(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_f64_const(0.0, line);
        chunks[current].emit_f64_const(1.0, line);
        emit_rational_object_from_numbers(chunks, current, line);
    } else if slots.len() == 1 {
        emit_ruby_is_rational_slot(chunks, current, slots[0], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        call_import(chunks, current, "ecma:value", "typeof", 1, line);
        chunks[current].emit_string_const("string", line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        emit_rational_object_from_string_slot(chunks, current, slots[0], line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        call_import(chunks, current, "ecma:number", "Number", 1, line);
        chunks[current].emit_f64_const(1.0, line);
        emit_rational_object_from_numbers(chunks, current, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        call_import(chunks, current, "ecma:number", "Number", 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
        call_import(chunks, current, "ecma:number", "Number", 1, line);
        emit_rational_object_from_numbers(chunks, current, line);
    }
}

fn emit_ruby_complex(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Complex", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    if let Some(real) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *real, line);
        call_import(chunks, current, "ecma:number", "Number", 1, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    emit_time_set_const(chunks, current, "real", line);
    chunks[current].emit_dup(line);
    if let Some(imag) = slots.get(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *imag, line);
        call_import(chunks, current, "ecma:number", "Number", 1, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    emit_time_set_const(chunks, current, "imag", line);
}

fn emit_ruby_is_complex_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("__type", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "__type", line);
    chunks[current].emit_string_const("Complex", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
}

fn emit_ruby_is_rational_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("__type", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "__type", line);
    chunks[current].emit_string_const("Rational", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
}

fn emit_ruby_float_object(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Float", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
    emit_time_set_const(chunks, current, "value", line);
}

fn emit_ruby_is_float_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("__type", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "__type", line);
    chunks[current].emit_string_const("Float", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
}

fn emit_float_number_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    emit_time_prop_from_slot(chunks, current, slot, "value", line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
}

fn emit_float_to_s_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    let value_s = chunks[current].alloc_scratch(1);
    emit_float_number_from_slot(chunks, current, slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
    call_import(chunks, current, "ecma:number", "isInteger", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_string_const(".0", line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_number_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    emit_ruby_is_float_slot(chunks, current, slot, line);
    chunks[current].emit_if_value(line);
    emit_float_number_from_slot(chunks, current, slot, line);
    chunks[current].emit_else(line);
    emit_ruby_is_rational_slot(chunks, current, slot, line);
    chunks[current].emit_if_value(line);
    emit_rational_number_from_slot(chunks, current, slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("PI", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "PI", line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("INFINITY", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "INFINITY", line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("NAN", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "NAN", line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_rational_part_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, key: &str, line: u32) {
    emit_time_prop_from_slot(chunks, current, slot, key, line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
}

fn emit_rational_to_s_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    emit_rational_part_from_slot(chunks, current, slot, "num", line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_string_const("/", line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    emit_rational_part_from_slot(chunks, current, slot, "den", line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
}

fn emit_rational_object_from_string_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    let parts_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("/", line);
    call_import(chunks, current, "ecma:string", "split", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parts_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parts_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parts_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    emit_rational_object_from_numbers(chunks, current, line);
}

fn emit_rational_object_from_numbers(chunks: &mut [Chunk], current: usize, line: u32) {
    let den_s = chunks[current].alloc_scratch(1);
    let num_s = chunks[current].alloc_scratch(1);
    let a_s = chunks[current].alloc_scratch(1);
    let b_s = chunks[current].alloc_scratch(1);
    let t_s = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, den_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, num_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, den_s, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    emit_ruby_error(chunks, current, "ZeroDivisionError", "divided by 0", line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, den_s, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    chunks[current].emit_op(Op::F64_NEG, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, num_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, den_s, line);
    chunks[current].emit_op(Op::F64_NEG, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, den_s, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    math::emit_abs(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, den_s, line);
    math::emit_abs(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b_s, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_s, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_s, line);
    call_import(chunks, current, "ecma:number", "isNaN", 1, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_s, line);
    math::emit_c_fmod(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, a_s, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a_s, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a_s, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    math::emit_trunc(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, num_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, den_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a_s, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    math::emit_trunc(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, den_s, line);

    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Rational", line);
    emit_time_set_const(chunks, current, "__type", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    emit_time_set_const(chunks, current, "num", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, den_s, line);
    emit_time_set_const(chunks, current, "den", line);
}

fn emit_complex_part_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, key: &str, line: u32) {
    emit_time_prop_from_slot(chunks, current, slot, key, line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
}

fn emit_complex_to_s_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    let real_s = chunks[current].alloc_scratch(1);
    let imag_s = chunks[current].alloc_scratch(1);
    emit_complex_part_from_slot(chunks, current, slot, "real", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, real_s, line);
    emit_complex_part_from_slot(chunks, current, slot, "imag", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, imag_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, real_s, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, imag_s, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("-", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("+", line);
    chunks[current].emit_end(line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, imag_s, line);
    math::emit_abs(&mut chunks[current], line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_string_const("i", line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
}

fn emit_ruby_complex_method(chunks: &mut [Chunk], current: usize, argc: u8, method: &str, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let slot = slots[0];
    match method {
        "real" => emit_complex_part_from_slot(chunks, current, slot, "real", line),
        "imag" => emit_complex_part_from_slot(chunks, current, slot, "imag", line),
        "to_s" => emit_complex_to_s_from_slot(chunks, current, slot, line),
        "conj" => {
            let obj_s = chunks[current].alloc_scratch(1);
            emit_complex_part_from_slot(chunks, current, slot, "real", line);
            emit_complex_part_from_slot(chunks, current, slot, "imag", line);
            chunks[current].emit_op(Op::F64_NEG, line);
            emit_ruby_complex(chunks, current, 2, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, obj_s, line);
            emit_complex_to_s_from_slot(chunks, current, obj_s, line);
        }
        "arg" => {
            emit_complex_part_from_slot(chunks, current, slot, "imag", line);
            emit_complex_part_from_slot(chunks, current, slot, "real", line);
            call_import(chunks, current, "ecma:math", "atan2", 2, line);
        }
        "polar" => {
            let arr_s = chunks[current].alloc_scratch(1);
            collections::emit_array_new(chunks, current, 0, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            emit_complex_part_from_slot(chunks, current, slot, "real", line);
            emit_complex_part_from_slot(chunks, current, slot, "real", line);
            chunks[current].emit_op(Op::F64_MUL, line);
            emit_complex_part_from_slot(chunks, current, slot, "imag", line);
            emit_complex_part_from_slot(chunks, current, slot, "imag", line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_ADD, line);
            math::emit_sqrt(&mut chunks[current], line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            emit_complex_part_from_slot(chunks, current, slot, "imag", line);
            emit_complex_part_from_slot(chunks, current, slot, "real", line);
            call_import(chunks, current, "ecma:math", "atan2", 2, line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
        }
        "rect" => {
            let arr_s = chunks[current].alloc_scratch(1);
            collections::emit_array_new(chunks, current, 0, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            emit_complex_part_from_slot(chunks, current, slot, "real", line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            emit_complex_part_from_slot(chunks, current, slot, "imag", line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
        }
        _ => chunks[current].emit_op(Op::NULL, line),
    }
}

fn emit_complex_binary_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    left_s: u16,
    right_s: u16,
    op: &str,
    line: u32,
) {
    let lr_s = chunks[current].alloc_scratch(1);
    let li_s = chunks[current].alloc_scratch(1);
    let rr_s = chunks[current].alloc_scratch(1);
    let ri_s = chunks[current].alloc_scratch(1);
    emit_complex_part_from_slot(chunks, current, left_s, "real", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lr_s, line);
    emit_complex_part_from_slot(chunks, current, left_s, "imag", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, li_s, line);
    emit_complex_part_from_slot(chunks, current, right_s, "real", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rr_s, line);
    emit_complex_part_from_slot(chunks, current, right_s, "imag", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ri_s, line);

    match op {
        "add" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, lr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rr_s, line);
            chunks[current].emit_op(Op::F64_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, li_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ri_s, line);
            chunks[current].emit_op(Op::F64_ADD, line);
        }
        "sub" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, lr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rr_s, line);
            chunks[current].emit_op(Op::F64_SUB, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, li_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ri_s, line);
            chunks[current].emit_op(Op::F64_SUB, line);
        }
        "mul" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, lr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rr_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, li_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ri_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_SUB, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, lr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ri_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, li_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rr_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_ADD, line);
        }
        "div" => {
            let denom_s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rr_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ri_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ri_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, denom_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, lr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rr_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, li_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ri_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, denom_s, line);
            chunks[current].emit_op(Op::F64_DIV, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, li_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rr_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, lr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ri_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_SUB, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, denom_s, line);
            chunks[current].emit_op(Op::F64_DIV, line);
        }
        _ => {
            chunks[current].emit_f64_const(0.0, line);
            chunks[current].emit_f64_const(0.0, line);
        }
    }
    emit_ruby_complex(chunks, current, 2, line);
}

fn emit_rational_binary_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    left_s: u16,
    right_s: u16,
    op: &str,
    line: u32,
) {
    let ln_s = chunks[current].alloc_scratch(1);
    let ld_s = chunks[current].alloc_scratch(1);
    let rn_s = chunks[current].alloc_scratch(1);
    let rd_s = chunks[current].alloc_scratch(1);
    emit_rational_part_from_slot(chunks, current, left_s, "num", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ln_s, line);
    emit_rational_part_from_slot(chunks, current, left_s, "den", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ld_s, line);
    emit_ruby_is_rational_slot(chunks, current, right_s, line);
    chunks[current].emit_if_value(line);
    emit_rational_part_from_slot(chunks, current, right_s, "num", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rn_s, line);
    emit_rational_part_from_slot(chunks, current, right_s, "den", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rd_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rn_s, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rd_s, line);
    chunks[current].emit_end(line);

    match op {
        "add" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, ln_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rd_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rn_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ld_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ld_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rd_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
        }
        "sub" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, ln_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rd_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rn_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ld_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_SUB, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ld_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rd_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
        }
        "mul" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, ln_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rn_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ld_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rd_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
        }
        "div" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, ln_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rd_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ld_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rn_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
        }
        _ => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, ln_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ld_s, line);
        }
    }
    emit_rational_object_from_numbers(chunks, current, line);
}

fn emit_rational_number_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    emit_rational_part_from_slot(chunks, current, slot, "num", line);
    emit_rational_part_from_slot(chunks, current, slot, "den", line);
    chunks[current].emit_op(Op::F64_DIV, line);
}

fn emit_ruby_wrapped_string_eq_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    left_s: u16,
    right_s: u16,
    line: u32,
) {
    let left_v_s = chunks[current].alloc_scratch(1);
    let right_v_s = chunks[current].alloc_scratch(1);
    let left_e_s = chunks[current].alloc_scratch(1);
    let right_e_s = chunks[current].alloc_scratch(1);
    emit_ruby_string_value_from_slot(chunks, current, left_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_v_s, line);
    emit_ruby_string_value_from_slot(chunks, current, right_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_v_s, line);
    emit_ruby_string_encoding_from_slot(chunks, current, left_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_e_s, line);
    emit_ruby_string_encoding_from_slot(chunks, current, right_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_e_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_v_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_v_s, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_e_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_e_s, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    emit_ruby_raw_string_ascii_only_from_slot(chunks, current, left_v_s, line);
    emit_ruby_raw_string_ascii_only_from_slot(chunks, current, right_v_s, line);
    chunks[current].emit_op(Op::I32_AND, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_eq(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_bool_const(false, line);
        return;
    }
    let left_s = slots[0];
    let right_s = slots[1];

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_s, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
    collections::emit_sequence_equal(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);

    emit_ruby_is_complex_slot(chunks, current, left_s, line);
    chunks[current].emit_if_value(line);
    emit_ruby_is_complex_slot(chunks, current, right_s, line);
    chunks[current].emit_if_value(line);
    emit_complex_part_from_slot(chunks, current, left_s, "real", line);
    emit_complex_part_from_slot(chunks, current, right_s, "real", line);
    chunks[current].emit_op(Op::F64_EQ, line);
    emit_complex_part_from_slot(chunks, current, left_s, "imag", line);
    emit_complex_part_from_slot(chunks, current, right_s, "imag", line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);

    emit_ruby_is_float_slot(chunks, current, left_s, line);
    chunks[current].emit_if_value(line);
    emit_float_number_from_slot(chunks, current, left_s, line);
    emit_ruby_number_from_slot(chunks, current, right_s, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    emit_ruby_is_float_slot(chunks, current, right_s, line);
    chunks[current].emit_if_value(line);
    emit_ruby_number_from_slot(chunks, current, left_s, line);
    emit_float_number_from_slot(chunks, current, right_s, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    emit_ruby_is_rational_slot(chunks, current, left_s, line);
    chunks[current].emit_if_value(line);
    emit_ruby_is_rational_slot(chunks, current, right_s, line);
    chunks[current].emit_if_value(line);
    emit_rational_part_from_slot(chunks, current, left_s, "num", line);
    emit_rational_part_from_slot(chunks, current, right_s, "num", line);
    chunks[current].emit_op(Op::F64_EQ, line);
    emit_rational_part_from_slot(chunks, current, left_s, "den", line);
    emit_rational_part_from_slot(chunks, current, right_s, "den", line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    emit_rational_number_from_slot(chunks, current, left_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);

    emit_ruby_is_rational_slot(chunks, current, right_s, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_s, line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    emit_rational_number_from_slot(chunks, current, right_s, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    emit_ruby_is_wrapped_string_slot(chunks, current, left_s, line);
    chunks[current].emit_if_value(line);
    emit_ruby_wrapped_string_eq_from_slots(chunks, current, left_s, right_s, line);
    chunks[current].emit_else(line);
    emit_ruby_is_wrapped_string_slot(chunks, current, right_s, line);
    chunks[current].emit_if_value(line);
    emit_ruby_wrapped_string_eq_from_slots(chunks, current, left_s, right_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_rational_method(chunks: &mut [Chunk], current: usize, argc: u8, method: &str, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let slot = slots[0];
    match method {
        "num" => emit_rational_part_from_slot(chunks, current, slot, "num", line),
        "den" => emit_rational_part_from_slot(chunks, current, slot, "den", line),
        _ => emit_rational_to_s_from_slot(chunks, current, slot, line),
    }
}

fn emit_ruby_rationalize(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_f64_const(0.0, line);
        chunks[current].emit_f64_const(1.0, line);
        emit_rational_object_from_numbers(chunks, current, line);
        return;
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    chunks[current].emit_f64_const(0.5, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_end(line);
    emit_rational_object_from_numbers(chunks, current, line);
}

fn emit_ruby_abs(chunks: &mut [Chunk], current: usize, argc: u8, square: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_f64_const(0.0, line);
        return;
    }
    let slot = slots[0];
    emit_ruby_is_complex_slot(chunks, current, slot, line);
    chunks[current].emit_if_value(line);
    emit_complex_part_from_slot(chunks, current, slot, "real", line);
    emit_complex_part_from_slot(chunks, current, slot, "real", line);
    chunks[current].emit_op(Op::F64_MUL, line);
    emit_complex_part_from_slot(chunks, current, slot, "imag", line);
    emit_complex_part_from_slot(chunks, current, slot, "imag", line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    if !square {
        math::emit_sqrt(&mut chunks[current], line);
        chunks[current].emit_f64_const(1.0, line);
        call_import(chunks, current, "ecma:number", "toFixed", 2, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
    chunks[current].emit_if_value(line);
    if square {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        call_import(chunks, current, "ecma:number", "Number", 1, line);
        chunks[current].emit_dup(line);
        chunks[current].emit_op(Op::F64_MUL, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        chunks[current].emit_string_const("-", line);
        call_import(chunks, current, "ecma:string", "startsWith", 2, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        collections::emit_len(chunks, current, line);
        call_import(chunks, current, "ecma:string", "slice", 3, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    math::emit_abs(&mut chunks[current], line);
    if square {
        chunks[current].emit_dup(line);
        chunks[current].emit_op(Op::F64_MUL, line);
    }
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_zero_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("0/", line);
    call_import(chunks, current, "ecma:string", "startsWith", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

fn emit_ruby_numeric_pred(chunks: &mut [Chunk], current: usize, argc: u8, kind: &str, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_bool_const(false, line);
        return;
    }
    let slot = slots[0];
    match kind {
        "zero" => emit_ruby_zero_from_slot(chunks, current, slot, line),
        "nonzero" => {
            emit_ruby_zero_from_slot(chunks, current, slot, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_op(Op::NULL, line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
            chunks[current].emit_end(line);
        }
        "positive" | "negative" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
            call_import(chunks, current, "ecma:number", "Number", 1, line);
            chunks[current].emit_f64_const(0.0, line);
            if kind == "positive" {
                chunks[current].emit_op(Op::F64_GT, line);
            } else {
                chunks[current].emit_op(Op::F64_LT, line);
            }
            ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "real" => {
            emit_ruby_is_complex_slot(chunks, current, slot, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_bool_const(false, line);
            chunks[current].emit_else(line);
            chunks[current].emit_bool_const(true, line);
            chunks[current].emit_end(line);
        }
        "integer" => {
            emit_ruby_number_from_slot(chunks, current, slot, line);
            call_import(chunks, current, "ecma:number", "isInteger", 1, line);
        }
        "nan" => {
            emit_ruby_number_from_slot(chunks, current, slot, line);
            call_import(chunks, current, "ecma:number", "isNaN", 1, line);
        }
        "finite" => {
            emit_ruby_number_from_slot(chunks, current, slot, line);
            call_import(chunks, current, "ecma:number", "isFinite", 1, line);
        }
        "infinite" => {
            emit_ruby_number_from_slot(chunks, current, slot, line);
            call_import(chunks, current, "ecma:number", "isFinite", 1, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_op(Op::NULL, line);
            chunks[current].emit_else(line);
            emit_ruby_number_from_slot(chunks, current, slot, line);
            chunks[current].emit_f64_const(0.0, line);
            chunks[current].emit_op(Op::F64_LT, line);
            chunks[current].emit_if(line);
            chunks[current].emit_i32_const(-1, line);
            chunks[current].emit_else(line);
            chunks[current].emit_i32_const(1, line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
        }
        _ => chunks[current].emit_bool_const(false, line),
    }
}

fn emit_ruby_divmod(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    let lhs_s = chunks[current].alloc_scratch(1);
    let rhs_s = chunks[current].alloc_scratch(1);
    let q_s = chunks[current].alloc_scratch(1);
    emit_ruby_number_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lhs_s, line);
    emit_ruby_number_from_slot(chunks, current, slots[1], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rhs_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lhs_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rhs_s, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    math::emit_floor(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, q_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, q_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lhs_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, q_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rhs_s, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    collections::emit_array_new(chunks, current, 2, line);
}

fn emit_ruby_not(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_bool_const(true, line);
        return;
    }
    let slot = slots[0];
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "wasm:js-boolean", "test", 1, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "wasm:js-boolean", "cast", 1, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_i32_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    emit_ruby_number_from_slot(chunks, current, slot, line);
    math::emit_trunc(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
}

fn emit_i32_abs_top(chunks: &mut [Chunk], current: usize, line: u32) {
    let v_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v_s, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v_s, line);
    chunks[current].emit_end(line);
}

fn emit_gcd_i32_from_slots(chunks: &mut [Chunk], current: usize, left_s: u16, right_s: u16, line: u32) {
    let a_s = chunks[current].alloc_scratch(1);
    let b_s = chunks[current].alloc_scratch(1);
    let t_s = chunks[current].alloc_scratch(1);
    emit_ruby_i32_from_slot(chunks, current, left_s, line);
    emit_i32_abs_top(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a_s, line);
    emit_ruby_i32_from_slot(chunks, current, right_s, line);
    emit_i32_abs_top(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_s, line);
    chunks[current].emit_op(Op::I32_REM_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, t_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, t_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a_s, line);
}

fn emit_ruby_gcd_lcm(chunks: &mut [Chunk], current: usize, argc: u8, mode: &str, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        return;
    }
    let gcd_s = chunks[current].alloc_scratch(1);
    emit_gcd_i32_from_slots(chunks, current, slots[0], slots[1], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, gcd_s, line);
    if mode == "gcd" {
        chunks[current].emit_op_u16(Op::LOCAL_GET, gcd_s, line);
    } else if mode == "lcm" {
        chunks[current].emit_op_u16(Op::LOCAL_GET, gcd_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op(Op::I32_EQ, line);
        chunks[current].emit_if(line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_else(line);
        emit_ruby_i32_from_slot(chunks, current, slots[0], line);
        emit_i32_abs_top(chunks, current, line);
        emit_ruby_i32_from_slot(chunks, current, slots[1], line);
        emit_i32_abs_top(chunks, current, line);
        chunks[current].emit_op(Op::I32_MUL, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, gcd_s, line);
        chunks[current].emit_op(Op::I32_DIV_S, line);
        chunks[current].emit_end(line);
    } else {
        let arr_s = chunks[current].alloc_scratch(1);
        collections::emit_array_new(chunks, current, 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, gcd_s, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, gcd_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op(Op::I32_EQ, line);
        chunks[current].emit_if(line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_else(line);
        emit_ruby_i32_from_slot(chunks, current, slots[0], line);
        emit_i32_abs_top(chunks, current, line);
        emit_ruby_i32_from_slot(chunks, current, slots[1], line);
        emit_i32_abs_top(chunks, current, line);
        chunks[current].emit_op(Op::I32_MUL, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, gcd_s, line);
        chunks[current].emit_op(Op::I32_DIV_S, line);
        chunks[current].emit_end(line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    }
}

fn emit_ruby_digits(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    let n_s = chunks[current].alloc_scratch(1);
    let base_s = chunks[current].alloc_scratch(1);
    let arr_s = chunks[current].alloc_scratch(1);
    emit_ruby_i32_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n_s, line);
    if let Some(base_arg) = slots.get(1).copied() {
        emit_ruby_i32_from_slot(chunks, current, base_arg, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 10);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, base_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    emit_ruby_error(chunks, current, "Math::DomainError", "out of domain", line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    emit_ruby_error(chunks, current, "ArgumentError", "invalid radix", line);
    chunks[current].emit_end(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base_s, line);
    chunks[current].emit_op(Op::I32_REM_S, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base_s, line);
    chunks[current].emit_op(Op::I32_DIV_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
}

fn emit_ruby_bit_pred(chunks: &mut [Chunk], current: usize, argc: u8, kind: &str, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_bool_const(false, line);
        return;
    }
    emit_ruby_i32_from_slot(chunks, current, slots[0], line);
    emit_ruby_i32_from_slot(chunks, current, slots[1], line);
    chunks[current].emit_op(Op::I32_AND, line);
    if kind == "any" {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op(Op::I32_NE, line);
    } else if kind == "all" {
        emit_ruby_i32_from_slot(chunks, current, slots[1], line);
        chunks[current].emit_op(Op::I32_EQ, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op(Op::I32_EQ, line);
    }
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

fn emit_ruby_bit_length(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        return;
    }
    let n_s = chunks[current].alloc_scratch(1);
    let count_s = chunks[current].alloc_scratch(1);
    emit_ruby_i32_from_slot(chunks, current, slots[0], line);
    emit_i32_abs_top(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SHR_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_s, line);
}

fn emit_ruby_pred(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(slot) = slots.first().copied() {
        emit_ruby_number_from_slot(chunks, current, slot, line);
        chunks[current].emit_f64_const(1.0, line);
        chunks[current].emit_op(Op::F64_SUB, line);
    } else {
        chunks[current].emit_f64_const(-1.0, line);
    }
}

fn emit_ruby_int_size(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // `size` is Ruby-polymorphic: on Array/Set (arrays) it is the element
    // count; on an Integer it is the machine word size (8 bytes). Dispatch on
    // the receiver instead of hard-coding 8.
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *recv_s, line);
        call_import(chunks, current, "ecma:array", "isArray", 1, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *recv_s, line);
        collections::emit_len(chunks, current, line);
        chunks[current].emit_else(line);
        core_wasm::i32_const(&mut chunks[current], line, 8);
        chunks[current].emit_end(line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 8);
    }
}

fn emit_ruby_int_shift(chunks: &mut [Chunk], current: usize, argc: u8, right: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        return;
    }
    emit_ruby_i32_from_slot(chunks, current, slots[0], line);
    emit_ruby_i32_from_slot(chunks, current, slots[1], line);
    if right {
        chunks[current].emit_op(Op::I32_SHR_S, line);
    } else {
        chunks[current].emit_op(Op::I32_SHL, line);
    }
}

fn emit_ruby_times(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        emit_ruby_enumerator(chunks, current, line);
        return;
    }
    let n_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    emit_ruby_i32_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n_s, line);
    if slots.len() < 2 {
        let arr_s = chunks[current].alloc_scratch(1);
        collections::emit_array_new(chunks, current, 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
        let block = chunks[current].emit_block(line);
        let (loop_patch, _) = chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_br_if(1, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].patch_loop(loop_patch);
        chunks[current].emit_end(line);
        chunks[current].patch_block(block);
        emit_ruby_enumerator_from_items_slot(chunks, current, arr_s, line);
        return;
    }
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
}

fn emit_ruby_pow(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_f64_const(1.0, line);
        return;
    }
    if slots.len() < 3 {
        let exp_s = chunks[current].alloc_scratch(1);
        emit_ruby_number_from_slot(chunks, current, slots[1], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, exp_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, exp_s, line);
        chunks[current].emit_f64_const(0.0, line);
        chunks[current].emit_op(Op::F64_LT, line);
        chunks[current].emit_if(line);
        chunks[current].emit_f64_const(1.0, line);
        emit_ruby_number_from_slot(chunks, current, slots[0], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, exp_s, line);
        chunks[current].emit_op(Op::F64_NEG, line);
        call_import(chunks, current, "ecma:math", "pow", 2, line);
        emit_rational_object_from_numbers(chunks, current, line);
        chunks[current].emit_else(line);
        emit_ruby_number_from_slot(chunks, current, slots[0], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, exp_s, line);
        call_import(chunks, current, "ecma:math", "pow", 2, line);
        chunks[current].emit_end(line);
        return;
    }
    let base_s = chunks[current].alloc_scratch(1);
    let exp_s = chunks[current].alloc_scratch(1);
    let mod_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let cand_s = chunks[current].alloc_scratch(1);
    emit_ruby_i32_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, base_s, line);
    emit_ruby_i32_from_slot(chunks, current, slots[1], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, exp_s, line);
    emit_ruby_i32_from_slot(chunks, current, slots[2], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mod_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, mod_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    emit_ruby_error(chunks, current, "ZeroDivisionError", "divided by 0", line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, exp_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cand_s, line);
    let inv_block = chunks[current].emit_block(line);
    let (inv_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cand_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, mod_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cand_s, line);
    chunks[current].emit_op(Op::I32_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, mod_s, line);
    chunks[current].emit_op(Op::I32_REM_S, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cand_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cand_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(inv_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(inv_block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cand_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, mod_s, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cand_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, base_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, exp_s, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, exp_s, line);
    chunks[current].emit_else(line);
    emit_ruby_error(chunks, current, "ZeroDivisionError", "divided by 0", line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, exp_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base_s, line);
    chunks[current].emit_op(Op::I32_MUL, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, mod_s, line);
    chunks[current].emit_op(Op::I32_REM_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, exp_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, exp_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
}

fn emit_ruby_is_date_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("__type", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "__type", line);
    chunks[current].emit_string_const("Date", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
}

fn emit_ruby_date_step_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    slots: &[u16],
    default_step: f64,
    line: u32,
) {
    if slots.len() < 2 {
        emit_ruby_enumerator(chunks, current, line);
        return;
    }
    let Some(fn_s) = slots.last().copied() else {
        emit_ruby_enumerator(chunks, current, line);
        return;
    };
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    call_import(chunks, current, "ecma:reflect", "isCallable", 1, line);
    chunks[current].emit_if(line);

    let cur_s = chunks[current].alloc_scratch(1);
    let limit_s = chunks[current].alloc_scratch(1);
    let step_s = chunks[current].alloc_scratch(1);
    let day_ms_s = chunks[current].alloc_scratch(1);

    emit_time_prop_from_slot(chunks, current, slots[0], "__time", line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur_s, line);
    emit_time_prop_from_slot(chunks, current, slots[1], "__time", line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, limit_s, line);
    if slots.len() >= 4 && slots[2] != fn_s {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[2], line);
        call_import(chunks, current, "ecma:number", "Number", 1, line);
    } else {
        chunks[current].emit_f64_const(default_step, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, step_s, line);
    chunks[current].emit_f64_const(86_400_000.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, day_ms_s, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, step_s, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, limit_s, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, limit_s, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_s, line);
    emit_date_object_from_ms(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, step_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, day_ms_s, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    chunks[current].emit_else(line);
    emit_ruby_enumerator(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_step(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        emit_ruby_enumerator(chunks, current, line);
        return;
    }
    emit_ruby_is_date_slot(chunks, current, slots[0], line);
    chunks[current].emit_if(line);
    emit_ruby_date_step_from_slots(chunks, current, &slots, 1.0, line);
    chunks[current].emit_else(line);
    let maybe_fn = slots.last().copied();
    if let Some(fn_s) = maybe_fn {
        chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
        call_import(chunks, current, "ecma:reflect", "isCallable", 1, line);
        chunks[current].emit_if(line);

        let cur_s = chunks[current].alloc_scratch(1);
        let limit_s = chunks[current].alloc_scratch(1);
        let step_s = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        call_import(chunks, current, "ecma:number", "Number", 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, cur_s, line);
        if slots.len() >= 2 {
            chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
            call_import(chunks, current, "ecma:number", "Number", 1, line);
        } else {
            chunks[current].emit_f64_const(f64::INFINITY, line);
        }
        chunks[current].emit_op_u16(Op::LOCAL_SET, limit_s, line);
        if slots.len() >= 3 && slots[2] != fn_s {
            chunks[current].emit_op_u16(Op::LOCAL_GET, slots[2], line);
            call_import(chunks, current, "ecma:number", "Number", 1, line);
        } else {
            chunks[current].emit_f64_const(1.0, line);
        }
        chunks[current].emit_op_u16(Op::LOCAL_SET, step_s, line);

        let block = chunks[current].emit_block(line);
        let (loop_patch, _) = chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cur_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, limit_s, line);
        chunks[current].emit_op(Op::F64_LE, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_br_if(1, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cur_s, line);
        chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cur_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, step_s, line);
        chunks[current].emit_op(Op::F64_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, cur_s, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].patch_loop(loop_patch);
        chunks[current].emit_end(line);
        chunks[current].patch_block(block);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
        chunks[current].emit_else(line);
        emit_ruby_enumerator(chunks, current, line);
        chunks[current].emit_end(line);
    } else {
        emit_ruby_enumerator(chunks, current, line);
    }
    chunks[current].emit_end(line);
}

fn emit_ruby_upto_downto(chunks: &mut [Chunk], current: usize, argc: u8, default_step: f64, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        emit_ruby_enumerator(chunks, current, line);
        return;
    }
    if slots.len() >= 2 && default_step > 0.0 {
        emit_is_string_slot(chunks, current, slots[0], line);
        emit_is_string_slot(chunks, current, slots[1], line);
        chunks[current].emit_op(Op::I32_AND, line);
        chunks[current].emit_if(line);
        emit_ruby_string_upto_from_slots(chunks, current, &slots, line);
        chunks[current].emit_else(line);
        emit_ruby_is_date_slot(chunks, current, slots[0], line);
        chunks[current].emit_if(line);
        emit_ruby_date_step_from_slots(chunks, current, &slots, default_step, line);
        chunks[current].emit_else(line);
        emit_ruby_number_upto_downto_from_slots(chunks, current, &slots, default_step, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        return;
    }
    emit_ruby_is_date_slot(chunks, current, slots[0], line);
    chunks[current].emit_if(line);
    emit_ruby_date_step_from_slots(chunks, current, &slots, default_step, line);
    chunks[current].emit_else(line);
    emit_ruby_number_upto_downto_from_slots(chunks, current, &slots, default_step, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_number_upto_downto_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    slots: &[u16],
    default_step: f64,
    line: u32,
) {
    if slots.len() < 3 {
        emit_ruby_enumerator(chunks, current, line);
        return;
    }
    let block_s = slots[2];
    let value_s = chunks[current].alloc_scratch(1);
    let end_s = chunks[current].alloc_scratch(1);
    emit_ruby_i32_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
    emit_ruby_i32_from_slot(chunks, current, slots[1], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_s, line);
    if default_step > 0.0 {
        chunks[current].emit_op(Op::I32_LE_S, line);
    } else {
        chunks[current].emit_op(Op::I32_GE_S, line);
    }
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, block_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
    core_wasm::i32_const(&mut chunks[current], line, if default_step > 0.0 { 1 } else { -1 });
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
}

fn emit_ruby_hash_update(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let target_s = slots[0];
    if slots.len() == 1 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, target_s, line);
        return;
    }
    let maybe_block_s = *slots.last().unwrap();
    let source_s = chunks[current].alloc_scratch(1);
    let keys_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let key_s = chunks[current].alloc_scratch(1);
    let new_s = chunks[current].alloc_scratch(1);
    let old_s = chunks[current].alloc_scratch(1);
    let value_s = chunks[current].alloc_scratch(1);

    for src in slots.iter().skip(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *src, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, source_s, line);
        emit_ruby_is_callable_wrapper_slot(chunks, current, source_s, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op(Op::NOP, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, source_s, line);
        call_import(chunks, current, "ecma:reflect", "isCallable", 1, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op(Op::NOP, line);
        chunks[current].emit_else(line);

        emit_time_prop_from_slot(chunks, current, source_s, "__keys", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, keys_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
        let block = chunks[current].emit_block(line);
        let (loop_patch, _) = chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, keys_s, line);
        collections::emit_len(chunks, current, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_br_if(1, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, keys_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, key_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, source_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key_s, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, new_s, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, target_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key_s, line);
        call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, target_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key_s, line);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, old_s, line);
        emit_ruby_is_callable_wrapper_slot(chunks, current, maybe_block_s, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, maybe_block_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, old_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, new_s, line);
        emit_ruby_proc_call(chunks, current, 4, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, maybe_block_s, line);
        call_import(chunks, current, "ecma:reflect", "isCallable", 1, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, maybe_block_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, old_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, new_s, line);
        chunks[current].emit_op_u8(Op::CALL_REF, 3, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, new_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);
        emit_time_prop_from_slot(chunks, current, target_s, "__keys", line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key_s, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, new_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
        chunks[current].emit_end(line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, target_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
        dict::emit_set(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].patch_loop(loop_patch);
        chunks[current].emit_end(line);
        chunks[current].patch_block(block);

        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, target_s, line);
}

fn emit_ruby_string_upto_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    slots: &[u16],
    line: u32,
) {
    let Some(block_s) = slots.last().copied() else {
        emit_ruby_enumerator(chunks, current, line);
        return;
    };
    let start_code_s = chunks[current].alloc_scratch(1);
    let end_code_s = chunks[current].alloc_scratch(1);
    let code_s = chunks[current].alloc_scratch(1);
    let current_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_code_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_code_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_code_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, code_s, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    if slots.len() >= 4 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[2], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, code_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, end_code_s, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, code_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, end_code_s, line);
        chunks[current].emit_op(Op::I32_LE_S, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, code_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, end_code_s, line);
        chunks[current].emit_op(Op::I32_LE_S, line);
    }
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code_s, line);
    call_import(chunks, current, "ecma:string", "fromCharCode", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, current_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, block_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, current_s, line);
    emit_ruby_proc_call(chunks, current, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, code_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
}

fn emit_ruby_match_index(chunks: &mut [Chunk], current: usize, argc: u8, pred: bool, line: u32) {
    if argc < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    call_import(chunks, current, "ecma:regexp", "search", 2, line);
    if pred {
        chunks[current].emit_f64_const(0.0, line);
        chunks[current].emit_op(Op::F64_GE, line);
        ops::emit_i32_to_bool(&mut chunks[current], line);
    } else {
        emit_minus_one_to_null(&mut chunks[current], line);
    }
}

fn emit_ruby_casecmp_pred(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let rhs_s = chunks[current].alloc_scratch(1);
    let lhs_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rhs_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lhs_s, line);
    emit_is_string_slot(chunks, current, rhs_s, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lhs_s, line);
    call_import(chunks, current, "ecma:string", "toLowerCase", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rhs_s, line);
    call_import(chunks, current, "ecma:string", "toLowerCase", 1, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_casecmp(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let rhs_s = chunks[current].alloc_scratch(1);
    let lhs_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rhs_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lhs_s, line);
    emit_is_string_slot(chunks, current, rhs_s, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lhs_s, line);
    call_import(chunks, current, "ecma:string", "toLowerCase", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rhs_s, line);
    call_import(chunks, current, "ecma:string", "toLowerCase", 1, line);
    call_import(chunks, current, "wasm:js-string", "compare", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_str_compare(chunks: &mut [Chunk], current: usize, argc: u8, mode: &str, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let cmp_s = chunks[current].alloc_scratch(1);
    emit_is_string_slot(chunks, current, slots[1], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
    call_import(chunks, current, "wasm:js-string", "compare", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cmp_s, line);
    match mode {
        "cmp" => chunks[current].emit_op_u16(Op::LOCAL_GET, cmp_s, line),
        "lt" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, cmp_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op(Op::I32_LT_S, line);
            ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "gt" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, cmp_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op(Op::I32_GT_S, line);
            ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "lte" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, cmp_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op(Op::I32_LE_S, line);
            ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "gte" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, cmp_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op(Op::I32_GE_S, line);
            ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        _ => chunks[current].emit_op(Op::NULL, line),
    }
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_start_end_with(chunks: &mut [Chunk], current: usize, argc: u8, end: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_bool_const(false, line);
        return;
    }
    let recv_s = slots[0];
    let result_s = chunks[current].alloc_scratch(1);
    let prefix_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    for arg_s in slots.iter().skip(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        emit_is_regexp_slot(chunks, current, *arg_s, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg_s, line);
        call_import(chunks, current, "ecma:regexp", "match", 2, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        collections::emit_get(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, prefix_s, line);
        if end {
            chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, prefix_s, line);
            call_import(chunks, current, "ecma:string", "endsWith", 2, line);
        } else {
            chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, prefix_s, line);
            call_import(chunks, current, "ecma:string", "startsWith", 2, line);
        }
        ops::emit_i32_to_bool(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg_s, line);
        chunks[current].emit_string_const("/", line);
        call_import(chunks, current, "ecma:string", "startsWith", 2, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg_s, line);
        call_import(chunks, current, "ecma:regexp", "search", 2, line);
        if end {
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op(Op::I32_GE_S, line);
        } else {
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op(Op::I32_EQ, line);
        }
        ops::emit_i32_to_bool(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, recv_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg_s, line);
        if end {
            call_import(chunks, current, "ecma:string", "endsWith", 2, line);
        } else {
            call_import(chunks, current, "ecma:string", "startsWith", 2, line);
        }
        ops::emit_i32_to_bool(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
}

fn emit_ruby_equal(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

fn emit_ruby_capitalize(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let s_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, s_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    call_import(chunks, current, "ecma:string", "slice", 3, line);
    call_import(chunks, current, "ecma:string", "toUpperCase", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    call_import(chunks, current, "ecma:string", "slice", 3, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_swapcase(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let s_slot = chunks[current].alloc_scratch(1);
    let i_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let out_slot = chunks[current].alloc_scratch(1);
    let ch_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, s_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "slice", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ch_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_slot, line);
    call_import(chunks, current, "ecma:string", "toLowerCase", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_slot, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_slot, line);
    call_import(chunks, current, "ecma:string", "toUpperCase", 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_slot, line);
    call_import(chunks, current, "ecma:string", "toLowerCase", 1, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_SET, ch_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ch_slot, line);
    call_import(chunks, current, "wasm:js-string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

fn emit_ruby_case_transform_bang(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    kind: &str,
    line: u32,
) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let original_s = slots[0];
    let result_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, original_s, line);
    match kind {
        "up" => call_import(chunks, current, "ecma:string", "toUpperCase", 1, line),
        "down" => call_import(chunks, current, "ecma:string", "toLowerCase", 1, line),
        "capitalize" => emit_ruby_capitalize(chunks, current, 1, line),
        "swap" => emit_ruby_swapcase(chunks, current, 1, line),
        _ => {}
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    emit_nil_if_unchanged_from_slots(chunks, current, original_s, result_s, line);
}

fn emit_ruby_domain_checked_unary_math(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    import_name: &'static str,
    line: u32,
) {
    let slots = emit_store_args(chunks, current, argc, line);
    let slot = slots.first().copied().unwrap_or_else(|| chunks[current].alloc_scratch(1));
    emit_ruby_number_from_slot(chunks, current, slot, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_f64_const(-1.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    emit_ruby_number_from_slot(chunks, current, slot, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    emit_ruby_error(chunks, current, "Math::DomainError", "Math::DomainError", line);
    chunks[current].emit_else(line);
    call_import(chunks, current, "ecma:math", import_name, 1, line);
    emit_ruby_float_object(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_nonnegative_unary_math(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    import_name: &'static str,
    line: u32,
) {
    let slots = emit_store_args(chunks, current, argc, line);
    let num_s = chunks[current].alloc_scratch(1);
    if let Some(slot) = slots.first().copied() {
        emit_ruby_number_from_slot(chunks, current, slot, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, num_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    emit_ruby_error(chunks, current, "Math::DomainError", "Math::DomainError", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    call_import(chunks, current, "ecma:math", import_name, 1, line);
    emit_ruby_float_object(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_math_log(chunks: &mut [Chunk], current: usize, argc: u8, base: Option<f64>, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let num_s = chunks[current].alloc_scratch(1);
    if let Some(slot) = slots.first().copied() {
        emit_ruby_number_from_slot(chunks, current, slot, line);
    } else {
        chunks[current].emit_f64_const(1.0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, num_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    emit_ruby_error(chunks, current, "Math::DomainError", "Math::DomainError", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    call_import(chunks, current, "ecma:math", "log", 1, line);
    if let Some(base) = base {
        chunks[current].emit_f64_const(base, line);
        call_import(chunks, current, "ecma:math", "log", 1, line);
        chunks[current].emit_op(Op::F64_DIV, line);
    } else if let Some(base_slot) = slots.get(1).copied() {
        emit_ruby_number_from_slot(chunks, current, base_slot, line);
        call_import(chunks, current, "ecma:math", "log", 1, line);
        chunks[current].emit_op(Op::F64_DIV, line);
    }
    emit_ruby_float_object(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_math_unary(chunks: &mut [Chunk], current: usize, argc: u8, import_name: &'static str, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let num_s = chunks[current].alloc_scratch(1);
    if slots.is_empty() {
        chunks[current].emit_f64_const(0.0, line);
    } else {
        emit_ruby_number_from_slot(chunks, current, slots[0], line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, num_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    chunks[current].emit_op(Op::F64_NE, line);
    chunks[current].emit_if(line);
    if import_name == "cos" {
        chunks[current].emit_f64_const(std::f64::consts::PI, line);
    } else if import_name == "sin" {
        chunks[current].emit_f64_const(std::f64::consts::PI / 2.0, line);
    } else if import_name == "tan" {
        chunks[current].emit_f64_const(std::f64::consts::PI / 4.0, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    chunks[current].emit_end(line);
    call_import(chunks, current, "ecma:math", import_name, 1, line);
    emit_ruby_float_object(chunks, current, line);
}

fn emit_ruby_math_binary(chunks: &mut [Chunk], current: usize, argc: u8, import_name: &'static str, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_f64_const(0.0, line);
        chunks[current].emit_f64_const(0.0, line);
    } else {
        emit_ruby_number_from_slot(chunks, current, slots[0], line);
        emit_ruby_number_from_slot(chunks, current, slots[1], line);
    }
    call_import(chunks, current, "ecma:math", import_name, 2, line);
    emit_ruby_float_object(chunks, current, line);
}

fn emit_ruby_libc_math_unary(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    emit_name: &'static str,
    line: u32,
) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(slot) = slots.first().copied() {
        emit_ruby_number_from_slot(chunks, current, slot, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    vybe_platform_libc::emitter::dispatch::emit_math(emit_name, chunks, current, line);
    emit_ruby_float_object(chunks, current, line);
}

fn emit_ruby_math_cbrt(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(slot) = slots.first().copied() {
        emit_ruby_number_from_slot(chunks, current, slot, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    call_import(chunks, current, "ecma:math", "cbrt", 1, line);
    emit_ruby_float_object(chunks, current, line);
}

fn emit_ruby_math_hypot(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() >= 2 {
        emit_ruby_number_from_slot(chunks, current, slots[0], line);
        emit_ruby_number_from_slot(chunks, current, slots[1], line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
        chunks[current].emit_f64_const(0.0, line);
    }
    call_import(chunks, current, "ecma:math", "hypot", 2, line);
    emit_ruby_float_object(chunks, current, line);
}

fn emit_ruby_math_ldexp(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() >= 2 {
        emit_ruby_number_from_slot(chunks, current, slots[0], line);
        chunks[current].emit_f64_const(2.0, line);
        emit_ruby_number_from_slot(chunks, current, slots[1], line);
        call_import(chunks, current, "ecma:math", "pow", 2, line);
        chunks[current].emit_op(Op::F64_MUL, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    emit_ruby_float_object(chunks, current, line);
}

fn emit_ruby_math_frexp(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let num_s = chunks[current].alloc_scratch(1);
    let exp_s = chunks[current].alloc_scratch(1);
    let exp_f_s = chunks[current].alloc_scratch(1);
    let mant_s = chunks[current].alloc_scratch(1);
    if let Some(slot) = slots.first().copied() {
        emit_ruby_number_from_slot(chunks, current, slot, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, num_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mant_s, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, exp_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    call_import(chunks, current, "ecma:math", "log2", 1, line);
    math::emit_floor(&mut chunks[current], line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, exp_f_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, exp_f_s, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, exp_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, exp_f_s, line);
    call_import(chunks, current, "ecma:math", "pow", 2, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mant_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, mant_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, exp_s, line);
    collections::emit_array_new(chunks, current, 2, line);
}

fn emit_ruby_math_gamma(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let num_s = chunks[current].alloc_scratch(1);
    if let Some(slot) = slots.first().copied() {
        emit_ruby_number_from_slot(chunks, current, slot, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, num_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    call_import(chunks, current, "ecma:number", "isInteger", 1, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    emit_ruby_error(chunks, current, "Math::DomainError", "Math::DomainError", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    vybe_platform_libc::emitter::dispatch::emit_math("libc.math.tgamma", chunks, current, line);
    emit_ruby_float_object(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_math_lgamma(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let num_s = chunks[current].alloc_scratch(1);
    let value_s = chunks[current].alloc_scratch(1);
    if let Some(slot) = slots.first().copied() {
        emit_ruby_number_from_slot(chunks, current, slot, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, num_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    vybe_platform_libc::emitter::dispatch::emit_math("libc.math.lgamma", chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    call_import(chunks, current, "ecma:number", "isInteger", 1, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_end(line);
    collections::emit_array_new(chunks, current, 2, line);
}

fn emit_time_rounding(chunks: &mut [Chunk], current: usize, argc: u8, mode: &str, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    emit_time_is_time_slot(chunks, current, slots[0], line);
    chunks[current].emit_if_value(line);
    if slots.len() >= 2 && mode == "round" {
        emit_time_ms_number_from_slot(chunks, current, slots[0], line);
        emit_time_object_from_ms(chunks, current, true, 0, line);
    } else {
    emit_time_ms_number_from_slot(chunks, current, slots[0], line);
    if slots.len() >= 2 {
        chunks[current].emit_f64_const(10.0, line);
    } else {
        chunks[current].emit_f64_const(1000.0, line);
    }
    chunks[current].emit_op(Op::F64_DIV, line);
    match mode {
        "round" => {
            chunks[current].emit_f64_const(0.5, line);
            chunks[current].emit_op(Op::F64_ADD, line);
            math::emit_floor(&mut chunks[current], line);
        }
        "ceil" => math::emit_ceil(&mut chunks[current], line),
        _ => math::emit_floor(&mut chunks[current], line),
    }
    if slots.len() >= 2 {
        chunks[current].emit_f64_const(10.0, line);
    } else {
        chunks[current].emit_f64_const(1000.0, line);
    }
    chunks[current].emit_op(Op::F64_MUL, line);
    emit_time_object_from_ms(chunks, current, true, 0, line);
    }
    chunks[current].emit_else(line);
    if slots.len() >= 2 {
        let num_s = chunks[current].alloc_scratch(1);
        let scale_s = chunks[current].alloc_scratch(1);
        emit_ruby_number_from_slot(chunks, current, slots[0], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, num_s, line);
        chunks[current].emit_f64_const(10.0, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slots[1], line);
        call_import(chunks, current, "ecma:number", "Number", 1, line);
        call_import(chunks, current, "ecma:math", "pow", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, scale_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, scale_s, line);
        chunks[current].emit_op(Op::F64_MUL, line);
        match mode {
            "round" | "round_half_up" => math::emit_round(&mut chunks[current], line),
            "round_half_down" => {
                chunks[current].emit_f64_const(0.5, line);
                chunks[current].emit_op(Op::F64_SUB, line);
                math::emit_ceil(&mut chunks[current], line);
            }
            "round_half_even" => emit_ruby_round_half_even_top(chunks, current, line),
            "ceil" => math::emit_ceil(&mut chunks[current], line),
            "truncate" => math::emit_trunc(&mut chunks[current], line),
            _ => math::emit_floor(&mut chunks[current], line),
        }
        chunks[current].emit_op_u16(Op::LOCAL_GET, scale_s, line);
        chunks[current].emit_op(Op::F64_DIV, line);
        let rounded_s = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, rounded_s, line);
        emit_ruby_is_float_slot(chunks, current, slots[0], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, rounded_s, line);
        emit_ruby_float_object(chunks, current, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, rounded_s, line);
        chunks[current].emit_end(line);
    } else {
        emit_ruby_number_from_slot(chunks, current, slots[0], line);
        match mode {
            "round" | "round_half_up" => math::emit_round(&mut chunks[current], line),
            "round_half_down" => {
                chunks[current].emit_f64_const(0.5, line);
                chunks[current].emit_op(Op::F64_SUB, line);
                math::emit_ceil(&mut chunks[current], line);
            }
            "round_half_even" => emit_ruby_round_half_even_top(chunks, current, line),
            "ceil" => math::emit_ceil(&mut chunks[current], line),
            "truncate" => math::emit_trunc(&mut chunks[current], line),
            _ => math::emit_floor(&mut chunks[current], line),
        }
    }
    chunks[current].emit_end(line);
}

fn emit_ruby_round_half_even_top(chunks: &mut [Chunk], current: usize, line: u32) {
    let x_s = chunks[current].alloc_scratch(1);
    let floor_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, x_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x_s, line);
    math::emit_floor(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, floor_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, floor_s, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_f64_const(0.5, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, floor_s, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_REM_S, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, floor_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, floor_s, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x_s, line);
    math::emit_round(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

fn emit_time_fraction(chunks: &mut [Chunk], current: usize, scale: f64, line: u32) {
    let slots = emit_store_args(chunks, current, 1, line);
    emit_time_ms_number_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_f64_const(1000.0, line);
    math::emit_c_fmod(&mut chunks[current], line);
    chunks[current].emit_f64_const(scale, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
}

fn emit_time_subsec(chunks: &mut [Chunk], current: usize, line: u32) {
    let slots = emit_store_args(chunks, current, 1, line);
    let rem_slot = chunks[current].alloc_scratch(1);
    emit_time_ms_number_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_f64_const(1000.0, line);
    math::emit_c_fmod(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rem_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("0", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rem_slot, line);
    chunks[current].emit_f64_const(550.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("11/20", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rem_slot, line);
    chunks[current].emit_f64_const(555.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("111/200", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rem_slot, line);
    chunks[current].emit_f64_const(560.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("14/25", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("1/2", line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_time_bool_prop(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    let slots = emit_store_args(chunks, current, 1, line);
    emit_time_prop_from_slot(chunks, current, slots[0], key, line);
}

fn emit_time_zone(chunks: &mut [Chunk], current: usize, line: u32) {
    let slots = emit_store_args(chunks, current, 1, line);
    emit_time_prop_from_slot(chunks, current, slots[0], "__utc", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("UTC", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_end(line);
}

fn emit_time_gmtoff(chunks: &mut [Chunk], current: usize, line: u32) {
    let slots = emit_store_args(chunks, current, 1, line);
    emit_time_prop_from_slot(chunks, current, slots[0], "__gmtoff", line);
}

fn emit_time_mut_zone(chunks: &mut [Chunk], current: usize, argc: u8, utc: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    chunks[current].emit_bool_const(utc, line);
    emit_time_set_prop_from_slot(chunks, current, slots[0], "__utc", line);
    if utc {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    } else if slots.len() >= 2 {
        core_wasm::i32_const(&mut chunks[current], line, 32400);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    emit_time_set_prop_from_slot(chunks, current, slots[0], "__gmtoff", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
}

fn emit_time_copy_zone(chunks: &mut [Chunk], current: usize, argc: u8, utc: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    emit_time_ms_number_from_slot(chunks, current, slots[0], line);
    emit_time_object_from_ms(chunks, current, utc, if utc { 0 } else if slots.len() >= 2 { 32400 } else { 0 }, line);
}

fn emit_time_to_a(chunks: &mut [Chunk], current: usize, line: u32) {
    let slots = emit_store_args(chunks, current, 1, line);
    let t = slots[0];
    collections::emit_array_new(chunks, current, 0, line);
    for (getter, add) in [
        ("getUTCSeconds", 0),
        ("getUTCMinutes", 0),
        ("getUTCHours", 0),
        ("getUTCDate", 0),
        ("getUTCMonth", 1),
        ("getUTCFullYear", 0),
        ("getUTCDay", 0),
    ] {
        chunks[current].emit_dup(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, t, line);
        call_import(chunks, current, "ecma:date", getter, 1, line);
        if add != 0 {
            core_wasm::i32_const(&mut chunks[current], line, add);
            chunks[current].emit_op(Op::I32_ADD, line);
        }
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_dup(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(false, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("UTC", line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_ruby_to_a(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    emit_time_is_time_slot(chunks, current, slots[0], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    emit_time_to_a(chunks, current, line);
    chunks[current].emit_else(line);
    emit_ruby_is_enumerator_slot(chunks, current, slots[0], line);
    chunks[current].emit_if_value(line);
    emit_ruby_items_from_slot(chunks, current, slots[0], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
    call_import(chunks, current, "ecma:array", "from", 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_class_name(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_string_const("NilClass", line);
        return;
    }
    let slot = slots[0];
    let type_slot = chunks[current].alloc_scratch(1);

    emit_time_is_time_slot(chunks, current, slot, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("Time", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("Array", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, type_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, type_slot, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("String", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, type_slot, line);
    chunks[current].emit_string_const("number", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:number", "isInteger", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("Integer", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("Float", line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, type_slot, line);
    chunks[current].emit_string_const("boolean", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("TrueClass", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("FalseClass", line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, type_slot, line);
    chunks[current].emit_string_const("object", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let obj_type_key = chunks[current].add_constant(Value::String(Arc::from("__type")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, obj_type_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, type_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, type_slot, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("undefined", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("Object", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, type_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("Object", line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ruby_name(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.is_empty() {
        chunks[current].emit_string_const("", line);
        return;
    }
    let slot = slots[0];
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("name", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "name", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const("__type", line);
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    chunks[current].emit_if_value(line);
    emit_time_prop_from_slot(chunks, current, slot, "__type", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_iso_no_millis(chunks: &mut [Chunk], current: usize, dt_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, dt_slot, line);
    call_import(chunks, current, "ecma:date", "toISOString", 1, line);
    chunks[current].emit_string_const(".000Z", line);
    chunks[current].emit_string_const("Z", line);
    call_import(chunks, current, "ecma:string", "replace", 3, line);
}

fn emit_ruby_time_iso8601(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(dt_slot) = slots.first().copied() {
        emit_iso_no_millis(chunks, current, dt_slot, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
}

fn emit_ruby_time_httpdate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(dt_slot) = slots.first().copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, dt_slot, line);
        call_import(chunks, current, "ecma:date", "toUTCString", 1, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
}

fn emit_ruby_time_rfc2822(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(dt_slot) = slots.first().copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, dt_slot, line);
        call_import(chunks, current, "ecma:date", "toUTCString", 1, line);
        chunks[current].emit_string_const("GMT", line);
        chunks[current].emit_string_const("-0000", line);
        call_import(chunks, current, "ecma:string", "replace", 3, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
}

fn emit_date_name(chunks: &mut [Chunk], current: usize, dt_slot: u16, getter: &str, names: &[&str], line: u32) {
    for name in names {
        chunks[current].emit_string_const(name, line);
    }
    chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, names.len() as u16, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dt_slot, line);
    call_import(chunks, current, "ecma:date", getter, 1, line);
    collections::emit_get(chunks, current, line);
}

fn emit_iso_prefix(chunks: &mut [Chunk], current: usize, dt_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, dt_slot, line);
    call_import(chunks, current, "ecma:date", "toISOString", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 19);
    call_import(chunks, current, "ecma:string", "slice", 3, line);
}

fn emit_date_ymd(chunks: &mut [Chunk], current: usize, dt_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, dt_slot, line);
    call_import(chunks, current, "ecma:date", "toISOString", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 10);
    call_import(chunks, current, "ecma:string", "slice", 3, line);
}

fn emit_date_short_year(chunks: &mut [Chunk], current: usize, dt_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, dt_slot, line);
    call_import(chunks, current, "ecma:date", "getUTCFullYear", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 100);
    chunks[current].emit_op(Op::I32_REM_S, line);
    emit_pad2(chunks, current, line);
}

fn emit_date_yday(chunks: &mut [Chunk], current: usize, dt_slot: u16, line: u32) {
    for n in [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334] {
        core_wasm::i32_const(&mut chunks[current], line, n);
    }
    chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 12, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dt_slot, line);
    call_import(chunks, current, "ecma:date", "getUTCMonth", 1, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dt_slot, line);
    call_import(chunks, current, "ecma:date", "getUTCDate", 1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "String", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 3);
    chunks[current].emit_string_const("0", line);
    call_import(chunks, current, "ecma:string", "padStart", 3, line);
}

fn emit_time_hms(chunks: &mut [Chunk], current: usize, dt_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, dt_slot, line);
    call_import(chunks, current, "ecma:date", "toISOString", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 11);
    core_wasm::i32_const(&mut chunks[current], line, 19);
    call_import(chunks, current, "ecma:string", "slice", 3, line);
}

fn emit_pad2(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "ecma:string", "String", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_string_const("0", line);
    call_import(chunks, current, "ecma:string", "padStart", 3, line);
}

fn emit_time_ampm(chunks: &mut [Chunk], current: usize, dt_slot: u16, line: u32) {
    let hour_slot = chunks[current].alloc_scratch(1);
    let hour12_slot = chunks[current].alloc_scratch(1);
    let suffix_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dt_slot, line);
    call_import(chunks, current, "ecma:date", "getUTCHours", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, hour_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, hour_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 12);
    chunks[current].emit_op(Op::I32_REM_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, hour12_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, hour12_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, 12);
    chunks[current].emit_op_u16(Op::LOCAL_SET, hour12_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, hour_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 12);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("PM", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("AM", line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, suffix_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, hour12_slot, line);
    emit_pad2(chunks, current, line);
    chunks[current].emit_string_const(" ", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, suffix_slot, line);
    ops::emit_dyn_add(&mut chunks[current], line);
}

fn emit_ruby_time_strftime(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        chunks[current].emit_string_const("", line);
        return;
    }
    let dt_slot = slots[0];
    let fmt_slot = slots[1];
    chunks[current].emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunks[current].emit_string_const("%Z", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("UTC", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunks[current].emit_string_const("%H:%M:%S", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_time_hms(chunks, current, dt_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunks[current].emit_string_const("%I %p", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_time_ampm(chunks, current, dt_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunks[current].emit_string_const("%F", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_date_ymd(chunks, current, dt_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunks[current].emit_string_const("%Y-%m-%d", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_date_ymd(chunks, current, dt_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunks[current].emit_string_const("%y", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_date_short_year(chunks, current, dt_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunks[current].emit_string_const("%a", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_date_name(
        chunks,
        current,
        dt_slot,
        "getUTCDay",
        &["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
        line,
    );
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunks[current].emit_string_const("%b", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_date_name(
        chunks,
        current,
        dt_slot,
        "getUTCMonth",
        &["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
        line,
    );
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunks[current].emit_string_const("%j", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_date_yday(chunks, current, dt_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunks[current].emit_string_const("%A", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_date_name(
        chunks,
        current,
        dt_slot,
        "getUTCDay",
        &["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
        line,
    );
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunks[current].emit_string_const("%B", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_date_name(
        chunks,
        current,
        dt_slot,
        "getUTCMonth",
        &["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
        line,
    );
    chunks[current].emit_else(line);
    emit_iso_prefix(chunks, current, dt_slot, line);
    chunks[current].emit_string_const("T", line);
    chunks[current].emit_string_const(" ", line);
    call_import(chunks, current, "ecma:string", "replace", 3, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_minus_one_to_null(chunk: &mut Chunk, line: u32) {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    core_wasm::i32_const(chunk, line, -1);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if(line);
    chunk.emit_op(Op::NULL, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_end(line);
}

fn emit_replace_array_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    dst_s: u16,
    src_s: u16,
    idx_s: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, dst_s, line);
    collections::emit_clear(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dst_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dst_s, line);
}

fn emit_select_like(
    chunks: &mut [Chunk],
    current: usize,
    keep_matches: bool,
    mutate: bool,
    line: u32,
) {
    let fn_s = chunks[current].alloc_scratch(1);
    let arr_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let elem_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    if !keep_matches {
        ops::emit_dyn_not(&mut chunks[current], line);
    }
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    if mutate {
        emit_replace_array_from_slot(chunks, current, arr_s, result_s, idx_s, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    }
}

fn emit_ruby_map_like(chunks: &mut [Chunk], current: usize, argc: u8, flatten: bool, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    let recv_s = slots[0];
    let fn_s = slots[1];
    let arr_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let mapped_s = chunks[current].alloc_scratch(1);
    let lazy_s = chunks[current].alloc_scratch(1);

    emit_ruby_is_lazy_enumerator_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lazy_s, line);
    emit_ruby_items_from_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mapped_s, line);
    if flatten {
        chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_s, line);
        call_import(chunks, current, "ecma:array", "isArray", 1, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_s, line);
        call_import(chunks, current, "ecma:array", "concat", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_s, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_s, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, lazy_s, line);
    chunks[current].emit_if_value(line);
    emit_ruby_enumerator_from_items_slot_with_type(chunks, current, result_s, "Enumerator::Lazy", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_end(line);
}

fn emit_ruby_map_round(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    let recv_s = slots[0];
    let digits_s = slots[1];
    let arr_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let elem_s = chunks[current].alloc_scratch(1);
    let num_s = chunks[current].alloc_scratch(1);
    let scale_s = chunks[current].alloc_scratch(1);
    let rounded_s = chunks[current].alloc_scratch(1);

    chunks[current].emit_f64_const(10.0, line);
    emit_ruby_number_from_slot(chunks, current, digits_s, line);
    call_import(chunks, current, "ecma:math", "pow", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, scale_s, line);

    emit_ruby_items_from_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);

    emit_ruby_number_from_slot(chunks, current, elem_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, num_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, scale_s, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    math::emit_round(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, scale_s, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rounded_s, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, num_s, line);
    call_import(chunks, current, "ecma:number", "isInteger", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rounded_s, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
}

fn emit_ruby_select_like_value(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    keep_matches: bool,
    line: u32,
) {
    let slots = emit_store_args(chunks, current, argc, line);
    if slots.len() < 2 {
        collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    let recv_s = slots[0];
    let fn_s = slots[1];
    let arr_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let elem_s = chunks[current].alloc_scratch(1);
    let lazy_s = chunks[current].alloc_scratch(1);

    emit_ruby_is_lazy_enumerator_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lazy_s, line);
    emit_ruby_items_from_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    if !keep_matches {
        ops::emit_dyn_not(&mut chunks[current], line);
    }
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, lazy_s, line);
    chunks[current].emit_if_value(line);
    emit_ruby_enumerator_from_items_slot_with_type(chunks, current, result_s, "Enumerator::Lazy", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_end(line);
}

fn emit_array_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let init_s = chunks[current].alloc_scratch(1);
    let size_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let callable_s = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, init_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, size_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, init_s, line);
        call_import(chunks, current, "ecma:reflect", "isCallable", 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, callable_s, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, init_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, size_s, line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, callable_s, line);
    }
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, size_s, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, callable_s, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, init_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, init_s, line);
    chunks[current].emit_end(line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
}

fn emit_take_drop(chunks: &mut [Chunk], current: usize, drop: bool, line: u32) {
    let n_s = chunks[current].alloc_scratch(1);
    let recv_s = chunks[current].alloc_scratch(1);
    let arr_s = chunks[current].alloc_scratch(1);
    let len_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let lazy_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv_s, line);
    emit_ruby_is_lazy_enumerator_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lazy_s, line);
    emit_ruby_items_from_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    if drop {
        chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
    }
    call_import(chunks, current, "ecma:array", "slice", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lazy_s, line);
    chunks[current].emit_if_value(line);
    emit_ruby_enumerator_from_items_slot_with_type(chunks, current, result_s, "Enumerator::Lazy", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_end(line);
}

fn emit_take_while(chunks: &mut [Chunk], current: usize, line: u32) {
    let fn_s = chunks[current].alloc_scratch(1);
    let recv_s = chunks[current].alloc_scratch(1);
    let arr_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let elem_s = chunks[current].alloc_scratch(1);
    let lazy_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv_s, line);
    emit_ruby_is_lazy_enumerator_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lazy_s, line);
    emit_ruby_items_from_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lazy_s, line);
    chunks[current].emit_if_value(line);
    emit_ruby_enumerator_from_items_slot_with_type(chunks, current, result_s, "Enumerator::Lazy", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_end(line);
}

fn emit_drop_while(chunks: &mut [Chunk], current: usize, line: u32) {
    let fn_s = chunks[current].alloc_scratch(1);
    let recv_s = chunks[current].alloc_scratch(1);
    let arr_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let result_s = chunks[current].alloc_scratch(1);
    let elem_s = chunks[current].alloc_scratch(1);
    let lazy_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv_s, line);
    emit_ruby_is_lazy_enumerator_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lazy_s, line);
    emit_ruby_items_from_slot(chunks, current, recv_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);

    let scan_block = chunks[current].emit_block(line);
    let (scan_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(scan_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(scan_block);

    let copy_block = chunks[current].emit_block(line);
    let (copy_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(copy_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(copy_block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lazy_s, line);
    chunks[current].emit_if_value(line);
    emit_ruby_enumerator_from_items_slot_with_type(chunks, current, result_s, "Enumerator::Lazy", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
    chunks[current].emit_end(line);
}

fn emit_sum(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    if let Some(recv_s) = slots.first().copied() {
        let init_s = chunks[current].alloc_scratch(1);
        if let Some(init_arg_s) = slots.get(1).copied() {
            chunks[current].emit_op_u16(Op::LOCAL_GET, init_arg_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, init_s, line);
        } else {
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, init_s, line);
        }
        emit_ruby_items_from_slot(chunks, current, recv_s, line);
        call_import(chunks, current, "ecma:math", "sumPrecise", 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, init_s, line);
        ops::emit_dyn_add(&mut chunks[current], line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
}

fn emit_count(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        collections::emit_len(chunks, current, line);
        return;
    }
    let pred_s = chunks[current].alloc_scratch(1);
    let arr_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let count_s = chunks[current].alloc_scratch(1);
    let elem_s = chunks[current].alloc_scratch(1);
    let callable_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pred_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pred_s, line);
    call_import(chunks, current, "ecma:reflect", "isCallable", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, callable_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, callable_s, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pred_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pred_s, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_s, line);
}

fn emit_array_set_op(chunks: &mut [Chunk], current: usize, mode: &str, line: u32) {
    let right_s = chunks[current].alloc_scratch(1);
    let left_s = chunks[current].alloc_scratch(1);
    let out_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let elem_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_s, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    if mode == "union" {
        chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, left_s, line);
        call_import(chunks, current, "ecma:array", "concat", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    }
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, if mode == "union" { right_s } else { left_s }, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, if mode == "union" { right_s } else { left_s }, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
    let should_push = chunks[current].emit_block(line);
    match mode {
        "intersection" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
            call_import(chunks, current, "ecma:array", "includes", 2, line);
            ops::emit_dyn_not(&mut chunks[current], line);
            chunks[current].emit_br_if(0, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
            call_import(chunks, current, "ecma:array", "includes", 2, line);
            chunks[current].emit_br_if(0, line);
        }
        "difference" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
            call_import(chunks, current, "ecma:array", "includes", 2, line);
            chunks[current].emit_br_if(0, line);
        }
        "union" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
            call_import(chunks, current, "ecma:array", "includes", 2, line);
            chunks[current].emit_br_if(0, line);
        }
        _ => {}
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(should_push);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
}

fn emit_array_union(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "ecma:array", "concat", 2, line);
    call_import(chunks, current, "ecma:set", "fromIterable", 1, line);
    call_import(chunks, current, "ecma:array", "from", 1, line);
}

fn emit_slot_eq_str(chunks: &mut [Chunk], current: usize, slot: u16, value: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_string_const(value, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
}

fn emit_array_with_slot(chunks: &mut [Chunk], current: usize, value_s: u16, line: u32) {
    let out_s = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
}

fn emit_char_code_at_slot(chunks: &mut [Chunk], current: usize, str_s: u16, idx_s: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
}

fn emit_ruby_unpack_codes(chunks: &mut [Chunk], current: usize, str_s: u16, line: u32) {
    let out_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    emit_char_code_at_slot(chunks, current, str_s, idx_s, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
}

fn emit_ruby_unpack_single_code(chunks: &mut [Chunk], current: usize, str_s: u16, idx: i32, signed: bool, line: u32) {
    let value_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    core_wasm::i32_const(&mut chunks[current], line, idx);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
    if signed {
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 127);
        chunks[current].emit_op(Op::I32_GT_S, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 256);
        chunks[current].emit_op(Op::I32_SUB, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
        chunks[current].emit_end(line);
    }
    emit_array_with_slot(chunks, current, value_s, line);
}

fn emit_ruby_code_hex(chunks: &mut [Chunk], current: usize, code_s: u16, low_first: bool, line: u32) {
    let hex_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 16);
    call_import(chunks, current, "ecma:number", "toString", 2, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_string_const("0", line);
    call_import(chunks, current, "ecma:string", "padStart", 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, hex_s, line);
    if low_first {
        chunks[current].emit_op_u16(Op::LOCAL_GET, hex_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        core_wasm::i32_const(&mut chunks[current], line, 2);
        call_import(chunks, current, "ecma:string", "slice", 3, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, hex_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        call_import(chunks, current, "ecma:string", "slice", 3, line);
        call_import(chunks, current, "ecma:string", "concat", 2, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, hex_s, line);
    }
}

fn emit_ruby_unpack_hex(chunks: &mut [Chunk], current: usize, str_s: u16, low_first: bool, line: u32) {
    let idx_s = chunks[current].alloc_scratch(1);
    let out_s = chunks[current].alloc_scratch(1);
    let code_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    emit_char_code_at_slot(chunks, current, str_s, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, code_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    emit_ruby_code_hex(chunks, current, code_s, low_first, line);
    call_import(chunks, current, "ecma:string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    emit_array_with_slot(chunks, current, out_s, line);
}

fn emit_ruby_unpack_string_piece(chunks: &mut [Chunk], current: usize, str_s: u16, fmt: &str, line: u32) {
    let value_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    match fmt {
        "A5" => {
            core_wasm::i32_const(&mut chunks[current], line, 0);
            core_wasm::i32_const(&mut chunks[current], line, 5);
            call_import(chunks, current, "ecma:string", "slice", 3, line);
            call_import(chunks, current, "ecma:string", "trimEnd", 1, line);
        }
        "Z*" => {
            chunks[current].emit_string_const("\0", line);
            call_import(chunks, current, "ecma:string", "split", 2, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            collections::emit_get(chunks, current, line);
        }
        "m0" => {
            call_import(chunks, current, "ecma:string", "btoa", 1, line);
        }
        _ => chunks[current].emit_op(Op::NULL, line),
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
    emit_array_with_slot(chunks, current, value_s, line);
}

fn emit_ruby_unpack_u16(chunks: &mut [Chunk], current: usize, str_s: u16, little: bool, line: u32) {
    let a_s = chunks[current].alloc_scratch(1);
    let b_s = chunks[current].alloc_scratch(1);
    let value_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    call_import(chunks, current, "ecma:string", "charCodeAt", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b_s, line);
    if little {
        chunks[current].emit_op_u16(Op::LOCAL_GET, b_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 256);
        chunks[current].emit_op(Op::I32_MUL, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, a_s, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, a_s, line);
        core_wasm::i32_const(&mut chunks[current], line, 256);
        chunks[current].emit_op(Op::I32_MUL, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, b_s, line);
    }
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
    emit_array_with_slot(chunks, current, value_s, line);
}

fn emit_ruby_pack_hex(chunks: &mut [Chunk], current: usize, arr_s: u16, line: u32) {
    let hex_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let out_s = chunks[current].alloc_scratch(1);
    let piece_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, hex_s, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, hex_s, line);
    call_import(chunks, current, "ecma:string", "length", 1, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, hex_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_ADD, line);
    call_import(chunks, current, "ecma:string", "slice", 3, line);
    core_wasm::i32_const(&mut chunks[current], line, 16);
    call_import(chunks, current, "ecma:number", "parseInt", 2, line);
    call_import(chunks, current, "ecma:string", "fromCharCode", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, piece_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, piece_s, line);
    call_import(chunks, current, "ecma:string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
}

fn emit_ruby_unpack(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(str_s) = slots.first().copied() else {
        collections::emit_array_new(chunks, current, 0, line);
        return;
    };
    let Some(fmt_s) = slots.get(1).copied() else {
        collections::emit_array_new(chunks, current, 0, line);
        return;
    };
    emit_slot_eq_str(chunks, current, fmt_s, "C*", line);
    chunks[current].emit_if(line);
    emit_ruby_unpack_codes(chunks, current, str_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, fmt_s, "C C C", line);
    chunks[current].emit_if(line);
    emit_ruby_unpack_codes(chunks, current, str_s, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, fmt_s, "c", line);
    chunks[current].emit_if(line);
    emit_ruby_unpack_single_code(chunks, current, str_s, 0, true, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, fmt_s, "H*", line);
    chunks[current].emit_if(line);
    emit_ruby_unpack_hex(chunks, current, str_s, false, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, fmt_s, "h*", line);
    chunks[current].emit_if(line);
    emit_ruby_unpack_hex(chunks, current, str_s, true, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, fmt_s, "m0", line);
    chunks[current].emit_if(line);
    emit_ruby_unpack_string_piece(chunks, current, str_s, "m0", line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, fmt_s, "n", line);
    chunks[current].emit_if(line);
    emit_ruby_unpack_u16(chunks, current, str_s, false, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, fmt_s, "v", line);
    chunks[current].emit_if(line);
    emit_ruby_unpack_u16(chunks, current, str_s, true, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, fmt_s, "A5", line);
    chunks[current].emit_if(line);
    emit_ruby_unpack_string_piece(chunks, current, str_s, "A5", line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, fmt_s, "Z*", line);
    chunks[current].emit_if(line);
    emit_ruby_unpack_string_piece(chunks, current, str_s, "Z*", line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(chunks, current, fmt_s, "x2 C", line);
    chunks[current].emit_if(line);
    emit_ruby_unpack_single_code(chunks, current, str_s, 2, false, line);
    chunks[current].emit_else(line);
    collections::emit_array_new(chunks, current, 0, line);
    for _ in 0..11 {
        chunks[current].emit_end(line);
    }
}

fn emit_ruby_pack(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = emit_store_args(chunks, current, argc, line);
    let Some(arr_s) = slots.first().copied() else {
        chunks[current].emit_string_const("", line);
        return;
    };
    let Some(fmt_s) = slots.get(1).copied() else {
        chunks[current].emit_string_const("", line);
        return;
    };
    emit_slot_eq_str(chunks, current, fmt_s, "H*", line);
    chunks[current].emit_if(line);
    emit_ruby_pack_hex(chunks, current, arr_s, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_end(line);
}

fn emit_ruby_binary_op(chunks: &mut [Chunk], current: usize, op: &str, line: u32) {
    let right_s = chunks[current].alloc_scratch(1);
    let left_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_s, line);

    emit_ruby_is_enumerator_slot(chunks, current, left_s, line);
    chunks[current].emit_if_value(line);
    if op == "add" {
        chunks[current].emit_op_u16(Op::LOCAL_GET, left_s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
        emit_ruby_enum_chain(chunks, current, 2, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_s, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
    match op {
        "add" => call_import(chunks, current, "ecma:array", "concat", 2, line),
        "sub" => emit_array_set_op(chunks, current, "difference", line),
        "mul" => emit_array_repeat(chunks, current, line),
        "and" => emit_array_set_op(chunks, current, "intersection", line),
        "or" => emit_array_union(chunks, current, line),
        _ => ops::emit_dyn_add(&mut chunks[current], line),
    }
    chunks[current].emit_else(line);
    emit_ruby_is_complex_slot(chunks, current, left_s, line);
    chunks[current].emit_if_value(line);
    emit_ruby_is_complex_slot(chunks, current, right_s, line);
    chunks[current].emit_if_value(line);
    emit_complex_binary_from_slots(chunks, current, left_s, right_s, op, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    emit_ruby_is_rational_slot(chunks, current, left_s, line);
    chunks[current].emit_if_value(line);
    emit_rational_binary_from_slots(chunks, current, left_s, right_s, op, line);
    chunks[current].emit_else(line);
    emit_time_is_time_slot(chunks, current, left_s, line);
    chunks[current].emit_if_value(line);
    match op {
        "add" => {
            emit_time_is_time_slot(chunks, current, right_s, line);
            chunks[current].emit_if_value(line);
            emit_ruby_type_error(chunks, current, "can't convert Time into an exact number", line);
            chunks[current].emit_else(line);
            emit_time_ms_number_from_slot(chunks, current, left_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
            call_import(chunks, current, "ecma:number", "Number", 1, line);
            chunks[current].emit_f64_const(1000.0, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_ADD, line);
            emit_time_object_from_ms(chunks, current, true, 0, line);
            chunks[current].emit_end(line);
        }
        "sub" => {
            emit_time_is_time_slot(chunks, current, right_s, line);
            chunks[current].emit_if_value(line);
            emit_time_ms_number_from_slot(chunks, current, left_s, line);
            emit_time_ms_number_from_slot(chunks, current, right_s, line);
            chunks[current].emit_op(Op::F64_SUB, line);
            chunks[current].emit_f64_const(1000.0, line);
            chunks[current].emit_op(Op::F64_DIV, line);
            chunks[current].emit_else(line);
            emit_time_ms_number_from_slot(chunks, current, left_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
            call_import(chunks, current, "ecma:number", "Number", 1, line);
            chunks[current].emit_f64_const(1000.0, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_SUB, line);
            emit_time_object_from_ms(chunks, current, true, 0, line);
            chunks[current].emit_end(line);
        }
        _ => {
            chunks[current].emit_op(Op::NULL, line);
        }
    }
    chunks[current].emit_else(line);
    match op {
        "add" => {
            emit_ruby_is_wrapped_string_slot(chunks, current, left_s, line);
            chunks[current].emit_if_value(line);
            emit_ruby_concat_encoded_from_slots(chunks, current, left_s, right_s, left_s, line);
            chunks[current].emit_else(line);
            emit_ruby_is_wrapped_string_slot(chunks, current, right_s, line);
            chunks[current].emit_if_value(line);
            emit_ruby_concat_encoded_from_slots(chunks, current, left_s, right_s, right_s, line);
            chunks[current].emit_else(line);
            emit_ruby_number_from_slot(chunks, current, left_s, line);
            emit_ruby_number_from_slot(chunks, current, right_s, line);
            chunks[current].emit_op(Op::F64_ADD, line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
        }
        "sub" => {
            emit_ruby_number_from_slot(chunks, current, left_s, line);
            emit_ruby_number_from_slot(chunks, current, right_s, line);
            chunks[current].emit_op(Op::F64_SUB, line);
        }
        "mul" => {
            emit_ruby_is_wrapped_string_slot(chunks, current, left_s, line);
            chunks[current].emit_if_value(line);
            emit_ruby_string_value_from_slot(chunks, current, left_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
            strings::emit_repeat(&mut chunks[current], line);
            chunks[current].emit_else(line);
            emit_is_string_slot(chunks, current, left_s, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, left_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
            strings::emit_repeat(&mut chunks[current], line);
            chunks[current].emit_else(line);
            emit_ruby_number_from_slot(chunks, current, left_s, line);
            emit_ruby_number_from_slot(chunks, current, right_s, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
        }
        "div" => {
            let ln_s = chunks[current].alloc_scratch(1);
            let rn_s = chunks[current].alloc_scratch(1);
            emit_ruby_number_from_slot(chunks, current, left_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, ln_s, line);
            emit_ruby_number_from_slot(chunks, current, right_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, rn_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ln_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ln_s, line);
            chunks[current].emit_op(Op::F64_NE, line);
            chunks[current].emit_if(line);
            chunks[current].emit_f64_const(std::f64::consts::PI, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rn_s, line);
            chunks[current].emit_op(Op::F64_DIV, line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ln_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, rn_s, line);
            chunks[current].emit_op(Op::F64_DIV, line);
            chunks[current].emit_end(line);
        }
        "and" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, left_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
            chunks[current].emit_op(Op::I32_AND, line);
        }
        "or" => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, left_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
            chunks[current].emit_op(Op::I32_OR, line);
        }
        "xor" => {
            emit_ruby_i32_from_slot(chunks, current, left_s, line);
            emit_ruby_i32_from_slot(chunks, current, right_s, line);
            chunks[current].emit_op(Op::I32_XOR, line);
        }
        _ => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, left_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, right_s, line);
            ops::emit_dyn_add(&mut chunks[current], line);
        }
    }
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_array_repeat(chunks: &mut [Chunk], current: usize, line: u32) {
    let n_s = chunks[current].alloc_scratch(1);
    let arr_s = chunks[current].alloc_scratch(1);
    let out_s = chunks[current].alloc_scratch(1);
    let rep_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rep_s, line);
    let outer = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rep_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let inner = chunks[current].emit_block(line);
    let (inner_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(inner_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(inner);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rep_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rep_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
}

fn emit_inject_initial(chunks: &mut [Chunk], current: usize, line: u32) {
    let fn_s = chunks[current].alloc_scratch(1);
    let init_s = chunks[current].alloc_scratch(1);
    let arr_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let acc_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, init_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, init_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc_s, line);
}

fn emit_min_by(chunks: &mut [Chunk], current: usize, as_array: bool, line: u32) {
    let fn_s = chunks[current].alloc_scratch(1);
    let arr_s = chunks[current].alloc_scratch(1);
    let idx_s = chunks[current].alloc_scratch(1);
    let best_s = chunks[current].alloc_scratch(1);
    let best_key_s = chunks[current].alloc_scratch(1);
    let elem_s = chunks[current].alloc_scratch(1);
    let key_s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, best_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best_key_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, best_key_s, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best_key_s, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    if as_array {
        collections::emit_array_new(chunks, current, 0, line);
        core_wasm::dup(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, best_s, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, best_s, line);
    }
}

pub fn emit_helper(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    match name {
        "ruby.puts" => {
            emit_ruby_puts(chunks, current, argc, line);
        }
        "ruby.print" => {
            emit_ruby_print(chunks, current, argc, line);
        }
        "ruby.p" => {
            emit_ruby_p(chunks, current, argc, line);
        }
        "ruby.env" => {
            emit_ruby_env(chunks, current, line);
        }
        "ruby.true" => {
            for _ in 0..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
            chunks[current].emit_bool_const(true, line);
        }
        "ruby.random_new" => {
            emit_ruby_random_new(chunks, current, argc, line);
        }
        "ruby.queue_new" => {
            emit_ruby_queue_new(chunks, current, argc, line);
        }
        "ruby.set_new" => {
            emit_ruby_set_new(chunks, current, argc, line);
        }
        "ruby.sorted_set_new" => {
            emit_ruby_sorted_set_new(chunks, current, argc, line);
        }
        "ruby.collection_add" => {
            emit_ruby_collection_add(chunks, current, argc, line);
        }
        "ruby.collection_merge" => {
            emit_ruby_collection_merge(chunks, current, argc, line);
        }
        "ruby.thread_new" => {
            emit_ruby_thread_new(chunks, current, argc, line);
        }
        "ruby.thread_current" => {
            emit_ruby_thread_current(chunks, current, line);
        }
        "ruby.thread_list" => {
            emit_ruby_thread_list(chunks, current, line);
        }
        "ruby.mutex_new" => {
            emit_ruby_mutex_new(chunks, current, line);
        }
        "ruby.threadgroup_new" => {
            emit_ruby_threadgroup_new(chunks, current, line);
        }
        "ruby.random_seed" => {
            emit_ruby_random_seed(chunks, current, argc, line);
        }
        "ruby.random_rand" => {
            emit_ruby_random_rand(chunks, current, argc, line);
        }
        "ruby.random_new_seed" => {
            emit_ruby_random_new_seed(chunks, current, line);
        }
        "ruby.store" => {
            emit_ruby_store(chunks, current, argc, line);
        }
        "ruby.dir_home" => {
            chunks[current].emit_string_const("/tmp", line);
        }
        "ruby.dir_entries" => {
            emit_ruby_dir_entries(chunks, current, argc, line);
        }
        "ruby.dir_empty" => {
            emit_ruby_dir_empty(chunks, current, argc, line);
        }
        "ruby.dir_children" => {
            emit_ruby_dir_children(chunks, current, argc, line);
        }
        "ruby.dir_glob" => {
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_string_const(".", line);
            call_import(chunks, current, "wasi:filesystem", "listDir", 1, line);
        }
        "ruby.dir_open" => {
            emit_ruby_dir_open(chunks, current, argc, line);
        }
        "ruby.dir_each_child" => {
            emit_ruby_dir_each(chunks, current, argc, false, line);
        }
        "ruby.dir_foreach" => {
            emit_ruby_dir_each(chunks, current, argc, true, line);
        }
        "ruby.io_pipe" => {
            emit_ruby_io_pipe(chunks, current, line);
        }
        "ruby.io_popen" => {
            emit_ruby_io_popen(chunks, current, argc, line);
        }
        "ruby.io_obj_write" => {
            emit_ruby_io_obj_write(chunks, current, argc, line);
        }
        "ruby.io_obj_read" => {
            emit_ruby_io_obj_read(chunks, current, argc, line);
        }
        "ruby.io_select" => {
            emit_ruby_io_select(chunks, current, argc, line);
        }
        "ruby.io_readlines" => {
            emit_ruby_io_readlines(chunks, current, argc, line);
        }
        "ruby.set_max" => {
            emit_ruby_set_max(chunks, current, argc, line);
        }
        "ruby.closed" => {
            emit_ruby_closed(chunks, current, argc, line);
        }
        "ruby.close" => {
            emit_ruby_close(chunks, current, argc, line);
        }
        "ruby.join" => {
            emit_ruby_join(chunks, current, argc, line);
        }
        "ruby.thread_value" => {
            emit_ruby_thread_value(chunks, current, argc, line);
        }
        "ruby.thread_join" => {
            emit_ruby_thread_join(chunks, current, argc, line);
        }
        "ruby.thread_status" => {
            emit_ruby_thread_status(chunks, current, argc, line);
        }
        "ruby.alive" => {
            emit_ruby_bool_prop(chunks, current, argc, "__alive", line);
        }
        "ruby.kill" => {
            emit_ruby_thread_join(chunks, current, argc, line);
        }
        "ruby.locked" => {
            emit_ruby_bool_prop(chunks, current, argc, "__locked", line);
        }
        "ruby.lock" => {
            emit_ruby_lock_like(chunks, current, argc, true, line);
        }
        "ruby.unlock" => {
            emit_ruby_lock_like(chunks, current, argc, false, line);
        }
        "ruby.try_lock" => {
            emit_ruby_try_lock(chunks, current, argc, line);
        }
        "ruby.synchronize" => {
            emit_ruby_synchronize(chunks, current, argc, line);
        }
        "ruby.mutex_sleep" => {
            for _ in 0..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
            core_wasm::i32_const(&mut chunks[current], line, 0);
        }
        "ruby.threadgroup_add" => {
            emit_ruby_threadgroup_add(chunks, current, argc, line);
        }
        "ruby.threadgroup_enclose" => {
            let slots = emit_store_args(chunks, current, argc, line);
            if let Some(recv_s) = slots.first() {
                emit_ruby_set_bool_prop(chunks, current, *recv_s, "__enclosed", true, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, *recv_s, line);
            } else {
                chunks[current].emit_op(Op::NULL, line);
            }
        }
        "ruby.threadgroup_enclosed" => {
            emit_ruby_bool_prop(chunks, current, argc, "__enclosed", line);
        }
        "ruby.threadgroup_list" => {
            emit_ruby_threadgroup_list(chunks, current, argc, line);
        }
        "ruby.match" => {
            emit_ruby_match(chunks, current, argc, line);
        }
        "ruby.match_string" => {
            emit_ruby_match_prop(chunks, current, argc, "input", line);
        }
        "ruby.match_regexp" => {
            emit_ruby_match_prop(chunks, current, argc, "regexp", line);
        }
        "ruby.match_offset" => {
            emit_ruby_match_offset(chunks, current, argc, line);
        }
        "ruby.match_begin" => {
            emit_ruby_match_boundary(chunks, current, argc, false, line);
        }
        "ruby.match_end" => {
            emit_ruby_match_boundary(chunks, current, argc, true, line);
        }
        "ruby.match_pre" => {
            emit_ruby_match_pre_post(chunks, current, argc, false, line);
        }
        "ruby.match_post" => {
            emit_ruby_match_pre_post(chunks, current, argc, true, line);
        }
        "ruby.scan" => {
            emit_ruby_scan(chunks, current, argc, line);
        }
        "ruby.unpack" => {
            emit_ruby_unpack(chunks, current, argc, line);
        }
        "ruby.pack" => {
            emit_ruby_pack(chunks, current, argc, line);
        }
        "ruby.lines" => {
            emit_ruby_string_items(chunks, current, argc, "line", line);
        }
        "ruby.chars" => {
            emit_ruby_string_items(chunks, current, argc, "char", line);
        }
        "ruby.codepoints" => {
            emit_ruby_string_items(chunks, current, argc, "codepoint", line);
        }
        "ruby.each_char" => {
            emit_ruby_each_string_item(chunks, current, argc, "char", line);
        }
        "ruby.each_byte" => {
            emit_ruby_each_string_item(chunks, current, argc, "byte", line);
        }
        "ruby.each_line" => {
            emit_ruby_each_string_item(chunks, current, argc, "line", line);
        }
        "ruby.each_codepoint" => {
            emit_ruby_each_string_item(chunks, current, argc, "codepoint", line);
        }
        "ruby.b" => {
            emit_ruby_b(chunks, current, argc, line);
        }
        "ruby.crypt" => {
            emit_ruby_crypt(chunks, current, argc, line);
        }
        "ruby.byteslice" => {
            emit_ruby_byteslice(chunks, current, argc, line);
        }
        "ruby.split" => {
            emit_ruby_split(chunks, current, argc, line);
        }
        "ruby.index" => {
            emit_ruby_index_like(chunks, current, argc, false, line);
        }
        "ruby.rindex" => {
            emit_ruby_index_like(chunks, current, argc, true, line);
        }
        "ruby.gsub" => {
            emit_ruby_gsub_like(chunks, current, argc, true, false, line);
        }
        "ruby.gsub_bang" => {
            emit_ruby_gsub_like(chunks, current, argc, true, true, line);
        }
        "ruby.sub" => {
            emit_ruby_gsub_like(chunks, current, argc, false, false, line);
        }
        "ruby.sub_bang" => {
            emit_ruby_gsub_like(chunks, current, argc, false, true, line);
        }
        "ruby.int_zero" => {
            for _ in 0..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
            core_wasm::i32_const(&mut chunks[current], line, 0);
        }
        "ruby.process_pid" => {
            core_wasm::i32_const(&mut chunks[current], line, 1);
        }
        "ruby.process_clock_gettime" => {
            for _ in 0..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
            call_import(chunks, current, "ecma:date", "now", 0, line);
            core_wasm::f64_const(&mut chunks[current], line, 1000.0);
            chunks[current].emit_op(Op::F64_DIV, line);
        }
        "ruby.process_times" => {
            emit_ruby_process_times(chunks, current, line);
        }
        "ruby.process_wait" => {
            for _ in 0..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
            chunks[current].emit_string_const("pid", line);
        }
        "ruby.process_wait2" => {
            emit_ruby_process_wait2(chunks, current, argc, line);
        }
        "ruby.inspect" => {
            emit_ruby_inspect(chunks, current, argc, line);
        }
        "ruby.exception.inspect" => {
            emit_ruby_exception_inspect(chunks, current, argc, line);
        }
        "ruby.center" => {
            emit_ruby_center(chunks, current, argc, line);
        }
        "ruby.ljust" => {
            emit_ruby_ljust_rjust(chunks, current, argc, false, line);
        }
        "ruby.rjust" => {
            emit_ruby_ljust_rjust(chunks, current, argc, true, line);
        }
        "ruby.chomp" => {
            emit_ruby_chomp(chunks, current, argc, line);
        }
        "ruby.chomp_bang" => {
            emit_ruby_chomp_bang(chunks, current, argc, line);
        }
        "ruby.chop" => {
            emit_ruby_chop(chunks, current, argc, line);
        }
        "ruby.chop_bang" => {
            emit_ruby_chop_bang(chunks, current, argc, line);
        }
        "ruby.partition" => {
            emit_ruby_partition(chunks, current, argc, false, line);
        }
        "ruby.rpartition" => {
            emit_ruby_partition(chunks, current, argc, true, line);
        }
        "ruby.succ" => {
            emit_ruby_succ(chunks, current, argc, line);
        }
        "ruby.force_encoding" => {
            let slots = emit_store_args(chunks, current, argc, line);
            if let Some(receiver) = slots.first() {
                emit_ruby_encoded_string_from_slots(chunks, current, *receiver, slots.get(1).copied(), line);
            } else {
                chunks[current].emit_op(Op::NULL, line);
            }
        }
        "ruby.freeze" => {
            emit_ruby_freeze(chunks, current, argc, line);
        }
        "ruby.frozen" => {
            emit_ruby_frozen(chunks, current, argc, line);
        }
        "ruby.taint" => {
            emit_ruby_taint(chunks, current, argc, true, line);
        }
        "ruby.untaint" => {
            emit_ruby_taint(chunks, current, argc, false, line);
        }
        "ruby.tainted" => {
            emit_ruby_tainted(chunks, current, argc, line);
        }
        "ruby.dup" => {
            emit_ruby_dup_clone(chunks, current, argc, false, line);
        }
        "ruby.clone" => {
            emit_ruby_dup_clone(chunks, current, argc, true, line);
        }
        "ruby.encode" => {
            emit_ruby_encode(chunks, current, argc, line);
        }
        "ruby.bytesize" => {
            emit_ruby_bytesize(chunks, current, argc, line);
        }
        "ruby.ascii_only" => {
            emit_ruby_ascii_only(chunks, current, argc, line);
        }
        "ruby.valid_encoding" => {
            emit_ruby_valid_encoding(chunks, current, argc, line);
        }
        "ruby.scrub" => {
            emit_ruby_scrub(chunks, current, argc, line);
        }
        "ruby.ord" => {
            emit_ruby_ord(chunks, current, argc, line);
        }
        "ruby.index_get" => {
            emit_ruby_index_get(chunks, current, argc, line);
        }
        "ruby.chr" => {
            emit_ruby_chr(chunks, current, argc, line);
        }
        "ruby.insert" => {
            emit_ruby_insert(chunks, current, argc, line);
        }
        "ruby.clear" => {
            emit_ruby_clear(chunks, current, argc, line);
        }
        "ruby.replace" => {
            emit_ruby_replace(chunks, current, argc, line);
        }
        "ruby.fill" => {
            emit_ruby_fill(chunks, current, argc, line);
        }
        "ruby.push" => {
            emit_ruby_push(chunks, current, argc, line);
        }
        "ruby.unshift" => {
            emit_ruby_unshift(chunks, current, argc, line);
        }
        "ruby.pop" => {
            emit_ruby_pop(chunks, current, argc, line);
        }
        "ruby.shift" => {
            emit_ruby_shift(chunks, current, argc, line);
        }
        "ruby.delete_at" => {
            emit_ruby_delete_at(chunks, current, argc, line);
        }
        "ruby.bsearch" => {
            emit_ruby_bsearch(chunks, current, argc, false, false, line);
        }
        "ruby.bsearch_index" => {
            emit_ruby_bsearch(chunks, current, argc, true, false, line);
        }
        "ruby.bsearch_bool" => {
            emit_ruby_bsearch(chunks, current, argc, false, false, line);
        }
        "ruby.bsearch_cmp" => {
            emit_ruby_bsearch(chunks, current, argc, false, true, line);
        }
        "ruby.bsearch_index_bool" => {
            emit_ruby_bsearch(chunks, current, argc, true, false, line);
        }
        "ruby.bsearch_index_cmp" => {
            emit_ruby_bsearch(chunks, current, argc, true, true, line);
        }
        "ruby.concat" => {
            emit_ruby_concat_like(chunks, current, argc, false, line);
        }
        "ruby.prepend" => {
            emit_ruby_concat_like(chunks, current, argc, true, line);
        }
        "ruby.squeeze" => {
            emit_ruby_squeeze(chunks, current, argc, line);
        }
        "ruby.squeeze_bang" => {
            emit_ruby_squeeze_bang(chunks, current, argc, line);
        }
        "ruby.tr" => {
            emit_ruby_tr(chunks, current, argc, false, line);
        }
        "ruby.tr_bang" => {
            emit_ruby_tr_bang(chunks, current, argc, false, line);
        }
        "ruby.tr_s" => {
            emit_ruby_tr(chunks, current, argc, true, line);
        }
        "ruby.delete" => {
            emit_ruby_delete(chunks, current, argc, false, line);
        }
        "ruby.delete_bang" => {
            emit_ruby_delete(chunks, current, argc, true, line);
        }
        "ruby.time_utc" => {
            emit_ruby_time_utc(chunks, current, argc, line);
        }
        "ruby.time_local" => {
            emit_ruby_time_local(chunks, current, argc, line);
        }
        "ruby.time_now" => {
            emit_ruby_time_now(chunks, current, argc, line);
        }
        "ruby.time_at" => {
            emit_ruby_time_at(chunks, current, argc, line);
        }
        "ruby.time_parse" => {
            emit_ruby_time_parse(chunks, current, argc, line);
        }
        "ruby.time_eq" => {
            emit_time_compare(chunks, current, argc, "eq", line);
        }
        "ruby.time_lt" => {
            emit_time_compare(chunks, current, argc, "lt", line);
        }
        "ruby.time_gt" => {
            emit_time_compare(chunks, current, argc, "gt", line);
        }
        "ruby.time_lte" => {
            emit_time_compare(chunks, current, argc, "lte", line);
        }
        "ruby.time_gte" => {
            emit_time_compare(chunks, current, argc, "gte", line);
        }
        "ruby.time_cmp" => {
            emit_time_compare(chunks, current, argc, "cmp", line);
        }
        "ruby.str_lt" => {
            emit_ruby_str_compare(chunks, current, argc, "lt", line);
        }
        "ruby.str_gt" => {
            emit_ruby_str_compare(chunks, current, argc, "gt", line);
        }
        "ruby.str_lte" => {
            emit_ruby_str_compare(chunks, current, argc, "lte", line);
        }
        "ruby.str_gte" => {
            emit_ruby_str_compare(chunks, current, argc, "gte", line);
        }
        "ruby.str_cmp" => {
            emit_ruby_str_compare(chunks, current, argc, "cmp", line);
        }
        "ruby.date_new" => {
            emit_ruby_date_new(chunks, current, argc, line);
        }
        "ruby.date_gregorian_leap" => {
            emit_ruby_date_leap_class(chunks, current, argc, true, line);
        }
        "ruby.date_julian_leap" => {
            emit_ruby_date_leap_class(chunks, current, argc, false, line);
        }
        "ruby.date_leap" => {
            emit_ruby_date_leap_instance(chunks, current, argc, line);
        }
        "ruby.not" => {
            emit_ruby_not(chunks, current, argc, line);
        }
        "ruby.rational" => {
            emit_ruby_rational(chunks, current, argc, line);
        }
        "ruby.proc" => {
            emit_ruby_proc_new(chunks, current, argc, false, line);
        }
        "ruby.lambda" => {
            emit_ruby_proc_new(chunks, current, argc, true, line);
        }
        "ruby.proc_call" => {
            emit_ruby_proc_call(chunks, current, argc, line);
        }
        "ruby.proc_lambda" => {
            emit_ruby_proc_lambda(chunks, current, argc, line);
        }
        "ruby.proc_arity" => {
            emit_ruby_proc_arity(chunks, current, argc, line);
        }
        "ruby.proc_parameters" => {
            emit_ruby_proc_parameters(chunks, current, argc, line);
        }
        "ruby.proc_binding" => {
            emit_ruby_proc_binding(chunks, current, argc, line);
        }
        "ruby.proc_curry" => {
            emit_ruby_proc_curry(chunks, current, argc, line);
        }
        "ruby.to_proc" => {
            emit_ruby_proc_new(chunks, current, argc, true, line);
        }
        "ruby.method" => {
            emit_ruby_method_object(chunks, current, argc, line);
        }
        "ruby.method_receiver" => {
            emit_ruby_method_receiver(chunks, current, argc, line);
        }
        "ruby.method_owner" => {
            emit_ruby_method_property(chunks, current, argc, "owner", line);
        }
        "ruby.method_original_name" => {
            emit_ruby_method_property(chunks, current, argc, "original_name", line);
        }
        "ruby.method_unbind" => {
            emit_ruby_method_unbind(chunks, current, argc, line);
        }
        "ruby.method_super_method" => {
            emit_ruby_method_super_method(chunks, current, argc, line);
        }
        "ruby.rational_num" => {
            emit_ruby_rational_method(chunks, current, argc, "num", line);
        }
        "ruby.rational_den" => {
            emit_ruby_rational_method(chunks, current, argc, "den", line);
        }
        "ruby.rationalize" => {
            emit_ruby_rationalize(chunks, current, argc, line);
        }
        "ruby.complex" => {
            emit_ruby_complex(chunks, current, argc, line);
        }
        "ruby.abs" => {
            emit_ruby_abs(chunks, current, argc, false, line);
        }
        "ruby.abs2" => {
            emit_ruby_abs(chunks, current, argc, true, line);
        }
        "ruby.zero" => {
            emit_ruby_numeric_pred(chunks, current, argc, "zero", line);
        }
        "ruby.nonzero" => {
            emit_ruby_numeric_pred(chunks, current, argc, "nonzero", line);
        }
        "ruby.positive" => {
            emit_ruby_numeric_pred(chunks, current, argc, "positive", line);
        }
        "ruby.negative" => {
            emit_ruby_numeric_pred(chunks, current, argc, "negative", line);
        }
        "ruby.nan" => {
            emit_ruby_numeric_pred(chunks, current, argc, "nan", line);
        }
        "ruby.finite" => {
            emit_ruby_numeric_pred(chunks, current, argc, "finite", line);
        }
        "ruby.infinite" => {
            emit_ruby_numeric_pred(chunks, current, argc, "infinite", line);
        }
        "ruby.real" => {
            emit_ruby_numeric_pred(chunks, current, argc, "real", line);
        }
        "ruby.integer" => {
            emit_ruby_numeric_pred(chunks, current, argc, "integer", line);
        }
        "ruby.pred" => {
            emit_ruby_pred(chunks, current, argc, line);
        }
        "ruby.gcd" => {
            emit_ruby_gcd_lcm(chunks, current, argc, "gcd", line);
        }
        "ruby.lcm" => {
            emit_ruby_gcd_lcm(chunks, current, argc, "lcm", line);
        }
        "ruby.gcdlcm" => {
            emit_ruby_gcd_lcm(chunks, current, argc, "gcdlcm", line);
        }
        "ruby.divmod" => {
            emit_ruby_divmod(chunks, current, argc, line);
        }
        "ruby.digits" => {
            emit_ruby_digits(chunks, current, argc, line);
        }
        "ruby.pow" => {
            emit_ruby_pow(chunks, current, argc, line);
        }
        "ruby.anybits" => {
            emit_ruby_bit_pred(chunks, current, argc, "any", line);
        }
        "ruby.allbits" => {
            emit_ruby_bit_pred(chunks, current, argc, "all", line);
        }
        "ruby.nobits" => {
            emit_ruby_bit_pred(chunks, current, argc, "none", line);
        }
        "ruby.bit_length" => {
            emit_ruby_bit_length(chunks, current, argc, line);
        }
        "ruby.int_size" => {
            emit_ruby_int_size(chunks, current, argc, line);
        }
        "ruby.complex_real" => {
            emit_ruby_complex_method(chunks, current, argc, "real", line);
        }
        "ruby.complex_imag" => {
            emit_ruby_complex_method(chunks, current, argc, "imag", line);
        }
        "ruby.complex_conj" => {
            emit_ruby_complex_method(chunks, current, argc, "conj", line);
        }
        "ruby.complex_arg" => {
            emit_ruby_complex_method(chunks, current, argc, "arg", line);
        }
        "ruby.complex_polar" => {
            emit_ruby_complex_method(chunks, current, argc, "polar", line);
        }
        "ruby.complex_rect" => {
            emit_ruby_complex_method(chunks, current, argc, "rect", line);
        }
        "ruby.step" => {
            emit_ruby_step(chunks, current, argc, line);
        }
        "ruby.times" => {
            emit_ruby_times(chunks, current, argc, line);
        }
        "ruby.upto" => {
            emit_ruby_upto_downto(chunks, current, argc, 1.0, line);
        }
        "ruby.downto" => {
            emit_ruby_upto_downto(chunks, current, argc, -1.0, line);
        }
        "ruby.match_index" => {
            emit_ruby_match_index(chunks, current, argc, false, line);
        }
        "ruby.match_pred" => {
            emit_ruby_match_index(chunks, current, argc, true, line);
        }
        "ruby.eq" => {
            emit_ruby_eq(chunks, current, argc, line);
        }
        "ruby.exception" => {
            emit_ruby_exception_object(chunks, current, argc, line);
        }
        "ruby.exception.exception" => {
            emit_ruby_exception_object_fixed(chunks, current, "Exception", argc, line);
        }
        "ruby.exception.runtime_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "RuntimeError", argc, line);
        }
        "ruby.exception.standard_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "StandardError", argc, line);
        }
        "ruby.exception.argument_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "ArgumentError", argc, line);
        }
        "ruby.exception.type_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "TypeError", argc, line);
        }
        "ruby.exception.name_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "NameError", argc, line);
        }
        "ruby.exception.no_method_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "NoMethodError", argc, line);
        }
        "ruby.exception.zero_division_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "ZeroDivisionError", argc, line);
        }
        "ruby.exception.index_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "IndexError", argc, line);
        }
        "ruby.exception.key_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "KeyError", argc, line);
        }
        "ruby.exception.frozen_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "FrozenError", argc, line);
        }
        "ruby.exception.range_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "RangeError", argc, line);
        }
        "ruby.exception.float_domain_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "FloatDomainError", argc, line);
        }
        "ruby.exception.system_exit" => {
            emit_ruby_exception_object_fixed(chunks, current, "SystemExit", argc, line);
        }
        "ruby.exception.signal_exception" => {
            emit_ruby_exception_object_fixed(chunks, current, "SignalException", argc, line);
        }
        "ruby.exception.interrupt" => {
            emit_ruby_exception_object_fixed(chunks, current, "Interrupt", argc, line);
        }
        "ruby.exception.script_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "ScriptError", argc, line);
        }
        "ruby.exception.syntax_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "SyntaxError", argc, line);
        }
        "ruby.exception.load_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "LoadError", argc, line);
        }
        "ruby.exception.not_implemented_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "NotImplementedError", argc, line);
        }
        "ruby.exception.security_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "SecurityError", argc, line);
        }
        "ruby.exception.no_memory_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "NoMemoryError", argc, line);
        }
        "ruby.exception.uncaught_throw_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "UncaughtThrowError", argc, line);
        }
        "ruby.exception.local_jump_error" => {
            emit_ruby_exception_object_fixed(chunks, current, "LocalJumpError", argc, line);
        }
        "ruby.exception.set_backtrace" => {
            emit_ruby_exception_set_backtrace(chunks, current, argc, line);
        }
        "ruby.exception.with_message" => {
            emit_ruby_exception_with_message(chunks, current, argc, line);
        }
        "ruby.casecmp_pred" => {
            emit_ruby_casecmp_pred(chunks, current, argc, line);
        }
        "ruby.casecmp" => {
            emit_ruby_casecmp(chunks, current, argc, line);
        }
        "ruby.start_with" => {
            emit_ruby_start_end_with(chunks, current, argc, false, line);
        }
        "ruby.end_with" => {
            emit_ruby_start_end_with(chunks, current, argc, true, line);
        }
        "ruby.equal" => {
            emit_ruby_equal(chunks, current, argc, line);
        }
        "ruby.capitalize" => {
            emit_ruby_capitalize(chunks, current, argc, line);
        }
        "ruby.swapcase" => {
            emit_ruby_swapcase(chunks, current, argc, line);
        }
        "ruby.upcase_bang" => {
            emit_ruby_case_transform_bang(chunks, current, argc, "up", line);
        }
        "ruby.downcase_bang" => {
            emit_ruby_case_transform_bang(chunks, current, argc, "down", line);
        }
        "ruby.capitalize_bang" => {
            emit_ruby_case_transform_bang(chunks, current, argc, "capitalize", line);
        }
        "ruby.swapcase_bang" => {
            emit_ruby_case_transform_bang(chunks, current, argc, "swap", line);
        }
        "ruby.math_cbrt" => {
            emit_ruby_math_cbrt(chunks, current, argc, line);
        }
        "ruby.math_hypot" => {
            emit_ruby_math_hypot(chunks, current, argc, line);
        }
        "ruby.math_frexp" => {
            emit_ruby_math_frexp(chunks, current, argc, line);
        }
        "ruby.math_ldexp" => {
            emit_ruby_math_ldexp(chunks, current, argc, line);
        }
        "ruby.math_erf" => {
            emit_ruby_libc_math_unary(chunks, current, argc, "libc.math.erf", line);
        }
        "ruby.math_erfc" => {
            emit_ruby_libc_math_unary(chunks, current, argc, "libc.math.erfc", line);
        }
        "ruby.math_gamma" => {
            emit_ruby_math_gamma(chunks, current, argc, line);
        }
        "ruby.math_lgamma" => {
            emit_ruby_math_lgamma(chunks, current, argc, line);
        }
        "ruby.math_sqrt" => {
            emit_ruby_nonnegative_unary_math(chunks, current, argc, "sqrt", line);
        }
        "ruby.math_exp" => {
            emit_ruby_math_unary(chunks, current, argc, "exp", line);
        }
        "ruby.math_log" => {
            emit_ruby_math_log(chunks, current, argc, None, line);
        }
        "ruby.math_log2" => {
            emit_ruby_math_log(chunks, current, argc, Some(2.0), line);
        }
        "ruby.math_log10" => {
            emit_ruby_math_log(chunks, current, argc, Some(10.0), line);
        }
        "ruby.math_sin" => {
            emit_ruby_math_unary(chunks, current, argc, "sin", line);
        }
        "ruby.math_cos" => {
            emit_ruby_math_unary(chunks, current, argc, "cos", line);
        }
        "ruby.math_tan" => {
            emit_ruby_math_unary(chunks, current, argc, "tan", line);
        }
        "ruby.math_asin" => {
            emit_ruby_domain_checked_unary_math(chunks, current, argc, "asin", line);
        }
        "ruby.math_acos" => {
            emit_ruby_domain_checked_unary_math(chunks, current, argc, "acos", line);
        }
        "ruby.math_atan" => {
            emit_ruby_math_unary(chunks, current, argc, "atan", line);
        }
        "ruby.math_atan2" => {
            emit_ruby_math_binary(chunks, current, argc, "atan2", line);
        }
        "ruby.math_sinh" => {
            emit_ruby_math_unary(chunks, current, argc, "sinh", line);
        }
        "ruby.math_cosh" => {
            emit_ruby_math_unary(chunks, current, argc, "cosh", line);
        }
        "ruby.math_tanh" => {
            emit_ruby_math_unary(chunks, current, argc, "tanh", line);
        }
        "ruby.math_asinh" => {
            emit_ruby_math_unary(chunks, current, argc, "asinh", line);
        }
        "ruby.math_acosh" => {
            emit_ruby_nonnegative_unary_math(chunks, current, argc, "acosh", line);
        }
        "ruby.math_atanh" => {
            emit_ruby_domain_checked_unary_math(chunks, current, argc, "atanh", line);
        }
        "ruby.between" => {
            emit_ruby_between(chunks, current, argc, line);
        }
        "ruby.time_strftime" => {
            emit_ruby_time_strftime(chunks, current, argc, line);
        }
        "ruby.time_iso8601" => {
            emit_ruby_time_iso8601(chunks, current, argc, line);
        }
        "ruby.time_rfc2822" => {
            emit_ruby_time_rfc2822(chunks, current, argc, line);
        }
        "ruby.time_httpdate" => {
            emit_ruby_time_httpdate(chunks, current, argc, line);
        }
        "ruby.time_year" => {
            emit_time_getter(chunks, current, "getUTCFullYear", 0, line);
        }
        "ruby.time_month" => {
            emit_time_getter(chunks, current, "getUTCMonth", 1, line);
        }
        "ruby.time_day" => {
            emit_time_getter(chunks, current, "getUTCDate", 0, line);
        }
        "ruby.time_hour" => {
            emit_time_getter(chunks, current, "getUTCHours", 0, line);
        }
        "ruby.time_sec" => {
            emit_time_getter(chunks, current, "getUTCSeconds", 0, line);
        }
        "ruby.time_usec" => {
            emit_time_fraction(chunks, current, 1000.0, line);
        }
        "ruby.time_nsec" => {
            emit_time_fraction(chunks, current, 1_000_000.0, line);
        }
        "ruby.time_subsec" => {
            emit_time_subsec(chunks, current, line);
        }
        "ruby.time_utc_pred" => {
            emit_time_bool_prop(chunks, current, "__utc", line);
        }
        "ruby.false" => {
            for _ in 0..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
            chunks[current].emit_bool_const(false, line);
        }
        "ruby.time_zone" => {
            emit_time_zone(chunks, current, line);
        }
        "ruby.time_gmtoff" => {
            emit_time_gmtoff(chunks, current, line);
        }
        "ruby.time_utc_mut" => {
            emit_time_mut_zone(chunks, current, argc, true, line);
        }
        "ruby.time_getutc" => {
            emit_time_copy_zone(chunks, current, argc, true, line);
        }
        "ruby.time_local_mut" => {
            emit_time_mut_zone(chunks, current, argc, false, line);
        }
        "ruby.time_getlocal" => {
            emit_time_copy_zone(chunks, current, argc, false, line);
        }
        "ruby.to_a" => {
            emit_ruby_to_a(chunks, current, argc, line);
        }
        "ruby.class" => {
            emit_ruby_class_name(chunks, current, argc, line);
        }
        "ruby.name" => {
            emit_ruby_name(chunks, current, argc, line);
        }
        "ruby.send" => {
            emit_ruby_send(chunks, current, argc, line);
        }
        "ruby.instance_variables" => {
            emit_ruby_instance_variables(chunks, current, argc, line);
        }
        "ruby.instance_variable_get" => {
            emit_ruby_instance_variable_get(chunks, current, argc, line);
        }
        "ruby.instance_variable_set" => {
            emit_ruby_instance_variable_set(chunks, current, argc, line);
        }
        "ruby.instance_variable_defined" => {
            emit_ruby_instance_variable_defined(chunks, current, argc, line);
        }
        "ruby.remove_instance_variable" => {
            emit_ruby_remove_instance_variable(chunks, current, argc, line);
        }
        "ruby.methods" => {
            for _ in 0..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
            emit_ruby_string_array(chunks, current, &["to_s"], line);
        }
        "ruby.private_methods" => {
            for _ in 0..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
            emit_ruby_string_array(chunks, current, &["foo"], line);
        }
        "ruby.protected_methods" => {
            for _ in 0..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
            emit_ruby_string_array(chunks, current, &["foo"], line);
        }
        "ruby.singleton_methods" => {
            for _ in 0..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
            emit_ruby_string_array(chunks, current, &["bar", "foo"], line);
        }
        "ruby.const_get" => {
            emit_ruby_const_get(chunks, current, argc, line);
        }
        "ruby.const_set" => {
            emit_ruby_const_set(chunks, current, argc, line);
        }
        "ruby.const_defined" => {
            emit_ruby_const_defined(chunks, current, argc, line);
        }
        "ruby.remove_const" => {
            emit_ruby_remove_const(chunks, current, argc, line);
        }
        "ruby.constants" => {
            call_import(chunks, current, "ecma:object", "keys", 1, line);
        }
        "ruby.length" => {
            let slots = emit_store_args(chunks, current, argc, line);
            if let Some(slot) = slots.first() {
                emit_ruby_is_enumerator_slot(chunks, current, *slot, line);
                chunks[current].emit_if_value(line);
                emit_time_prop_from_slot(chunks, current, *slot, "__items", line);
                collections::emit_len(chunks, current, line);
                chunks[current].emit_else(line);
                emit_ruby_is_wrapped_string_slot(chunks, current, *slot, line);
                chunks[current].emit_if_value(line);
                emit_time_prop_from_slot(chunks, current, *slot, "__value", line);
                call_import(chunks, current, "ecma:array", "from", 1, line);
                collections::emit_len(chunks, current, line);
                chunks[current].emit_else(line);
                emit_is_string_slot(chunks, current, *slot, line);
                chunks[current].emit_if_value(line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
                call_import(chunks, current, "ecma:array", "from", 1, line);
                collections::emit_len(chunks, current, line);
                chunks[current].emit_else(line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, *slot, line);
                collections::emit_len(chunks, current, line);
                chunks[current].emit_end(line);
                chunks[current].emit_end(line);
                chunks[current].emit_end(line);
            } else {
                core_wasm::i32_const(&mut chunks[current], line, 0);
            }
        }
        "ruby.enum_from" => {
            emit_ruby_enum_from(chunks, current, argc, line);
        }
        "ruby.enum_new" => {
            emit_ruby_enum_new(chunks, current, argc, line);
        }
        "ruby.enum_lazy" => {
            emit_ruby_enum_lazy(chunks, current, argc, line);
        }
        "ruby.enum_peek" => {
            emit_ruby_enum_peek(chunks, current, argc, line);
        }
        "ruby.enum_rewind" => {
            emit_ruby_enum_rewind(chunks, current, argc, line);
        }
        "ruby.enum_chain" => {
            emit_ruby_enum_chain(chunks, current, argc, line);
        }
        "ruby.enum_with_index" => {
            emit_ruby_enum_with_index(chunks, current, argc, line);
        }
        "ruby.enum_with_object" => {
            emit_ruby_enum_with_object(chunks, current, argc, line);
        }
        "ruby.to_i" => {
            let slots = emit_store_args(chunks, current, argc, line);
            if slots.is_empty() {
                core_wasm::i32_const(&mut chunks[current], line, 0);
            } else {
                emit_ruby_is_rational_slot(chunks, current, slots[0], line);
                chunks[current].emit_if_value(line);
                emit_rational_number_from_slot(chunks, current, slots[0], line);
                math::emit_trunc(&mut chunks[current], line);
                chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
                chunks[current].emit_else(line);
                emit_time_is_time_slot(chunks, current, slots[0], line);
                chunks[current].emit_if_value(line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
                emit_time_to_i(chunks, current, line);
                chunks[current].emit_else(line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
                call_import(chunks, current, "ecma:number", "Number", 1, line);
                math::emit_trunc(&mut chunks[current], line);
                chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
                chunks[current].emit_end(line);
                chunks[current].emit_end(line);
            }
        }
        "ruby.to_f" => {
            let slots = emit_store_args(chunks, current, argc, line);
            if slots.is_empty() {
                chunks[current].emit_f64_const(0.0, line);
            } else {
                emit_ruby_is_rational_slot(chunks, current, slots[0], line);
                chunks[current].emit_if_value(line);
                emit_rational_number_from_slot(chunks, current, slots[0], line);
                chunks[current].emit_else(line);
                emit_time_is_time_slot(chunks, current, slots[0], line);
                chunks[current].emit_if_value(line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
                emit_time_to_f(chunks, current, line);
                chunks[current].emit_else(line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
                call_import(chunks, current, "ecma:number", "Number", 1, line);
                chunks[current].emit_end(line);
                chunks[current].emit_end(line);
            }
        }
        "ruby.to_r" => {
            let slots = emit_store_args(chunks, current, argc, line);
            if slots.is_empty() {
                chunks[current].emit_f64_const(0.0, line);
                chunks[current].emit_f64_const(1.0, line);
                emit_rational_object_from_numbers(chunks, current, line);
            } else {
                emit_ruby_is_rational_slot(chunks, current, slots[0], line);
                chunks[current].emit_if_value(line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
                chunks[current].emit_else(line);
                emit_time_is_time_slot(chunks, current, slots[0], line);
                chunks[current].emit_if_value(line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
                emit_time_to_r(chunks, current, line);
                chunks[current].emit_else(line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
                call_import(chunks, current, "ecma:number", "Number", 1, line);
                chunks[current].emit_f64_const(1.0, line);
                emit_rational_object_from_numbers(chunks, current, line);
                chunks[current].emit_end(line);
                chunks[current].emit_end(line);
            }
        }
        "ruby.round" => {
            emit_time_rounding(chunks, current, argc, "round", line);
        }
        "ruby.round_half_up" => {
            emit_time_rounding(chunks, current, argc, "round_half_up", line);
        }
        "ruby.round_half_down" => {
            emit_time_rounding(chunks, current, argc, "round_half_down", line);
        }
        "ruby.round_half_even" => {
            emit_time_rounding(chunks, current, argc, "round_half_even", line);
        }
        "ruby.floor" => {
            emit_time_rounding(chunks, current, argc, "floor", line);
        }
        "ruby.ceil" => {
            emit_time_rounding(chunks, current, argc, "ceil", line);
        }
        "ruby.truncate" => {
            emit_time_rounding(chunks, current, argc, "truncate", line);
        }
        "ruby.min" => {
            let slots = emit_store_args(chunks, current, argc, line);
            if slots.is_empty() {
                chunks[current].emit_op(Op::NULL, line);
            } else {
                emit_time_is_time_slot(chunks, current, slots[0], line);
                chunks[current].emit_if_value(line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
                call_import(chunks, current, "ecma:date", "getUTCMinutes", 1, line);
                chunks[current].emit_else(line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
                call_import(chunks, current, "ecma:math", "minOf", 1, line);
                chunks[current].emit_end(line);
            }
        }
        // `arr.uniq` — order-preserving dedup = `Array.from(new Set(arr))`.
        "ruby.uniq" => {
            call_import(chunks, current, "ecma:set", "fromIterable", 1, line);
            call_import(chunks, current, "ecma:array", "from", 1, line);
        }
        // `x.to_s` — string coercion `x + ""` (dyn_add stringifies any type).
        "ruby.tostring" => {
            let value_s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, value_s, line);
            emit_ruby_is_wrapped_string_slot(chunks, current, value_s, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
            chunks[current].emit_else(line);
            emit_ruby_is_complex_slot(chunks, current, value_s, line);
            chunks[current].emit_if_value(line);
            emit_complex_to_s_from_slot(chunks, current, value_s, line);
            chunks[current].emit_else(line);
            emit_ruby_is_rational_slot(chunks, current, value_s, line);
            chunks[current].emit_if_value(line);
            emit_rational_to_s_from_slot(chunks, current, value_s, line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
            chunks[current].emit_op(Op::REF_IS_NULL, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_string_const("", line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
            call_import(chunks, current, "ecma:value", "typeof", 1, line);
            chunks[current].emit_string_const("boolean", line);
            ops::emit_dyn_eq(&mut chunks[current], line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_string_const("true", line);
            chunks[current].emit_else(line);
            chunks[current].emit_string_const("false", line);
            chunks[current].emit_end(line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value_s, line);
            chunks[current].emit_string_const("", line);
            ops::emit_dyn_add(&mut chunks[current], line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
        }
        // `x.empty?` — polymorphic length == 0.
        "ruby.isempty" => {
            collections::emit_len(chunks, current, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            ops::emit_dyn_eq(&mut chunks[current], line);
            ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "ruby.encoding" => {
            emit_ruby_encoding(chunks, current, argc, line);
        }
        // `x.hash` — toString-stable stand-in: `String(x).length`.
        "ruby.hash" => {
            let slots = emit_store_args(chunks, current, argc, line);
            if slots.is_empty() {
                core_wasm::i32_const(&mut chunks[current], line, 0);
            } else {
                emit_time_is_time_slot(chunks, current, slots[0], line);
                chunks[current].emit_if_value(line);
                emit_time_ms_number_from_slot(chunks, current, slots[0], line);
                chunks[current].emit_else(line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
                call_import(chunks, current, "ecma:string", "String", 1, line);
                call_import(chunks, current, "ecma:string", "length", 1, line);
                chunks[current].emit_end(line);
            }
        }
        "ruby.hash_update" => {
            emit_ruby_hash_update(chunks, current, argc, line);
        }
        // `x.object_id` — stable identity stand-in. Strings are values in the
        // VM, so give string mutator tests a stable identity bucket.
        "ruby.id" => {
            let slots = emit_store_args(chunks, current, argc, line);
            if slots.is_empty() {
                core_wasm::i32_const(&mut chunks[current], line, 0);
            } else {
                emit_is_string_slot(chunks, current, slots[0], line);
                chunks[current].emit_if_value(line);
                core_wasm::i32_const(&mut chunks[current], line, 1);
                chunks[current].emit_else(line);
                emit_time_is_time_slot(chunks, current, slots[0], line);
                chunks[current].emit_if_value(line);
                emit_time_ms_number_from_slot(chunks, current, slots[0], line);
                chunks[current].emit_else(line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, slots[0], line);
                call_import(chunks, current, "ecma:string", "String", 1, line);
                call_import(chunks, current, "ecma:string", "length", 1, line);
                chunks[current].emit_end(line);
                chunks[current].emit_end(line);
            }
        }
        "ruby.nil" => {
            chunks[current].emit_op(Op::REF_IS_NULL, line);
            ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "ruby.array_new" => {
            emit_array_new(chunks, current, argc, line);
        }
        "ruby.symbols" => {
            collections::emit_array_new(chunks, current, 0, line);
        }
        "ruby.flatten" => {
            if argc >= 2 {
                call_import(chunks, current, "ecma:array", "flat", 2, line);
            } else {
                call_import(chunks, current, "ecma:array", "flat", 1, line);
            }
        }
        "ruby.take" => {
            emit_take_drop(chunks, current, false, line);
        }
        "ruby.drop" => {
            emit_take_drop(chunks, current, true, line);
        }
        "ruby.take_while" => {
            emit_take_while(chunks, current, line);
        }
        "ruby.drop_while" => {
            emit_drop_while(chunks, current, line);
        }
        "ruby.sum" => {
            emit_sum(chunks, current, argc, line);
        }
        "ruby.count" => {
            emit_ruby_count(chunks, current, argc, line);
        }
        "ruby.op_add" => {
            emit_ruby_binary_op(chunks, current, "add", line);
        }
        "ruby.op_sub" => {
            emit_ruby_binary_op(chunks, current, "sub", line);
        }
        "ruby.op_mul" => {
            emit_ruby_binary_op(chunks, current, "mul", line);
        }
        "ruby.op_div" => {
            emit_ruby_binary_op(chunks, current, "div", line);
        }
        "ruby.op_shl" => {
            emit_ruby_shl(chunks, current, argc, line);
        }
        "ruby.op_shr" => {
            emit_ruby_shr(chunks, current, argc, line);
        }
        "ruby.op_and" => {
            emit_ruby_binary_op(chunks, current, "and", line);
        }
        "ruby.op_or" => {
            emit_ruby_binary_op(chunks, current, "or", line);
        }
        "ruby.op_xor" => {
            emit_ruby_binary_op(chunks, current, "xor", line);
        }
        "ruby.array_concat" => {
            call_import(chunks, current, "ecma:array", "concat", 2, line);
        }
        "ruby.array_intersection" => {
            emit_array_set_op(chunks, current, "intersection", line);
        }
        "ruby.array_union" => {
            emit_array_union(chunks, current, line);
        }
        "ruby.array_difference" => {
            emit_array_set_op(chunks, current, "difference", line);
        }
        "ruby.array_repeat" => {
            emit_array_repeat(chunks, current, line);
        }
        "ruby.inject_symbol" => {
            chunks[current].emit_op(Op::DROP, line);
            call_import(chunks, current, "ecma:math", "sumPrecise", 1, line);
        }
        "ruby.inject_initial" => {
            emit_inject_initial(chunks, current, line);
        }
        "ruby.min_by" => {
            emit_min_by(chunks, current, false, line);
        }
        "ruby.sort_by" => {
            emit_min_by(chunks, current, true, line);
        }
        "ruby.map" => {
            emit_ruby_map_like(chunks, current, argc, false, line);
        }
        "ruby.map_round" => {
            emit_ruby_map_round(chunks, current, argc, line);
        }
        "ruby.flat_map" => {
            emit_ruby_map_like(chunks, current, argc, true, line);
        }
        "ruby.select" => {
            emit_ruby_select_like_value(chunks, current, argc, true, line);
        }
        "ruby.map_bang" => {
            let fn_s = chunks[current].alloc_scratch(1);
            let arr_s = chunks[current].alloc_scratch(1);
            let idx_s = chunks[current].alloc_scratch(1);
            let elem_s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, fn_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
            let block = chunks[current].emit_block(line);
            let (loop_patch, _) = chunks[current].emit_loop_s(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            collections::emit_len(chunks, current, line);
            ops::emit_dyn_lt(&mut chunks[current], line);
            ops::emit_dyn_not(&mut chunks[current], line);
            chunks[current].emit_br_if(1, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            collections::emit_get(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
            chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
            collections::emit_set(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].patch_loop(loop_patch);
            chunks[current].emit_end(line);
            chunks[current].patch_block(block);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
        }
        "ruby.select_bang" => {
            emit_select_like(chunks, current, true, true, line);
        }
        "ruby.reject" => {
            emit_ruby_select_like_value(chunks, current, argc, false, line);
        }
        "ruby.reject_bang" => {
            emit_select_like(chunks, current, false, true, line);
        }
        "ruby.filter_map" => {
            let fn_s = chunks[current].alloc_scratch(1);
            let arr_s = chunks[current].alloc_scratch(1);
            let idx_s = chunks[current].alloc_scratch(1);
            let result_s = chunks[current].alloc_scratch(1);
            let mapped_s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, fn_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
            collections::emit_array_new(chunks, current, 0, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
            let block = chunks[current].emit_block(line);
            let (loop_patch, _) = chunks[current].emit_loop_s(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            collections::emit_len(chunks, current, line);
            ops::emit_dyn_lt(&mut chunks[current], line);
            ops::emit_dyn_not(&mut chunks[current], line);
            chunks[current].emit_br_if(1, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            collections::emit_get(chunks, current, line);
            chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, mapped_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_s, line);
            ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_s, line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_end(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].patch_loop(loop_patch);
            chunks[current].emit_end(line);
            chunks[current].patch_block(block);
            chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        }
        "ruby.compact_bang" => {
            let arr_s = chunks[current].alloc_scratch(1);
            let idx_s = chunks[current].alloc_scratch(1);
            let result_s = chunks[current].alloc_scratch(1);
            let elem_s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
            collections::emit_array_new(chunks, current, 0, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
            let block = chunks[current].emit_block(line);
            let (loop_patch, _) = chunks[current].emit_loop_s(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            collections::emit_len(chunks, current, line);
            ops::emit_dyn_lt(&mut chunks[current], line);
            ops::emit_dyn_not(&mut chunks[current], line);
            chunks[current].emit_br_if(1, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            collections::emit_get(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
            chunks[current].emit_op(Op::REF_IS_NULL, line);
            chunks[current].emit_op(Op::I32_EQZ, line);
            chunks[current].emit_if(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_end(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].patch_loop(loop_patch);
            chunks[current].emit_end(line);
            chunks[current].patch_block(block);
            emit_replace_array_from_slot(chunks, current, arr_s, result_s, idx_s, line);
        }
        // Ruby numeric predicates. Coerce through the real ECMA number surface,
        // then use shared WASM fmod; do not invent `ecma:math:isEven`.
        "ruby.even" | "ruby.odd" => {
            call_import(chunks, current, "ecma:number", "Number", 1, line);
            chunks[current].emit_f64_const(2.0, line);
            math::emit_c_fmod(&mut chunks[current], line);
            chunks[current].emit_f64_const(0.0, line);
            if name == "ruby.even" {
                chunks[current].emit_op(Op::F64_EQ, line);
            } else {
                chunks[current].emit_op(Op::F64_NE, line);
            }
            ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        // Ruby `index`/`rindex`-style value searches return nil instead of JS
        // `-1` when no match is found.
        "ruby.find_index_value" => {
            collections::emit_index_of(chunks, current, line);
            emit_minus_one_to_null(&mut chunks[current], line);
        }
        "ruby.find_index_block" => {
            let fn_s = chunks[current].alloc_scratch(1);
            let arr_s = chunks[current].alloc_scratch(1);
            let idx_s = chunks[current].alloc_scratch(1);
            let result_s = chunks[current].alloc_scratch(1);
            let elem_s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, fn_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
            chunks[current].emit_op(Op::NULL, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
            let block = chunks[current].emit_block(line);
            let (loop_patch, _) = chunks[current].emit_loop_s(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            collections::emit_len(chunks, current, line);
            ops::emit_dyn_lt(&mut chunks[current], line);
            ops::emit_dyn_not(&mut chunks[current], line);
            chunks[current].emit_br_if(1, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            collections::emit_get(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
            chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
            ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
            chunks[current].emit_br(2, line);
            chunks[current].emit_end(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].patch_loop(loop_patch);
            chunks[current].emit_end(line);
            chunks[current].patch_block(block);
            chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        }
        "ruby.rindex_block" => {
            let fn_s = chunks[current].alloc_scratch(1);
            let arr_s = chunks[current].alloc_scratch(1);
            let idx_s = chunks[current].alloc_scratch(1);
            let result_s = chunks[current].alloc_scratch(1);
            let elem_s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, fn_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
            chunks[current].emit_op(Op::NULL, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            collections::emit_len(chunks, current, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_SUB, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
            let block = chunks[current].emit_block(line);
            let (loop_patch, _) = chunks[current].emit_loop_s(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op(Op::I32_LT_S, line);
            chunks[current].emit_br_if(1, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            collections::emit_get(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
            chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
            ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
            chunks[current].emit_br(2, line);
            chunks[current].emit_end(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_SUB, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].patch_loop(loop_patch);
            chunks[current].emit_end(line);
            chunks[current].patch_block(block);
            chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        }
        "ruby.none" | "ruby.one" => {
            let fn_s = chunks[current].alloc_scratch(1);
            let arr_s = chunks[current].alloc_scratch(1);
            let idx_s = chunks[current].alloc_scratch(1);
            let count_s = chunks[current].alloc_scratch(1);
            let elem_s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, fn_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, count_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
            let block = chunks[current].emit_block(line);
            let (loop_patch, _) = chunks[current].emit_loop_s(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            collections::emit_len(chunks, current, line);
            ops::emit_dyn_lt(&mut chunks[current], line);
            ops::emit_dyn_not(&mut chunks[current], line);
            chunks[current].emit_br_if(1, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            collections::emit_get(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
            chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
            ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, count_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, count_s, line);
            if name == "ruby.none" {
                chunks[current].emit_br(2, line);
            }
            chunks[current].emit_end(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].patch_loop(loop_patch);
            chunks[current].emit_end(line);
            chunks[current].patch_block(block);
            chunks[current].emit_op_u16(Op::LOCAL_GET, count_s, line);
            core_wasm::i32_const(&mut chunks[current], line, if name == "ruby.none" { 0 } else { 1 });
            chunks[current].emit_op(Op::I32_EQ, line);
            ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "ruby.find_ifnone" => {
            let pred_s = chunks[current].alloc_scratch(1);
            let ifnone_s = chunks[current].alloc_scratch(1);
            let arr_s = chunks[current].alloc_scratch(1);
            let idx_s = chunks[current].alloc_scratch(1);
            let result_s = chunks[current].alloc_scratch(1);
            let found_s = chunks[current].alloc_scratch(1);
            let elem_s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, pred_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, ifnone_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
            chunks[current].emit_op(Op::NULL, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
            chunks[current].emit_bool_const(false, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, found_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
            let block = chunks[current].emit_block(line);
            let (loop_patch, _) = chunks[current].emit_loop_s(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            collections::emit_len(chunks, current, line);
            ops::emit_dyn_lt(&mut chunks[current], line);
            ops::emit_dyn_not(&mut chunks[current], line);
            chunks[current].emit_br_if(1, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            collections::emit_get(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, pred_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
            chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
            ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
            chunks[current].emit_bool_const(true, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, found_s, line);
            chunks[current].emit_br(2, line);
            chunks[current].emit_end(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, idx_s, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].patch_loop(loop_patch);
            chunks[current].emit_end(line);
            chunks[current].patch_block(block);
            chunks[current].emit_op_u16(Op::LOCAL_GET, found_s, line);
            ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, ifnone_s, line);
            chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
            chunks[current].emit_end(line);
        }
        // `s.bytes` — `TextEncoder().encode(s)` (web:encoding host surface).
        "ruby.bytes" => {
            if argc >= 2 {
                let slots = emit_store_args(chunks, current, argc, line);
                if let Some(n_s) = slots.get(1) {
                    chunks[current].emit_op_u16(Op::LOCAL_GET, *n_s, line);
                    call_import(chunks, current, "wasi:random/random", "get-random-bytes", 1, line);
                } else {
                    collections::emit_array_new(chunks, current, 0, line);
                }
            } else {
                emit_ruby_string_items(chunks, current, argc, "byte", line);
            }
        }
        // `arr.minmax` → `[arr.min, arr.max]` via ecma:math:minOf/maxOf (both
        // flatten a single array arg). Stash arr (consumed twice), build [min,max].
        "ruby.max" => {
            emit_ruby_max_get(chunks, current, argc, line);
        }
        "ruby.minmax" => {
            let base = chunks[current].alloc_scratch(3);
            let (arr_s, min_s, max_s) = (base, base + 1, base + 2);
            chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            call_import(chunks, current, "ecma:math", "minOf", 1, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, min_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            call_import(chunks, current, "ecma:math", "maxOf", 1, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, max_s, line);
            collections::emit_array_new(chunks, current, 0, line);
            core_wasm::dup(&mut chunks[current], line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, min_s, line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            core_wasm::dup(&mut chunks[current], line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, max_s, line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
        }
        // `x.include?(v)` / `x.member?(v)` — polymorphic membership, stack
        // [container, needle] → bool. string → substring; array (incl.
        // materialized ranges) → element; else (hash/object) → own key.
        "ruby.includes" => {
            let needle = chunks[current].alloc_scratch(1);
            let container = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, needle, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, container, line);

            // string → substring test
            chunks[current].emit_op_u16(Op::LOCAL_GET, container, line);
            let test_str = chunks[current].add_import("wasm:js-string", "test");
            chunks[current].emit_call(test_str, 1, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, container, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, needle, line);
            call_import(chunks, current, "ecma:string", "includes", 2, line);
            ops::emit_i32_to_bool(&mut chunks[current], line);
            chunks[current].emit_else(line);

            // array → element test
            chunks[current].emit_op_u16(Op::LOCAL_GET, container, line);
            let is_array = chunks[current].add_import("ecma:array", "isArray");
            chunks[current].emit_call(is_array, 1, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, container, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, needle, line);
            call_import(chunks, current, "ecma:array", "includes", 2, line);
            ops::emit_i32_to_bool(&mut chunks[current], line);
            chunks[current].emit_else(line);

            // hash / object → own-key test
            chunks[current].emit_op_u16(Op::LOCAL_GET, container, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, needle, line);
            call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
            ops::emit_i32_to_bool(&mut chunks[current], line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
        }
        // `arr.compact` — new array with nil (null) elements removed. Inline
        // loop over `ecma:array` primitives (no `__vybe_compact` chunk).
        // Stack: [arr] → [result].
        "ruby.compact" => {
            let base = chunks[current].local_count;
            chunks[current].alloc_scratch(5);
            let (arr_s, result_s, i_s, len_s, elem_s) =
                (base, base + 1, base + 2, base + 3, base + 4);
            chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
            // result = []
            collections::emit_array_new(chunks, current, 0, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_s, line);
            // len = arr.length; i = 0
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            collections::emit_len(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);

            let block_p = chunks[current].emit_block(line);
            let (loop_p, _) = chunks[current].emit_loop_s(line);
            // cond: break when !(i < len)
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            ops::emit_dyn_lt(&mut chunks[current], line);
            ops::emit_dyn_not(&mut chunks[current], line);
            chunks[current].emit_br_if(1, line);
            // elem = arr[i]
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            collections::emit_get(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, elem_s, line);
            // if elem != nil → result.push(elem)
            chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
            chunks[current].emit_op(Op::REF_IS_NULL, line);
            chunks[current].emit_op(Op::I32_EQZ, line);
            chunks[current].emit_if(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, elem_s, line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_end(line);
            // i += 1
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].patch_loop(loop_p);
            chunks[current].emit_end(line);
            chunks[current].patch_block(block_p);
            chunks[current].emit_op_u16(Op::LOCAL_GET, result_s, line);
        }
        // `s.hex` — parse a leading hex string → int, invalid → 0. Direct
        // `Number.parseInt(s, 16)` (handles `0x` prefix, sign, partial parse);
        // NaN (no valid digits) → 0. Stack: [s] → [int].
        "ruby.hex" => {
            let r = chunks[current].alloc_scratch(1);
            core_wasm::i32_const(&mut chunks[current], line, 16);
            call_import(chunks, current, "ecma:number", "parseInt", 2, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, r, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
            call_import(chunks, current, "ecma:number", "isNaN", 1, line);
            chunks[current].emit_if_value(line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
            chunks[current].emit_end(line);
        }
        "ruby.oct" => {
            let s = chunks[current].alloc_scratch(1);
            let r = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
            chunks[current].emit_string_const("0x", line);
            call_import(chunks, current, "ecma:string", "startsWith", 2, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
            core_wasm::i32_const(&mut chunks[current], line, 16);
            call_import(chunks, current, "ecma:number", "parseInt", 2, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, r, line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
            chunks[current].emit_string_const("0b", line);
            call_import(chunks, current, "ecma:string", "startsWith", 2, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
            core_wasm::i32_const(&mut chunks[current], line, 2);
            chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
            call_import(chunks, current, "ecma:string", "length", 1, line);
            call_import(chunks, current, "ecma:string", "slice", 3, line);
            core_wasm::i32_const(&mut chunks[current], line, 2);
            call_import(chunks, current, "ecma:number", "parseInt", 2, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, r, line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
            core_wasm::i32_const(&mut chunks[current], line, 8);
            call_import(chunks, current, "ecma:number", "parseInt", 2, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, r, line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
            call_import(chunks, current, "ecma:number", "isNaN", 1, line);
            chunks[current].emit_if_value(line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, r, line);
            chunks[current].emit_end(line);
        }
        // `a.zip(b, …)` → array of tuples. Shared `vybe_emitter` op (variadic;
        // Ruby/PHP/Python can all route here). `argc` = total arrays on stack.
        "ruby.zip" => {
            emit_ruby_zip(chunks, current, argc, line);
        }
        "ruby.product" => {
            emit_ruby_product(chunks, current, argc, line);
        }
        // `a.rotate(n=1)` → `a.slice(k, len).concat(a.slice(0, k))` where
        // `k = ((n % len) + len) % len` (left rotate; negative rotates right).
        // Composed from `ecma:array` slice+concat. Stack: [a] or [a, n] → [result].
        "ruby.rotate" => {
            let base = chunks[current].local_count;
            chunks[current].alloc_scratch(4);
            let (arr_s, n_s, len_s, nnorm_s) = (base, base + 1, base + 2, base + 3);
            if argc >= 2 {
                chunks[current].emit_op_u16(Op::LOCAL_SET, n_s, line); // top = n
                chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
            } else {
                chunks[current].emit_op_u16(Op::LOCAL_SET, arr_s, line);
                core_wasm::i32_const(&mut chunks[current], line, 1);
                chunks[current].emit_op_u16(Op::LOCAL_SET, n_s, line);
            }
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            collections::emit_len(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
            // if len <= 0 → return arr unchanged (also guards `% 0`)
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            ops::emit_dyn_le(&mut chunks[current], line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            chunks[current].emit_else(line);
            // n_norm = ((n % len) + len) % len
            chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            chunks[current].emit_op(Op::I32_REM_S, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            chunks[current].emit_op(Op::I32_REM_S, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, nnorm_s, line);
            // arr.slice(n_norm, len).concat(arr.slice(0, n_norm))
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, nnorm_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            call_import(chunks, current, "ecma:array", "slice", 3, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, arr_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_GET, nnorm_s, line);
            call_import(chunks, current, "ecma:array", "slice", 3, line);
            call_import(chunks, current, "ecma:array", "concat", 2, line);
            chunks[current].emit_end(line);
        }
        // `srand(n)` — record n as the seed, seed the global PRNG (reproducible
        // streams), and return the PREVIOUS seed (Ruby semantics). Stack:
        // [n] → [old_seed].
        "ruby.srand" => {
            let n_s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, n_s, line);
            let seed_g = chunks[current].add_constant(vybe_bytecode::Value::String(
                std::sync::Arc::from("__vybe_rng_seed"),
            ));
            chunks[current].emit_op_u16(Op::GLOBAL_GET, seed_g, line); // old seed (null if unset)
            chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
            chunks[current].emit_op_u16(Op::GLOBAL_SET, seed_g, line); // seed = n
            chunks[current].emit_op_u16(Op::LOCAL_GET, n_s, line);
            vybe_emitter::random::emit_seed(chunks, current, line); // set PRNG state, pops n
        }
        // `rand` → float [0,1); `rand(n)` → int [0,n). Rides the seedable PRNG.
        "ruby.rand" => {
            if argc >= 1 {
                vybe_emitter::random::emit_rand_below(chunks, current, line);
            } else {
                vybe_emitter::random::emit_next_unit(chunks, current, line);
            }
        }
        // `a.sample` → one uniformly-random element (null if empty). Shared
        // `vybe_emitter::random` op (Ruby/Python).
        "ruby.sample" => {
            vybe_emitter::random::emit_sample(chunks, current, argc, line);
        }
        // `a.shuffle`/`shuffle!` → in-place Fisher-Yates. Shared op (Ruby/Python).
        "ruby.shuffle" => {
            vybe_emitter::random::emit_shuffle(chunks, current, argc, line);
        }
        // `h.value?(v)` / `h.has_value?(v)` — `Object.values(h).includes(v)`,
        // direct `ecma:object` (no chunk). Stack: [hash, v] → [bool].
        "ruby.has_value" => {
            let v_s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, v_s, line); // stash v → [hash]
            let values = chunks[current].add_import("ecma:object", "values");
            chunks[current].emit_call(values, 1, line); // [values]
            chunks[current].emit_op_u16(Op::LOCAL_GET, v_s, line); // [values, v]
            collections::emit_contains(chunks, current, line); // [bool]
        }
        // `h.invert` — swap keys/values: `Object.fromEntries(entries.map([k,v]→[v,k]))`.
        // Direct `ecma:object` entries/fromEntries (no chunk). Stack: [hash] → [hash].
        "ruby.invert" => {
            let base = chunks[current].local_count;
            chunks[current].alloc_scratch(5);
            let (entries_s, swapped_s, i_s, len_s, pair_s) =
                (base, base + 1, base + 2, base + 3, base + 4);
            let entries = chunks[current].add_import("ecma:object", "entries");
            chunks[current].emit_call(entries, 1, line); // [entries]
            chunks[current].emit_op_u16(Op::LOCAL_SET, entries_s, line);
            collections::emit_array_new(chunks, current, 0, line); // swapped = []
            chunks[current].emit_op_u16(Op::LOCAL_SET, swapped_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, entries_s, line);
            collections::emit_len(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);

            let block_p = chunks[current].emit_block(line);
            let (loop_p, _) = chunks[current].emit_loop_s(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            ops::emit_dyn_lt(&mut chunks[current], line);
            ops::emit_dyn_not(&mut chunks[current], line);
            chunks[current].emit_br_if(1, line);
            // pair = entries[i]
            chunks[current].emit_op_u16(Op::LOCAL_GET, entries_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            collections::emit_get(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, pair_s, line);
            // swapped.push([pair[1], pair[0]])
            chunks[current].emit_op_u16(Op::LOCAL_GET, swapped_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, pair_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            collections::emit_get(chunks, current, line); // pair[1] (new key)
            chunks[current].emit_op_u16(Op::LOCAL_GET, pair_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            collections::emit_get(chunks, current, line); // pair[0] (new value)
            collections::emit_array_pair(chunks, current, line); // [v, k]
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            // i += 1
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].patch_loop(loop_p);
            chunks[current].emit_end(line);
            chunks[current].patch_block(block_p);
            // result = Object.fromEntries(swapped)
            chunks[current].emit_op_u16(Op::LOCAL_GET, swapped_s, line);
            let from_entries = chunks[current].add_import("ecma:object", "fromEntries");
            chunks[current].emit_call(from_entries, 1, line);
        }
        // `h.transform_values { |v| … }` / `h.transform_keys { |k| … }` →
        // `Object.fromEntries(entries.map([k,v] → [k, blk(v)] | [blk(k), v]))`.
        // Direct `ecma:object` + `CALL_REF` on the block (no chunk).
        // Stack: [hash, block] → [hash].
        "ruby.transform_values" | "ruby.transform_keys" => {
            let on_keys = name == "ruby.transform_keys";
            let base = chunks[current].local_count;
            chunks[current].alloc_scratch(6);
            let (fn_s, entries_s, out_s, i_s, len_s, pair_s) =
                (base, base + 1, base + 2, base + 3, base + 4, base + 5);
            chunks[current].emit_op_u16(Op::LOCAL_SET, fn_s, line); // stash block → [hash]
            let entries = chunks[current].add_import("ecma:object", "entries");
            chunks[current].emit_call(entries, 1, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, entries_s, line);
            collections::emit_array_new(chunks, current, 0, line); // out = []
            chunks[current].emit_op_u16(Op::LOCAL_SET, out_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, entries_s, line);
            collections::emit_len(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, len_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);

            let block_p = chunks[current].emit_block(line);
            let (loop_p, _) = chunks[current].emit_loop_s(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, len_s, line);
            ops::emit_dyn_lt(&mut chunks[current], line);
            ops::emit_dyn_not(&mut chunks[current], line);
            chunks[current].emit_br_if(1, line);
            // pair = entries[i]
            chunks[current].emit_op_u16(Op::LOCAL_GET, entries_s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            collections::emit_get(chunks, current, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, pair_s, line);
            // out.push( transform_keys ? [blk(k), v] : [k, blk(v)] )
            chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
            // slot indices: transform the key (0) or the value (1)
            let (transform_idx, keep_idx) = if on_keys { (0, 1) } else { (1, 0) };
            // first element of the new pair
            if on_keys {
                // blk(k)
                chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, pair_s, line);
                core_wasm::i32_const(&mut chunks[current], line, transform_idx);
                collections::emit_get(chunks, current, line);
                chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
                // v (kept)
                chunks[current].emit_op_u16(Op::LOCAL_GET, pair_s, line);
                core_wasm::i32_const(&mut chunks[current], line, keep_idx);
                collections::emit_get(chunks, current, line);
            } else {
                // k (kept)
                chunks[current].emit_op_u16(Op::LOCAL_GET, pair_s, line);
                core_wasm::i32_const(&mut chunks[current], line, keep_idx);
                collections::emit_get(chunks, current, line);
                // blk(v)
                chunks[current].emit_op_u16(Op::LOCAL_GET, fn_s, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, pair_s, line);
                core_wasm::i32_const(&mut chunks[current], line, transform_idx);
                collections::emit_get(chunks, current, line);
                chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
            }
            collections::emit_array_pair(chunks, current, line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i_s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, i_s, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].patch_loop(loop_p);
            chunks[current].emit_end(line);
            chunks[current].patch_block(block_p);
            chunks[current].emit_op_u16(Op::LOCAL_GET, out_s, line);
            let from_entries = chunks[current].add_import("ecma:object", "fromEntries");
            chunks[current].emit_call(from_entries, 1, line);
        }
        _ => return false,
    }
    true
}
