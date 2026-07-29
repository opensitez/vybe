//! PHP stdout / output-buffer helpers.
//!
//! Normal output still goes through the existing WASI stdout stream path. When
//! `ob_start()` is active, writes append to a PHP-local buffer global instead.

use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

const OB_ACTIVE: &str = "__php_ob_active";
const OB_BUFFER: &str = "__php_ob_buffer";
const OB_LEVEL: &str = "__php_ob_level";
const OB_PREV_BUFFER: &str = "__php_ob_prev_buffer";

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn key_idx(chunk: &mut Chunk, key: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(key)))
}

fn global_get(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = key_idx(chunk, key);
    chunk.emit_op_u16(Op::GLOBAL_GET, idx, line);
}

fn global_set(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = key_idx(chunk, key);
    chunk.emit_op_u16(Op::GLOBAL_SET, idx, line);
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    chunk.emit_string_const(value, line);
}

fn push_level(chunk: &mut Chunk, value: f64, line: u32) {
    chunk.emit_f64_const(value, line);
    global_set(chunk, OB_LEVEL, line);
}

fn emit_level_gt_one(chunk: &mut Chunk, line: u32) {
    global_get(chunk, OB_LEVEL, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_GT, line);
}

fn decrement_level(chunk: &mut Chunk, line: u32) {
    global_get(chunk, OB_LEVEL, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_SUB, line);
    global_set(chunk, OB_LEVEL, line);
}

fn write_stdout_slot(chunk: &mut Chunk, val_slot: u16, line: u32) {
    let write_idx = chunk.add_import("wasi:cli/stdout", "write-via-stream");
    let rd_slot = alloc_local(chunk);
    let wr_slot = alloc_local(chunk);
    vybe_compiler::primitives::io::emit_write_stdout_with_imports(
        chunk,
        write_idx,
        rd_slot,
        wr_slot,
        line,
        |chunk| {
            chunk.emit_op_u16(Op::LOCAL_GET, val_slot, line);
        },
    );
}

fn direct_stdout_value_from_slot(chunks: &mut [Chunk], current: usize, val_slot: u16, line: u32) {
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, val_slot, line);
    }
    super::string_adapter::emit_echo_stringify(chunks, current, 1, line);
    let chunk = &mut chunks[current];
    let str_slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, str_slot, line);
    emit_php_stdout_write_string_slot(chunk, str_slot, line);
}

fn emit_php_stdout_write_string_slot(chunk: &mut Chunk, str_slot: u16, line: u32) {
    global_get(chunk, OB_ACTIVE, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    global_get(chunk, OB_BUFFER, line);
    chunk.emit_op_u16(Op::LOCAL_GET, str_slot, line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    global_set(chunk, OB_BUFFER, line);

    chunk.emit_else(line);
    write_stdout_slot(chunk, str_slot, line);
    chunk.emit_end(line);
}

pub fn emit_php_stdout_write(chunks: &mut [Chunk], current: usize, line: u32) {
    {
        let chunk = &mut chunks[current];
        let val_slot = alloc_local(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, val_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, val_slot, line);
    }

    super::string_adapter::emit_echo_stringify(chunks, current, 1, line);

    let chunk = &mut chunks[current];
    let str_slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, str_slot, line);
    emit_php_stdout_write_string_slot(chunk, str_slot, line);
}

pub fn emit_php_echo(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let mut slots = Vec::with_capacity(argc as usize);
    for _ in 0..argc {
        let s = alloc_local(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        slots.push(s);
    }
    slots.reverse();
    for s in slots {
        direct_stdout_value_from_slot(chunks, current, s, line);
    }
}

pub fn emit_php_print_expr(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_php_stdout_write(chunks, current, line);
    chunks[current].emit_i32_const(1, line);
}

pub fn emit_ob_start(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    global_get(chunk, OB_ACTIVE, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    global_get(chunk, OB_BUFFER, line);
    global_set(chunk, OB_PREV_BUFFER, line);
    global_get(chunk, OB_LEVEL, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    global_set(chunk, OB_LEVEL, line);
    chunk.emit_else(line);
    push_level(chunk, 1.0, line);
    chunk.emit_end(line);
    chunk.emit_bool_const(true, line);
    global_set(chunk, OB_ACTIVE, line);
    push_str(chunk, "", line);
    global_set(chunk, OB_BUFFER, line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_ob_get_clean(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let buf_slot = alloc_local(chunk);
    global_get(chunk, OB_BUFFER, line);
    chunk.emit_op_u16(Op::LOCAL_SET, buf_slot, line);
    global_get(chunk, OB_ACTIVE, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    emit_level_gt_one(chunk, line);
    chunk.emit_if(line);
    global_get(chunk, OB_PREV_BUFFER, line);
    global_set(chunk, OB_BUFFER, line);
    push_str(chunk, "", line);
    global_set(chunk, OB_PREV_BUFFER, line);
    decrement_level(chunk, line);
    chunk.emit_bool_const(true, line);
    global_set(chunk, OB_ACTIVE, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    global_set(chunk, OB_ACTIVE, line);
    push_str(chunk, "", line);
    global_set(chunk, OB_BUFFER, line);
    push_str(chunk, "", line);
    global_set(chunk, OB_PREV_BUFFER, line);
    push_level(chunk, 0.0, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    global_set(chunk, OB_ACTIVE, line);
    push_level(chunk, 0.0, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
}

pub fn emit_ob_get_contents(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    global_get(&mut chunks[current], OB_BUFFER, line);
}

pub fn emit_ob_end_clean(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    global_get(chunk, OB_ACTIVE, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    emit_level_gt_one(chunk, line);
    chunk.emit_if(line);
    global_get(chunk, OB_PREV_BUFFER, line);
    global_set(chunk, OB_BUFFER, line);
    push_str(chunk, "", line);
    global_set(chunk, OB_PREV_BUFFER, line);
    decrement_level(chunk, line);
    chunk.emit_bool_const(true, line);
    global_set(chunk, OB_ACTIVE, line);
    chunk.emit_else(line);
    global_get(chunk, OB_BUFFER, line);
    vybe_compiler::primitives::strings::emit_length(chunk, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if(line);
    let one_byte_slot = alloc_local(chunk);
    global_get(chunk, OB_BUFFER, line);
    chunk.emit_op_u16(Op::LOCAL_SET, one_byte_slot, line);
    write_stdout_slot(chunk, one_byte_slot, line);
    chunk.emit_end(line);
    chunk.emit_bool_const(false, line);
    global_set(chunk, OB_ACTIVE, line);
    push_str(chunk, "", line);
    global_set(chunk, OB_BUFFER, line);
    push_str(chunk, "", line);
    global_set(chunk, OB_PREV_BUFFER, line);
    push_level(chunk, 0.0, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    global_set(chunk, OB_ACTIVE, line);
    push_level(chunk, 0.0, line);
    chunk.emit_end(line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_ob_end_flush(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let out_slot = {
        let chunk = &mut chunks[current];
        let buf_slot = alloc_local(chunk);
        let out_slot = alloc_local(chunk);
        push_str(chunk, "", line);
        chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
        global_get(chunk, OB_BUFFER, line);
        chunk.emit_op_u16(Op::LOCAL_SET, buf_slot, line);
        global_get(chunk, OB_ACTIVE, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        emit_level_gt_one(chunk, line);
        chunk.emit_if(line);
        global_get(chunk, OB_PREV_BUFFER, line);
        chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
        vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
        global_set(chunk, OB_BUFFER, line);
        push_str(chunk, "", line);
        global_set(chunk, OB_PREV_BUFFER, line);
        decrement_level(chunk, line);
        chunk.emit_bool_const(true, line);
        global_set(chunk, OB_ACTIVE, line);
        chunk.emit_else(line);
        chunk.emit_bool_const(false, line);
        global_set(chunk, OB_ACTIVE, line);
        push_str(chunk, "", line);
        global_set(chunk, OB_BUFFER, line);
        push_str(chunk, "", line);
        global_set(chunk, OB_PREV_BUFFER, line);
        push_level(chunk, 0.0, line);
        chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
        out_slot
    };
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    emit_php_stdout_write(chunks, current, line);
    chunks[current].emit_bool_const(true, line);
}

pub fn emit_ob_get_level(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let level_slot = alloc_local(chunk);
    global_get(chunk, OB_ACTIVE, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    global_get(chunk, OB_LEVEL, line);
    chunk.emit_else(line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, level_slot, line);

    global_get(chunk, OB_ACTIVE, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    direct_stdout_value_from_slot(chunks, current, level_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, level_slot, line);
}

pub fn emit_ob_clean(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    push_str(chunk, "", line);
    global_set(chunk, OB_BUFFER, line);
    chunk.emit_bool_const(true, line);
}

pub fn emit_ob_get_length(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        global_get(chunk, OB_BUFFER, line);
        vybe_compiler::primitives::strings::emit_length(chunk, line);
    }
    let (len_slot, len_str_slot, out_slot) = {
        let chunk = &mut chunks[current];
        let len_slot = alloc_local(chunk);
        let len_str_slot = alloc_local(chunk);
        let out_slot = alloc_local(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
        (len_slot, len_str_slot, out_slot)
    };
    super::string_adapter::emit_echo_stringify(chunks, current, 1, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, len_str_slot, line);
        global_get(chunk, OB_BUFFER, line);
        chunk.emit_op_u16(Op::LOCAL_GET, len_str_slot, line);
        vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
        write_stdout_slot(chunk, out_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    }
}

pub fn emit_ob_implicit_flush(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    chunks[current].emit_bool_const(true, line);
}

pub fn emit_ob_list_handlers(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    push_str(&mut chunks[current], "default output handler", line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 1, line);
}

pub fn emit_ob_gzhandler(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    chunks[current].emit_bool_const(false, line);
}
