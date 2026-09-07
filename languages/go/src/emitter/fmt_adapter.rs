use vybe_compiler::primitives::instructions::host;
use vybe_compiler::primitives::{collections, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "go.fmt_println" => emit_fmt_joined(chunks, current, argc, line, " "),
        "go.fmt_print" => emit_fmt_joined(chunks, current, argc, line, ""),
        "go.fmt_printf" => emit_fmt_printf(chunks, current, argc, line),
        "go.fmt_sprint" => emit_fmt_sprint(chunks, current, argc, line),
        "go.fmt_sprintf" => {
            vybe_compiler::primitives::sprintf::emit_sprintf(chunks, current, argc, line);
        }
        "go.fmt_string" => {
            if argc != 1 {
                return false;
            }
            emit_fmt_string(chunks, current, line);
        }
        "go.fmt_quote" => {
            if argc != 1 {
                return false;
            }
            emit_fmt_quote(chunks, current, line);
        }
        "go.fmt_fix_exp" => {
            if argc != 1 {
                return false;
            }
            emit_fmt_fix_exp(chunks, current, line);
        }
        "go.fmt_slice" => {
            if argc != 1 {
                return false;
            }
            emit_fmt_slice(chunks, current, line);
        }
        _ => return false,
    }
    true
}

fn emit_fmt_sprint(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        emit_string(chunks, current, "", line);
        return;
    }

    let base = alloc_locals(&mut chunks[current], argc as u16);
    for offset in (0..argc as u16).rev() {
        local_set(&mut chunks[current], base + offset, line);
    }

    emit_formatted_local(chunks, current, base, line);
    for offset in 1..argc as u16 {
        emit_formatted_local(chunks, current, base + offset, line);
        host::emit(&mut chunks[current], "wasm:js-string", "concat", 2, line);
    }
}

fn emit_fmt_joined(chunks: &mut [Chunk], current: usize, argc: u8, line: u32, sep: &str) {
    if argc == 0 {
        emit_string(chunks, current, "", line);
        emit_log(chunks, current, line);
        return;
    }

    let base = alloc_locals(&mut chunks[current], argc as u16);
    for offset in (0..argc as u16).rev() {
        local_set(&mut chunks[current], base + offset, line);
    }

    emit_formatted_local(chunks, current, base, line);
    for offset in 1..argc as u16 {
        if !sep.is_empty() {
            emit_string(chunks, current, sep, line);
            host::emit(&mut chunks[current], "wasm:js-string", "concat", 2, line);
        }
        emit_formatted_local(chunks, current, base + offset, line);
        host::emit(&mut chunks[current], "wasm:js-string", "concat", 2, line);
    }

    emit_log(chunks, current, line);
}

fn emit_fmt_printf(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    vybe_compiler::primitives::sprintf::emit_sprintf(chunks, current, argc, line);
    emit_log(chunks, current, line);
}

fn emit_formatted_local(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_array, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    {
        chunks[current].emit_string_const("[", line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        chunks[current].emit_string_const(" ", line);
        let join = chunks[current].add_import("ecma:array", "join");
        chunks[current].emit_call(join, 2, line);
        host::emit(&mut chunks[current], "wasm:js-string", "concat", 2, line);
        chunks[current].emit_string_const("]", line);
        host::emit(&mut chunks[current], "wasm:js-string", "concat", 2, line);
    }
    chunks[current].emit_else(line);
    {
        crate::emitter::errors_adapter::emit_go_error_string_value(
            &mut chunks[current],
            slot,
            line,
        );
    }
    chunks[current].emit_end(line);
}

fn emit_fmt_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    crate::emitter::errors_adapter::emit_go_error_string_value(&mut chunks[current], slot, line);
}

fn emit_fmt_quote(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:json", "stringify", 1, line);
}

fn emit_fmt_fix_exp(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let s = base;
    let n = base + 1;
    let sign = base + 2;
    let digit = base + 3;

    chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_i32_const(3, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_if_value(line);
    {
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
        chunks[current].emit_i32_const(2, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        host::emit(
            &mut chunks[current],
            "wasm:js-string",
            "charCodeAt",
            2,
            line,
        );
        chunks[current].emit_op_u16(Op::LOCAL_SET, sign, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        host::emit(
            &mut chunks[current],
            "wasm:js-string",
            "charCodeAt",
            2,
            line,
        );
        chunks[current].emit_op_u16(Op::LOCAL_SET, digit, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, sign, line);
        chunks[current].emit_i32_const('+' as i32, line);
        chunks[current].emit_op(Op::I32_EQ, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, sign, line);
        chunks[current].emit_i32_const('-' as i32, line);
        chunks[current].emit_op(Op::I32_EQ, line);
        chunks[current].emit_op(Op::I32_OR, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, digit, line);
        chunks[current].emit_i32_const('0' as i32, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_op(Op::I32_AND, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, digit, line);
        chunks[current].emit_i32_const('9' as i32, line);
        chunks[current].emit_op(Op::I32_LE_S, line);
        chunks[current].emit_op(Op::I32_AND, line);
        chunks[current].emit_if_value(line);
        {
            chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
            chunks[current].emit_i32_const(1, line);
            chunks[current].emit_op(Op::I32_SUB, line);
            host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
            chunks[current].emit_string_const("0", line);
            strings::emit_str_concat(&mut chunks[current], line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
            chunks[current].emit_i32_const(1, line);
            chunks[current].emit_op(Op::I32_SUB, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
            host::emit(&mut chunks[current], "ecma:string", "slice", 3, line);
            strings::emit_str_concat(&mut chunks[current], line);
        }
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_end(line);
}

fn emit_fmt_slice(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let value = base;
    let out = base + 1;
    let len = base + 2;
    let idx = base + 3;
    let elem = base + 4;

    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_string_const("[", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond_from_i32(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_string_const(" ", line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    crate::emitter::errors_adapter::emit_go_error_string_value(&mut chunks[current], elem, line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_string_const("]", line);
    strings::emit_str_concat(&mut chunks[current], line);
}

fn emit_log(chunks: &mut [Chunk], current: usize, line: u32) {
    let log = chunks[current].add_import("web:console", "log");
    chunks[current].emit_call(log, 1, line);
}

fn emit_string(chunks: &mut [Chunk], current: usize, value: &str, line: u32) {
    chunks[current].emit_string_const(value, line);
}

fn alloc_locals(chunk: &mut Chunk, count: u16) -> u16 {
    chunk.alloc_scratch(count)
}

fn local_set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}
