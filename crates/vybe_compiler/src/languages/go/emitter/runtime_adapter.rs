//! Go runtime-surface helpers routed via `common:go.*`.

use crate::emitter::collections;
use crate::emitter::instructions::host;
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

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
        // fmt.Sprintf / __go_sprintf — format to a string (no output). Same
        // runtime formatter as Printf but leaves the result on the stack
        // instead of logging it.
        "go.fmt_sprintf" => {
            crate::emitter::sprintf::emit_sprintf(chunks, current, argc, line);
        }
        "go.regex_split_pat_first" => {
            collections::emit_runtime_helper_call(
                chunks,
                current,
                "__ecma_regexp_split_pat_first",
                argc,
                line,
            );
        }
        _ => return false,
    }
    true
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
    crate::emitter::sprintf::emit_sprintf(chunks, current, argc, line);
    emit_log(chunks, current, line);
}

fn emit_formatted_local(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    // Go prints slices/arrays as "[e1 e2 e3]" (space-separated, bracketed),
    // unlike JS's comma join. Everything else goes through ecma:string.String.
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_array, 1, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    {
        // then: "[" + join(v, " ") + "]"
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
        // else: ecma:string.String(v)
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        let to_string = chunks[current].add_import("ecma:string", "String");
        chunks[current].emit_call(to_string, 1, line);
    }
    chunks[current].emit_end(line);
}

fn emit_log(chunks: &mut [Chunk], current: usize, line: u32) {
    let log = chunks[current].add_import("wasi:logging/logging", "log");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, log, line);
    chunks[current].emit(1, line);
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
