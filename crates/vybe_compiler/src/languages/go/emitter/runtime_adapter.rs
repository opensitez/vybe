//! Go runtime-surface helpers routed via `common:go.*`.

use crate::emitter::instructions::host;
use crate::emitter::collections;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::Chunk;

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
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let to_string = chunks[current].add_import("ecma:string", "String");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, to_string, line);
    chunks[current].emit(1, line);
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
    let base = chunk.local_count;
    chunk.local_count += count;
    base
}

fn local_set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}
