//! Go runtime-surface helpers routed via `common:go.*`.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_compiler::primitives::instructions::host;

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
        "go.panic" => vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line),
        // fmt.Sprintf / __go_sprintf — format to a string (no output). Same
        // runtime formatter as Printf but leaves the result on the stack
        // instead of logging it.
        "go.fmt_sprintf" => {
            vybe_compiler::primitives::sprintf::emit_sprintf(chunks, current, argc, line);
        }
        "go.regex_split_pat_first" => {
            // `regexp.Split(str, pat)` (Go source is pattern-first) → ecma
            // `ecma:regexp.split(str, pat)`. Args arrive as [pat, str] (str on
            // top); swap through scratch locals then call the host fn directly
            // instead of routing through the `__ecma_regexp_split_pat_first`
            // bundle chunk (which was itself just this arg-swap wrapper).
            let base = alloc_locals(&mut chunks[current], 2);
            local_set(&mut chunks[current], base + 1, line); // str
            local_set(&mut chunks[current], base, line); // pat
            chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line); // str
            chunks[current].emit_op_u16(Op::LOCAL_GET, base, line); // pat
            host::emit(&mut chunks[current], "ecma:regexp", "split", 2, line);
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
    // Go prints slices/arrays as "[e1 e2 e3]" (space-separated, bracketed),
    // unlike JS's comma join. Everything else goes through ecma:string.String.
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_array, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
        // else: the value's ToString ROLE if it fills one, otherwise the ECMA
        // `String()` coercion. A Go type says "here is my text form" by having
        // `String() string` — that is all `fmt.Stringer` is — and the
        // normalizer records it as `ProtocolSlot::ToString`, so this reaches it
        // without the walker having to infer the receiver's type first (which
        // is why `go_stringer_call_expr` only ever fired at some call sites).
        vybe_compiler::primitives::expressions::emit_rich_to_string(
            &mut chunks[current],
            slot,
            line,
        );
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
