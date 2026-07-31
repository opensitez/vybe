//! PHP stdout / output-buffer helpers.
//!
//! PHP's stringification of a value, then the SHARED writer. The buffer stack
//! itself lives in `vybe_compiler::primitives::io` — it is not a PHP concept,
//! and keeping a private copy here is what let `var_dump` escape `ob_start()`.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

/// Stringify one value PHP's way, then hand it to the SHARED writer.
///
/// The buffer check is deliberately not repeated here: `emit_write_or_buffer`
/// owns it, so `echo`, `print` and `var_dump` cannot drift apart on whether
/// output gets captured — which is precisely how `var_dump` used to escape
/// `ob_start()`.
fn direct_stdout_value_from_slot(chunks: &mut [Chunk], current: usize, val_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, val_slot, line);
    super::string_adapter::emit_echo_stringify(chunks, current, 1, line);
    vybe_compiler::primitives::io::emit_write_or_buffer(chunks, current, line);
}

pub fn emit_php_stdout_write(chunks: &mut [Chunk], current: usize, line: u32) {
    super::string_adapter::emit_echo_stringify(chunks, current, 1, line);
    vybe_compiler::primitives::io::emit_write_or_buffer(chunks, current, line);
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

/// `var_dump($a, $b, …)` — one type-annotated dump record per argument.
///
/// Same argc-driven spill as [`emit_php_echo`]: the arguments are already on
/// the stack, so they come off in reverse and are replayed in source order.
/// Like `echo`, it goes through the shared writer, so `ob_start(); var_dump($x);`
/// captures. It previously wrote straight to `wasi:logging` and escaped every
/// active buffer.
pub fn emit_php_var_dump(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let mut slots = Vec::with_capacity(argc as usize);
    for _ in 0..argc {
        let s = alloc_local(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        slots.push(s);
    }
    slots.reverse();
    for s in slots {
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        super::string_adapter::emit_var_dump_stringify(chunks, current, 1, line);
        // One dump record per line, as PHP writes it. This used to come for
        // free from `wasi:logging`, which is line-oriented; the stdout stream
        // is not, so the newline has to be explicit — it is part of the format,
        // not an artifact of the old sink.
        chunks[current].emit_string_const("\n", line);
        vybe_compiler::primitives::strings::emit_concat(&mut chunks[current], 2, line);
        vybe_compiler::primitives::io::emit_write_or_buffer(chunks, current, line);
    }
}

pub fn emit_php_print_expr(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_php_stdout_write(chunks, current, line);
    chunks[current].emit_i32_const(1, line);
}









pub fn emit_ob_implicit_flush(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    chunks[current].emit_bool_const(true, line);
}

pub fn emit_ob_list_handlers(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    vybe_compiler::primitives::io::emit_ob_list_handlers(
        chunks,
        current,
        "default output handler",
        line,
    );
}

pub fn emit_ob_gzhandler(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    chunks[current].emit_bool_const(false, line);
}

/// PHP measures buffered output in UTF-8 BYTES (`strlen`), not code units, so
/// `ob_get_length()` on "éclair" is 7 and not 6. Reuses the same emitter
/// `strlen` itself uses rather than spelling the UTF-8 walk a second time.
fn php_buffer_len(chunks: &mut [Chunk], current: usize, line: u32) {
    super::string_adapter::emit_strlen(chunks, current, 1, line);
}

/// `ob_get_length()` — PHP's byte count, not the shared code-unit count.
pub fn emit_ob_get_length(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    vybe_compiler::primitives::io::emit_ob_get_length(chunks, current, Some(php_buffer_len), line);
}

/// `ob_get_status()` — PHP supplies its own word for a frame with no handler
/// and its own byte count; the shared primitive has neither.
pub fn emit_ob_get_status(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    vybe_compiler::primitives::io::emit_ob_get_status(
        chunks,
        current,
        argc,
        "default output handler",
        Some(php_buffer_len),
        line,
    );
}
