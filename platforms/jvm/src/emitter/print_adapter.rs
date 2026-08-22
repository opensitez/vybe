//! `java.io.PrintStream` / `java.util.Formatter` output surface.
//!
//! Moved from `languages/java` — the print contract is JDK spec
//! (`System.out.println(Object)` IS `String.valueOf(x)`, JLS/PrintStream),
//! identical for every JVM frontend, so the platform owns it.

use vybe_compiler::primitives::instructions::host;
use vybe_compiler::primitives::io;
use vybe_compiler::primitives::strings;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// Write the string on top of the stack to stdout. Stack: `[text] → []`.
///
/// This used to be 0.2's pair — `wasi:cli/stdout.get-stdout()` handing back an
/// `output-stream` resource, then `wasi:io/streams.[method]output-stream.
/// blocking-write-and-flush`. `wasi:io` does not exist in WASI 0.3.1: streams
/// became a Component Model TYPE, so there is no interface left to declare a
/// stream resource, and `get-stdout` went with it — `wasi:cli/stdout` declares
/// only `write-via-stream(data: stream<u8>)`.
///
/// Both halves resolved against this host, which is why nothing failed and this
/// was the LAST emitter of `wasi:io` in the tree. `primitives::io` already had
/// the 0.3.1 transport; this file simply never adopted it.
fn emit_stdout_text(chunk: &mut Chunk, line: u32) {
    let text_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    io::emit_write_stdout_slot(chunk, text_slot, line);
}

/// `System.out` evaluates to the `PrintStream` identity sentinel, and every
/// write returns it (JLS: `append`/`format` return `this`).
fn emit_print_stream_sentinel(chunk: &mut Chunk, line: u32) {
    chunk.emit_string_const("__java_out", line);
}

/// `println(Object)` — `String.valueOf(x)` then a line write.
pub fn emit_println(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
    } else {
        crate::emitter::string_adapter::emit_value_of(chunks, current, line);
    }
    host::emit(&mut chunks[current], "web:console", "log", 1, line);
    emit_print_stream_sentinel(&mut chunks[current], line);
}

/// `print(Object)` — `String.valueOf(x)` to real WASI stdout, no newline.
pub fn emit_print_no_newline(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
    }
    crate::emitter::string_adapter::emit_value_of(chunks, current, line);
    emit_stdout_text(&mut chunks[current], line);
    emit_print_stream_sentinel(&mut chunks[current], line);
}

/// `printf(fmt, args...)` — the shared sprintf engine, then stdout.
pub fn emit_printf(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    vybe_compiler::primitives::sprintf::emit_sprintf(chunks, current, argc, line);
    emit_stdout_text(&mut chunks[current], line);
    emit_print_stream_sentinel(&mut chunks[current], line);
}

/// `printf(fmt, argsArray)` — args pre-packed by the runtime prelude.
pub fn emit_printf_array(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    vybe_compiler::primitives::sprintf::emit_sprintf_from_array(chunks, current, line);
    emit_stdout_text(&mut chunks[current], line);
    emit_print_stream_sentinel(&mut chunks[current], line);
}

/// `%,d` — grouped integer rendering.
pub fn emit_format_grouped_int(chunks: &mut [Chunk], current: usize, line: u32) {
    let to_locale = chunks[current].add_import("ecma:number", "toLocaleString");
    chunks[current].emit_call(to_locale, 1, line);
}

/// `%e` / `%E` — Java's two-digit exponent form.
pub fn emit_format_exp(chunks: &mut [Chunk], current: usize, upper: bool, line: u32) {
    chunks[current].emit_i32_const(6, line);
    let to_exp = chunks[current].add_import("ecma:number", "toExponential");
    chunks[current].emit_call(to_exp, 2, line);
    if upper {
        let to_upper = chunks[current].add_import("ecma:string", "toUpperCase");
        chunks[current].emit_call(to_upper, 1, line);
    }

    let (plus, plus_padded, minus, minus_padded) = if upper {
        ("E+", "E+0", "E-", "E-0")
    } else {
        ("e+", "e+0", "e-", "e-0")
    };
    chunks[current].emit_string_const(plus, line);
    chunks[current].emit_string_const(plus_padded, line);
    strings::emit_replace(&mut chunks[current], line);
    chunks[current].emit_string_const(minus, line);
    chunks[current].emit_string_const(minus_padded, line);
    strings::emit_replace(&mut chunks[current], line);
}
