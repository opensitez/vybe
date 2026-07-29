//! .NET `Console.Write` / `Console.WriteLine` — Rust inline emitters.
//!
//! Output goes through proper WASI I/O (`wasi:cli/stdout.get-stdout` +
//! `wasi:io/streams.[method]output-stream.blocking-write-and-flush`), NOT the
//! line-oriented `wasi:logging/logging.log`. That distinction is what lets
//! `Console.Write` emit its text with NO trailing newline while
//! `Console.WriteLine` appends exactly one `\n` — logging always terminated a
//! record, so the two were indistinguishable before.
//!
//! Each value is first run through a .NET-style stringifier:
//! - `bool` → `"True"` / `"False"` (capitalised per .NET spec, vs. JS-style
//!   lowercase `true`/`false` from the default Display impl).
//! - `null` → `""` (matches `Console.WriteLine((string)null)`).
//! - everything else → `String(v)`.
//!
//! Without this conversion, `Console.WriteLine(true)` prints `true` and
//! `is_constant_pattern` etc. fail their .NET-shaped assertions.

use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

/// Stage the .NET string form of the value in `v_local` into `result_local`.
///
/// Structured control flow with a single outer block as the exit; each arm
/// stages its converted string and `br exit`. Emitting a RETURN inside a
/// structured block would leak the block label to the caller's `label_stack`
/// (the trap iter_drain hit), so every path branches to the shared exit.
pub(crate) fn emit_dotnet_stringify(chunk: &mut Chunk, v_local: u16, result_local: u16, line: u32) {
    use vybe_compiler::primitives::instructions::host;

    let exit_block = chunk.emit_block(line);

    // Bool branch → "True" / "False".
    let not_bool = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, v_local, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_string_const("boolean", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, v_local, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_string_const("True", line);
    chunk.emit_else(line);
    chunk.emit_string_const("False", line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);
    chunk.emit_br(1, line); // exit
    chunk.emit_end(line);
    chunk.patch_block(not_bool);

    // Null branch — `Console.WriteLine((string)null)` prints "".
    let not_null = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, v_local, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);
    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);
    chunk.emit_br(1, line);
    chunk.emit_end(line);
    chunk.patch_block(not_null);

    // Default: direct ECMA String(v) coercion. User-defined `ToString`
    // overrides on .NET-shape classes are NOT picked up here yet — that
    // requires routing through method dispatch; call `Console.WriteLine(
    // p.ToString())` explicitly to get the override.
    chunk.emit_op_u16(Op::LOCAL_GET, v_local, line);
    vybe_compiler::primitives::strings::emit_to_string(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);

    chunk.emit_end(line);
    chunk.patch_block(exit_block);
}

/// Write the string in `text_local` to an output stream, then leave `null` on
/// the stack (the call site DROPs print results uniformly).
///
/// Mirrors the proven libc stdout path: `get-<stream>` → output-stream handle,
/// then `blocking-write-and-flush(stream, contents)` — byte-faithful, no
/// implicit newline. Imports are registered on the CURRENT chunk (via
/// `add_import` + `emit_call`) so the `CALL_IMPORT` indices resolve against the
/// same per-chunk table the runtime uses — `chunks[0]` would only be correct at
/// top level.
fn emit_stream_write(
    chunk: &mut Chunk,
    text_local: u16,
    stream_module: &str,
    stream_getter: &str,
    line: u32,
) {
    let get_idx = chunk.add_import(stream_module, stream_getter);
    let write_idx = chunk.add_import(
        "wasi:io/streams",
        "[method]output-stream.blocking-write-and-flush",
    );
    // stream = get-stdout()
    chunk.emit_call(get_idx, 0, line);
    // blocking-write-and-flush(stream, contents)
    chunk.emit_op_u16(Op::LOCAL_GET, text_local, line);
    chunk.emit_call(write_idx, 2, line);
    chunk.emit_op(Op::DROP, line); // discard result<_, stream-error>
    chunk.emit_op(Op::NULL, line);
}

/// `Console.WriteLine(v)` — stringify, append one `\n`, write to stdout.
/// Stack: [v] → [null].
pub fn emit_console_writeline(chunks: &mut [Chunk], current: usize, line: u32) {
    let v_local = alloc_local(&mut chunks[current]);
    let result_local = alloc_local(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v_local, line);
    emit_dotnet_stringify(&mut chunks[current], v_local, result_local, line);

    // result += "\n"
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, result_local, line);
    chunk.emit_string_const("\n", line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);

    emit_stream_write(
        &mut chunks[current],
        result_local,
        "wasi:cli/stdout",
        "get-stdout",
        line,
    );
}

/// `Console.Write(v)` — stringify and write to stdout with NO newline.
/// Stack: [v] → [null].
pub fn emit_console_write(chunks: &mut [Chunk], current: usize, line: u32) {
    let v_local = alloc_local(&mut chunks[current]);
    let result_local = alloc_local(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v_local, line);
    emit_dotnet_stringify(&mut chunks[current], v_local, result_local, line);
    emit_stream_write(
        &mut chunks[current],
        result_local,
        "wasi:cli/stdout",
        "get-stdout",
        line,
    );
}

/// `Console.WriteLine()` — bare newline, no argument. Stack: [] → [null].
pub fn emit_console_writeline_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    let nl_local = alloc_local(&mut chunks[current]);
    chunks[current].emit_string_const("\n", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nl_local, line);
    emit_stream_write(
        &mut chunks[current],
        nl_local,
        "wasi:cli/stdout",
        "get-stdout",
        line,
    );
}

/// `Console.ReadLine()` — wasi:cli/stdin.get-stdin → [method]input-stream.blocking-read.
/// Stack: [] → [string]
pub fn emit_console_readline(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::io::emit_input(&mut chunks[current], line);
}

fn emit_console_stderr(chunks: &mut [Chunk], current: usize, append_newline: bool, line: u32) {
    let v_local = alloc_local(&mut chunks[current]);
    let result_local = alloc_local(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v_local, line);
    emit_dotnet_stringify(&mut chunks[current], v_local, result_local, line);

    if append_newline {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_GET, result_local, line);
        chunk.emit_string_const("\n", line);
        vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);
    }

    emit_stream_write(
        &mut chunks[current],
        result_local,
        "wasi:cli/stderr",
        "get-stderr",
        line,
    );
}

/// `Console.Error.Write(v)` — stringify and write to stderr with NO newline.
/// Stack: [v] → [null].
pub fn emit_console_error_write(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_console_stderr(chunks, current, false, line);
}

/// `Console.Error.WriteLine(v)` — stringify, append `\n`, write to stderr.
/// Stack: [v] → [null].
pub fn emit_console_error_writeline(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_console_stderr(chunks, current, true, line);
}

/// Legacy shared surface name kept for profile/back-compat during migration.
pub fn emit_console_error(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_console_error_writeline(chunks, current, line);
}

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}
