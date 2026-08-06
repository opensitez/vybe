//! I/O compilation — WASI-compatible print, input, file operations.
//!
//! Print uses `web:console.log` (WHATWG Console Standard — `log(...data)`
//! is VARIADIC BY SPEC; each datum rendered, space-joined). The strict
//! `wasi:logging/logging.log(level, context, message)` remains for code
//! calling the WASI interface explicitly.
//! Input uses `wasi:cli/stdin.get-stdin` + `[method]input-stream.blocking-read`.
//! File I/O uses `wasi:filesystem/*` imports.
//!
//! Output BUFFERING lives here too, for one structural reason: this module owns
//! the write. A buffer that some writers respect and others bypass is not a
//! buffer, and that is exactly the bug that existed while each language kept its
//! own — PHP's `echo` checked the buffer and its `var_dump` did not. Routing
//! every write through [`emit_write_or_buffer`] makes capture correct by
//! construction for every writer in every language, present and future.

use std::sync::Arc;

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

/// Emit print/log. Stack: [arg1, ..., argN] → []
/// Routes to `web:console.log` — WHATWG `log(...data)`, variadic by spec
/// (the host renders each datum and joins with a single space).
pub fn emit_print(chunk: &mut Chunk, arg_count: u8, line: u32) {
    let idx = chunk.add_import("web:console", "log");
    chunk.emit_call(idx, arg_count, line);
}

/// Emit print using a pre-resolved import index.
pub fn emit_print_with_import(chunk: &mut Chunk, import_idx: u16, arg_count: u8, line: u32) {
    chunk.emit_call(import_idx, arg_count, line);
}

/// Emit a raw byte write to stdout — NO implicit newline, unlike
/// `wasi:logging/logging.log` which is one line-oriented record per call.
///
/// Composes the WASI 0.3 surface: `canon stream.new` (STREAM_NEW) creates
/// a `stream<u8>` as (readable, writable) i32 handles, the contents go in
/// via STREAM_WRITE, the writable end closes (EOF), and the readable end
/// is passed to `wasi:cli/stdout.write-via-stream(data: stream<u8>)`.
/// The returned `future<result<_, error-code>>` is discarded; both handle
/// table entries are dropped afterwards per the canonical ABI.
///
/// Stack: [] → []. `push_contents` emits the string to write while the
/// writable handle is on the stack. `rd_slot`/`wr_slot` are caller-defined
/// scratch locals for the two canon handles; `write_idx` is the resolved
/// `wasi:cli/stdout.write-via-stream` import.
pub fn emit_write_stdout_with_imports(
    chunk: &mut Chunk,
    write_idx: u16,
    rd_slot: u16,
    wr_slot: u16,
    line: u32,
    push_contents: impl FnOnce(&mut Chunk),
) {
    // canon stream.new → [rd, wr]
    chunk.emit_op(Op::STREAM_NEW, line);
    chunk.emit_op_u16(Op::LOCAL_SET, wr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, rd_slot, line);
    // stream.write(wr, contents)
    chunk.emit_op_u16(Op::LOCAL_GET, wr_slot, line);
    push_contents(chunk);
    chunk.emit_op(Op::STREAM_WRITE, line);
    // stream.drop-writable(wr) — signals EOF
    chunk.emit_op_u16(Op::LOCAL_GET, wr_slot, line);
    chunk.emit_op(Op::STREAM_DROP_WR, line);
    // wasi:cli/stdout.write-via-stream(rd) → future (discard)
    chunk.emit_op_u16(Op::LOCAL_GET, rd_slot, line);
    chunk.emit_call(write_idx, 1, line);
    chunk.emit_op(Op::DROP, line);
    // stream.drop-readable(rd)
    chunk.emit_op_u16(Op::LOCAL_GET, rd_slot, line);
    chunk.emit_op(Op::STREAM_DROP_RD, line);
}

/// Emit print to stderr. Stack: [message] → []
/// WHATWG `console.error(...data)` — the stderr stream of the same
/// console surface.
pub fn emit_print_error(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("web:console", "error");
    chunk.emit_call(idx, 1, line);
}

/// Emit readline. Stack: [] → [string]
/// Calls wasi:cli/stdin.get-stdin → [method]input-stream.blocking-read(stream, 65536).
/// Host returns a string for fd=0 (stdin line input).
pub fn emit_input(chunk: &mut Chunk, line: u32) {
    let stdin_idx = chunk.add_import("wasi:cli/stdin", "get-stdin");
    chunk.emit_call(stdin_idx, 0, line);
    chunk.emit_i32_const(65536, line);
    let read_idx = chunk.add_import("wasi:io/streams", "[method]input-stream.blocking-read");
    chunk.emit_call(read_idx, 2, line);
}

/// Emit readline using a pre-resolved import index.
pub fn emit_input_with_import(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 0, line);
}

/// Emit print + input (prompt then read). Stack: [prompt_string] → [string]
pub fn emit_prompt_input(chunk: &mut Chunk, line: u32) {
    emit_print(chunk, 1, line);
    chunk.emit_op(Op::DROP, line);
    emit_input(chunk, line);
}

/// Read file contents. Stack: [filename] → [contents_string]
pub fn emit_read_file(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "readFile");
    chunk.emit_call(idx, 1, line);
}

/// Write to file. Stack: [target, data...] → [null]
pub fn emit_write_file(chunk: &mut Chunk, argc: u8, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "writeFile");
    chunk.emit_call(idx, argc, line);
}

// ── File handle I/O (wasi:filesystem) ─────────────────────────────────

/// Open file handle. Stack: [filename, mode] → [handle]
pub fn emit_open_file(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "openFile");
    chunk.emit_call(idx, 2, line);
}

/// Close file handle. Stack: [handle] → []
pub fn emit_close_file(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "closeFile");
    chunk.emit_call(idx, 1, line);
}

/// Print to file handle. Stack: [handle, data...] → []
pub fn emit_print_file(chunk: &mut Chunk, argc: u8, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "printFile");
    chunk.emit_call(idx, argc, line);
}

/// Read from file handle (Input). Stack: [handle] → [string]
pub fn emit_input_file(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "inputFile");
    chunk.emit_call(idx, 1, line);
}

/// Read line from file handle. Stack: [handle] → [string]
pub fn emit_line_input(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "lineInput");
    chunk.emit_call(idx, 1, line);
}

// ── Output buffering ────────────────────────────────────────────────────────
//
// A STACK of buffer frames, not a flag plus a "previous" slot. Nesting is the
// normal case (`ob_start()` inside a template inside a handler), and the depth
// is only known at runtime, so the representation has to be a real stack — the
// single-level shape could not express `ob_get_level() == 3` at all.
//
// One frame per active buffer, each carrying its own handler and options,
// because PHP's `ob_list_handlers()` / `ob_get_status(true)` report them
// per-level. Frames are Maps rather than a parallel set of arrays so that
// adding a field is a key, not a fifth global to keep in sync.

/// Global holding the buffer stack: an array of frames, innermost LAST.
/// Absent (null) until the first `ob_start` — see [`emit_ob_stack`].
const OB_STACK: &str = "__vybe_ob_stack";

/// Frame keys. Language-neutral: these are the fields of a buffer, not PHP's
/// spelling of them.
const OB_BUFFER: &str = "buffer";
const OB_HANDLER: &str = "handler";
const OB_CHUNK_SIZE: &str = "chunk_size";
const OB_FLAGS: &str = "flags";

fn key_const(chunk: &mut Chunk, key: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(key)))
}

fn global_get(chunk: &mut Chunk, key: &str, line: u32) {
    crate::primitives::globals::emit_read(chunk, key, line);
}

fn global_set(chunk: &mut Chunk, key: &str, line: u32) {
    crate::primitives::globals::emit_write(chunk, key, line);
}

/// Push the buffer stack, creating it on first use. Stack: [] → [array].
///
/// Lazily created rather than emitted as a module init so that a program which
/// never buffers pays nothing but the null check on its first write.
pub fn emit_ob_stack(chunks: &mut [Chunk], current: usize, line: u32) {
    global_get(&mut chunks[current], OB_STACK, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    super::collections::emit_array_new(chunks, current, 0, line);
    global_set(&mut chunks[current], OB_STACK, line);
    chunks[current].emit_end(line);
    global_get(&mut chunks[current], OB_STACK, line);
}

/// Number of active buffers. Stack: [] → [i32].
pub fn emit_ob_depth(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_ob_stack(chunks, current, line);
    super::collections::emit_array_length(&mut chunks[current], line);
}

/// The innermost frame. Stack: [] → [frame]. Callers must have established
/// depth > 0; there is no empty-stack frame to return.
pub fn emit_ob_top_frame(chunks: &mut [Chunk], current: usize, line: u32) {
    let stack_slot = chunks[current].alloc_scratch(1);
    emit_ob_stack(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stack_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, stack_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stack_slot, line);
    super::collections::emit_array_length(&mut chunks[current], line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    super::collections::emit_get(chunks, current, line);
}

/// Read a field of the innermost frame. Stack: [] → [value].
pub fn emit_ob_top_field(chunks: &mut [Chunk], current: usize, field: &str, line: u32) {
    emit_ob_top_frame(chunks, current, line);
    chunks[current].emit_string_const(field, line);
    super::collections::emit_get(chunks, current, line);
}

/// Write a field of the innermost frame. Stack: [value] → [].
pub fn emit_ob_set_top_field(chunks: &mut [Chunk], current: usize, field: &str, line: u32) {
    let val_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, val_slot, line);

    emit_ob_top_frame(chunks, current, line);
    chunks[current].emit_string_const(field, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, val_slot, line);
    super::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Write a string straight to stdout, bypassing any buffer. Stack: [] → [],
/// contents taken from `str_slot`.
pub fn emit_write_stdout_slot(chunk: &mut Chunk, str_slot: u16, line: u32) {
    let write_idx = chunk.add_import("wasi:cli/stdout", "write-via-stream");
    let rd_slot = chunk.alloc_scratch(1);
    let wr_slot = chunk.alloc_scratch(1);
    emit_write_stdout_with_imports(chunk, write_idx, rd_slot, wr_slot, line, |chunk| {
        chunk.emit_op_u16(Op::LOCAL_GET, str_slot, line);
    });
}

/// THE write. Stack: [string] → [].
///
/// Appends to the innermost buffer when one is active, otherwise goes to
/// stdout. Every language's print/echo/dump should route through this rather
/// than calling stdout directly — that is what makes output capture work
/// uniformly instead of per-builtin.
pub fn emit_write_or_buffer(chunks: &mut [Chunk], current: usize, line: u32) {
    let str_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, str_slot, line);

    emit_ob_depth(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);

    emit_ob_top_field(chunks, current, OB_BUFFER, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, str_slot, line);
    super::strings::emit_concat(&mut chunks[current], 2, line);
    emit_ob_set_top_field(chunks, current, OB_BUFFER, line);

    chunks[current].emit_else(line);
    emit_write_stdout_slot(&mut chunks[current], str_slot, line);
    chunks[current].emit_end(line);
}

/// Open a new buffer. Stack: [handler, chunk_size, flags] (per `argc`, missing
/// trailing args defaulted) → [true].
///
/// Always succeeds, matching PHP: `ob_start()` returns `false` only when a
/// handler refuses to start, which cannot happen for the default handler.
pub fn emit_ob_start(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // Spill the supplied arguments; they arrive in source order so they come
    // off reversed.
    let flags_slot = chunks[current].alloc_scratch(1);
    let chunk_slot = chunks[current].alloc_scratch(1);
    let handler_slot = chunks[current].alloc_scratch(1);
    let slots = [handler_slot, chunk_slot, flags_slot];
    for i in (0..argc as usize).rev() {
        if i < slots.len() {
            chunks[current].emit_op_u16(Op::LOCAL_SET, slots[i], line);
        } else {
            chunks[current].emit_op(Op::DROP, line);
        }
    }
    for (i, slot) in slots.iter().enumerate() {
        if i >= argc as usize {
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, *slot, line);
        }
    }

    let frame_slot = chunks[current].alloc_scratch(1);
    super::collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, frame_slot, line);

    for (field, slot) in [
        (OB_HANDLER, Some(handler_slot)),
        (OB_CHUNK_SIZE, Some(chunk_slot)),
        (OB_FLAGS, Some(flags_slot)),
        (OB_BUFFER, None),
    ] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, frame_slot, line);
        chunks[current].emit_string_const(field, line);
        match slot {
            Some(s) => chunks[current].emit_op_u16(Op::LOCAL_GET, s, line),
            None => chunks[current].emit_string_const("", line) }
        super::collections::emit_set(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }

    emit_ob_stack(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, frame_slot, line);
    super::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_bool_const(true, line);
}

/// Number of active buffers, as the numeric value languages report.
/// Stack: [] → [f64].
pub fn emit_ob_get_level(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_ob_depth(chunks, current, line);
    // Unsigned: a buffer count is never negative, so this is exact.
    chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
}

/// Emit `if depth > 0 { on_active } else { push false }`, leaving exactly one
/// value on the stack either way.
///
/// Every `ob_*` accessor has this shape — PHP returns `false` from all of them
/// when no buffer is active — so the branch is written once here rather than
/// re-derived (and mis-derived) per operation.
fn emit_when_buffering(
    chunks: &mut [Chunk],
    current: usize,
    line: u32,
    on_active: impl FnOnce(&mut [Chunk], usize),
) {
    let result_slot = chunks[current].alloc_scratch(1);
    emit_ob_depth(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    on_active(chunks, current);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Contents of the innermost buffer, or `false`. Stack: [] → [string|false].
pub fn emit_ob_get_contents(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_when_buffering(chunks, current, line, |chunks, current| {
        emit_ob_top_field(chunks, current, OB_BUFFER, line);
    });
}

/// How a language measures the SIZE of buffered output. Stack: [string] → [num].
///
/// Output is bytes, but "length of a string" is not the same question in every
/// language — PHP's `strlen` counts UTF-8 bytes while the shared `emit_len`
/// counts code units, so `"éclair"` is 7 or 6 depending who asks. The language
/// supplies its own, exactly as it supplies `default_name`.
pub type LengthEmit = fn(&mut [Chunk], usize, u32);

fn emit_buffer_len(chunks: &mut [Chunk], current: usize, len: Option<LengthEmit>, line: u32) {
    match len {
        Some(emit) => emit(chunks, current, line),
        None => super::collections::emit_len(chunks, current, line) }
}

/// Size of the innermost buffer, or `false`. Stack: [] → [num|false].
pub fn emit_ob_get_length(
    chunks: &mut [Chunk],
    current: usize,
    len: Option<LengthEmit>,
    line: u32,
) {
    emit_when_buffering(chunks, current, line, |chunks, current| {
        emit_ob_top_field(chunks, current, OB_BUFFER, line);
        emit_buffer_len(chunks, current, len, line);
    });
}

/// Discard the innermost buffer's contents, keeping it open. Stack: [] → [bool].
pub fn emit_ob_clean(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_when_buffering(chunks, current, line, |chunks, current| {
        chunks[current].emit_string_const("", line);
        emit_ob_set_top_field(chunks, current, OB_BUFFER, line);
        chunks[current].emit_bool_const(true, line);
    });
}

/// Pop the innermost frame and leave its buffer contents on the stack.
/// Stack: [] → [string]. Caller must have established depth > 0.
fn emit_ob_pop(chunks: &mut [Chunk], current: usize, line: u32) {
    let contents_slot = chunks[current].alloc_scratch(1);
    emit_ob_top_field(chunks, current, OB_BUFFER, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, contents_slot, line);
    emit_ob_stack(chunks, current, line);
    super::collections::emit_pop(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, contents_slot, line);
}

/// Close the innermost buffer, discarding its contents. Stack: [] → [bool].
pub fn emit_ob_end_clean(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_when_buffering(chunks, current, line, |chunks, current| {
        emit_ob_pop(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_bool_const(true, line);
    });
}

/// Pop the innermost frame and write its contents outward, through the frame's
/// handler if it has one. Stack: [] → [raw_contents].
///
/// The RAW buffer is what comes back, while the HANDLED text is what gets
/// written — `ob_get_flush()` returns the former and outputs the latter, so the
/// two must not be conflated. The write happens after the pop, so
/// `emit_write_or_buffer` naturally targets the enclosing buffer (or stdout at
/// depth 0), which is where a flushed buffer's contents belong.
fn emit_ob_pop_and_flush(chunks: &mut [Chunk], current: usize, line: u32) {
    let handler_slot = chunks[current].alloc_scratch(1);
    let raw_slot = chunks[current].alloc_scratch(1);

    emit_ob_top_field(chunks, current, OB_HANDLER, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, handler_slot, line);
    emit_ob_pop(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, raw_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, raw_slot, line);
    emit_write_or_buffer(chunks, current, line);
    chunks[current].emit_else(line);
    // Call convention: [func_ref, arg0] then CALL_REF. One argument — the
    // buffer — matching the single-parameter handlers languages actually write.
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, raw_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    emit_write_or_buffer(chunks, current, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, raw_slot, line);
}

/// Close the innermost buffer and write its contents to the next target out —
/// the enclosing buffer if there is one, otherwise stdout. Stack: [] → [bool].
pub fn emit_ob_end_flush(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_when_buffering(chunks, current, line, |chunks, current| {
        emit_ob_pop_and_flush(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_bool_const(true, line);
    });
}

/// Write the innermost buffer outward and empty it, WITHOUT closing it.
/// Stack: [] → [bool].
///
/// The frame is popped, written past, then pushed back — that is what makes the
/// write land in the ENCLOSING buffer rather than in the frame being flushed,
/// which is where a flush is supposed to go. The Map is a reference, so the
/// frame that comes back is the same one.
pub fn emit_ob_flush(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_when_buffering(chunks, current, line, |chunks, current| {
        let frame_slot = chunks[current].alloc_scratch(1);
        let handler_slot = chunks[current].alloc_scratch(1);
        let contents_slot = chunks[current].alloc_scratch(1);

        emit_ob_top_frame(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, frame_slot, line);
        emit_ob_top_field(chunks, current, OB_HANDLER, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, handler_slot, line);
        emit_ob_top_field(chunks, current, OB_BUFFER, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, contents_slot, line);

        chunks[current].emit_string_const("", line);
        emit_ob_set_top_field(chunks, current, OB_BUFFER, line);

        emit_ob_stack(chunks, current, line);
        super::collections::emit_pop(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, contents_slot, line);
        emit_write_or_buffer(chunks, current, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, contents_slot, line);
        chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
        emit_write_or_buffer(chunks, current, line);
        chunks[current].emit_end(line);

        emit_ob_stack(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, frame_slot, line);
        super::collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);

        chunks[current].emit_bool_const(true, line);
    });
}

/// Flush every still-open buffer, innermost first. Stack: [] → [].
///
/// Emitted at the end of a program: an unclosed buffer is flushed, not thrown
/// away. Each iteration pops one frame and writes it outward, so the innermost
/// contents land in its enclosing buffer and the outermost reach stdout.
pub fn emit_ob_flush_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let block_p = chunks[current].emit_block(line);
    let (loop_p, _) = chunks[current].emit_loop_s(line);
    emit_ob_depth(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_br_if(1, line);
    emit_ob_pop_and_flush(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_p);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_p);
}

/// Build the status record for the frame at `index_slot`. Stack: [] → [map].
fn emit_ob_status_record(
    chunks: &mut [Chunk],
    current: usize,
    stack_slot: u16,
    index_slot: u16,
    default_name: &str,
    len: Option<LengthEmit>,
    line: u32,
) {
    let frame_slot = chunks[current].alloc_scratch(1);
    let out_slot = chunks[current].alloc_scratch(1);
    let name_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, stack_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
    super::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, frame_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, frame_slot, line);
    chunks[current].emit_string_const(OB_HANDLER, line);
    super::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const(default_name, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunks[current].emit_end(line);

    super::collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);

    // `level` is the frame's index; `buffer_used` its current byte count. The
    // rest are the fixed fields the status record is documented to carry.
    for field in [
        "name",
        "type",
        "flags",
        "level",
        "chunk_size",
        "buffer_size",
        "buffer_used",
    ] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
        chunks[current].emit_string_const(field, line);
        match field {
            "name" => chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line),
            "level" => {
                chunks[current].emit_op_u16(Op::LOCAL_GET, index_slot, line);
                chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
            }
            "chunk_size" => {
                chunks[current].emit_op_u16(Op::LOCAL_GET, frame_slot, line);
                chunks[current].emit_string_const(OB_CHUNK_SIZE, line);
                super::collections::emit_get(chunks, current, line);
            }
            "flags" => {
                chunks[current].emit_op_u16(Op::LOCAL_GET, frame_slot, line);
                chunks[current].emit_string_const(OB_FLAGS, line);
                super::collections::emit_get(chunks, current, line);
            }
            "buffer_used" => {
                chunks[current].emit_op_u16(Op::LOCAL_GET, frame_slot, line);
                chunks[current].emit_string_const(OB_BUFFER, line);
                super::collections::emit_get(chunks, current, line);
                emit_buffer_len(chunks, current, len, line);
            }
            _ => chunks[current].emit_f64_const(0.0, line) }
        super::collections::emit_set(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

/// Buffer status. Stack: [] → [map] for the innermost buffer (empty when none),
/// or [array-of-maps], outermost first, when `full` is requested.
pub fn emit_ob_get_status(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    default_name: &str,
    len: Option<LengthEmit>,
    line: u32,
) {
    let full_slot = chunks[current].alloc_scratch(1);
    if argc >= 1 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, full_slot, line);
        for _ in 1..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    } else {
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, full_slot, line);
    }

    let stack_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    let i_slot = chunks[current].alloc_scratch(1);
    emit_ob_stack(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stack_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, full_slot, line);
    super::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    // Full: one record per frame, outermost first.
    super::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let block_p = chunks[current].emit_block(line);
    let (loop_p, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stack_slot, line);
    super::collections::emit_array_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    emit_ob_status_record(chunks, current, stack_slot, i_slot, default_name, len, line);
    super::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_p);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_p);

    chunks[current].emit_else(line);

    // Innermost only — an empty map when nothing is buffering.
    chunks[current].emit_op_u16(Op::LOCAL_GET, stack_slot, line);
    super::collections::emit_array_length(&mut chunks[current], line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stack_slot, line);
    super::collections::emit_array_length(&mut chunks[current], line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    emit_ob_status_record(chunks, current, stack_slot, i_slot, default_name, len, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    super::collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Names of the active handlers, innermost LAST. Stack: [] → [array].
///
/// `default_name` is what a frame with no handler reports. The shared primitive
/// has no word for "no handler" — PHP spells it `"default output handler"` —
/// so the language supplies its own.
pub fn emit_ob_list_handlers(chunks: &mut [Chunk], current: usize, default_name: &str, line: u32) {
    let stack_slot = chunks[current].alloc_scratch(1);
    let out_slot = chunks[current].alloc_scratch(1);
    let i_slot = chunks[current].alloc_scratch(1);

    emit_ob_stack(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stack_slot, line);
    super::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let block_p = chunks[current].emit_block(line);
    let (loop_p, _) = chunks[current].emit_loop_s(line);
    // Break once the index reaches the depth.
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stack_slot, line);
    super::collections::emit_array_length(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    let handler_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stack_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    super::collections::emit_get(chunks, current, line);
    chunks[current].emit_string_const(OB_HANDLER, line);
    super::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, handler_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const(default_name, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, handler_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
    super::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_p);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_p);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

/// Contents of the innermost buffer, then close it discarding. Stack: []
/// → [string|false].
pub fn emit_ob_get_clean(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_when_buffering(chunks, current, line, |chunks, current| {
        // The handler runs on a clean as well as on a flush — it is how a
        // throwing handler surfaces rather than being swallowed — and its
        // result is what comes back. Nothing is WRITTEN: that is the whole
        // difference between `ob_get_clean` and `ob_get_flush`.
        let handler_slot = chunks[current].alloc_scratch(1);
        emit_ob_top_field(chunks, current, OB_HANDLER, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, handler_slot, line);

        let raw_slot = chunks[current].alloc_scratch(1);
        emit_ob_pop(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, raw_slot, line);

        let out_slot = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, raw_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, handler_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, raw_slot, line);
        chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
        chunks[current].emit_end(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    });
}

/// Contents of the innermost buffer, then close it flushing outward.
/// Stack: [] → [string|false].
pub fn emit_ob_get_flush(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_when_buffering(chunks, current, line, |chunks, current| {
        emit_ob_pop_and_flush(chunks, current, line);
    });
}
