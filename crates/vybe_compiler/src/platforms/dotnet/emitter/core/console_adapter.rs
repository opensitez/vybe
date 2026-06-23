//! .NET `Console.Write` / `Console.WriteLine` — Rust inline emitters.
//!
//! Maps to `wasi:cli.log` like the previous direct host-call pattern,
//! but inserts a .NET-style stringifier first:
//!
//! - `bool` → `"True"` / `"False"` (capitalised per .NET spec, vs.
//!   JS-style lowercase `true`/`false` that the default Display impl
//!   emits).
//! - `null` → `""` (matches `Console.WriteLine((string)null)`).
//! - everything else → `String(v)` via `ecma:string.String`.
//!
//! Without this conversion, `Console.WriteLine(true)` prints `true`
//! and `is_constant_pattern` etc. fail their .NET-shaped assertions.

use crate::emitter::instructions::host;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::Chunk;

/// `Console.WriteLine(v)` / `Console.Write(v)` — emit the bool/null
/// fixup then dispatch to `wasi:cli.log`. Stack: [v] → [null].
///
/// Single outer block as the structured-control-flow exit; each arm
/// stages its converted string in `result_local` and `br exit`. One
/// log call at the end. Avoids emitting RETURN inside a structured
/// block (which would leak the block label to the caller's
/// `label_stack` — the same trap iter_drain hit).
pub fn emit_console_writeline(chunks: &mut [Chunk], current: usize, line: u32) {
    let log_idx = chunks[0].add_import("wasi:logging/logging", "log");
    let chunk = &mut chunks[current];
    // Direct ECMA String coercion. This still gives primitive/stringifiable
    // behavior, but user-defined .NET-style ToString overrides need to be
    // called explicitly by the frontend when desired.
    let v_local = alloc_local(chunk);
    let result_local = alloc_local(chunk);

    // Stash v.
    chunk.emit_op_u16(Op::LOCAL_SET, v_local, line);
    chunk.emit_op(Op::DROP, line);

    let exit_block = chunk.emit_block(line);

    // Bool branch
    let not_bool = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, v_local, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_string_const("boolean", line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, v_local, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_string_const("True", line);
    chunk.emit_else(line);
    chunk.emit_string_const("False", line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_br(1, line); // exit
    chunk.emit_end(line);
    chunk.patch_block(not_bool);

    // Null branch — `Console.WriteLine((string)null)` prints "".
    let not_null = chunk.emit_block(line);
    chunk.emit_op_u16(Op::LOCAL_GET, v_local, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);
    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_br(1, line);
    chunk.emit_end(line);
    chunk.patch_block(not_null);

    // Default: direct ECMA String(v) coercion.
    // User-defined `ToString` overrides on .NET-shape classes are NOT
    // picked up here yet — that requires routing through
    // `ecma:value.invokeMethod` which is itself a method-dispatch
    // problem (the receiver might not be a class instance). For now,
    // tests that need `Console.WriteLine(p)` to call `p.ToString()`
    // fail with "[object]" — call `Console.WriteLine(p.ToString())`
    // explicitly to get the override.
    chunk.emit_op_u16(Op::LOCAL_GET, v_local, line);
    crate::emitter::strings::emit_to_string(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_local, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line);
    chunk.patch_block(exit_block);

    // Single log call with the staged string. Push null after so the
    // call site (which DROPs print results uniformly) sees a value.
    chunk.emit_op_u16(Op::LOCAL_GET, result_local, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, log_idx, line);
    chunk.emit(1, line);
    chunk.emit_op(Op::NULL, line);
}

/// `Console.ReadLine()` — wasi:cli/stdin.get-stdin → [method]input-stream.blocking-read.
/// Stack: [] → [string]
pub fn emit_console_readline(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::emitter::io::emit_input(&mut chunks[current], line);
}

/// `Console.Error.WriteLine(v)` / `Console.Error.Write(v)` — log at error level.
/// Prepends "error" level arg before calling wasi:logging/logging.log.
/// Stack: [v] → [null]
pub fn emit_console_error(chunks: &mut [Chunk], current: usize, line: u32) {
    let log_idx = chunks[0].add_import("wasi:logging/logging", "log");
    let chunk = &mut chunks[current];
    // Stack currently has [v]. We need to call log(level, v) = 2 args.
    // Push level UNDER v: stash v, push level, restore v.
    let v_local = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, v_local, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_string_const("error", line);
    chunk.emit_op_u16(Op::LOCAL_GET, v_local, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, log_idx, line);
    chunk.emit(2, line);
    chunk.emit_op(Op::NULL, line);
}

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let s = chunk.local_count;
    chunk.local_count = s + 1;
    s
}
