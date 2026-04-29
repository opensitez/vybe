//! .NET `System.Diagnostics.Process` / `ProcessStartInfo` adapter.
//!
//! Lowers .NET process-management APIs to `node:child_process.*`
//! (the spec-aligned Node-shaped child process surface). Multi-arity
//! constructors use the threaded `argc` to branch.
//!
//! No `vybe:types` involvement — pure compile-time adapter, runtime
//! work happens in `node:child_process.spawnSync` etc.

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use std::sync::Arc;

const FILENAME_KEY: &str = "filename";
const ARGUMENTS_KEY: &str = "arguments";
const TYPE_KEY: &str = "__type";

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    let idx = chunk.add_constant(val);
    chunk.emit_op_u16(Op::CONST, idx, line);
}

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    let slot = chunk.local_count;
    chunk.local_count = slot + 1;
    slot
}

/// `new ProcessStartInfo()` / `new ProcessStartInfo(cmd)` /
/// `new ProcessStartInfo(cmd, args)` — multi-arity ctor that builds
/// an `Object {filename, arguments}` plain record.
///
/// Stack on entry varies by `argc`:
///   argc=0 : `[]`            → `{filename: "", arguments: ""}`
///   argc=1 : `[cmd]`         → `{filename: cmd, arguments: ""}`
///   argc≥2 : `[cmd, args]`   → `{filename: cmd, arguments: args}`
pub fn emit_process_start_info_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let filename_key = chunk.add_constant(Value::String(Arc::from(FILENAME_KEY)));
    let args_key = chunk.add_constant(Value::String(Arc::from(ARGUMENTS_KEY)));

    // Stash whatever args we got into scratch slots so we can reorder
    // for STRUCT_SET (which pops [obj, val]).
    let cmd_slot = reserve_slot(chunk);
    let args_slot = reserve_slot(chunk);

    match argc {
        0 => {
            push_const(chunk, Value::String(Arc::from("")), line);
            chunk.emit_op_u16(Op::LOCAL_SET, cmd_slot, line); chunk.emit_op(Op::DROP, line);
            push_const(chunk, Value::String(Arc::from("")), line);
            chunk.emit_op_u16(Op::LOCAL_SET, args_slot, line); chunk.emit_op(Op::DROP, line);
        }
        1 => {
            // Stack: [cmd]
            chunk.emit_op_u16(Op::LOCAL_SET, cmd_slot, line); chunk.emit_op(Op::DROP, line);
            push_const(chunk, Value::String(Arc::from("")), line);
            chunk.emit_op_u16(Op::LOCAL_SET, args_slot, line); chunk.emit_op(Op::DROP, line);
        }
        _ => {
            // Stack: [cmd, args, ...] — defensive drop of any extras.
            for _ in 2..argc { chunk.emit_op(Op::DROP, line); }
            chunk.emit_op_u16(Op::LOCAL_SET, args_slot, line); chunk.emit_op(Op::DROP, line);
            chunk.emit_op_u16(Op::LOCAL_SET, cmd_slot, line); chunk.emit_op(Op::DROP, line);
        }
    }

    // Build the record: {filename: cmd_slot, arguments: args_slot}
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cmd_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, filename_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, args_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, args_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// `new Process()` — empty Process object. Most users actually call
/// `Process.Start(...)` which runs the process and returns a populated
/// Process; the bare `new Process()` is a placeholder. Stack: `[]` →
/// `[process]`.
pub fn emit_process_new(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op(Op::DUP, line);
    push_const(chunk, Value::String(Arc::from("Process")), line);
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// `Process.Start(cmd)` static method. The arg may be a string
/// filename or a `ProcessStartInfo` object; we read the `filename`
/// field if it exists, else use the arg directly. Lowers to
/// `node:child_process.spawnSync(filename, [])`.
///
/// Stack on entry: `[arg]` (string or ProcessStartInfo)
/// Stack on exit:  `[process_info]` from spawnSync
pub fn emit_process_start(chunks: &mut [Chunk], current: usize, line: u32) {
    let spawn_idx = chunks[0].add_import("node:child_process", "spawnSync");
    let chunk = &mut chunks[current];
    let filename_key = chunk.add_constant(Value::String(Arc::from(FILENAME_KEY)));
    let arg_slot = reserve_slot(chunk);

    // Stash the arg
    chunk.emit_op_u16(Op::LOCAL_SET, arg_slot, line); chunk.emit_op(Op::DROP, line);

    // Push filename (arg.filename if Object; else arg as-is). The VM's
    // `STRUCT_GET` returns null when the field is missing on a String
    // value (per its dispatch contract), so we DUP the arg, attempt
    // STRUCT_GET, and fall back to the arg itself if null.
    chunk.emit_op_u16(Op::LOCAL_GET, arg_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, filename_key, line);
    // If null, replace with the arg itself.
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let skip_fallback = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arg_slot, line);
    chunk.patch_jump(skip_fallback);
    // Stack: [filename_str]

    // Push empty array as args
    crate::emitter::collections::emit_array_new(chunks, current, 0, line);
    // Stack: [filename_str, []]

    // Call spawnSync(filename, [])
    chunks[current].emit_op_u16(Op::CALL_IMPORT, spawn_idx, line);
    chunks[current].emit(2, line);
}

/// `Process.GetCurrentProcess()` — return a Process-shaped object
/// populated with the host process's PID. Lowers to a struct-build
/// reading `node:process.pid`.
///
/// Stack on entry: `[]` ; Stack on exit: `[process_info]`
pub fn emit_process_get_current(chunks: &mut [Chunk], current: usize, line: u32) {
    let pid_idx = chunks[0].add_import("node:process", "pid");
    let chunk = &mut chunks[current];
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let pid_key = chunk.add_constant(Value::String(Arc::from("pid")));

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op(Op::DUP, line);
    push_const(chunk, Value::String(Arc::from("Process")), line);
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, pid_idx, line);
    chunk.emit(0, line);
    chunk.emit_op_u16(Op::STRUCT_SET, pid_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// `process.WaitForExit()` — `node:child_process.spawnSync` is already
/// blocking, so this is effectively a no-op for our model. Returns
/// the process untouched.
///
/// Stack on entry: `[process]` ; Stack on exit: `[process]`
pub fn emit_process_wait_for_exit(_chunks: &mut [Chunk], _current: usize, _line: u32) {
    // No-op: process already on stack from caller.
}
