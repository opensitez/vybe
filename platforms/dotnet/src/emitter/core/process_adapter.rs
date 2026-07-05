//! .NET `System.Diagnostics.Process` / `ProcessStartInfo` adapter.
//!
//! Lowers .NET process-management APIs to `node:child_process.*`
//! (the spec-aligned Node-shaped child process surface). Multi-arity
//! constructors use the threaded `argc` to branch.
//!
//! No `vybe:types` involvement — pure compile-time adapter, runtime
//! work happens in `node:child_process.spawnSync` etc.

use vybe_emitter::classes::emit_bind_method;
use vybe_emitter::functions::create_function_chunk;
use vybe_emitter::instructions::{core_wasm, host};
use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

const FILENAME_KEY: &str = "filename";
const ARGUMENTS_KEY: &str = "arguments";
const TYPE_KEY: &str = "__type";

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn bind_process_wait_for_exit(chunks: &mut Vec<Chunk>, current: usize, this_slot: u16, line: u32) {
    let mut method = create_function_chunk("__process_waitforexit", 1);
    method.emit_op(Op::DROP, line);
    method.emit_op(Op::NULL, line);
    method.emit_op(Op::RETURN, line);
    method.local_count = 1;
    chunks.push(method);
    let method_idx = chunks.len() - 1;
    emit_bind_method(
        &mut chunks[current],
        this_slot,
        "waitforexit",
        method_idx,
        line,
    );
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
            chunk.emit_op_u16(Op::LOCAL_SET, cmd_slot, line);
            push_const(chunk, Value::String(Arc::from("")), line);
            chunk.emit_op_u16(Op::LOCAL_SET, args_slot, line);
        }
        1 => {
            // Stack: [cmd]
            chunk.emit_op_u16(Op::LOCAL_SET, cmd_slot, line);
            push_const(chunk, Value::String(Arc::from("")), line);
            chunk.emit_op_u16(Op::LOCAL_SET, args_slot, line);
        }
        _ => {
            // Stack: [cmd, args, ...] — defensive drop of any extras.
            for _ in 2..argc {
                chunk.emit_op(Op::DROP, line);
            }
            chunk.emit_op_u16(Op::LOCAL_SET, args_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, cmd_slot, line);
        }
    }

    // Build the record: {filename: cmd_slot, arguments: args_slot}
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cmd_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, filename_key, line);
    chunk.emit_op(Op::DROP, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, args_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, args_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// `new Process()` — empty Process object. Most users actually call
/// `Process.Start(...)` which runs the process and returns a populated
/// Process; the bare `new Process()` is a placeholder. Stack: `[]` →
/// `[process]`.
pub fn emit_process_new(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let process_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, process_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, process_slot, line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Process")), line);
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);
    bind_process_wait_for_exit(chunks, current, process_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, process_slot, line);
}

/// `Process.Start(cmd)` static method. The arg may be a string
/// filename or a `ProcessStartInfo` object; we read the `filename`
/// field if it exists, else use the arg directly. Lowers to
/// `node:child_process.spawnSync(filename, [])`, then wraps the
/// raw host result in a .NET-shaped Process struct.
///
/// Stack on entry: `[arg]` (string or ProcessStartInfo)
/// Stack on exit:  `[Process { HasExited, ExitCode, __type }]`
pub fn emit_process_start(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let spawn_idx = chunks[0].add_import("node:child_process", "spawnSync");
    let chunk = &mut chunks[current];
    let filename_key = chunk.add_constant(Value::String(Arc::from(FILENAME_KEY)));
    let args_key = chunk.add_constant(Value::String(Arc::from(ARGUMENTS_KEY)));
    let arg_slot = reserve_slot(chunk);
    let args_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);
    let process_slot = reserve_slot(chunk);

    // Stash the arg
    chunk.emit_op_u16(Op::LOCAL_SET, arg_slot, line);

    // Resolve filename: arg.filename if Object; else arg itself.
    chunk.emit_op_u16(Op::LOCAL_GET, arg_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, filename_key, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let filename_fallback = chunk.emit_block(line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arg_slot, line);
    chunk.emit_end(line);
    chunk.patch_block(filename_fallback);
    // Stack: [filename_str]

    // Resolve arguments: arg.arguments if Object; else "".
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, arg_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, args_key, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let args_fallback = chunk.emit_block(line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::DROP, line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_end(line);
    chunk.patch_block(args_fallback);
    chunk.emit_op_u16(Op::LOCAL_SET, args_slot, line);

    // ProcessStartInfo.Arguments is a raw string. For the current
    // adapter surface, split on spaces into the argv array that the
    // Node bridge expects. Keep [] for the empty-string default.
    chunk.emit_op_u16(Op::LOCAL_GET, args_slot, line);
    core_wasm::dup(chunk, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    push_const(chunk, Value::I32(0), line);
    vybe_emitter::ops::emit_dyn_eq(chunk, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op(Op::DROP, line);
    vybe_emitter::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_else(line);
    let chunk = &mut chunks[current];
    push_const(chunk, Value::String(Arc::from(" ")), line);
    host::emit(chunk, "ecma:string", "split", 2, line);
    chunk.emit_end(line);
    // Stack: [filename_str, argv_array]

    // Call spawnSync(filename, []) → raw host result on stack
    chunks[current].emit_op_u16(Op::CALL_IMPORT, spawn_idx, line);
    chunks[current].emit(2, line);
    // Stack: [raw_result]

    // Stash the raw result
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    // Build a .NET-shaped Process struct.
    // Fields: __type="Process", HasExited=true, ExitCode=raw.status (or 0).
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let he_key = chunk.add_constant(Value::String(Arc::from("hasexited")));
    let ec_key = chunk.add_constant(Value::String(Arc::from("exitcode")));
    let status_key = chunk.add_constant(Value::String(Arc::from("status")));

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);

    // __type = "Process"
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Process")), line);
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);

    // HasExited = true (spawnSync is synchronous; process is always done)
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_op_u16(Op::STRUCT_SET, he_key, line);
    chunk.emit_op(Op::DROP, line);

    // ExitCode = raw_result.status ?? 0
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, status_key, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let status_fallback = chunk.emit_block(line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::DROP, line);
    push_const(chunk, Value::I32(0), line);
    chunk.emit_end(line);
    chunk.patch_block(status_fallback);
    chunk.emit_op_u16(Op::STRUCT_SET, ec_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, process_slot, line);
    bind_process_wait_for_exit(chunks, current, process_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, process_slot, line);
    // Stack: [Process struct]
}

/// `Process.GetCurrentProcess()` — return a Process-shaped object
/// populated with the host process's PID. Lowers to a struct-build
/// reading `node:process.pid`.
///
/// Stack on entry: `[]` ; Stack on exit: `[process_info]`
pub fn emit_process_get_current(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let pid_idx = chunks[0].add_import("node:process", "pid");
    let chunk = &mut chunks[current];
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let pid_key = chunk.add_constant(Value::String(Arc::from("pid")));
    let process_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, process_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, process_slot, line);
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Process")), line);
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, pid_idx, line);
    chunk.emit(0, line);
    chunk.emit_op_u16(Op::STRUCT_SET, pid_key, line);
    chunk.emit_op(Op::DROP, line);
    bind_process_wait_for_exit(chunks, current, process_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, process_slot, line);
}

/// `process.WaitForExit()` — `node:child_process.spawnSync` is already
/// blocking, so the process is done by the time Start() returns.
/// Drop the receiver and return null.
///
/// Stack on entry: `[process]` ; Stack on exit: `[null]`
pub fn emit_process_wait_for_exit(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}
