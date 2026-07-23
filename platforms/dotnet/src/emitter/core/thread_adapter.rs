//! Thread.Sleep adapter — WASI spec-compliant blocking sleep.
//!
//! Replaces the retired `vybe:clocks.sleep` host function. Thread.Sleep(ms)
//! compiles to the two-step WASI 0.2 pollable pattern:
//!   1. `wasi:clocks/monotonic-clock.subscribe-duration(ns)` → pollable
//!   2. `wasi:io/poll.[method]pollable.block(pollable)` → ()
//!
//! Both functions are WASI 0.2 spec and already registered in vybe_host.

use vybe_bytecode::{Chunk, opcode::Op};
use vybe_emitter::collections;

// Thread.Sleep bytecode lives in vybe_emitter::threading::emit_thread_sleep
// (pure WASI, platform-neutral).

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(argc, line);
}

/// `task.Wait()` — async runs eagerly in this VM, so the task is already
/// complete when awaited; `.Wait()` is a no-op. Discard the receiver on the
/// stack and yield void. Stack: [task] → [null].
pub fn emit_task_wait(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// `Task.FromResult(v)` — a task that is already completed with `v`. Since async
/// runs eagerly, that is just `Promise.resolve(v)`. Stack: [value] → [task].
pub fn emit_task_from_result(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "ecma:promise", "resolve", 1, line);
}

/// Collapse the combinator's args into a single iterable of tasks: one array
/// argument (the `IEnumerable<Task>` overload) is used directly; otherwise the
/// N task args are packed into an array.
fn emit_pack_task_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 1 {
        let slot = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        call_import(chunks, current, "ecma:array", "isArray", 1, line);
        vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        collections::emit_array_new(chunks, current, 1, line);
        chunks[current].emit_end(line);
    } else {
        collections::emit_array_new(chunks, current, argc as u16, line);
    }
}

/// `Task.WhenAll(t1, …)` — completes with the array of every task's result.
/// `Promise.all` over the (eagerly-resolved) tasks. Stack: [t1 .. tN] → [task].
pub fn emit_task_when_all(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_pack_task_args(chunks, current, argc, line);
    call_import(chunks, current, "ecma:promise", "all", 1, line);
}

/// `Task.WhenAny(t1, …)` — completes with the first task to finish
/// (`Promise.race`). Stack: [t1 .. tN] → [task].
pub fn emit_task_when_any(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_pack_task_args(chunks, current, argc, line);
    call_import(chunks, current, "ecma:promise", "race", 1, line);
}

/// `antecedent.ContinueWith(fn)` — run `fn(antecedent)` once the antecedent has
/// completed and return its result as a new task. `fn` receives the completed
/// task, so its `t.Result` reads the antecedent's value (unwrap/join). async
/// runs eagerly, so no explicit wait is needed. Stack: [antecedent, fn] → [task].
pub fn emit_task_continue_with(chunks: &mut [Chunk], current: usize, line: u32) {
    let fn_slot = chunks[current].alloc_scratch(1);
    let ante_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ante_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ante_slot, line);
    vybe_emitter::functions::emit_call(&mut chunks[current], 1, line);
    call_import(chunks, current, "ecma:promise", "resolve", 1, line);
}
