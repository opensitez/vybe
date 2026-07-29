//! Threading helpers — shared bytecode patterns for WASM Threads.
//!
//! WASM Threads spec: shared linear memory + atomic operations + wait/notify.
//! Thread spawning is host-provided (not a WASM opcode).
//!
//! All languages compile to the same atomic opcodes:
//! - Python `threading.Lock` → atomic_rmw_xchg (spinlock)
//! - JS `Atomics.load()` → i32_atomic_load
//! - Dart `Isolate` → thread_spawn host call
//! - C# `lock {}` → atomic_rmw_xchg + atomic_store
//! - VB `SyncLock` → same as C#
//!
//! The VM implements these with real Rust atomics on shared memory.

use crate::primitives::functions::create_function_chunk;
use crate::primitives::instructions::core_wasm;
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

// ── Atomic memory operations (WASM Threads spec) ────────────────────────

/// Emit atomic load: read i32 from shared memory at address.
/// Stack before: [addr]  Stack after: [i32_value]
pub fn emit_atomic_load(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_ATOMIC_LOAD, line);
}

/// Emit atomic store: write i32 to shared memory at address.
/// Stack before: [addr, value]  Stack after: []
pub fn emit_atomic_store(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_ATOMIC_STORE, line);
}

/// Emit atomic read-modify-write add: atomically add value to memory[addr].
/// Stack before: [addr, value]  Stack after: [old_value]
pub fn emit_atomic_add(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_ATOMIC_RMW_ADD, line);
}

/// Emit atomic read-modify-write sub.
/// Stack before: [addr, value]  Stack after: [old_value]
pub fn emit_atomic_sub(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_ATOMIC_RMW_SUB, line);
}

/// Emit atomic exchange: atomically swap memory[addr] with value.
/// Stack before: [addr, value]  Stack after: [old_value]
pub fn emit_atomic_xchg(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_ATOMIC_RMW_XCHG, line);
}

/// Emit atomic compare-and-swap: if memory[addr] == expected, set to replacement.
/// Stack before: [addr, expected, replacement]  Stack after: [old_value]
pub fn emit_atomic_cmpxchg(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_ATOMIC_RMW_CMPXCHG, line);
}

/// Emit atomic fence: memory barrier.
/// Stack: unchanged
pub fn emit_atomic_fence(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::ATOMIC_FENCE, line);
}

/// Emit atomic wait: block thread until memory[addr] != expected or timeout.
/// Stack before: [addr, expected, timeout_ns]  Stack after: [0=ok, 1=not_equal, 2=timed_out]
pub fn emit_atomic_wait(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::MEMORY_ATOMIC_WAIT32, line);
}

/// Emit atomic notify: wake N threads waiting on memory[addr].
/// Stack before: [addr, count]  Stack after: [num_woken]
pub fn emit_atomic_notify(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::MEMORY_ATOMIC_NOTIFY, line);
}

// ── Lock pattern (spinlock via atomics) ─────────────────────────────────
//
// Standard WASM spinlock using atomic_rmw_xchg:
//   lock:   while atomic_xchg(addr, 1) != 0 { atomic_wait(addr, 1, -1) }
//   unlock: atomic_store(addr, 0); atomic_notify(addr, 1)
//
// All languages use this same pattern: Python Lock, JS Mutex, C# lock{}, VB SyncLock

/// Emit lock acquisition on a mutex at memory address.
/// `addr_slot`: local slot containing the memory address of the lock word.
/// Stack: unchanged
pub fn emit_lock_acquire(chunk: &mut Chunk, addr_slot: u16, line: u32) {
    // Spin loop: block { loop { if atomic_xchg(addr, 1) == 0 { br 1 } wait; br 0 } }
    let outer = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, addr_slot, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ATOMIC_RMW_XCHG, line);
    // If old value was 0, we acquired the lock
    core_wasm::i32_const(chunk, line, 0);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_br_if(1, line);
    // Not acquired — wait and retry
    chunk.emit_op_u16(Op::LOCAL_GET, addr_slot, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_i64_const(-1, line);
    chunk.emit_op(Op::MEMORY_ATOMIC_WAIT32, line);
    chunk.emit_op(Op::DROP, line); // drop wait result
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(outer);
}

/// Emit lock release on a mutex at memory address.
/// `addr_slot`: local slot containing the memory address of the lock word.
/// Stack: unchanged
pub fn emit_lock_release(chunk: &mut Chunk, addr_slot: u16, line: u32) {
    // Store 0 (unlocked)
    chunk.emit_op_u16(Op::LOCAL_GET, addr_slot, line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op(Op::I32_ATOMIC_STORE, line);
    // Notify one waiter
    chunk.emit_op_u16(Op::LOCAL_GET, addr_slot, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::MEMORY_ATOMIC_NOTIFY, line);
    chunk.emit_op(Op::DROP, line); // drop notify count
}

// ── Thread/Task (WASM stack switching) ──────────────────────────────────
//
// WASM stack switching primitives:
//   cont_new  — create a continuation (fiber) from a function reference
//   resume    — start/resume the continuation, passing a value
//   suspend   — pause the current continuation
//
// Two patterns:
//   1. Task.Run(fn) — create AND run immediately: cont_new + resume
//   2. New Thread(fn) — create only: cont_new. Start later with resume.

/// Emit Task.Run(fn) — spawn OS thread and return Task handle.
/// Stack before: [func_ref]  Stack after: [task_object]
///
/// THREAD_SPAWN takes `[start_arg, func_ref]` per wasi-threads. Task.Run
/// has no start arg, so we slot `null` underneath the func_ref via a
/// scratch local. The spawned function lands `null` in slot 0 and
/// ignores it.
pub fn emit_task_run(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_thread_spawn_no_arg(chunks, current, line);
}

/// Emit New Thread(fn) — spawn OS thread, return Task handle.
/// Stack before: [func_ref]  Stack after: [task_object]
pub fn emit_thread_new(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_thread_spawn_no_arg(chunks, current, line);
}

/// Lower a `THREAD_SPAWN` whose caller supplied a func_ref on the stack
/// but no start_arg. We need `[null, func_ref]` for THREAD_SPAWN; given
/// `[func_ref]`, the cheapest reorder is via a scratch local.
fn emit_thread_spawn_no_arg(chunks: &mut [Chunk], current: usize, line: u32) {
    let slot = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    // [func_ref] → stash to scratch slot, drop from stack
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    // Push null (start_arg), then func_ref back from scratch
    chunk.emit_op(Op::NULL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op(Op::THREAD_SPAWN, line);
}

/// Emit thread.Start() — for pre-created threads.
/// In wasi-threads, thread_spawn both creates AND starts, so Start is a no-op.
/// Stack before: [thread_id]  Stack after: [thread_id] (passthrough)
pub fn emit_thread_start(_chunk: &mut Chunk, _line: u32) {
    // thread_spawn already started the thread. Nothing to do.
}

/// Emit thread.Join() — wait for thread to complete.
/// Stack before: [thread_id]  Stack after: [status: i32]
pub fn emit_thread_join(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::THREAD_JOIN, line);
}

/// Emit suspend — yield from current continuation (stack switching).
/// Stack before: [value]  Stack after: (suspended — caller gets value)
pub fn emit_suspend(chunk: &mut Chunk, line: u32) {
    crate::primitives::generators::emit_suspend(chunk, line);
}

/// Emit Thread.Sleep(ms) — WASI pollable blocking sleep.
/// Stack before: [ms_value]  Stack after: []
pub fn emit_sleep(chunk: &mut Chunk, sub_dur_idx: u16, block_idx: u16, line: u32) {
    emit_thread_sleep(chunk, sub_dur_idx, block_idx, line);
}

/// Blocking sleep via WASI pollables — `subscribe-duration(ns)` then
/// `pollable.block`. Platform-neutral (pure WASI); every platform's
/// sleep-shaped surface (Thread.Sleep, usleep, …) routes here.
/// Stack before: [ms_value]  Stack after: []
pub fn emit_thread_sleep(chunk: &mut Chunk, sub_dur_idx: u16, block_idx: u16, line: u32) {
    // ms × 1_000_000 = nanoseconds
    chunk.emit_f64_const(1_000_000.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    // subscribe-duration(ns) → pollable
    chunk.emit_op_u16(Op::CALL_IMPORT, sub_dur_idx, line);
    chunk.emit(1u8, line);
    // [method]pollable.block(pollable) → ()
    chunk.emit_op_u16(Op::CALL_IMPORT, block_idx, line);
    chunk.emit(1u8, line);
}

/// Emit Task.Delay(ms) — spawn a worker that sleeps for `ms`, returning
/// the Task object that `Op::THREAD_SPAWN` constructs natively.
///
/// Pure WASM, zero host fns: `THREAD_SPAWN` matches the wasi-threads
/// `thread.spawn(start_arg)` shape, so we pass `ms` as the start arg and
/// the worker reads it from slot 0. The Task object's `iscompleted` /
/// `isalive` / `result` / `status` fields are populated by the VM's
/// THREAD_SPAWN handler when the worker fiber returns.
///
/// Stack before: [ms]  Stack after: [task_object]
pub fn emit_task_delay(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    // WASI pollable sleep imports — all imports flow through chunks[0].
    let sub_dur_idx =
        chunks[current].add_import("wasi:clocks/monotonic-clock", "subscribe-duration");
    let block_idx = chunks[current].add_import("wasi:io/poll", "[method]pollable.block");

    // Worker chunk: arity=1 (start_arg = ms), body sleeps via WASI pollable,
    // returns null. The Task.result field reflects this null on completion.
    let mut worker = create_function_chunk("__task_delay_worker", 1);
    worker.emit_op_u16(Op::LOCAL_GET, 0, line);
    emit_thread_sleep(&mut worker, sub_dur_idx, block_idx, line);
    worker.emit_op(Op::NULL, line);
    worker.emit_op(Op::RETURN, line);
    worker.local_count = 1;
    chunks.push(worker);
    let worker_idx = chunks.len() - 1;

    let chunk = &mut chunks[current];
    // Stack already has [ms] from the caller. Push the worker func_ref;
    // THREAD_SPAWN pops [ms, func_ref] and spawns the worker with ms as
    // its slot-0 arg.
    chunk.emit_op_u16(Op::REF_FUNC, worker_idx as u16, line);
    chunk.emit(0, line); // 0 upvalues — no closure capture needed
    chunk.emit_op(Op::THREAD_SPAWN, line);
}

// Backward compat
pub fn emit_thread_spawn(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_thread_new(chunks, current, line);
}
