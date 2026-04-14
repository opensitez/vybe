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

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::Value;

// ── Atomic memory operations (WASM Threads spec) ────────────────────────

/// Emit atomic load: read i32 from shared memory at address.
/// Stack before: [addr]  Stack after: [i32_value]
pub fn emit_atomic_load(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::i32_atomic_load, line);
}

/// Emit atomic store: write i32 to shared memory at address.
/// Stack before: [addr, value]  Stack after: []
pub fn emit_atomic_store(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::i32_atomic_store, line);
}

/// Emit atomic read-modify-write add: atomically add value to memory[addr].
/// Stack before: [addr, value]  Stack after: [old_value]
pub fn emit_atomic_add(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::i32_atomic_rmw_add, line);
}

/// Emit atomic read-modify-write sub.
/// Stack before: [addr, value]  Stack after: [old_value]
pub fn emit_atomic_sub(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::i32_atomic_rmw_sub, line);
}

/// Emit atomic exchange: atomically swap memory[addr] with value.
/// Stack before: [addr, value]  Stack after: [old_value]
pub fn emit_atomic_xchg(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::i32_atomic_rmw_xchg, line);
}

/// Emit atomic compare-and-swap: if memory[addr] == expected, set to replacement.
/// Stack before: [addr, expected, replacement]  Stack after: [old_value]
pub fn emit_atomic_cmpxchg(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::i32_atomic_rmw_cmpxchg, line);
}

/// Emit atomic fence: memory barrier.
/// Stack: unchanged
pub fn emit_atomic_fence(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::atomic_fence, line);
}

/// Emit atomic wait: block thread until memory[addr] != expected or timeout.
/// Stack before: [addr, expected, timeout_ns]  Stack after: [0=ok, 1=not_equal, 2=timed_out]
pub fn emit_atomic_wait(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::memory_atomic_wait32, line);
}

/// Emit atomic notify: wake N threads waiting on memory[addr].
/// Stack before: [addr, count]  Stack after: [num_woken]
pub fn emit_atomic_notify(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::memory_atomic_notify, line);
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
    // Spin loop: while atomic_xchg(addr, 1) != 0 { wait }
    let loop_start = chunk.current_offset();
    chunk.emit_op_u16(Op::local_get, addr_slot, line);
    let one = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::r#const, one, line);
    chunk.emit_op(Op::i32_atomic_rmw_xchg, line);
    // If old value was 0, we acquired the lock
    chunk.emit_op(Op::i32_const_0, line);
    chunk.emit_op(Op::dyn_eq, line);
    let acquired = chunk.emit_jump(Op::br_if_true, line);
    // Not acquired — wait and retry
    chunk.emit_op_u16(Op::local_get, addr_slot, line);
    let one2 = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::r#const, one2, line);
    let neg1 = chunk.add_constant(Value::I64(-1));
    chunk.emit_op_u16(Op::r#const, neg1, line);
    chunk.emit_op(Op::memory_atomic_wait32, line);
    chunk.emit_op(Op::drop, line); // drop wait result
    chunk.emit_loop(loop_start, line);
    chunk.patch_jump(acquired);
}

/// Emit lock release on a mutex at memory address.
/// `addr_slot`: local slot containing the memory address of the lock word.
/// Stack: unchanged
pub fn emit_lock_release(chunk: &mut Chunk, addr_slot: u16, line: u32) {
    // Store 0 (unlocked)
    chunk.emit_op_u16(Op::local_get, addr_slot, line);
    chunk.emit_op(Op::i32_const_0, line);
    chunk.emit_op(Op::i32_atomic_store, line);
    // Notify one waiter
    chunk.emit_op_u16(Op::local_get, addr_slot, line);
    let one = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::r#const, one, line);
    chunk.emit_op(Op::memory_atomic_notify, line);
    chunk.emit_op(Op::drop, line); // drop notify count
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

/// Emit Task.Run(fn) — spawn OS thread and return thread handle.
/// Stack before: [func_ref]  Stack after: [thread_id: i32]
pub fn emit_task_run(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::thread_spawn, line);
}

/// Emit New Thread(fn) — spawn OS thread, return thread handle.
/// Stack before: [func_ref]  Stack after: [thread_id: i32]
pub fn emit_thread_new(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::thread_spawn, line);
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
    chunk.emit_op(Op::thread_join, line);
}

/// Emit suspend — yield from current continuation (stack switching).
/// Stack before: [value]  Stack after: (suspended — caller gets value)
pub fn emit_suspend(chunk: &mut Chunk, line: u32) {
    let tag = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::suspend, tag, line);
}

/// Emit Thread.Sleep(ms) — blocks the current thread for the given duration.
/// `sleep_import_idx` must be obtained from chunk 0 via `chunks[0].add_import("wasi:clocks", "sleep")`.
/// Stack before: [ms_value]  Stack after: []
pub fn emit_sleep(chunk: &mut Chunk, sleep_import_idx: u16, line: u32) {
    chunk.emit_op_u16(Op::call_import, sleep_import_idx, line);
    chunk.emit(1, line);
    chunk.emit_op(Op::drop, line);
}

// Backward compat
pub fn emit_thread_spawn(chunk: &mut Chunk, line: u32) {
    emit_thread_new(chunk, line);
}
