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
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

// ── Atomic memory operations (WASM Threads spec) ────────────────────────
//
// Every atomic op carries a MEMARG — the VM's handlers unconditionally
// parse one (`pop_atomic_addr`), so the old memarg-less emissions here
// misparsed the following instruction as align/offset. The memarg is
// grid-padded to 4 bytes with non-minimal LEBs AND spec-aligned: the
// threads spec requires the natural alignment (2 for the 32-bit class),
// which a real validator enforces.

/// Emit the 32-bit-class atomic memarg: align 2, offset 0, grid-padded.
fn atomic_memarg32(chunk: &mut Chunk, line: u32) {
    chunk.emit(0x82, line);
    chunk.emit(0x00, line);
    chunk.emit(0x80, line);
    chunk.emit(0x00, line);
}

/// Emit atomic load: read i32 from shared memory at address.
/// Stack before: [addr]  Stack after: [i32_value]
pub fn emit_atomic_load(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_ATOMIC_LOAD, line);
    atomic_memarg32(chunk, line);
}

/// Emit atomic store: write i32 to shared memory at address.
/// Stack before: [addr, value]  Stack after: []
pub fn emit_atomic_store(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_ATOMIC_STORE, line);
    atomic_memarg32(chunk, line);
}

/// Emit atomic read-modify-write add: atomically add value to memory[addr].
/// Stack before: [addr, value]  Stack after: [old_value]
pub fn emit_atomic_add(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_ATOMIC_RMW_ADD, line);
    atomic_memarg32(chunk, line);
}

/// Emit atomic read-modify-write sub.
/// Stack before: [addr, value]  Stack after: [old_value]
pub fn emit_atomic_sub(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_ATOMIC_RMW_SUB, line);
    atomic_memarg32(chunk, line);
}

/// Emit atomic exchange: atomically swap memory[addr] with value.
/// Stack before: [addr, value]  Stack after: [old_value]
pub fn emit_atomic_xchg(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_ATOMIC_RMW_XCHG, line);
    atomic_memarg32(chunk, line);
}

/// Emit atomic compare-and-swap: if memory[addr] == expected, set to replacement.
/// Stack before: [addr, expected, replacement]  Stack after: [old_value]
pub fn emit_atomic_cmpxchg(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_ATOMIC_RMW_CMPXCHG, line);
    atomic_memarg32(chunk, line);
}

/// Emit atomic fence: memory barrier. Spec: one u8 immediate, must be 0.
/// Stack: unchanged
pub fn emit_atomic_fence(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::ATOMIC_FENCE, line);
    chunk.emit(0x00, line);
}

/// Emit atomic wait: block thread until memory[addr] != expected or timeout.
/// Stack before: [addr, expected, timeout_ns]  Stack after: [0=ok, 1=not_equal, 2=timed_out]
pub fn emit_atomic_wait(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::MEMORY_ATOMIC_WAIT32, line);
    atomic_memarg32(chunk, line);
}

/// Emit atomic notify: wake N threads waiting on memory[addr].
/// Stack before: [addr, count]  Stack after: [num_woken]
pub fn emit_atomic_notify(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::MEMORY_ATOMIC_NOTIFY, line);
    atomic_memarg32(chunk, line);
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
    atomic_memarg32(chunk, line);
    // If old value was 0, we acquired the lock
    core_wasm::i32_const(chunk, line, 0);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_br_if(1, line);
    // Not acquired — wait and retry
    chunk.emit_op_u16(Op::LOCAL_GET, addr_slot, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_i64_const(-1, line);
    chunk.emit_op(Op::MEMORY_ATOMIC_WAIT32, line);
    atomic_memarg32(chunk, line);
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
    atomic_memarg32(chunk, line);
    // Notify one waiter
    chunk.emit_op_u16(Op::LOCAL_GET, addr_slot, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::MEMORY_ATOMIC_NOTIFY, line);
    atomic_memarg32(chunk, line);
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

/// Emit Task.Run(fn) — spawn an OS thread via the `wasi:threads` import.
/// Stack before: [func_ref]  Stack after: [task_object]
pub fn emit_task_run(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_wasi_spawn(chunks, current, line, false);
}

/// Emit New Thread(fn) — spawn OS thread, return Task handle.
/// Stack before: [func_ref]  Stack after: [task_object]
pub fn emit_thread_new(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_wasi_spawn(chunks, current, line, false);
}

/// The ONE spawn lowering — pure spec surface, NO thread opcode:
///
///   1. the start function goes into funcref TABLE 0 (`table.grow` with
///      the closure as init returns the old size = its index),
///   2. a 16-byte record `{fn_idx, status, user_arg}` is reserved in the
///      shared futex page (`__vybe_futex_alloc16`; the wasi-libc
///      `pthread_create` packing),
///   3. `call` the `wasi:threads`.`thread-spawn(start_arg) -> tid` IMPORT
///      (the VM is the embedder-side implementation, as wasmtime is),
///   4. the Task object is ordinary bytecode (`__vybe_task_new`).
///
/// Join is `__vybe_task_wait` — a futex wait on the record's status word,
/// wasi-threads' sanctioned user-code join.
///
/// Stack before: `[func_ref]`, or `[user_arg, func_ref]` when
/// `has_user_arg` (the arg lands in the record's third word and reaches
/// the start function as its slot-0 parameter).
fn emit_wasi_spawn(chunks: &mut [Chunk], current: usize, line: u32, has_user_arg: bool) {
    let scratch = chunks[current].alloc_scratch(3);
    let (idx_slot, base_slot, arg_slot) = (scratch, scratch + 1, scratch + 2);
    let spawn_idx = chunks[current].add_import("wasi:threads", "thread-spawn");
    let chunk = &mut chunks[current];
    // [.., func_ref] → table.grow(t0, func_ref, 1) → old size = fn index
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::TABLE_GROW, line);
    chunk.emit(0u8, line); // table 0
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    if has_user_arg {
        // [user_arg] → the record's third word (the atomic store's i32
        // coercion handles f64-shaped args like Task.Delay's ms)
        chunk.emit_op_u16(Op::LOCAL_SET, arg_slot, line);
    } else {
        chunk.emit_i32_const(0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, arg_slot, line);
    }
    crate::primitives::bundle::emit_call_push_func(chunk, "__vybe_futex_alloc16", line);
    crate::primitives::bundle::emit_call_invoke(chunk, 0, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, base_slot, line);
    // record[0] = fn index
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    emit_atomic_store_aligned(chunk, line);
    // record[+8] = user_arg
    chunk.emit_op_u16(Op::LOCAL_GET, base_slot, line);
    chunk.emit_i32_const(8, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arg_slot, line);
    emit_atomic_store_aligned(chunk, line);
    // tid = wasi:threads.thread-spawn(base)
    chunk.emit_op_u16(Op::LOCAL_GET, base_slot, line);
    chunk.emit_call(spawn_idx, 1u8, line);
    // task = __vybe_task_new(tid, base) — push_func must come FIRST, so
    // stash tid (idx_slot's job is done).
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    crate::primitives::bundle::emit_call_push_func(chunk, "__vybe_task_new", line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, base_slot, line);
    crate::primitives::bundle::emit_call_invoke(chunk, 2, line);
}

fn emit_atomic_store_aligned(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::I32_ATOMIC_STORE, line);
    chunk.emit(0x82, line); // natural align 2, grid-padded
    chunk.emit(0x00, line);
    chunk.emit(0x80, line); // offset 0
    chunk.emit(0x00, line);
}

/// Emit thread.Start() — for pre-created threads.
/// In wasi-threads, thread_spawn both creates AND starts, so Start is a no-op.
/// Stack before: [thread_id]  Stack after: [thread_id] (passthrough)
pub fn emit_thread_start(_chunk: &mut Chunk, _line: u32) {
    // thread_spawn already started the thread. Nothing to do.
}

/// Emit thread.Join() — wait for thread to complete: `__vybe_task_wait`,
/// a spec-bytecode futex wait on the task's status word (wasi-threads'
/// sanctioned user-code join; the proposal has no join primitive).
/// Stack before: [task]  Stack after: [status: i32]
pub fn emit_thread_join(chunk: &mut Chunk, line: u32) {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    crate::primitives::bundle::emit_call_push_func(chunk, "__vybe_task_wait", line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    crate::primitives::bundle::emit_call_invoke(chunk, 1, line);
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
    chunk.emit_call(sub_dur_idx, 1u8, line);
    // [method]pollable.block(pollable) → ()
    chunk.emit_call(block_idx, 1u8, line);
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
    worker.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    worker.emit_op(Op::RETURN, line);
    worker.local_count = 1;
    chunks.push(worker);
    let worker_idx = chunks.len() - 1;

    // Stack already has [ms] from the caller. Push the worker func_ref;
    // the spawn lowering packs ms into the record's user_arg word and the
    // worker reads it as its slot-0 parameter.
    chunks[current].emit_op_u16(Op::REF_FUNC, worker_idx as u16, line);
    chunks[current].emit(0, line); // 0 upvalues — no closure capture needed
    emit_wasi_spawn(chunks, current, line, true);
}

// Backward compat
/// The two-operand spawn form: `[start_arg, func_ref]` — start_arg lands
/// in the record's user_arg word (i32; objects go via table 0 first).
pub fn emit_thread_spawn(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_wasi_spawn(chunks, current, line, true);
}
