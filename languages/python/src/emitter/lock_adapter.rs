//! Python `threading.Lock` / `RLock` / `Semaphore`, on the SHARED atomics.
//!
//! `primitives::threading::emit_lock_acquire` / `emit_lock_release` are the
//! WASM-threads spinlock every language compiles to — `atomic_rmw_xchg` to
//! take, `atomic_store` + `atomic_notify` to drop, `memory.atomic.wait32` to
//! block. C# `lock {}`, VB `SyncLock` and JS `Atomics` all land on the same
//! opcodes; python's `Lock` had a Python-level `self.locked = True` instead,
//! which is not a lock at all once a second thread exists.
//!
//! The monitor word is allocated lazily per object by
//! `emit_object_monitor_addr` and cached on it as `__vybe_monitor_addr`, so a
//! `Lock` needs no storage of its own and two `Lock`s never share a word.

use vybe_compiler::primitives::threading;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// Stash the receiver, resolve its monitor address into a local, and answer
/// that slot. Stack: `[obj] -> []`.
fn monitor_slot(chunk: &mut Chunk, argc: u8, line: u32) -> u16 {
    // Drop everything but the receiver — `acquire(blocking, timeout)` carries
    // arguments the spinlock has no use for.
    let recv = chunk.alloc_scratch(1);
    for _ in 1..argc.max(1) {
        chunk.emit_op(Op::DROP, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, recv, line);
    chunk.emit_op_u16(Op::LOCAL_GET, recv, line);
    threading::emit_object_monitor_addr(chunk, line);
    let addr = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, addr, line);
    addr
}

/// `lock.acquire(...)` → the shared spinlock. Answers `True`, as CPython does
/// for a blocking acquire that succeeded.
pub fn emit_lock_acquire(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let addr = monitor_slot(chunk, argc, line);
    threading::emit_lock_acquire(chunk, addr, line);
    chunk.emit_bool_const(true, line);
}

/// `lock.release()` → `atomic_store(0)` + `atomic_notify`.
pub fn emit_lock_release(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let addr = monitor_slot(chunk, argc, line);
    threading::emit_lock_release(chunk, addr, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}
