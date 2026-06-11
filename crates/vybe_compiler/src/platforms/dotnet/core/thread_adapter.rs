//! Thread.Sleep adapter — WASI spec-compliant blocking sleep.
//!
//! Replaces the retired `vybe:clocks.sleep` host function. Thread.Sleep(ms)
//! compiles to the two-step WASI 0.2 pollable pattern:
//!   1. `wasi:clocks/monotonic-clock.subscribe-duration(ns)` → pollable
//!   2. `wasi:io/poll.[method]pollable.block(pollable)` → ()
//!
//! Both functions are WASI 0.2 spec and already registered in vybe_host.

use vybe_bytecode::{Chunk, Value, opcode::Op};

/// Emit Thread.Sleep(ms) — blocking WASI pollable sleep.
///
/// `sub_dur_idx` = import index for `wasi:clocks/monotonic-clock.subscribe-duration`
/// `block_idx`   = import index for `wasi:io/poll.[method]pollable.block`
///
/// Stack before: [ms: f64]  Stack after: []
pub fn emit_thread_sleep(chunk: &mut Chunk, sub_dur_idx: u16, block_idx: u16, line: u32) {
    // ms × 1_000_000 = nanoseconds
    let ns_mul = chunk.add_constant(Value::F64(1_000_000.0));
    chunk.emit_op_u16(Op::CONST, ns_mul, line);
    chunk.emit_op(Op::F64_MUL, line);
    // subscribe-duration(ns) → pollable
    chunk.emit_op_u16(Op::CALL_IMPORT, sub_dur_idx, line);
    chunk.emit(1u8, line);
    // [method]pollable.block(pollable) → ()
    chunk.emit_op_u16(Op::CALL_IMPORT, block_idx, line);
    chunk.emit(1u8, line);
}
