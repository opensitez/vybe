//! Thread.Sleep adapter — WASI spec-compliant blocking sleep.
//!
//! Replaces the retired `vybe:clocks.sleep` host function. Thread.Sleep(ms)
//! compiles to the two-step WASI 0.2 pollable pattern:
//!   1. `wasi:clocks/monotonic-clock.subscribe-duration(ns)` → pollable
//!   2. `wasi:io/poll.[method]pollable.block(pollable)` → ()
//!
//! Both functions are WASI 0.2 spec and already registered in vybe_host.

use vybe_bytecode::{Chunk, opcode::Op};

// Thread.Sleep bytecode lives in vybe_emitter::threading::emit_thread_sleep
// (pure WASI, platform-neutral).
