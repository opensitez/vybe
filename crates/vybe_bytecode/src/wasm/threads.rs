//! # threads proposal
//!
//! Spec: <https://github.com/WebAssembly/threads>. Adds shared linear
//! memory, atomic memory operations, and wait/notify primitives.
//! Prefix `0xFE`.
//!
//! ## Status in Vybe
//!
//! | Feature                    | Status | Notes |
//! |----------------------------|--------|-------|
//! | Shared memory              | ⚠  | memory flag emitted when `shared` requested; host must cooperate |
//! | `atomic.load` / `store`    | ✅ | i32, i64, i32_8u, i32_16u |
//! | `atomic.rmw.*`             | ✅ | add, sub, and, or, xor, xchg, cmpxchg |
//! | `memory.atomic.wait32/64`  | ✅ | VM no-ops unless host driver is attached |
//! | `memory.atomic.notify`     | ✅ | VM no-ops |
//! | `atomic.fence`             | ✅ | |
//!
//! All 0xFE-prefix opcodes are emitted by the generic pass-through in
//! `code.rs` — no special handling needed. Actual atomic semantics are
//! implemented VM-side in `vm.rs`.

use crate::Chunk;

pub const IMPORTS: &[(&str, &str)] = &[];
pub const GLOBAL_IMPORTS: &[(&str, &str)] = &[];
pub fn custom_sections(_chunks: &[Chunk]) -> Vec<(&'static str, Vec<u8>)> {
    Vec::new()
}
