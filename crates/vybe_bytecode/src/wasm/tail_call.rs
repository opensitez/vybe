//! # tail-call proposal
//!
//! Spec: <https://github.com/WebAssembly/tail-call>. Adds
//! `return_call` (0x12), `return_call_indirect` (0x13), and
//! `return_call_ref` (`0xFC 0x15`) — tail-call variants that reuse the
//! caller's frame, enabling unbounded mutual recursion without stack
//! growth.
//!
//! ## Status in Vybe
//!
//! | Op                     | Status | Notes |
//! |------------------------|--------|-------|
//! | `return_call`          | ✅ | `RETURN_CALL` opcode — VM reuses the current frame |
//! | `return_call_indirect` | ✅ | `RETURN_CALL_INDIRECT` with funcref table lookup |
//! | `return_call_ref`      | ✅ | `RETURN_CALL_REF` for closure tail-calls |
//!
//! ## VM semantics
//!
//! On a tail call, `vm.rs` pops the current frame's locals, pushes the
//! new arguments, and replaces the frame's chunk index + local count —
//! no new frame is allocated. This is essential for functional-style
//! recursion in all our compiled languages (VB → JS all compile tail
//! positions to `RETURN_CALL`).

use crate::Chunk;

pub const IMPORTS: &[(&str, &str)] = &[];
pub const GLOBAL_IMPORTS: &[(&str, &str)] = &[];
pub fn custom_sections(_chunks: &[Chunk]) -> Vec<(&'static str, Vec<u8>)> { Vec::new() }
