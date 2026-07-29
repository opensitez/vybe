//! # multi-value proposal
//!
//! Spec: `proposals/multi-value/`. Allows
//! functions, blocks, and control instructions to produce multiple
//! results rather than at-most-one. Merged into core WASM in 2020.
//!
//! ## Status in Vybe
//!
//! | Feature                         | Status |
//! |---------------------------------|--------|
//! | Multi-result function types     | ✅ `chunk.result_arity` drives the type-section signature |
//! | Multi-result function returns   | ✅ `RETURN` pops N values and surfaces them on the caller stack |
//! | Multi-result block / loop types | ✅ `emit_block_typed(n)` / `emit_loop_typed(n)` register a `() -> externref^N` blocktype |
//! | Multi-result IF types           | ⚠  Only via block/loop so far — no dedicated IF helper |
//!
//! Pathway: a compiler sets `chunk.result_arity = N` (for function
//! multi-return) or calls `emit_block_typed(line, N)` / `emit_loop_typed`
//! (for multi-result blocks). The type section auto-registers the
//! required `() -> externref^N` function types — see `types.rs`'s block
//! pre-scan. The code emitter writes the corresponding typeidx
//! blocktype (signed-LEB128) per the spec.

use vybe_runtime::Chunk;

pub const IMPORTS: &[(&str, &str)] = &[];
pub const GLOBAL_IMPORTS: &[(&str, &str)] = &[];
pub fn custom_sections(_chunks: &[Chunk]) -> Vec<(&'static str, Vec<u8>)> {
    Vec::new()
}
