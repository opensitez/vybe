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
//! | Multi-result function types     | ⚠  Our ABI is `(externref*) -> externref` — multi-result never emitted |
//! | Multi-result block types        | ⚠  Blocks use type `externref` or `void` only |
//! | Multi-result control flow       | ❌ |
//!
//! We don't currently exploit multi-value on the emit side — every
//! chunk returns exactly one externref, and all blocks / loops are
//! either `void` or `(result externref)`. The type-section encoder in
//! `types.rs` does correctly handle multi-result function types (the
//! format allows a `vec(valtype)` for results), so adding multi-value
//! support would be additive: nothing to undo.
//!
//! ## Why not?
//!
//! Our universal-externref representation makes multi-value mostly
//! unnecessary — returns are already dynamic. The one legitimate use
//! (returning both success + status from a host call) could be added
//! per-import in the future.

use crate::Chunk;

pub const IMPORTS: &[(&str, &str)] = &[];
pub const GLOBAL_IMPORTS: &[(&str, &str)] = &[];
pub fn custom_sections(_chunks: &[Chunk]) -> Vec<(&'static str, Vec<u8>)> { Vec::new() }
