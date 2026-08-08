//! # bulk-memory-operations proposal
//!
//! Spec: `proposals/spec/proposals/bulk-memory-operations/`.
//! Adds bulk operations on memory and tables:
//! `memory.init`, `memory.copy`, `memory.fill`, `data.drop`,
//! `table.init`, `table.copy`, `elem.drop`. All under prefix `0xFC`.
//!
//! ## Status in Vybe
//!
//! | Op                     | Status |
//! |------------------------|--------|
//! | `memory.copy`          | ✅ real byte copy, memory64-aware, unsigned operands |
//! | `memory.fill`          | ✅ |
//! | `memory.init`          | ✅ reads real data segments; respects `data.drop`; unsigned operands trap OOB |
//! | `data.drop`            | ✅ marks the segment dropped (`dropped_data`) |
//! | `table.copy`           | ✅ overlap-aware copy on `func_table` |
//! | `table.init`           | ✅ bounds-checked against element segments |
//! | `elem.drop`            | ✅ |
//!
//! The 0xFC-prefix emitter in `code.rs` is NOT pass-through: it reads the
//! internal fixed-width immediates (u16 BE indices, including the
//! multi-memory memidx on memory.init/copy/fill) and re-serializes them
//! as the spec's LEB immediates. The reader does the inverse, and its
//! validation pass rejects indices above the internal u16 width instead
//! of truncating.

use vybe_runtime::Chunk;

pub const IMPORTS: &[(&str, &str)] = &[];
pub const GLOBAL_IMPORTS: &[(&str, &str)] = &[];
pub fn custom_sections(_chunks: &[Chunk]) -> Vec<(&'static str, Vec<u8>)> {
    Vec::new()
}
