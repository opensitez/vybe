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
//! | `memory.copy`          | ✅  VM handler copies real bytes |
//! | `memory.fill`          | ✅  |
//! | `memory.init`          | ✅ (no data section support yet — always acts on empty init region) |
//! | `data.drop`            | ✅ VM no-op (no runtime data segments) |
//! | `table.copy`           | ✅ overlap-aware copy on `func_table` |
//! | `table.init`           | ✅ bounds-checked no-op (no runtime element segments) |
//! | `elem.drop`            | ✅ VM no-op |
//!
//! The 0xFC-prefix emitter in `code.rs` treats these as pass-through
//! (prefix + sub + operands), which is correct for the arithmetic and
//! memory-ref sub-opcodes. Table operations would need our element
//! section to expose drop-able segments, which we don't currently do.

use crate::Chunk;

pub const IMPORTS: &[(&str, &str)] = &[];
pub const GLOBAL_IMPORTS: &[(&str, &str)] = &[];
pub fn custom_sections(_chunks: &[Chunk]) -> Vec<(&'static str, Vec<u8>)> {
    Vec::new()
}
