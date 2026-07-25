//! Compiler-side class normalisation. The language-agnostic IR + builders live
//! in `vybe_bytecode::class_normalize` (so language crates can produce them); they
//! are re-exported here so existing `crate::compiler::class_normalize::…` paths
//! keep resolving. The `emit` orchestrator (needs `Compiler` + walkers) stays.

pub use vybe_bytecode::class_normalize::*;

pub mod emit;
pub use emit::emit_class;
