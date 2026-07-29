//! Compiler-side class normalisation. The language-agnostic IR + builders live
//! in `vybe_ast::class_normalize` (so language crates can produce them); they
//! are re-exported here so existing `crate::primitives::class_normalize::…` paths
//! keep resolving. The `emit` orchestrator (needs `Compiler` + walkers) stays.

pub use vybe_ast::class_normalize::*;

pub mod emit;
pub use emit::emit_class;
