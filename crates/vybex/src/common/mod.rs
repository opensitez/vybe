//! Cross-language compile-time helpers that sit **above** the emit
//! primitives in `crate::emitter::*` and **below** the per-language
//! walkers. Everything here consumes a normalised AST shape (produced
//! by language walkers) and emits bytecode via `crate::emitter::*`,
//! with no language-specific branching.
//!
//! - `classes` — class normalisation (see `classnormalization.md`).

pub mod classes;
