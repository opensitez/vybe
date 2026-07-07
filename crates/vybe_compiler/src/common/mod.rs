//! Cross-language compile-time helpers that sit **above** the emit
//! primitives in `crate::emitter::*` and **below** the per-language
//! walkers. Everything here consumes a normalised AST shape (produced
//! by language walkers) and emits bytecode via `crate::emitter::*`,
//! with no language-specific branching.
//!
//! - `channels` — shared channel AST lowering helpers.
//! - `classes` — class normalisation (see `classnormalization.md`).
//! - `tuples` — named-tuple normalisation onto one canonical runtime shape.

pub mod channels;
pub mod classes;
pub mod events;
pub mod tuples;
