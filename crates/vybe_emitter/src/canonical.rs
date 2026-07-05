//! Canonical builtin operations — the language-agnostic API surface.
//!
//! Walkers normalize language-specific syntax to these canonical names.
//! The compiler looks up canonical names in this module to get the bytecode emitter.
//!
//! ## Naming convention (Python-style dunders)
//!
//! - `__len__`      — length/size of collection or string
//! - `__str__`      — string representation
//! - `__upper__`    — uppercase string
//! - `__lower__`    — lowercase string
//! - `__trim__`     — trim whitespace
//! - `__contains__` — membership test
//!
//! ## Why dunders?
//!
//! Python pioneered this convention. Using consistent canonical names across all languages
//! makes cross-language interop trivial: a method bound under `__len__` is callable from
//! any language regardless of its surface syntax (Python `len()`, JS `.length`, C# `.Length`,
//! VB `Length()`, Pascal `Length()`, Ruby `.size`).
//!
//! ## Adding a new canonical builtin
//!
//! 1. Add the canonical name to the [`CanonicalOp`] enum below
//! 2. Add the dispatch arm to [`emit_canonical`]
//! 3. Update walker(s) to map the language-specific syntax to the canonical name
//!
//! No changes needed in the language-agnostic compiler.

use crate::{collections, strings};
use vybe_bytecode::Chunk;

/// A canonical builtin operation. Walkers normalize language-specific syntax to these.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CanonicalOp {
    /// Length/size of collection or string. Stack: [obj] → [int]
    Len,
    /// String representation. Stack: [obj] → [string]
    /// Uses stdlib __vybe_tostring (pure WASM, no host dependency).
    Str,
    /// Uppercase string. Stack: [str] → [str]
    Upper,
    /// Lowercase string. Stack: [str] → [str]
    Lower,
    /// Trim whitespace. Stack: [str] → [str]
    Trim,
}

impl CanonicalOp {
    /// Look up a canonical operation by its dunder name.
    pub fn from_name(name: &str) -> Option<CanonicalOp> {
        match name {
            "__len__" => Some(CanonicalOp::Len),
            "__str__" => Some(CanonicalOp::Str),
            "__upper__" => Some(CanonicalOp::Upper),
            "__lower__" => Some(CanonicalOp::Lower),
            "__trim__" => Some(CanonicalOp::Trim),
            _ => None,
        }
    }

    /// How many args does this operation take?
    pub fn arity(&self) -> u8 {
        match self {
            CanonicalOp::Len
            | CanonicalOp::Str
            | CanonicalOp::Upper
            | CanonicalOp::Lower
            | CanonicalOp::Trim => 1,
        }
    }
}

/// Emit the bytecode for a canonical operation.
/// The args must already be on the stack in the correct order.
///
/// Takes `chunks` + `current` because `__len__` compiles to a
/// `ecma:array.length` / `wasm:js-string.length` runtime dispatch
/// and the import must register on `chunks[0]` (the module-level
/// imports section) while the code emits on `chunks[current]`.
pub fn emit_canonical(op: CanonicalOp, chunks: &mut [Chunk], current: usize, line: u32) {
    match op {
        // Use Op::ARRAY_LENGTH (WASM GC array.len) — handles both
        // ObjectKind::Array and Value::String in the interpreter and
        // emits the native array.len byte in the wasm tier.  No host
        // call, no import-table indirection; avoids the import-index
        // collision that emit_len (chunks[current]-relative indices vs
        // the VM's chunks[0]-based resolution) causes when .length is
        // accessed inside a nested function chunk.
        CanonicalOp::Len => collections::emit_array_length(&mut chunks[current], line),
        CanonicalOp::Str => {
            // __vybe_tostring is populated by bundle::finalize_with_runtime_helpers.
            strings::emit_to_string(&mut chunks[current], line);
        }
        CanonicalOp::Upper => strings::emit_to_upper(&mut chunks[current], line),
        CanonicalOp::Lower => strings::emit_to_lower(&mut chunks[current], line),
        CanonicalOp::Trim => strings::emit_trim(&mut chunks[current], line),
    }
}
