//! # reference-types proposal
//!
//! Spec: `proposals/reference-types/` (merged into core
//! WASM spec in 2022 — no longer strictly a "proposal"). Adds:
//!
//! * Two new reference types: `externref` (0x6F) and `funcref` (0x70)
//! * `ref.null t`  (0xD0) — null reference of type `t`
//! * `ref.is_null` (0xD1) — test a reference for null, returns `i32`
//! * `ref.func f`  (0xD2) — construct a funcref for function index `f`
//! * Tables of reference types (as opposed to only funcref)
//! * `table.grow`, `table.size`, `table.fill`, `table.copy`, `table.init`,
//!   `elem.drop` (all in the 0xFC prefix)
//! * Multiple tables per module
//!
//! ## Status in Vybe
//!
//! | Feature                     | Status |
//! |-----------------------------|--------|
//! | `externref`                 | ✅ used as universal value representation |
//! | `funcref` via table         | ✅ emitted in element section (chunks → funcref table) |
//! | `ref.null extern`           | ✅ `Op::NULL` → `0xD0 0x6F` |
//! | `ref.is_null`               | ✅ `Op::REF_IS_NULL` → `0xD1` |
//! | `ref.func N`                | ✅ `Op::REF_FUNC` → emits `i32.const N; emit_box_i32` (treated as table index) |
//! | externref params/results    | ✅ all chunk params and results are externref |
//! | externref locals            | ✅ chunk.local_count externref locals |
//! | `table.grow`                | ✅ `TABLE_GROW` opcode (VM grows `func_table`) |
//! | `table.size`                | ✅ `TABLE_SIZE` opcode |
//! | `table.fill` / `copy` / `init` | ✅ `TABLE_FILL` / `TABLE_COPY` / `TABLE_INIT` opcodes (overlap-safe) |
//! | Multiple tables             | ⚠ single function-table only; multi-table emission is structural and not yet wired |
//!
//! The unimplemented features only matter for dynamic table manipulation
//! at runtime — something neither VB nor JS compilers currently emit.

use crate::{Chunk, Op};

/// Reference-types imports nothing — all ops are core (prefix 0x00).
pub fn declare_imports() -> &'static [(&'static str, &'static str)] { &[] }

/// Reference-types declares no globals.
pub fn declare_globals() -> &'static [(&'static str, &'static str)] { &[] }

/// The reference-types ops that are **actually emitted** by our code
/// generator. Used by the test suite to audit coverage.
pub const EMITTED_OPCODES: &[Op] = &[Op::NULL, Op::REF_IS_NULL, Op::REF_FUNC];

/// Proposal adds no custom sections — reference types ride in the
/// standard type / table / element sections.
pub fn custom_sections(_chunks: &[Chunk]) -> Vec<(&'static str, Vec<u8>)> { Vec::new() }
