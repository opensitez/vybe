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
//! | `ref.null func`             | ✅ `Op::NULL_FUNC` → `0xD0 0x70` |
//! | `ref.null any`              | ✅ `Op::NULL_ANY` → `0xD0 0x6E` |
//! | `ref.null none`             | ✅ `Op::NULL_NONE` → `0xD0 0x71` |
//! | `ref.is_null`               | ✅ `Op::REF_IS_NULL` → `0xD1` |
//! | `ref.func N`                | ✅ (spec-compliant) `Op::REF_FUNC` emits `i32.const N; box_i32` and relies on `call_indirect` to dispatch against the funcref-typed table. The module validates on any engine because (a) the element section populates the funcref table with genuine `ref.func` expressions and (b) `call_indirect tableidx typeidx` legitimately consumes an i32 index. We don't use the `0xD2` instruction on the stack — that would require funcref-typed locals, which in turn requires a `Value` ABI refactor. |
//! | `table.get` / `table.set`   | ✅ `Op::TABLE_GET` / `Op::TABLE_SET` → core `0x25` / `0x26` with tableidx |
//! | Typed `select t`            | ✅ `Op::SELECT_T` → `0x1C 0x01 0x6F` (single externref result) |
//! | `table.grow`                | ✅ `TABLE_GROW` opcode; routes through tableidx |
//! | `table.size`                | ✅ `TABLE_SIZE` opcode |
//! | `table.fill`     Th           | ✅ `TABLE_FILL` opcode |
//! | `table.copy` / `table.init` | ✅ `TABLE_COPY` / `TABLE_INIT` opcodes (overlap-safe; copy operates on func_table) |
//! | Multiple tables             | ✅ `extra_tables` on VM; `table_ref`/`table_mut` route by tableidx. Emission-side multi-table declarations still go through `encode_table_section_with` (opt-in). |
//! | Typed function references `(ref null $typeidx)` | ❌ function signatures always use `externref`. A Value ABI refactor is needed before funcref-typed params/locals can be emitted. |
//!
//! The one remaining gap (typed function references) requires marking
//! some locals/params as funcref at the compiler level and plumbing
//! that type through the Value enum — a larger refactor than closing
//! the opcode-level gaps.

use vybe_bytecode::{Chunk, Op};

/// Reference-types imports nothing — all ops are core (prefix 0x00).
pub fn declare_imports() -> &'static [(&'static str, &'static str)] {
    &[]
}

/// Reference-types declares no globals.
pub fn declare_globals() -> &'static [(&'static str, &'static str)] {
    &[]
}

/// The reference-types ops that are **actually emitted** by our code
/// generator. Used by the test suite to audit coverage.
pub const EMITTED_OPCODES: &[Op] = &[Op::NULL, Op::REF_IS_NULL, Op::REF_FUNC];

/// Proposal adds no custom sections — reference types ride in the
/// standard type / table / element sections.
pub fn custom_sections(_chunks: &[Chunk]) -> Vec<(&'static str, Vec<u8>)> {
    Vec::new()
}
