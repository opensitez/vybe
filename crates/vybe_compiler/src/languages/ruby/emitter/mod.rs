//! Ruby adapter — bytecode-only.
//!
//! Mirrors the `emitter/php/` and `emitter/fortran/` patterns:
//! Ruby-specific surfaces that aren't covered by a single WASM opcode
//! or a single ECMA host fn are described as `emit_*` opcode-emitter
//! functions that compose pre-existing host fns and core WASM ops into
//! the Ruby-shaped surface.
//!
//! No new host fns are registered. Everything is bytecode emission
//! over the spec-conformant ECMA-262 / WASI surface plus core WASM
//! opcodes.
//!
//! The Ruby profile binds language fns to `common:ruby.<name>` emit
//! targets; `emitter::dispatch::emit_common` routes those names to the
//! `dispatch` function in this module.

pub mod dispatch;
pub mod runtime_adapter;
