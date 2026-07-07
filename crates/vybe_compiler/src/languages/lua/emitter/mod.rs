//! Lua language adapter — bytecode-only.
//!
//! Lua-specific runtime semantics that can't be expressed as pure
//! walker normalisation go here as `emit_*` opcode-emitter functions.
//! No new host fns; no JS polyfills. Everything composes existing
//! `ecma:*` / `wasi:*` host surfaces.
//!
//! The Lua profile binds language fns to `common:lua.<name>` emit
//! targets; `emitter::dispatch::emit_common` routes those names here
//! via the per-language dispatch table registered in `languages/mod.rs`.

pub mod dispatch;
pub mod string_adapter;
pub mod metamethods_adapter;

