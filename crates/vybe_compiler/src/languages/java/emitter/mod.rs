//! Java adapter — bytecode-only.
//!
//! Java-specific surfaces that aren't a single WASM opcode or single ECMA
//! host fn are described as `emit_*` opcode-emitter functions composing
//! pre-existing host fns and core WASM ops.
//!
//! No new host fns are registered. The Java profile binds language fns to
//! `common:java.<name>` emit targets; `emitter::dispatch::emit_common`
//! routes those names to the `emit_*` functions in this module.

pub mod arrays_adapter;
pub mod biginteger_adapter;
pub mod bitset_adapter;
pub mod dispatch;
pub mod enum_set_adapter;
pub mod instant_adapter;
pub mod list_adapter;
pub mod optional_adapter;
pub mod stream_adapter;
pub mod string_adapter;
pub mod uuid_adapter;
