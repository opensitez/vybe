//! Fortran adapter — bytecode-only.
//!
//! Mirrors the `emitter/php/` and `emitter/dotnet/` patterns:
//! Fortran-specific intrinsics that aren't covered by a single WASM
//! opcode or a single ECMA host fn are described as `emit_*`
//! opcode-emitter functions that compose pre-existing host fns
//! (`ecma:math.*`, `ecma:string.*`, `wasi:*`, etc.) and WASM ops
//! into the Fortran-shaped surface.
//!
//! No new host fns are registered. Everything is bytecode emission
//! over the spec-conformant ECMA-262 / WASI surface plus core WASM
//! opcodes.
//!
//! The Fortran profile binds language fns to `common:fortran.<name>`
//! emit targets; `emitter::dispatch::emit_common` routes those names
//! to the `emit_*` functions in this module.

pub mod bit_adapter;
pub mod dispatch;
pub mod math_adapter;
pub mod string_adapter;
