//! Dart adapter — bytecode-only.
//!
//! Mirrors the `emitter/fortran/`, `emitter/php/`, and `emitter/dotnet/`
//! patterns: Dart-specific surfaces that aren't a single WASM opcode or
//! single ECMA host fn are described as `emit_*` opcode-emitter functions
//! that compose pre-existing host fns and core WASM ops into the Dart
//! shape.
//!
//! No new host fns are registered. The Dart profile binds language fns
//! to `common:dart.<name>` emit targets; `emitter::dispatch::emit_common`
//! routes those names to the `emit_*` functions in this module.

pub mod dispatch;
pub mod string_adapter;
