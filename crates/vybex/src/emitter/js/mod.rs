//! JS adapter — Rust inline opcode emitters.
//!
//! Mirrors `emitter/dart/`, `emitter/php/`, and `emitter/dotnet/`: JS-
//! specific surfaces (`new Proxy(...)`, proxy member access, etc.) that
//! aren't a single WASM opcode are described as `emit_*` functions
//! that compose pre-existing host fns and core WASM ops into the JS
//! shape. No new host fns are registered. No JS-source polyfills.

pub mod proxy_adapter;
