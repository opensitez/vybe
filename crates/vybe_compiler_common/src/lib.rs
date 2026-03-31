//! Shared compilation helpers for all Vybe language compilers.
//!
//! Emits portable WASM-compatible bytecode sequences for common patterns:
//! - Dict/map operations (built from GC struct ops)
//! - Math builtins (host imports with standard names)
//! - Type conversions (host imports)
//! - Print/IO (WASI imports)
//!
//! Language compilers call these instead of reimplementing the same patterns.
//! Everything emitted is standard WASM — no custom opcodes.

pub mod dict;
pub mod math;
pub mod convert;
pub mod io;
pub mod collections;
