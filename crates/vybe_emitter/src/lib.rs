//! Shared emit surface for all Vybe language compilers and platforms.
//!
//! Emits portable WASM-compatible bytecode sequences for common patterns:
//! - Dict/map operations (built from GC struct ops)
//! - Math builtins (host imports with standard names)
//! - Type conversions (host imports)
//! - Print/IO (WASI imports)
//!
//! Language compilers and platform packages call these instead of
//! reimplementing the same patterns. Everything emitted is standard
//! WASM — no custom opcodes. This crate is pure chunk-level emission:
//! it depends only on the bytecode data model (and vybe_host for host
//! function names), never on languages, platforms, or the compiler —
//! routing of `common:*` names lives upstairs in vybe_compiler's
//! dispatcher, where all parties are visible.

pub mod canonical;
pub mod channels;
pub mod classes;
pub mod closures;
pub mod collections;
pub mod components;
pub mod convert;
pub mod delegates;
pub mod dict;
pub mod errors;
pub mod expressions;
pub mod functions;
pub mod generators;
pub mod gui;
pub mod imports;
pub mod instructions;
pub mod invoke;
pub mod io;
pub mod loops;
pub mod math;
pub mod ops;
pub mod promises;
pub mod prototypes;
pub mod references;
pub mod sprintf;
pub mod strings;
pub mod target;
pub mod threading;
pub mod type_registry;

pub use target::Target;
pub use type_registry::CompileTimeTypes;
