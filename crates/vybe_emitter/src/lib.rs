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
#[path = "enum.rs"]
pub mod r#enum;
pub mod errors;
pub mod events;
pub mod expressions;
pub mod functions;
pub mod generators;
pub mod gui;
pub mod heap;
pub mod imports;
pub mod instructions;
pub mod invoke;
pub mod io;
pub mod json;
pub mod loops;
pub mod math;
pub mod multivalue;
pub mod namespaces;
pub mod object;
pub mod ops;
pub mod packing;
pub mod promises;
pub mod prototypes;
pub mod random;
pub mod references;
pub mod reflection;
pub mod slices;
pub mod sorted_collection;
pub mod sprintf;
pub mod strings;
pub mod target;
pub mod threading;
pub mod tuples;
pub mod type_registry;
pub mod xml;

pub use target::Target;
pub use type_registry::CompileTimeTypes;
