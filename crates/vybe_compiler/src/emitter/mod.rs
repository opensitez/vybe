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
pub mod references;
pub mod math;
pub mod convert;
pub mod io;
pub mod collections;
pub mod classes;
pub mod loops;
pub mod errors;
pub mod strings;
pub mod functions;
pub mod delegates;
pub mod expressions;
pub mod threading;
pub mod components;
pub mod target;
pub mod stdlib;
pub mod bundle;
pub mod imports;
pub mod gui;
pub mod dotnet;
pub mod dispatch;
pub mod type_registry;
pub mod canonical;
pub mod invoke;

pub use crate::languages::php::emitter as php;
pub use crate::languages::python::emitter as python;
pub use crate::languages::fortran::emitter as fortran;
pub use crate::languages::dart::emitter as dart;
pub use crate::languages::js::emitter as js;
pub use crate::languages::vb::emitter as vb;

pub use target::Target;
pub use type_registry::CompileTimeTypes;
pub use stdlib::StdLib;
