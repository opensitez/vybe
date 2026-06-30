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

pub mod bundle;
pub mod canonical;
pub mod channels;
pub mod classes;
pub mod closures;
pub mod collections;
pub mod components;
pub mod convert;
pub mod delegates;
pub mod dict;
pub mod dispatch;
pub mod dotnet;
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
pub mod references;
pub mod runtime_helpers;
pub mod sprintf;
pub mod strings;
pub mod target;
pub mod threading;
pub mod type_registry;

pub use crate::languages::cobol::emitter as cobol;
pub use crate::languages::dart::emitter as dart;
pub use crate::languages::fortran::emitter as fortran;
pub use crate::languages::js::emitter as js;
pub use crate::languages::php::emitter as php;
pub use crate::languages::python::emitter as python;
pub use crate::languages::ruby::emitter as ruby;
pub use crate::languages::vb::emitter as vb;

pub use runtime_helpers::RuntimeHelpers;
pub use target::Target;
pub use type_registry::CompileTimeTypes;

/// Resolve a shared *platform* emit dispatcher by its `common:<prefix>.*`
/// prefix. Platforms are emit surfaces shared by more than one language —
/// currently `dotnet` (VB / C# / JS) and `libc` (C). Each platform module under
/// `emitter/` registers its prefix here; languages register their own via
/// [`crate::languages::Language::emit_dispatch`]. Returns `None` for
/// non-platform prefixes.
pub fn platform_emit_dispatch(prefix: &str) -> Option<crate::languages::EmitDispatch> {
    match prefix {
        "dotnet" => Some(dotnet::dispatch::dispatch),
        "libc" => Some(crate::platforms::libc::emitter::dispatch::dispatch),
        _ => None,
    }
}
