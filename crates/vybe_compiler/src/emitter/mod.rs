//! Emit-layer facade.
//!
//! The pure chunk-level emit surface lives in the `vybe_emitter` crate
//! (adapters, instructions, ops, collections, …) and is re-exported
//! here wholesale, so `crate::emitter::ops::…` paths keep resolving.
//! What actually lives at this level is the routing and bundling that
//! must see languages and platforms:
//! - `dispatch` — the `common:<name>` router (languages + platforms)
//! - `runtime_helpers` — bundled stdlib chunks + polyglot polyfills
//!   (compiles snippets through registered language frontends)
//! - `bundle` — links runtime helpers into compiled programs

pub use vybe_emitter::*;

pub mod bundle;
pub mod dispatch;
pub mod runtime_helpers;

pub use crate::platforms::dotnet::emitter as dotnet;

pub use runtime_helpers::RuntimeHelpers;

/// Resolve a shared *platform* emit dispatcher by its `common:<prefix>.*`
/// prefix. Platforms are emit surfaces shared by more than one language —
/// currently `dotnet` (VB / C# / JS) and `libc` (C). Languages register
/// their own via [`crate::languages::Language::emit_dispatch`]. Returns
/// `None` for non-platform prefixes.
pub fn platform_emit_dispatch(prefix: &str) -> Option<crate::languages::EmitDispatch> {
    match prefix {
        "dotnet" => Some(dotnet::dispatch::dispatch),
        "libc" => Some(crate::platforms::libc::emitter::dispatch::dispatch),
        _ => None,
    }
}
