//! Compatibility re-export.
//!
//! The canonical .NET platform surface now lives under `crate::platforms::dotnet`.
//! Keep this module as a shim so existing `crate::emitter::dotnet::*` call sites
//! continue to compile during the migration.

pub mod dispatch;
pub use crate::platforms::dotnet::*;
