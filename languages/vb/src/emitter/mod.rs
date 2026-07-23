//! VB-specific emitter dispatch.
//!
//! Dotnet/BCL surface is resolved through `platforms/dotnet`; this module
//! remains as the language hook for any truly VB-local lowering.

pub mod dispatch;
