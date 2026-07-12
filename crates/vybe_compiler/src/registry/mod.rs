//! Platform class-library registration passes.
//!
//! These install a platform's *class wrappers* (as opposed to its namespace
//! tree — see the `tree_register` family and
//! `vybe_emitter::namespaces::register_namespace_tree`) into the compiler's
//! class tables. They are `impl Compiler` extensions and therefore
//! compiler-coupled, so they live here rather than in the platform crates.
//!
//! - `dotnet` — .NET BCL wrapper classes (VB / C#).
//! - `plib` — Pascal GCL/plib adapter classes.

pub mod dotnet;
pub mod plib;
