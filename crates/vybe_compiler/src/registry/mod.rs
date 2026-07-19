//! Platform class-library registration passes.
//!
//! These install a platform's *class wrappers* (as opposed to its namespace
//! tree — see the `tree_register` family and
//! `vybe_emitter::namespaces::register_namespace_tree`) into the compiler's
//! class tables. They are `impl Compiler` extensions and therefore
//! compiler-coupled, so they live here rather than in the platform crates.
//!
//! - `plib` — Pascal GCL/plib adapter classes.
//!
//! The former `dotnet` pass (per-class .NET ctor globals for VB / C#) is
//! retired — control/value/drawing types resolve through the component
//! descriptor + the GUI-direct `vybe:gui` path instead of an emitted prelude.

pub mod plib;
