//! Platform class-library registration passes.
//!
//! These install a platform's *class wrappers* (as opposed to its namespace
//! tree — see the `tree_register` family and
//! `crate::primitives::namespaces::register_namespace_tree`) into the compiler's
//! class tables. They are `impl Compiler` extensions and therefore
//! compiler-coupled, so they live here rather than in the platform crates.
//!
//!
//! The former `dotnet` pass (per-class .NET ctor globals for VB / C#) is
//! retired — control/value/drawing types resolve through the component
//! descriptor + the GUI-direct element path instead of an emitted prelude.
