//! Common AST targeting WASM bytecode.
//!
//! Every language parser (pest-based) produces this AST. One compiler consumes it.
//! The AST represents semantic operations, not syntax — language-specific syntax
//! is resolved by each language's tree walker before reaching this level.
//!
//! ## Modules
//! * [`emitter`] — the WASM bytecode emission layer (dict, collections, strings,
//!   loops, classes, stdlib, invoke, dotnet). Was the `vybe_compiler_common`
//!   crate; moved in-tree now that it has a single consumer.
//! * [`compiler`] — AST-driven dispatch that feeds the emitter.
//! * [`languages`] — per-language walkers + profiles.
//! * [`platforms`] — reusable runtime/framework surfaces such as .NET.

pub use vybe_ast as ast;
pub mod bundle;
pub mod compiler;
pub use compiler::Compiler; // avoid the `vybe_compiler::compiler::Compiler` stutter
pub mod emitter;
pub mod languages;
pub mod lsp;
pub mod platforms {
    //! Facade over the platform crates so `crate::platforms::…` paths
    //! keep resolving; the packages live at `platforms/*` in the workspace.
    pub use vybe_platform_dotnet as dotnet;
    pub use vybe_platform_libc as libc;
    pub use vybe_platform_plib as plib;
}
pub mod profile;
pub mod projects;
pub mod registry; // platform class-library registration (dotnet BCL, pascal plib)
