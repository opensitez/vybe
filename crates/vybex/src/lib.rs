//! Common AST targeting WASM bytecode via vybe_compiler_common.
//!
//! Every language parser (pest-based) produces this AST. One compiler consumes it.
//! The AST represents semantic operations, not syntax — language-specific syntax
//! is resolved by each language's tree walker before reaching this level.
//!
//! Design principle: if compiler_common can emit it, the AST has a typed node for it.

pub mod ast;
pub mod compiler;
pub mod dotnet_register;
pub mod languages;
pub mod profile;
pub mod scope;
pub mod bundle;
pub mod projects;
pub mod lsp;
pub mod gui_launch;
