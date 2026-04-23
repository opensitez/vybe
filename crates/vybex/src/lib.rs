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

pub mod emitter;
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
pub mod server;
