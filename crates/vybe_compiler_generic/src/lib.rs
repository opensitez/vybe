//! vybe_compiler_generic — Compiles vybe_parser_generic::Module to bytecode.
//!
//! One compiler for all languages. Consumes the common AST produced by the
//! grammar-driven parser. Uses vybe_compiler_common for all emission helpers.

mod compiler;
mod scope;

pub use compiler::Compiler;
