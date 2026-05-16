//! COBOL language support.
//!
//! Pest grammar in `grammar.pest` parses COBOL 85/2002/2023 source.
//! The walker in `walker.rs` converts the parse tree into the common
//! `vybe_compiler::ast::Module`. From there everything goes through the
//! shared compiler — no COBOL-specific code in `compiler.rs`.

pub mod normalize_class;
pub mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "languages/cobol/grammar.pest"]
pub(crate) struct CobolParser;

pub fn parse(source: &str) -> Result<crate::ast::Module, String> {
    walker::parse(source)
}

/// Embedded TOML profile source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}
