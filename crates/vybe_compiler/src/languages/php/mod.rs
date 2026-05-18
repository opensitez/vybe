//! PHP language support.
//!
//! Pest grammar in `grammar.pest` parses PHP 8 source. The walker in
//! `walker.rs` converts the parse tree into the common
//! `vybe_compiler::ast::Module`. From there everything goes through the
//! shared compiler — no PHP-specific code in `compiler.rs`.

mod walker;
pub mod emitter;
pub mod normalize_class;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "languages/php/grammar.pest"]
pub(crate) struct PhpParser;

pub fn parse(source: &str) -> Result<crate::ast::Module, String> {
    walker::parse(source)
}

pub(crate) fn normalize_source_for_parser(source: &str) -> String {
    walker::normalize_source_for_parser(source)
}

/// Embedded TOML profile source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}
