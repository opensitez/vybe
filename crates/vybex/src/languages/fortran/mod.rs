//! Fortran language support (Modern Fortran 90/95/2003/2008/2018).
//!
//! Pest grammar in `grammar.pest` parses free-form Fortran source.
//! The walker in `walker.rs` converts the parse tree into the common
//! `vybex::ast::Module`. From there everything goes through the
//! shared compiler — no Fortran-specific code in `compiler.rs`.

pub mod walker;
pub mod emitter;
pub mod normalize_class;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "languages/fortran/grammar.pest"]
pub(crate) struct FortranParser;

pub fn parse(source: &str) -> Result<crate::ast::Module, String> {
    walker::parse(source)
}

/// Embedded TOML profile source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}
