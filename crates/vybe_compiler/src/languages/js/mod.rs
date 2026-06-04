pub mod emitter;
pub mod normalize_class;
mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "languages/js/grammar.pest"]
pub(crate) struct JsParser;

/// Parse JavaScript source into the common AST.
pub fn parse(source: &str) -> Result<crate::ast::Module, String> {
    walker::parse(source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}
