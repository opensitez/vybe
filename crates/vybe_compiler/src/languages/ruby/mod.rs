mod walker;
pub mod normalize_class;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "languages/ruby/grammar.pest"]
pub(crate) struct RubyParser;

/// Parse Ruby source into the common AST.
pub fn parse(source: &str) -> Result<crate::ast::Module, String> {
    walker::parse(source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}
