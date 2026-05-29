mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "languages/wast/grammar.pest"]
pub(crate) struct WastParser;

/// Parse WAST/WAT source into the common AST.
/// Both .wast (full script with assertions) and .wat (module only) are handled
/// by the same grammar — WAT is a strict subset of WAST.
pub fn parse(source: &str) -> Result<crate::ast::Module, String> {
    walker::parse(source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}
