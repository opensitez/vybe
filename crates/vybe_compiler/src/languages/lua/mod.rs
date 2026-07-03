mod normalize;
mod walker;
pub mod emitter;


use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "languages/lua/grammar.pest"]
pub(crate) struct LuaParser;

/// Parse Lua source into the common AST.
pub fn parse(source: &str) -> Result<crate::ast::Module, String> {
    walker::parse(source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}
