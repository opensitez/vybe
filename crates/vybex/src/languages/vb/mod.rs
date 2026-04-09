mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "languages/vb/grammar.pest"]
pub(crate) struct VbParser;

/// Parse VB source into the common AST.
pub fn parse(source: &str) -> Result<crate::ast::Module, String> {
    walker::parse(source)
}
