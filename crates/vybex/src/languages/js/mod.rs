mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "languages/js/grammar.pest"]
pub(crate) struct JsParser;

/// Parse JavaScript source into the common AST.
pub fn parse(source: &str) -> Result<crate::ast::Module, String> {
    walker::parse(source)
}
