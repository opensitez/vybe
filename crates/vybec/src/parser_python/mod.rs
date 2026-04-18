pub mod token;
pub mod lexer;
pub mod ast;
pub mod parser;

pub use ast::*;

/// Parse Python source into a Module AST.
pub fn parse(source: &str) -> Result<ast::Module, String> {
    let mut parser = parser::Parser::new(source)?;
    parser.parse_module()
}
