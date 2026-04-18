pub mod token;
pub mod lexer;
pub mod ast;
pub mod parser;

pub use ast::*;
pub use parser::Parser;

/// Parse a JavaScript source string into a Program AST.
pub fn parse(source: &str) -> Result<Program, String> {
    let mut parser = Parser::new(source);
    parser.parse_program()
}
