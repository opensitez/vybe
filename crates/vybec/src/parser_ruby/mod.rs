pub mod token;
pub mod lexer;
pub mod ast;
pub mod parser;

pub use ast::*;
pub use parser::Parser;

/// Parse a Ruby source string into a Program AST.
pub fn parse(source: &str) -> Result<Program, String> {
    Parser::new(source)?.parse_program()
}
