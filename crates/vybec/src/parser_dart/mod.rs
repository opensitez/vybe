pub mod token;
pub mod lexer;
pub mod ast;
pub mod parser;

pub use ast::*;
pub use parser::Parser;

/// Parse a Dart source string into a Program AST.
pub fn parse(source: &str) -> Result<Program, String> {
    let mut p = Parser::new(source);
    p.parse_program()
}
