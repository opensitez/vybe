pub mod token;
pub mod lexer;
pub mod ast;
pub mod parser;

pub use ast::*;
pub use parser::Parser;

pub fn parse(source: &str) -> Result<Program, String> {
    let mut p = Parser::new(source);
    p.parse_program()
}
