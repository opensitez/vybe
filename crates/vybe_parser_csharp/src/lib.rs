pub mod ast;
pub mod token;
pub mod lexer;
pub mod parser;

pub use ast::CompilationUnit;

/// Parse C# source code into an AST.
pub fn parse(source: &str) -> Result<CompilationUnit, String> {
    let tokens = lexer::tokenize(source);
    let mut parser = parser::Parser::new(tokens);
    parser.parse()
}
