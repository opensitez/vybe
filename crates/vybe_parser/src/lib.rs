//! Vybe-owned grammar compilation, independent of the compiler and VM.
//!
//! This first milestone reads grammar definitions into a located arena IR and
//! resolves references. It does not yet execute grammars against guest source.
pub mod diagnostic;
pub mod grammar;
mod lexer;
pub mod resolve;

pub use diagnostic::{Diagnostic, Span};
pub use grammar::{
    Expr, ExprId, ExprKind, GrammarSyntax, ParseOptions, Rule, RuleMode, parse, parse_with_options,
};
pub use resolve::{Builtin, CompiledGrammar, Reference, compile, resolve};
