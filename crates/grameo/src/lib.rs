//! Grameo: PEG grammar compilation and source parsing.
//!
//! Grammar definitions compile into located IR. The iterative reference engine
//! executes that IR without depending on pest, the compiler, or the VM.
pub mod analysis;
pub mod arena;
pub mod ast;
pub mod builder;
pub mod compat;
pub mod diagnostic;
pub mod editor;
pub mod engine;
pub mod grammar;
pub mod islands;
mod lexer;
pub mod lexical;
pub mod modules;
pub mod optimization;
pub mod pratt;
pub mod program;
pub mod quoted;
mod repetition_tails;
pub mod resolve;
pub mod source;
pub mod source_units;
pub mod streaming;
pub mod tree;

pub use diagnostic::{Diagnostic, Span};
pub use engine::{MatchOptions, ParseError, ParseErrorKind, Recognition};
pub use grammar::{
    Expr, ExprId, ExprKind, GrammarSyntax, ParseOptions, Rule, RuleMode, parse, parse_with_options,
};
pub use resolve::{Builtin, CompiledGrammar, Reference, compile, compile_with_ast_bindings, compile_with_options, CompileOptions, resolve};
pub use tree::{CaptureRule, Pair, Pairs, ParseTree};

mod wrapper_prefix;
mod group_prefix;
mod inner_prefix;

mod empty_success;

mod empty_suffix;
