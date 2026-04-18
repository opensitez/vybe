//! vybe_lsp — AST-based language intelligence for the Vybe IDE.
//!
//! Runs parsers in a background thread. Sends back symbols, diagnostics,
//! completions, and hover info via crossbeam channels. Non-blocking.

mod symbols;
mod extract;
mod engine;

pub use symbols::*;
pub use engine::{AnalysisEngine, AnalysisRequest, AnalysisEvent, analyze};
