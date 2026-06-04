//! Common symbol types — language-agnostic.

/// A symbol extracted from a parsed AST.
#[derive(Debug, Clone)]
pub struct Symbol {
    pub name: String,
    pub kind: SymbolKind,
    pub detail: String, // e.g. "(a: Integer, b: Integer): Integer" for a function
    pub line: u32,      // 0-based line where the symbol is defined
    pub end_line: u32,  // 0-based end line (for folding/outline)
    pub children: Vec<Symbol>, // nested symbols (methods inside a class, etc.)
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum SymbolKind {
    Function,
    Procedure,
    Class,
    Interface,
    Enum,
    EnumMember,
    Variable,
    Constant,
    Field,
    Property,
    Method,
    Constructor,
    Module,
    Struct,
    Event,
    Type,
}

/// A diagnostic (error/warning) from parsing.
#[derive(Debug, Clone)]
pub struct LspDiagnostic {
    pub line: u32, // 0-based
    pub col: u32,  // 0-based
    pub end_col: u32,
    pub message: String,
    pub severity: DiagSeverity,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum DiagSeverity {
    Error,
    Warning,
    Info,
}

/// A completion item for the autocomplete popup.
#[derive(Debug, Clone)]
pub struct CompletionItem {
    pub label: String,
    pub detail: String,
    pub insert_text: String,
    pub kind: SymbolKind,
}

/// The full analysis result for a single file.
#[derive(Debug, Clone)]
pub struct AnalysisResult {
    pub uri: String,
    pub version: u64,
    pub symbols: Vec<Symbol>,
    pub diagnostics: Vec<LspDiagnostic>,
    pub keywords: &'static [&'static str],
}
