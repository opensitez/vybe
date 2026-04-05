//! Grammar definition types — loaded from .grammar files.
//!
//! A GrammarDef fully describes a language's syntax. The engine
//! uses it to tokenize, parse expressions, and parse statements.

use std::collections::HashMap;

/// Complete grammar definition for one language.
#[derive(Debug, Clone)]
pub struct GrammarDef {
    pub language: LanguageSpec,
    pub lexer: LexerSpec,
    pub operators: OperatorTable,
    pub blocks: BlockSpec,
    pub types: TypeSpec,
    pub statements: Vec<PatternRule>,
    pub declarations: Vec<PatternRule>,
    pub expressions: ExpressionSpec,
    pub params: ParamSpec,
    pub assignment: AssignmentSpec,
    pub program: ProgramSpec,
}

// ── Language basics ──────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct LanguageSpec {
    pub name: String,
    pub case_sensitive: bool,
    pub statement_terminator: Terminator,
    pub indentation_based: bool,
    pub expression_language: bool,   // Lisp-like: everything is an expression
}

#[derive(Debug, Clone, PartialEq)]
pub enum Terminator {
    Char(char),       // ';'
    Newline,          // Python
    None,             // Lisp, Ruby
    Asi,              // JS: automatic semicolon insertion
}

// ���─ Lexer spec ───────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct LexerSpec {
    pub comment_line: Vec<String>,           // "//" or "#" or ";" etc.
    pub comment_block: Vec<(String, String)>, // ("/*", "*/"), ("{", "}") etc.
    pub string_delimiters: Vec<String>,       // "'", "\"", "`"
    pub string_escape: Option<String>,        // "\\" or "''"
    pub triple_string: Vec<String>,           // "'''" or "\"\"\""
    pub string_prefixes: Vec<String>,         // "f", "r", "b"
    pub interpolation: Option<(String, String)>, // ("${", "}") or ("{", "}")
    pub template_string: Option<String>,      // "`" for JS template literals
    pub char_prefix: Option<String>,          // "#" for Pascal #65
    pub hex_prefix: Option<String>,           // "$" for Pascal, "0x" for most
    pub keywords: Vec<String>,
    pub operators: Vec<String>,               // sorted longest-first for matching
}

// ── Operator table ───────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct OperatorTable {
    pub prefix: Vec<String>,
    pub postfix: Vec<String>,
    pub infix: Vec<InfixLevel>,
}

#[derive(Debug, Clone)]
pub struct InfixLevel {
    pub precedence: u8,
    pub ops: Vec<String>,
    pub assoc: Assoc,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Assoc { Left, Right }

// ── Block spec ───────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct BlockSpec {
    pub open: String,        // "{" or "begin" or "INDENT"
    pub close: String,       // "}" or "end" or "DEDENT"
    pub prefix: Option<String>,  // ":" for Python (before INDENT)
    /// VB-style: close is "end" followed by the block kind keyword.
    /// e.g. "End Sub", "End Function", "End If", "End Module"
    pub close_with_kind: bool,
}

// ── Type annotation spec ─────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct TypeSpec {
    pub position: TypePosition,
    pub separator: Option<String>,         // ":" or "As"
    pub return_separator: Option<String>,  // "->" for Python, ":" for Pascal
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum TypePosition {
    Before,   // C#, Dart: int x
    After,    // Pascal, VB: x: Integer / x As Integer
    None,     // JS, Ruby: no types
}

// ── Pattern rules ────────────────────────────────────────────────────────────

/// A syntax pattern for a statement or declaration.
#[derive(Debug, Clone)]
pub struct PatternRule {
    pub name: String,
    pub pattern: Vec<PatternElement>,
    pub maps_to: String,   // AST node kind name
    pub extra: HashMap<String, String>,
}

/// An element in a pattern.
#[derive(Debug, Clone)]
pub enum PatternElement {
    /// Literal keyword or operator: "if", "then", "(", etc.
    Keyword(String),
    /// Parse an expression.
    Expr,
    /// Parse an identifier.
    Ident,
    /// Parse a block (compound or single statement).
    Block,
    /// Parse a list of statements (until a closing keyword).
    StmtList,
    /// Parse a type annotation.
    Type,
    /// Parse a parameter list (with parens, types, defaults).
    Params,
    /// Parse a list of identifiers, comma-separated.
    IdentList,
    /// Parse a list of expressions, comma-separated.
    ExprList,
    /// Parse the declaration section (for Pascal-style languages).
    DeclSection,
    /// Parse case arms (for switch/case).
    CaseArms,
    /// Parse catch/except clauses.
    CatchClauses,
    /// Parse class members.
    ClassMembers,
    /// Parse interface members.
    InterfaceMembers,
    /// Parse record/struct members.
    RecordMembers,
    /// Parse enum members.
    EnumMembers,
    /// Optional group: (X)?
    Optional(Vec<PatternElement>),
    /// Alternatives: (X | Y)
    Alternatives(Vec<Vec<PatternElement>>),
    /// Repetition: (X)*
    Repeat(Vec<PatternElement>),
    /// A string literal token (for import paths etc.)
    StringLit,
    /// Newline token (for Python decorators etc.)
    Newline,
}

// ── Expression spec ──────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct ExpressionSpec {
    pub member_access: Option<String>,        // "."
    pub optional_chain: Option<String>,       // "?."
    pub index_open: Option<String>,           // "["
    pub index_close: Option<String>,          // "]"
    pub call_open: Option<String>,            // "("
    pub call_close: Option<String>,           // ")"
    pub deref: Option<String>,               // "^" for Pascal
    pub primary_forms: Vec<PrimaryForm>,
}

#[derive(Debug, Clone)]
pub enum PrimaryForm {
    ArrayLiteral(String, String),     // "[", "]"
    ObjectLiteral(String, String),    // "{", "}"
    ParenGroup(String, String),       // "(", ")"
    Lambda(Vec<PatternElement>),      // language-specific lambda syntax
}

// ── Param spec ───────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct ParamSpec {
    pub open: String,
    pub close: String,
    pub separator: String,             // "," or ";"
    pub name_type_sep: Option<String>, // ":" or "As"
    pub type_position: TypePosition,
    pub default_value: Option<String>, // "="
    pub rest_prefix: Option<String>,   // "..." or "*"
    pub kwargs_prefix: Option<String>, // "**"
    pub multi_name: bool,              // Pascal: (a, b: Integer)
    pub multi_name_sep: Option<String>,// ","
    pub pass_by: HashMap<String, String>, // "var" → "ref", "const" → "const"
}

// ── Assignment spec ──────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct AssignmentSpec {
    pub operator: Option<String>,      // ":=" or "=" or None (Lisp)
    pub compound: HashMap<String, String>, // "+=" → "Add", "-=" → "Sub"
    pub walrus: Option<String>,        // ":=" for Python
}

// ── Program structure ────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct ProgramSpec {
    pub header: Option<Vec<PatternElement>>,
    pub uses: Option<Vec<PatternElement>>,
    pub body: Option<Vec<PatternElement>>,
}
