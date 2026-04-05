//! vybe_parser_generic — Unified AST and grammar-driven parser engine.
//!
//! Every language parser produces a `Module` containing `Statement`s and `Expression`s.
//! Compilers, LSP, and tooling consume this common format instead of per-language ASTs.
//!
//! The grammar engine reads `.grammar` files (one per language) and uses a single
//! generic parser to tokenize and parse any language into the common AST.

pub mod grammar;
pub mod lexer;
pub mod parser;
pub mod profile;

/// A parsed source file.
#[derive(Debug, Clone)]
pub struct Module {
    pub name: String,
    pub language: Lang,
    pub body: Vec<Statement>,
    /// Imports/uses/requires at the top of the file.
    pub imports: Vec<Import>,
}

/// Source language tag.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Lang {
    VB, JavaScript, CSharp, Python, Ruby, PHP, Dart, Pascal, Cobol, Unknown,
}

/// An import/use/require declaration.
#[derive(Debug, Clone)]
pub struct Import {
    pub path: String,
    pub alias: Option<String>,
    pub names: Vec<String>, // specific imports (empty = import all)
    pub span: Span,
}

/// Source location (0-based lines and columns).
#[derive(Debug, Clone, Copy, Default)]
pub struct Span {
    pub start_line: u32,
    pub start_col: u32,
    pub end_line: u32,
    pub end_col: u32,
}

// ── Statements ───────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct Statement {
    pub kind: StmtKind,
    pub span: Span,
}

impl Statement {
    pub fn new(kind: StmtKind) -> Self { Self { kind, span: Span::default() } }
    pub fn with_span(kind: StmtKind, span: Span) -> Self { Self { kind, span } }
}

#[derive(Debug, Clone)]
pub enum StmtKind {
    /// Expression used as a statement (call, assignment expression, etc.)
    Expr(Expression),

    /// `begin ... end` / `{ ... }` / indented block
    Block(Vec<Statement>),

    // ── Declarations ──────────────────────────────────────────────────────

    /// Variable/local declaration.
    VarDecl {
        name: String,
        type_hint: Option<String>,
        init: Option<Expression>,
        is_const: bool,
        /// `var` (by-reference) parameter, mutable, etc.
        mutable: bool,
    },

    /// Function/method/procedure/sub declaration.
    FunctionDecl {
        name: String,
        params: Vec<Param>,
        return_type: Option<String>,
        body: Vec<Statement>,
        modifiers: Modifiers,
    },

    /// Class declaration.
    ClassDecl {
        name: String,
        parent: Option<String>,
        interfaces: Vec<String>,
        members: Vec<Statement>,   // fields, methods, properties as nested stmts
        modifiers: Modifiers,
    },

    /// Interface declaration (methods only, no body).
    InterfaceDecl {
        name: String,
        parent: Option<String>,
        members: Vec<Statement>,
    },

    /// Enum declaration.
    EnumDecl {
        name: String,
        members: Vec<EnumMember>,
    },

    /// Struct / record / value type.
    StructDecl {
        name: String,
        members: Vec<Statement>,
    },

    /// Module / namespace.
    ModuleDecl {
        name: String,
        body: Vec<Statement>,
    },

    /// Type alias: `type Foo = Bar`
    TypeAlias {
        name: String,
        target: String,
    },

    /// Property declaration (class member with getter/setter).
    PropertyDecl {
        name: String,
        type_hint: Option<String>,
        getter: Option<String>,  // method name or field name
        setter: Option<String>,
    },

    /// Event declaration (VB, C#).
    EventDecl {
        name: String,
        type_hint: Option<String>,
    },

    // ── Control flow ──────────────────────────────────────────────────────

    /// if / elif / else
    If {
        cond: Expression,
        then: Vec<Statement>,
        elifs: Vec<(Expression, Vec<Statement>)>,
        else_: Option<Vec<Statement>>,
    },

    /// for i := 0 to 10 / for (init; cond; update)
    For {
        init: Option<Box<Statement>>,
        cond: Option<Expression>,
        update: Option<Expression>,
        body: Vec<Statement>,
    },

    /// for x in collection / for each x in arr
    ForIn {
        var: String,
        iter: Expression,
        body: Vec<Statement>,
    },

    /// while cond do body
    While {
        cond: Expression,
        body: Vec<Statement>,
    },

    /// do body while cond / repeat body until cond
    DoWhile {
        body: Vec<Statement>,
        cond: Expression,
        /// Pascal `repeat..until` uses `until` (exit when true), not `while`.
        until: bool,
    },

    /// switch/case/evaluate
    Switch {
        expr: Expression,
        cases: Vec<SwitchCase>,
        default: Option<Vec<Statement>>,
    },

    /// try/catch/finally / try/except/finally / begin/rescue/ensure
    Try {
        body: Vec<Statement>,
        catches: Vec<CatchClause>,
        else_: Option<Vec<Statement>>,   // Python's `else` on try
        finally: Option<Vec<Statement>>,
    },

    // ── Jumps ─────────────────────────────────────────────────────────────

    Return(Option<Expression>),
    Break(Option<Expression>),      // Ruby `break value`
    Continue,
    Throw(Option<Expression>),

    /// `exit` (Pascal), `pass` (Python), no-op
    Exit(Option<Expression>),

    // ── Assignment ────────────────────────────────────────────────────────

    Assign {
        target: Expression,
        value: Expression,
    },

    /// `+=`, `-=`, `*=`, `/=`, etc.
    CompoundAssign {
        target: Expression,
        op: BinOp,
        value: Expression,
    },

    // ── Language-specific ─────────────────────────────────────────────────

    /// `with obj do` (Pascal), `with` (Python/JS), `using` (C#)
    With {
        expr: Expression,
        body: Vec<Statement>,
    },

    /// `raise`/`throw` with class creation: `raise Exception.Create('msg')`
    Raise(Option<Expression>),

    /// Empty statement / `pass` / `;`
    Empty,

    /// Anything that doesn't fit the common model.
    /// The string tag identifies it; the expressions/statements carry data.
    Extra {
        tag: String,
        exprs: Vec<Expression>,
        stmts: Vec<Statement>,
    },
}

// ── Expressions ──────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct Expression {
    pub kind: ExprKind,
    pub span: Span,
}

impl Expression {
    pub fn new(kind: ExprKind) -> Self { Self { kind, span: Span::default() } }
    pub fn with_span(kind: ExprKind, span: Span) -> Self { Self { kind, span } }
    pub fn ident(name: &str) -> Self { Self::new(ExprKind::Ident(name.to_string())) }
    pub fn int(n: i64) -> Self { Self::new(ExprKind::Lit(Literal::Int(n))) }
    pub fn float(n: f64) -> Self { Self::new(ExprKind::Lit(Literal::Float(n))) }
    pub fn string(s: &str) -> Self { Self::new(ExprKind::Lit(Literal::Str(s.to_string()))) }
    pub fn bool(b: bool) -> Self { Self::new(ExprKind::Lit(Literal::Bool(b))) }
    pub fn null() -> Self { Self::new(ExprKind::Lit(Literal::Null)) }
}

#[derive(Debug, Clone)]
pub enum ExprKind {
    /// Literal value.
    Lit(Literal),

    /// Variable / name reference.
    Ident(String),

    /// `this` / `self` / `Self` / `Me`
    This,

    /// `super` / `base` / `MyBase` / `parent`
    Super,

    /// Binary operation: `a + b`, `a == b`, `a and b`
    Binary { op: BinOp, left: Box<Expression>, right: Box<Expression> },

    /// Unary operation: `-x`, `not x`, `!x`
    Unary { op: UnaryOp, expr: Box<Expression> },

    /// Conditional / ternary: `cond ? then : else` / `if cond then x else y`
    Ternary { cond: Box<Expression>, then: Box<Expression>, else_: Box<Expression> },

    /// Function / method call: `f(args)`
    Call { callee: Box<Expression>, args: Vec<Expression> },

    /// Member access: `obj.field`, `obj?.field`
    Member { object: Box<Expression>, field: String, null_safe: bool },

    /// Index access: `arr[i]`, `dict[key]`
    Index { object: Box<Expression>, index: Box<Expression> },

    /// Object creation: `new Foo(args)` / `Foo.Create(args)` / `Foo(args)`
    New { class: Box<Expression>, args: Vec<Expression> },

    /// Assignment expression (languages where assignment is an expression).
    Assign { target: Box<Expression>, value: Box<Expression> },

    /// Lambda / anonymous function / closure.
    Lambda { params: Vec<Param>, body: Vec<Statement>, is_async: bool },

    /// Array literal: `[1, 2, 3]`
    Array(Vec<Expression>),

    /// Object/dict/hash literal: `{key: val, ...}`
    Object(Vec<(Expression, Expression)>),

    /// String interpolation: `f"hello {name}"` / `$"hello {name}"`
    Interpolation(Vec<InterpolPart>),

    /// Type check: `x is T` / `isinstance(x, T)`
    IsType { expr: Box<Expression>, type_name: String },

    /// Type cast: `x as T` / `(T)x` / `T(x)`
    AsCast { expr: Box<Expression>, type_name: String },

    /// Null coalescing: `x ?? default`
    NullCoalesce { left: Box<Expression>, right: Box<Expression> },

    /// Spread/splat: `...args` / `*args`
    Spread(Box<Expression>),

    /// Await: `await expr`
    Await(Box<Expression>),

    /// Yield: `yield expr` / `yield from expr`
    Yield(Option<Box<Expression>>),

    /// `inherited Create(args)` / `super(args)` / `base(args)`
    Inherited { method: Option<String>, args: Vec<Expression> },

    /// Range: `1..10` / `1...10`
    Range { start: Box<Expression>, end: Box<Expression>, inclusive: bool },

    /// Anything language-specific that doesn't fit.
    Extra { tag: String, exprs: Vec<Expression> },
}

// ── Literals ─────────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub enum Literal {
    Int(i64),
    Float(f64),
    Str(String),
    Bool(bool),
    Char(char),
    Null,
}

// ── Interpolation ────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub enum InterpolPart {
    Text(String),
    Expr(Expression),
}

// ── Parameters ───────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct Param {
    pub name: String,
    pub type_hint: Option<String>,
    pub default: Option<Expression>,
    pub pass_by: PassBy,
    pub is_rest: bool,     // `...args` / `*args` / `params`
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum PassBy { Value, Ref, Const, Out }

// ── Operators ────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum BinOp {
    // Arithmetic
    Add, Sub, Mul, Div, IDiv, Mod, Pow,
    // Comparison
    Eq, NotEq, Lt, Gt, Le, Ge,
    // Logical
    And, Or, Xor,
    // Bitwise
    BitAnd, BitOr, BitXor, Shl, Shr,
    // String
    Concat,
    // Membership
    In, NotIn,
    // Null coalescing (as binary op)
    NullCoalesce,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum UnaryOp {
    Neg, Pos,
    Not, BitNot,
    PreInc, PreDec,
    PostInc, PostDec,
    Typeof, Deref, AddrOf,
}

// ── Modifiers ────────────────────────────────────────────────────────────────

/// Modifiers applicable to declarations. Languages use different subsets.
#[derive(Debug, Clone, Default)]
pub struct Modifiers {
    pub visibility: Option<Visibility>,
    pub is_static: bool,
    pub is_abstract: bool,
    pub is_virtual: bool,
    pub is_override: bool,
    pub is_async: bool,
    pub is_readonly: bool,
    pub is_const: bool,
    pub decorators: Vec<String>,  // Python @decorators, Java annotations
    pub extra: Vec<String>,       // language-specific: "sealed", "partial", etc.
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Visibility { Public, Private, Protected, Internal }

// ── Switch / Case ────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct SwitchCase {
    pub values: Vec<Expression>,  // empty = default case
    pub body: Vec<Statement>,
}

// ── Catch / Except ───────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct CatchClause {
    pub type_name: Option<String>,
    pub var_name: Option<String>,
    pub body: Vec<Statement>,
}

// ── Enum members ─────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct EnumMember {
    pub name: String,
    pub value: Option<Expression>,
}
