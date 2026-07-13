//! Common AST — the language-neutral IR every walker produces and the
//! compiler/emitters consume.
//!
//! Maps to emitter modules:
//!   classes    → ClassDecl, ClassMember, Property
//!   functions  → FunctionDecl, Param, Lambda
//!   collections → Array, ForIn, destructuring
//!   dict       → Object literals
//!   loops      → For, ForIn, While, DoWhile
//!   errors     → Try, CatchClause, Throw
//!   expressions → Binary, Unary, Ternary, NullCoalesce
//!   strings    → builtins (profile-driven)
//!   math       → builtins (profile-driven)
//!   io         → Echo, builtins (profile-driven)
//!   threading  → async/await
//!   dotnet     → namespace resolution (profile-driven)

// ════════════════════════════════════════════════════════════════════════════
// Module (top-level compilation unit)
// ════════════════════════════════════════════════════════════════════════════

#[derive(Debug, Clone)]
pub struct Module {
    pub name: String,
    pub language: Lang,
    pub body: Vec<Statement>,
    pub imports: Vec<Import>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Lang {
    VB,
    JavaScript,
    CSharp,
    Python,
    Ruby,
    PHP,
    Dart,
    Pascal,
    Cobol,
    Fortran,
    Go,
    Lua,
    Java,
    Unknown,
}

#[derive(Debug, Clone)]
pub struct Import {
    pub kind: ImportKind,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub enum ImportKind {
    /// `Imports System.IO` / `import os` / `use Foo\Bar`
    Simple { path: String, alias: Option<String> },
    /// `from os import path, getcwd` / `import { x, y } from "mod"`
    /// Python: level > 0 for relative imports (`from ...pkg import x` → level=3)
    Named {
        path: String,
        names: Vec<ImportName>,
        level: usize,
    },
    /// `import * as ns from "mod"` / `from os import *`
    Wildcard { path: String, alias: Option<String> },
    /// `import defaultExport from "mod"` (JS)
    Default { path: String, local: String },
}

#[derive(Debug, Clone)]
pub struct ImportName {
    pub name: String,
    pub alias: Option<String>,
}

#[derive(Debug, Clone, Copy, Default)]
pub struct Span {
    pub start_line: u32,
    pub start_col: u32,
    pub end_line: u32,
    pub end_col: u32,
}

// ════════════════════════════════════════════════════════════════════════════
// Statements
// ════════════════════════════════════════════════════════════════════════════

#[derive(Debug, Clone)]
pub struct Statement {
    pub kind: StmtKind,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub struct RecordFieldFormat {
    pub decimal_places: usize,
}

impl Statement {
    pub fn new(kind: StmtKind) -> Self {
        Self {
            kind,
            span: Span::default(),
        }
    }
    pub fn with_span(kind: StmtKind, span: Span) -> Self {
        Self { kind, span }
    }
}

#[derive(Debug, Clone)]
pub enum StmtKind {
    /// Expression used as a statement.
    Expr(Expression),

    /// Block of statements.
    Block(Vec<Statement>),

    // ── Variable declarations ────────────────────────────────────────────
    VarDecl {
        declarations: Vec<VarDeclarator>,
        kind: VarDeclKind,
    },

    // ── Functions (compiler_common::functions) ───────────────────────────
    FunctionDecl {
        name: String,
        params: Vec<Param>,
        return_type: Option<String>,
        body: Vec<Statement>,
        modifiers: Modifiers,
        handles: Vec<String>,
        is_async: bool,
        /// Dart: `sync*`/`async*` generator functions
        is_generator: bool,
        is_sub: bool,
    },

    // ── Classes (compiler_common::classes) ───────────────────────────────
    ClassDecl {
        name: String,
        /// VB: single parent. Python: multiple bases. JS: single extends (as expression).
        parents: Vec<String>,
        interfaces: Vec<String>,
        members: Vec<ClassMember>,
        modifiers: ClassModifiers,
        decorators: Vec<Expression>,
    },

    InterfaceDecl {
        name: String,
        parents: Vec<String>,
        members: Vec<InterfaceMember>,
        decorators: Vec<Expression>,
    },

    EnumDecl {
        name: String,
        members: Vec<EnumMember>,
        visibility: Visibility,
        is_flags: bool,
        backing_type: Option<String>,
        interfaces: Vec<String>,
        body_members: Vec<ClassMember>,
        decorators: Vec<Expression>,
    },

    StructDecl {
        name: String,
        interfaces: Vec<String>,
        members: Vec<ClassMember>,
        visibility: Visibility,
        decorators: Vec<Expression>,
    },

    ModuleDecl {
        name: String,
        members: Vec<ClassMember>,
        visibility: Visibility,
    },

    NamespaceDecl {
        name: String,
        body: Vec<Statement>,
    },

    DelegateDecl {
        name: String,
        params: Vec<Param>,
        return_type: Option<String>,
        is_sub: bool,
        visibility: Visibility,
    },

    // ── Control flow (br_if, br, loop opcodes) ──────────────────────────
    If {
        cond: Expression,
        then_body: Vec<Statement>,
        elifs: Vec<(Expression, Vec<Statement>)>,
        else_body: Option<Vec<Statement>>,
    },

    For {
        init: Option<Box<Statement>>,
        cond: Option<Expression>,
        update: Option<Expression>,
        body: Vec<Statement>,
    },

    /// compiler_common::loops::emit_for_in
    ForIn {
        var: String,
        /// PHP: `foreach ($arr as $key => $value)` — key variable
        key: Option<String>,
        iter: Expression,
        body: Vec<Statement>,
        /// JS for-of vs for-in
        of: bool,
        /// Python: `for x in items: ... else: ...`
        else_body: Option<Vec<Statement>>,
        /// Python: `async for`
        is_async: bool,
    },

    While {
        cond: Expression,
        body: Vec<Statement>,
        /// Python: `while cond: ... else: ...`
        else_body: Option<Vec<Statement>>,
    },

    DoWhile {
        body: Vec<Statement>,
        cond: Expression,
        until: bool,
    },

    /// Switch / Select Case. Cases are in source order. A case with
    /// empty `conditions` is the default arm. The separate `default`
    /// field is kept for backward compatibility with walkers that
    /// still split it out — the compiler merges both.
    Switch {
        expr: Expression,
        cases: Vec<SwitchCase>,
        default: Option<Vec<Statement>>,
    },

    // ── Exception handling (compiler_common::errors) ─────────────────────
    Try {
        body: Vec<Statement>,
        catches: Vec<CatchClause>,
        else_body: Option<Vec<Statement>>,
        finally: Option<Vec<Statement>>,
    },

    // ── With / Using / Lock ─────────────────────────────────────────────
    /// Python: `with a() as x, b() as y:` — multiple items
    /// VB: `With obj ... End With` — single item, no var
    With {
        items: Vec<WithItem>,
        body: Vec<Statement>,
        is_async: bool,
    },

    Using {
        var: String,
        resource: Expression,
        body: Vec<Statement>,
    },

    Lock {
        expr: Expression,
        body: Vec<Statement>,
    },

    // ── Jumps ────────────────────────────────────────────────────────────
    Return(Option<Expression>),
    Break(BreakTarget),
    Continue(ContinueTarget),
    /// Python: `raise X from Y` — cause is the chained exception
    Throw {
        expr: Option<Expression>,
        cause: Option<Expression>,
    },

    // ── Assignment ───────────────────────────────────────────────────────
    /// Python: `a = b = c = 5` — multiple targets
    Assign {
        targets: Vec<Expression>,
        value: Expression,
    },

    CompoundAssign {
        target: Expression,
        op: CompoundOp,
        value: Expression,
    },

    // ── Array operations ─────────────────────────────────────────────────
    ReDim {
        preserve: bool,
        array: String,
        bounds: Vec<Expression>,
    },

    Erase {
        array: String,
    },

    // ── Events ───────────────────────────────────────────────────────────
    //
    // Canonical event binding. Every language frontend produces this for its
    // own surface syntax:
    //   VB:     `Sub Click_Btn() Handles btn1.Click` (injected at end of New)
    //           `AddHandler btn1.Click, AddressOf Click_Btn`
    //   C#:     `btn1.Click += Click_Btn;`
    //   JS:     `btn1.addEventListener("click", click_btn)`
    //   Python: `btn1.bind("<Button-1>", click_btn)`
    //
    // The compiler routes these through `compiler_common::gui::emit_bind_event`
    // so the bytecode is identical regardless of source language.
    AddHandler {
        /// Control instance — any expression that evaluates to a GUI object.
        control: Expression,
        /// Event name (e.g. "click", "textchanged") — already lowercased.
        event: String,
        /// Handler — any expression that evaluates to a function reference
        /// (a global function, a method on `me`, a lambda, etc.).
        handler: Expression,
    },

    RemoveHandler {
        control: Expression,
        event: String,
        handler: Expression,
    },

    RaiseEvent {
        event_name: String,
        args: Vec<Expression>,
    },

    // ── Error handling (VB6 legacy) ──────────────────────────────────────
    OnErrorResumeNext,
    OnErrorGoTo(String),
    GoTo(String),
    Label(String),

    // ── File I/O (VB6 legacy) ────────────────────────────────────────────
    OpenFile {
        path: Expression,
        mode: FileMode,
        file_number: Expression,
    },
    CloseFile(Option<Expression>),
    PrintFile {
        file_number: Expression,
        items: Vec<Expression>,
    },
    WriteFile {
        file_number: Expression,
        items: Vec<Expression>,
    },
    InputFile {
        file_number: Expression,
        variables: Vec<Expression>,
    },
    LineInput {
        file_number: Expression,
        variable: String,
    },
    StartFile {
        file_number: Expression,
        key_index: usize,
        key_value: Expression,
        relation: FileKeyRelation,
    },
    InputRecordFile {
        file_number: Expression,
        variables: Vec<String>,
        key_index: Option<usize>,
        key_value: Option<Expression>,
    },
    RewriteRecordFile {
        file_number: Expression,
        items: Vec<Expression>,
        field_formats: Vec<Option<RecordFieldFormat>>,
    },

    // ── Module system (JS) ───────────────────────────────────────────────
    Export {
        declaration: Option<Box<Statement>>,
        names: Vec<ExportName>,
        default: Option<Box<Expression>>,
        /// Re-export source module. `None` for local exports. `Some(path)`
        /// for `export { X } from "m"` and `export * from "m"` — the
        /// Linker resolves these as Indirect exports per ECMA-262
        /// §16.2.1.6.2 `ResolveExport`.
        from: Option<String>,
        /// `true` for `export * from "m"` (all re-exported under source
        /// names). `false` for named re-exports or local exports.
        star: bool,
    },

    // ── Labeled statement (JS) ──────────────────────────────────────────
    /// `myLabel: for (...) {}` — wraps a statement with a label for break/continue
    Labeled {
        label: String,
        body: Box<Statement>,
    },

    // ── Other ────────────────────────────────────────────────────────────
    Empty,
    ScopeDecl {
        kind: ScopeDeclKind,
        names: Vec<String>,
    },
    Delete(Vec<Expression>),
    Assert {
        test: Expression,
        msg: Option<Expression>,
    },
    Echo(Vec<Expression>),

    /// Python `match subject: case pattern: ...` — statement-level with full patterns
    MatchStatement {
        subject: Expression,
        cases: Vec<MatchCase>,
    },

    // ── WASM module structure (linear memory / data segments) ────────────
    //
    // WAT is the raw WASM module format, so it is the one frontend that
    // declares these directly. Higher-level languages reach linear memory
    // only indirectly — heap allocation lowers to it — so they never emit
    // these nodes. The compiler lowers them into the script chunk's memory /
    // data tables, which the VM instantiates (allocates pages, writes active
    // data) before `_start` runs, per the WASM spec.
    /// `(memory $id? min max?)` — a linear-memory declaration. Declaration
    /// order is the memory index.
    MemoryDecl {
        /// Minimum size in 64 KiB pages.
        min_pages: u64,
        /// Maximum size in pages; `None` = unbounded.
        max_pages: Option<u64>,
    },

    /// `(table $id? min max? funcref)` — a reference-table declaration.
    /// Declaration order is the table index.
    TableDecl {
        /// Minimum element count.
        min_size: u64,
        /// Maximum element count; `None` = unbounded.
        max_size: Option<u64>,
    },

    /// `(data (offset) "bytes")` (active) or `(data "bytes")` (passive). Active
    /// segments are copied into linear memory at instantiation; passive ones
    /// are held for `memory.init`.
    DataSegment {
        /// Target memory index for active segments (0 by default).
        memory_index: u32,
        /// Constant offset expression (e.g. `i32.const N`) for active segments;
        /// `None` marks a passive segment.
        offset: Option<Expression>,
        /// Initializer bytes, with WAT string escapes already decoded.
        bytes: Vec<u8>,
    },
}

// ════════════════════════════════════════════════════════════════════════════
// Class members (compiler_common::classes)
// ════════════════════════════════════════════════════════════════════════════

#[derive(Debug, Clone)]
pub enum ClassMember {
    Field {
        name: String,
        type_hint: Option<String>,
        init: Option<Expression>,
        modifiers: Modifiers,
        with_events: bool,
        array_bounds: Option<Vec<Expression>>,
    },

    Method(Box<Statement>),

    Constructor {
        params: Vec<Param>,
        body: Vec<Statement>,
        base_args: Option<Vec<Expression>>,
        initializer_target: ConstructorInitializerTarget,
        visibility: Visibility,
    },

    /// __get_ / __set_ closures on the struct
    Property {
        name: String,
        type_hint: Option<String>,
        getter: Option<Vec<Statement>>,
        setter: Option<PropertySetter>,
        is_auto: bool,
        modifiers: Modifiers,
    },

    Event {
        name: String,
        type_hint: Option<String>,
        params: Vec<Param>,
        visibility: Visibility,
    },

    Const {
        name: String,
        type_hint: Option<String>,
        value: Expression,
        visibility: Visibility,
    },

    NestedType(Box<Statement>),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ConstructorInitializerTarget {
    Base,
    This,
}

#[derive(Debug, Clone)]
pub struct PropertySetter {
    pub param: Param,
    pub body: Vec<Statement>,
}

// ════════════════════════════════════════════════════════════════════════════
// Interface members
// ════════════════════════════════════════════════════════════════════════════

#[derive(Debug, Clone)]
pub enum InterfaceMember {
    Method {
        name: String,
        params: Vec<Param>,
        return_type: Option<String>,
        is_sub: bool,
        signature_source: Option<String>,
    },
    Property {
        name: String,
        type_hint: Option<String>,
        is_readonly: bool,
        is_writeonly: bool,
    },
    Event {
        name: String,
        type_hint: Option<String>,
    },
}

// ════════════════════════════════════════════════════════════════════════════
// Expressions
// ════════════════════════════════════════════════════════════════════════════

#[derive(Debug, Clone)]
pub struct Expression {
    pub kind: ExprKind,
    pub span: Span,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ArrayIndexSemantics {
    pub first_index: i64,
}

impl ArrayIndexSemantics {
    pub const ZERO_BASED: Self = Self { first_index: 0 };
    pub const ONE_BASED: Self = Self { first_index: 1 };
}

impl Expression {
    pub fn new(kind: ExprKind) -> Self {
        Self {
            kind,
            span: Span::default(),
        }
    }
    pub fn with_span(kind: ExprKind, span: Span) -> Self {
        Self { kind, span }
    }
    pub fn ident(name: &str) -> Self {
        Self::new(ExprKind::Ident(name.to_string()))
    }
    pub fn int(n: i64) -> Self {
        Self::new(ExprKind::Lit(Literal::Int(n)))
    }
    pub fn float(n: f64) -> Self {
        Self::new(ExprKind::Lit(Literal::Float(n)))
    }
    pub fn string(s: &str) -> Self {
        Self::new(ExprKind::Lit(Literal::Str(s.to_string())))
    }
    pub fn bool(b: bool) -> Self {
        Self::new(ExprKind::Lit(Literal::Bool(b)))
    }
    pub fn null() -> Self {
        Self::new(ExprKind::Lit(Literal::Null))
    }
}

#[derive(Debug, Clone)]
pub enum PlaceExpr {
    Ident(String),
    Member {
        object: Box<Expression>,
        field: String,
        null_safe: bool,
    },
    Index {
        object: Box<Expression>,
        index: Box<Expression>,
        null_safe: bool,
    },
    Deref(Box<Expression>),
}

#[derive(Debug, Clone)]
pub enum ExprKind {
    Lit(Literal),
    Ident(String),
    This,
    Super,

    // ── Operators (compiler_common::expressions) ─────────────────────────
    Binary {
        op: BinOp,
        left: Box<Expression>,
        right: Box<Expression>,
    },
    Unary {
        op: UnaryOp,
        expr: Box<Expression>,
    },
    RefOf(Box<PlaceExpr>),
    RefLoad(Box<Expression>),
    Ternary {
        cond: Box<Expression>,
        then: Box<Expression>,
        else_: Box<Expression>,
    },

    // ── Access (struct_get / struct_set) ─────────────────────────────────
    Member {
        object: Box<Expression>,
        field: String,
        null_safe: bool,
    },
    Index {
        object: Box<Expression>,
        index: Box<Expression>,
        null_safe: bool,
    },

    // ── Calls (call opcode) ─────────────────────────────────────────────
    Call {
        callee: Box<Expression>,
        args: Vec<Argument>,
        optional: bool,
    },

    // ── Object creation (compiler_common::classes) ──────────────────────
    New {
        class: Box<Expression>,
        args: Vec<Argument>,
    },

    // ── Assignment as expression ─────────────────────────────────────────
    Assign {
        target: Box<Expression>,
        value: Box<Expression>,
    },

    // ── Functions as values (ref_func) ──────────────────────────────────
    Lambda {
        params: Vec<Param>,
        body: LambdaBody,
        is_async: bool,
        /// PHP: `function() use ($x, $y) { }` — explicitly captured variables
        captures: Vec<String>,
    },

    // ── Collections (compiler_common::collections, dict) ────────────────
    Array(Vec<ArrayElement>),
    /// Python `(1, 2, 3)` — immutable sequence
    Tuple(Vec<Expression>),
    /// Python `{1, 2, 3}` — unordered unique collection
    Set(Vec<Expression>),
    Object(Vec<ObjectProperty>),

    // ── String interpolation ─────────────────────────────────────────────
    Interpolation(Vec<InterpolPart>),

    // ── Type operations ──────────────────────────────────────────────────
    IsType {
        expr: Box<Expression>,
        type_name: String,
    },
    Cast {
        expr: Box<Expression>,
        type_name: String,
    },
    TypeOf(Box<Expression>),
    DefaultOf(String),

    // ── Null handling (compiler_common::expressions) ─────────────────────
    NullCoalesce {
        left: Box<Expression>,
        right: Box<Expression>,
    },

    // ── Spread / rest ────────────────────────────────────────────────────
    Spread(Box<Expression>),

    // ── Async (compiler_common::functions) ───────────────────────────────
    Await(Box<Expression>),
    Yield(Option<Box<Expression>>),
    YieldFrom(Box<Expression>),

    // ── VB / .NET ────────────────────────────────────────────────────────
    AddressOf(String),
    SuperCall {
        method: Option<String>,
        args: Vec<Argument>,
    },

    // ── Python ───────────────────────────────────────────────────────────
    Comprehension {
        kind: ComprehensionKind,
        element: Box<Expression>,
        generators: Vec<ComprehensionGen>,
    },
    Slice {
        lower: Option<Box<Expression>>,
        upper: Option<Box<Expression>>,
        step: Option<Box<Expression>>,
    },
    Walrus {
        target: Box<Expression>,
        value: Box<Expression>,
    },

    // ── JS ───────────────────────────────────────────────────────────────
    Void(Box<Expression>),
    Delete(Box<Expression>),
    Destructure(DestructurePattern),
    /// `(a, b, c)` — evaluates all, result is last value
    Sequence(Vec<Expression>),
    /// `class { ... }` as an expression: `let C = class { ... }`
    ClassExpr {
        name: Option<String>,
        parent: Option<Box<Expression>>,
        /// Interfaces the (anonymous) class implements. Only PHP populates
        /// this today (`new class implements I {}`); other languages leave it
        /// empty, matching the previous hardcoded `&[]`.
        interfaces: Vec<String>,
        members: Vec<ClassMember>,
    },
    /// `function(...) { }` as an expression: `let f = function() { ... }`
    FunctionExpr(Box<Statement>), // always StmtKind::FunctionDecl

    // ── Range ────────────────────────────────────────────────────────────
    Range {
        start: Box<Expression>,
        end: Box<Expression>,
        inclusive: bool,
    },

    // ── PHP ──────────────────────────────────────────────────────────────
    StaticAccess {
        class: Box<Expression>,
        member: Box<Expression>,
    },

    // ── Match expression (PHP/Python) ────────────────────────────────────
    Match {
        subject: Box<Expression>,
        arms: Vec<MatchArm>,
    },
}

// ════════════════════════════════════════════════════════════════════════════
// Supporting types
// ════════════════════════════════════════════════════════════════════════════

#[derive(Debug, Clone)]
pub enum Literal {
    Int(i64),
    Float(f64),
    BigInt(i64),
    Str(String),
    Bool(bool),
    Char(char),
    Null,
    Undefined,
    /// Python `...`
    Ellipsis,
}

// ── Variables ───────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct VarDeclarator {
    pub pattern: BindingPattern,
    pub type_hint: Option<String>,
    pub init: Option<Expression>,
    pub array_bounds: Option<Vec<Expression>>,
    pub with_events: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub enum VarDeclKind {
    Dim,
    Let,
    Const,
    Var,
    Static,
}

#[derive(Debug, Clone)]
pub enum BindingPattern {
    Ident(String),
    Object(Vec<ObjectPatternProp>),
    Array(Vec<ArrayPatternElem>),
}

#[derive(Debug, Clone)]
pub struct ObjectPatternProp {
    pub key: String,
    pub value: Option<BindingPattern>,
    pub default: Option<Expression>,
    pub is_rest: bool,
}

#[derive(Debug, Clone)]
pub enum ArrayPatternElem {
    Pattern(BindingPattern, Option<Expression>),
    Rest(String),
    Hole,
}

// ── Parameters ──────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct Param {
    pub name: String,
    pub type_hint: Option<String>,
    pub default: Option<Expression>,
    pub pass_by: PassBy,
    pub is_rest: bool,
    pub is_kwargs: bool,
    pub is_optional: bool,
    pub is_nullable: bool,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum PassBy {
    Value,
    Ref,
    Out,
    Const,
}

// ── Arguments ───────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct Argument {
    pub value: Expression,
    pub name: Option<String>,
    pub by_ref: bool,
    pub spread: bool,
}

impl Argument {
    pub fn positional(value: Expression) -> Self {
        Self {
            value,
            name: None,
            by_ref: false,
            spread: false,
        }
    }
}

// ── Lambda body ─────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub enum LambdaBody {
    Expr(Box<Expression>),
    Block(Vec<Statement>),
}

// ── Array/Object literals ───────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct ArrayElement {
    /// PHP: `"key" => value` — associative array key
    pub key: Option<Expression>,
    pub value: Expression,
    pub spread: bool,
    /// PHP: `&$var` — by-reference element
    pub by_ref: bool,
}

#[derive(Debug, Clone)]
pub enum ObjectProperty {
    KeyValue {
        key: Expression,
        value: Expression,
    },
    Shorthand(String),
    Spread(Expression),
    Method {
        key: String,
        value: Box<Statement>,
    },
    Accessor {
        kind: AccessorKind,
        key: String,
        value: Box<Statement>,
    },
    Computed {
        key: Expression,
        value: Expression,
    },
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum AccessorKind {
    Get,
    Set,
}

// ── Interpolation ───────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub enum InterpolPart {
    Text(String),
    Expr(Expression),
    Formatted(Expression, String),
}

// ── Switch / Case ───────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct SwitchCase {
    pub conditions: Vec<CaseCondition>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub enum CaseCondition {
    Value(Expression),
    Range { from: Expression, to: Expression },
    Comparison { op: ComparisonOp, expr: Expression },
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ComparisonOp {
    Eq,
    NotEq,
    Lt,
    LtEq,
    Gt,
    GtEq,
}

// ── Catch ───────────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct CatchClause {
    pub types: Vec<String>,
    pub var_name: Option<String>,
    /// Dart: `catch (e, stackTrace)` — second variable for stack trace
    pub stack_var: Option<String>,
    pub body: Vec<Statement>,
    pub when_clause: Option<Expression>,
}

// ── Match (PHP expression-level) ─────────────────────────────────────────────

/// PHP: `match($x) { val1, val2 => expr, default => expr }`
#[derive(Debug, Clone)]
pub struct MatchArm {
    /// None = default arm
    pub conditions: Option<Vec<Expression>>,
    pub body: Expression,
}

// ── Match (Python statement-level with patterns) ─────────────────────────────

#[derive(Debug, Clone)]
pub struct MatchCase {
    pub pattern: Pattern,
    pub guard: Option<Expression>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub enum Pattern {
    /// `case 42:` / `case "hello":`
    Value(Expression),
    /// `case None:` / `case True:`
    Singleton(Expression),
    /// `case [a, b, c]:`
    Sequence(Vec<Pattern>),
    /// `case {"key": value}:`
    Mapping(Vec<(Expression, Pattern)>),
    /// `case MyClass(x, y):` / `case Point(x=1, y=2):`
    Class {
        cls: Expression,
        patterns: Vec<Pattern>,
        kw_patterns: Vec<(String, Pattern)>,
    },
    /// `case [first, *rest]:`
    Star(Option<String>),
    /// `case pattern as name:`
    As {
        pattern: Option<Box<Pattern>>,
        name: Option<String>,
    },
    /// `case a | b | c:`
    Or(Vec<Pattern>),
    /// `case _:`
    Wildcard,
}

// ── With items ──────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct WithItem {
    pub expr: Expression,
    pub var: Option<String>,
}

// ── Comprehension ───────────────────────────────────────────────────────────

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ComprehensionKind {
    List,
    Set,
    Dict,
    Generator,
}

#[derive(Debug, Clone)]
pub struct ComprehensionGen {
    pub target: Expression,
    pub iter: Expression,
    pub conditions: Vec<Expression>,
    pub is_async: bool,
}

// ── Destructuring ───────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub enum DestructurePattern {
    Object(Vec<ObjectPatternProp>),
    Array(Vec<ArrayPatternElem>),
}

// ── Break/Continue targets ──────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub enum BreakTarget {
    Implicit,
    Label(String),
    Kind(ExitKind),
    Value(Expression),
    /// PHP: `break 2;` — skip N levels
    Level(u32),
}

#[derive(Debug, Clone)]
pub enum ContinueTarget {
    Implicit,
    Label(String),
    Kind(ContinueKind),
    /// PHP: `continue 2;` — skip N levels
    Level(u32),
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ExitKind {
    Sub,
    Function,
    For,
    Do,
    While,
    Select,
    Try,
    Property,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ContinueKind {
    Do,
    For,
    While,
}

// ── Misc enums ──────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ScopeDeclKind {
    Global,
    Nonlocal,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum FileMode {
    Input,
    Output,
    Append,
    Binary,
    Random,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum FileKeyRelation {
    Equal,
    Greater,
    GreaterOrEqual,
    Less,
    LessOrEqual,
}

#[derive(Debug, Clone)]
pub struct ExportName {
    pub name: String,
    pub alias: Option<String>,
}

// ════════════════════════════════════════════════════════════════════════════
// Operators
// ════════════════════════════════════════════════════════════════════════════

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum BinOp {
    Add,
    Sub,
    Mul,
    Div,
    IDiv,
    Mod,
    Pow,
    Eq,
    NotEq,
    StrictEq,
    StrictNotEq,
    Lt,
    Gt,
    LtEq,
    GtEq,
    Spaceship,
    And,
    Or,
    Xor,
    Eqv,
    Imp,
    BitAnd,
    BitOr,
    BitXor,
    Shl,
    Shr,
    UShr,
    Concat,
    In,
    NotIn,
    InstanceOf,
    NullCoalesce,
    MatMul,
    FloorDiv,
    Like,
    Is,
    IsNot,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum UnaryOp {
    Neg,
    Pos,
    Not,
    BitNot,
    PreInc,
    PreDec,
    PostInc,
    PostDec,
    Typeof,
    Void,
    Delete,
    Deref,
    AddrOf,
    Await,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum CompoundOp {
    Add,
    Sub,
    Mul,
    Div,
    IDiv,
    Mod,
    Pow,
    Concat,
    BitAnd,
    BitOr,
    BitXor,
    Shl,
    Shr,
    UShr,
    And,
    Or,
    NullCoalesce,
}

// ════════════════════════════════════════════════════════════════════════════
// Modifiers
// ════════════════════════════════════════════════════════════════════════════

#[derive(Debug, Clone, Default)]
pub struct Modifiers {
    pub visibility: Visibility,
    pub is_static: bool,
    pub is_abstract: bool,
    pub is_virtual: bool,
    pub is_override: bool,
    pub is_readonly: bool,
    pub is_shared: bool,
    pub is_extension: bool,
    pub is_overloads: bool,
    pub is_not_overridable: bool,
    pub decorators: Vec<Expression>,
}

#[derive(Debug, Clone, Default)]
pub struct ClassModifiers {
    pub visibility: Visibility,
    pub is_partial: bool,
    pub is_abstract: bool,
    pub is_sealed: bool,
    pub is_static: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum Visibility {
    #[default]
    Public,
    Private,
    Protected,
    Internal,
}

// ════════════════════════════════════════════════════════════════════════════
// Enum members
// ════════════════════════════════════════════════════════════════════════════

#[derive(Debug, Clone)]
pub struct EnumMember {
    pub name: String,
    pub value: Option<Expression>,
    pub constructor_args: Vec<Expression>,
}

pub fn body_has_yield(stmts: &[Statement]) -> bool {
    stmts.iter().any(statement_has_yield)
}

fn statement_has_yield(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::FunctionDecl { .. }
        | StmtKind::ClassDecl { .. }
        | StmtKind::InterfaceDecl { .. }
        | StmtKind::EnumDecl { .. }
        | StmtKind::StructDecl { .. }
        | StmtKind::ModuleDecl { .. }
        | StmtKind::NamespaceDecl { .. }
        | StmtKind::DelegateDecl { .. } => false,
        StmtKind::Expr(expr) => expr_has_yield(expr),
        StmtKind::Block(body) => body_has_yield(body),
        StmtKind::VarDecl { declarations, .. } => declarations
            .iter()
            .any(|decl| decl.init.as_ref().map_or(false, expr_has_yield)),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            expr_has_yield(cond)
                || body_has_yield(then_body)
                || elifs
                    .iter()
                    .any(|(cond, body)| expr_has_yield(cond) || body_has_yield(body))
                || else_body
                    .as_ref()
                    .map_or(false, |body| body_has_yield(body))
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            init.as_ref()
                .map_or(false, |stmt| statement_has_yield(stmt))
                || cond.as_ref().map_or(false, expr_has_yield)
                || update.as_ref().map_or(false, expr_has_yield)
                || body_has_yield(body)
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            expr_has_yield(iter)
                || body_has_yield(body)
                || else_body
                    .as_ref()
                    .map_or(false, |body| body_has_yield(body))
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            expr_has_yield(cond)
                || body_has_yield(body)
                || else_body
                    .as_ref()
                    .map_or(false, |body| body_has_yield(body))
        }
        StmtKind::DoWhile { body, cond, .. } => body_has_yield(body) || expr_has_yield(cond),
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            expr_has_yield(expr)
                || cases.iter().any(|case| {
                    case.conditions.iter().any(case_condition_has_yield)
                        || body_has_yield(&case.body)
                })
                || default.as_ref().map_or(false, |body| body_has_yield(body))
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            body_has_yield(body)
                || catches.iter().any(|catch| {
                    catch.when_clause.as_ref().map_or(false, expr_has_yield)
                        || body_has_yield(&catch.body)
                })
                || else_body
                    .as_ref()
                    .map_or(false, |body| body_has_yield(body))
                || finally.as_ref().map_or(false, |body| body_has_yield(body))
        }
        StmtKind::With { items, body, .. } => {
            items.iter().any(|item| expr_has_yield(&item.expr)) || body_has_yield(body)
        }
        StmtKind::Using { resource, body, .. } => expr_has_yield(resource) || body_has_yield(body),
        StmtKind::Lock { expr, body } => expr_has_yield(expr) || body_has_yield(body),
        StmtKind::Return(expr) => expr.as_ref().map_or(false, expr_has_yield),
        StmtKind::Throw { expr, cause } => {
            expr.as_ref().map_or(false, expr_has_yield)
                || cause.as_ref().map_or(false, expr_has_yield)
        }
        StmtKind::Assign { targets, value } => {
            targets.iter().any(expr_has_yield) || expr_has_yield(value)
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            expr_has_yield(target) || expr_has_yield(value)
        }
        StmtKind::AddHandler {
            control, handler, ..
        }
        | StmtKind::RemoveHandler {
            control, handler, ..
        } => expr_has_yield(control) || expr_has_yield(handler),
        StmtKind::RaiseEvent { args, .. } | StmtKind::Delete(args) | StmtKind::Echo(args) => {
            args.iter().any(expr_has_yield)
        }
        StmtKind::OpenFile {
            path, file_number, ..
        } => expr_has_yield(path) || expr_has_yield(file_number),
        StmtKind::CloseFile(expr) => expr.as_ref().map_or(false, expr_has_yield),
        StmtKind::PrintFile { file_number, items } | StmtKind::WriteFile { file_number, items } => {
            expr_has_yield(file_number) || items.iter().any(expr_has_yield)
        }
        StmtKind::StartFile {
            file_number,
            key_value,
            ..
        } => expr_has_yield(file_number) || expr_has_yield(key_value),
        StmtKind::InputRecordFile {
            file_number,
            key_value,
            ..
        } => expr_has_yield(file_number) || key_value.as_ref().map_or(false, expr_has_yield),
        StmtKind::RewriteRecordFile {
            file_number, items, ..
        } => expr_has_yield(file_number) || items.iter().any(expr_has_yield),
        StmtKind::ReDim { bounds, .. } => bounds.iter().any(expr_has_yield),
        StmtKind::Export {
            declaration,
            default,
            ..
        } => {
            declaration
                .as_ref()
                .map_or(false, |stmt| statement_has_yield(stmt))
                || default.as_ref().map_or(false, |expr| expr_has_yield(expr))
        }
        StmtKind::Labeled { body, .. } => statement_has_yield(body),
        StmtKind::Assert { test, msg } => {
            expr_has_yield(test) || msg.as_ref().map_or(false, expr_has_yield)
        }
        StmtKind::MatchStatement { subject, cases } => {
            expr_has_yield(subject)
                || cases.iter().any(|case| {
                    case.guard.as_ref().map_or(false, expr_has_yield) || body_has_yield(&case.body)
                })
        }
        _ => false,
    }
}

fn expr_has_yield(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Yield(_) | ExprKind::YieldFrom(_) => true,
        ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_) | ExprKind::ClassExpr { .. } => false,
        ExprKind::Unary { expr, .. }
        | ExprKind::IsType { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::TypeOf(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Await(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr) => expr_has_yield(expr),
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Assign {
            target: left,
            value: right,
        }
        | ExprKind::Walrus {
            target: left,
            value: right,
        }
        | ExprKind::Range {
            start: left,
            end: right,
            ..
        } => expr_has_yield(left) || expr_has_yield(right),
        ExprKind::StaticAccess { class, member } => expr_has_yield(class) || expr_has_yield(member),
        ExprKind::Ternary { cond, then, else_ } => {
            expr_has_yield(cond) || expr_has_yield(then) || expr_has_yield(else_)
        }
        ExprKind::Member { object, .. } => expr_has_yield(object),
        ExprKind::Index { object, index, .. } => expr_has_yield(object) || expr_has_yield(index),
        ExprKind::Call { callee, args, .. } => {
            expr_has_yield(callee) || args.iter().any(|arg| expr_has_yield(&arg.value))
        }
        ExprKind::New { class, args } => {
            expr_has_yield(class) || args.iter().any(|arg| expr_has_yield(&arg.value))
        }
        ExprKind::SuperCall { args, .. } => args.iter().any(|arg| expr_has_yield(&arg.value)),
        ExprKind::Array(items) => items.iter().any(|item| {
            item.key.as_ref().map_or(false, expr_has_yield) || expr_has_yield(&item.value)
        }),
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            items.iter().any(expr_has_yield)
        }
        ExprKind::Object(props) => props.iter().any(|prop| match prop {
            ObjectProperty::KeyValue { key, value } | ObjectProperty::Computed { key, value } => {
                expr_has_yield(key) || expr_has_yield(value)
            }
            ObjectProperty::Spread(expr) => expr_has_yield(expr),
            _ => false,
        }),
        ExprKind::Interpolation(parts) => parts.iter().any(|part| match part {
            InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => expr_has_yield(expr),
            InterpolPart::Text(_) => false,
        }),
        ExprKind::Match { subject, arms } => {
            expr_has_yield(subject)
                || arms.iter().any(|arm| {
                    arm.conditions
                        .as_ref()
                        .map_or(false, |conditions| conditions.iter().any(expr_has_yield))
                        || expr_has_yield(&arm.body)
                })
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            expr_has_yield(element)
                || generators.iter().any(|generator| {
                    expr_has_yield(&generator.target)
                        || expr_has_yield(&generator.iter)
                        || generator.conditions.iter().any(expr_has_yield)
                })
        }
        ExprKind::Slice { lower, upper, step } => {
            lower.as_ref().map_or(false, |expr| expr_has_yield(expr))
                || upper.as_ref().map_or(false, |expr| expr_has_yield(expr))
                || step.as_ref().map_or(false, |expr| expr_has_yield(expr))
        }
        _ => false,
    }
}

fn case_condition_has_yield(condition: &CaseCondition) -> bool {
    match condition {
        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => expr_has_yield(expr),
        CaseCondition::Range { from, to } => expr_has_yield(from) || expr_has_yield(to),
    }
}

// ── Rest-parameter arity collection ──────────────────────────────────────────
//
// The runtime rest-arg packing at a call site is only emitted when the
// compiler already knows some callable has a rest parameter of that fixed
// arity (`Compiler::rest_fixed_arities`). That set is populated lazily as each
// function/lambda body is compiled — but hoisted function declarations are
// compiled before later `const f = (...xs) => ...` arrow bindings, so a call to
// such an arrow from inside a hoisted function would miss the packing and drop
// arguments. This walker pre-collects every rest arity in the program so the
// set is complete before any body is compiled. Over-collecting is harmless: an
// arity that never matches a callee's `__vybe_rest_fixed_arity` stamp just adds
// an inert packing branch.

/// Collect the fixed arity (`params.len() - 1`) of every function, lambda,
/// method, or constructor that ends in a rest parameter, recursing through all
/// bodies and nested closures.
pub fn collect_rest_param_arities(stmts: &[Statement], out: &mut Vec<u8>) {
    for stmt in stmts {
        collect_rest_in_stmt(stmt, out);
    }
}

fn push_rest_arity(params: &[Param], out: &mut Vec<u8>) {
    if params.last().is_some_and(|p| p.is_rest) {
        out.push((params.len().saturating_sub(1)) as u8);
    }
}

fn collect_rest_in_stmt(stmt: &Statement, out: &mut Vec<u8>) {
    match &stmt.kind {
        StmtKind::FunctionDecl { params, body, .. } => {
            push_rest_arity(params, out);
            collect_rest_param_arities(body, out);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &decl.init {
                    collect_rest_in_expr(init, out);
                }
            }
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                collect_rest_in_member(member, out);
            }
        }
        StmtKind::EnumDecl { body_members, .. } => {
            for member in body_members {
                collect_rest_in_member(member, out);
            }
        }
        StmtKind::NamespaceDecl { body, .. } => collect_rest_param_arities(body, out),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            collect_rest_in_expr(cond, out);
            collect_rest_param_arities(then_body, out);
            for (econd, ebody) in elifs {
                collect_rest_in_expr(econd, out);
                collect_rest_param_arities(ebody, out);
            }
            if let Some(eb) = else_body {
                collect_rest_param_arities(eb, out);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                collect_rest_in_stmt(init, out);
            }
            if let Some(cond) = cond {
                collect_rest_in_expr(cond, out);
            }
            if let Some(update) = update {
                collect_rest_in_expr(update, out);
            }
            collect_rest_param_arities(body, out);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            collect_rest_in_expr(iter, out);
            collect_rest_param_arities(body, out);
            if let Some(eb) = else_body {
                collect_rest_param_arities(eb, out);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            collect_rest_in_expr(cond, out);
            collect_rest_param_arities(body, out);
            if let Some(eb) = else_body {
                collect_rest_param_arities(eb, out);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            collect_rest_param_arities(body, out);
            collect_rest_in_expr(cond, out);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            collect_rest_in_expr(expr, out);
            for case in cases {
                collect_rest_param_arities(&case.body, out);
            }
            if let Some(d) = default {
                collect_rest_param_arities(d, out);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            collect_rest_param_arities(body, out);
            for catch in catches {
                collect_rest_param_arities(&catch.body, out);
            }
            if let Some(eb) = else_body {
                collect_rest_param_arities(eb, out);
            }
            if let Some(f) = finally {
                collect_rest_param_arities(f, out);
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                collect_rest_in_expr(&item.expr, out);
            }
            collect_rest_param_arities(body, out);
        }
        StmtKind::Using { resource, body, .. } => {
            collect_rest_in_expr(resource, out);
            collect_rest_param_arities(body, out);
        }
        StmtKind::Lock { expr, body } => {
            collect_rest_in_expr(expr, out);
            collect_rest_param_arities(body, out);
        }
        StmtKind::Return(expr) => {
            if let Some(e) = expr {
                collect_rest_in_expr(e, out);
            }
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(e) = expr {
                collect_rest_in_expr(e, out);
            }
            if let Some(c) = cause {
                collect_rest_in_expr(c, out);
            }
        }
        StmtKind::Expr(expr) => collect_rest_in_expr(expr, out),
        StmtKind::Block(body) => collect_rest_param_arities(body, out),
        StmtKind::Assign { targets, value } => {
            for t in targets {
                collect_rest_in_expr(t, out);
            }
            collect_rest_in_expr(value, out);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            collect_rest_in_expr(target, out);
            collect_rest_in_expr(value, out);
        }
        StmtKind::Export {
            declaration,
            default,
            ..
        } => {
            if let Some(d) = declaration {
                collect_rest_in_stmt(d, out);
            }
            if let Some(e) = default {
                collect_rest_in_expr(e, out);
            }
        }
        StmtKind::Labeled { body, .. } => collect_rest_in_stmt(body, out),
        StmtKind::MatchStatement { subject, cases } => {
            collect_rest_in_expr(subject, out);
            for case in cases {
                collect_rest_param_arities(&case.body, out);
            }
        }
        _ => {}
    }
}

fn collect_rest_in_member(member: &ClassMember, out: &mut Vec<u8>) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            collect_rest_in_stmt(stmt, out)
        }
        ClassMember::Constructor { params, body, .. } => {
            push_rest_arity(params, out);
            collect_rest_param_arities(body, out);
        }
        ClassMember::Field { init, .. } => {
            if let Some(e) = init {
                collect_rest_in_expr(e, out);
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(g) = getter {
                collect_rest_param_arities(g, out);
            }
            if let Some(s) = setter {
                collect_rest_param_arities(&s.body, out);
            }
        }
        ClassMember::Const { value, .. } => collect_rest_in_expr(value, out),
        _ => {}
    }
}

fn collect_rest_in_expr(expr: &Expression, out: &mut Vec<u8>) {
    match &expr.kind {
        ExprKind::Lambda { params, body, .. } => {
            push_rest_arity(params, out);
            match body {
                LambdaBody::Expr(e) => collect_rest_in_expr(e, out),
                LambdaBody::Block(stmts) => collect_rest_param_arities(stmts, out),
            }
        }
        ExprKind::FunctionExpr(stmt) => collect_rest_in_stmt(stmt, out),
        ExprKind::ClassExpr { members, .. } => {
            for member in members {
                collect_rest_in_member(member, out);
            }
        }
        ExprKind::Call { callee, args, .. }
        | ExprKind::New {
            class: callee,
            args,
            ..
        } => {
            collect_rest_in_expr(callee, out);
            for arg in args {
                collect_rest_in_expr(&arg.value, out);
            }
        }
        ExprKind::SuperCall { args, .. } => {
            for arg in args {
                collect_rest_in_expr(&arg.value, out);
            }
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Assign {
            target: left,
            value: right,
        }
        | ExprKind::Walrus {
            target: left,
            value: right,
        }
        | ExprKind::Range {
            start: left,
            end: right,
            ..
        } => {
            collect_rest_in_expr(left, out);
            collect_rest_in_expr(right, out);
        }
        ExprKind::StaticAccess { class, member } => {
            collect_rest_in_expr(class, out);
            collect_rest_in_expr(member, out);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::IsType { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::TypeOf(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Await(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::RefLoad(expr) => collect_rest_in_expr(expr, out),
        ExprKind::Yield(expr) => {
            if let Some(e) = expr {
                collect_rest_in_expr(e, out);
            }
        }
        ExprKind::Ternary { cond, then, else_ } => {
            collect_rest_in_expr(cond, out);
            collect_rest_in_expr(then, out);
            collect_rest_in_expr(else_, out);
        }
        ExprKind::Member { object, .. } => collect_rest_in_expr(object, out),
        ExprKind::Index { object, index, .. } => {
            collect_rest_in_expr(object, out);
            collect_rest_in_expr(index, out);
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(k) = &item.key {
                    collect_rest_in_expr(k, out);
                }
                collect_rest_in_expr(&item.value, out);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                collect_rest_in_expr(item, out);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        collect_rest_in_expr(key, out);
                        collect_rest_in_expr(value, out);
                    }
                    ObjectProperty::Spread(e) => collect_rest_in_expr(e, out),
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => collect_rest_in_stmt(value, out),
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(e) | InterpolPart::Formatted(e, _) => {
                        collect_rest_in_expr(e, out)
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        ExprKind::Match { subject, arms } => {
            collect_rest_in_expr(subject, out);
            for arm in arms {
                if let Some(conds) = &arm.conditions {
                    for c in conds {
                        collect_rest_in_expr(c, out);
                    }
                }
                collect_rest_in_expr(&arm.body, out);
            }
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            collect_rest_in_expr(element, out);
            for comp_gen in generators {
                collect_rest_in_expr(&comp_gen.target, out);
                collect_rest_in_expr(&comp_gen.iter, out);
                for c in &comp_gen.conditions {
                    collect_rest_in_expr(c, out);
                }
            }
        }
        ExprKind::Slice { lower, upper, step } => {
            for e in [lower, upper, step].into_iter().flatten() {
                collect_rest_in_expr(e, out);
            }
        }
        _ => {}
    }
}

pub fn normalize_array_index_operand(
    index: Expression,
    semantics: ArrayIndexSemantics,
) -> Expression {
    match index.kind {
        ExprKind::Slice { lower, upper, step } => Expression::with_span(
            ExprKind::Slice {
                lower: lower.map(|expr| Box::new(normalize_array_subscript(*expr, semantics))),
                upper,
                step,
            },
            index.span,
        ),
        other => normalize_array_subscript(Expression::with_span(other, index.span), semantics),
    }
}

fn normalize_array_subscript(index: Expression, semantics: ArrayIndexSemantics) -> Expression {
    if semantics.first_index == 0 {
        return index;
    }

    let span = index.span;
    Expression::with_span(
        ExprKind::Binary {
            left: Box::new(index),
            op: BinOp::Sub,
            right: Box::new(Expression::int(semantics.first_index)),
        },
        span,
    )
}
