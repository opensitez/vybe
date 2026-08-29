//! Common AST — the language-neutral IR every walker produces and the
//! primitives/emitters consume.
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

pub mod builtin_slots;
pub mod builtin_types;
pub mod canon;
pub mod class_normalize;
pub mod datetime;

// ════════════════════════════════════════════════════════════════════════════
// Module (top-level compilation unit)
// ════════════════════════════════════════════════════════════════════════════

#[derive(Debug, Clone)]
pub struct Module {
    pub name: String,
    pub language: Lang,
    pub body: Vec<Statement>,
    pub imports: Vec<Import>,
    /// This module's declared policy, in force from its first statement. The
    /// walker states its language's defaults here; a [`StmtKind::Directive`]
    /// in the body changes them from that point on. See [`Directives`].
    pub directives: Directives,
    /// The component's CANON SECTION, in canonidx order — empty for every
    /// language that is not a Component Model front end.
    ///
    /// Module-level metadata, not code: a canon definition has no execution
    /// position and cannot be branched over, which is why it is a table here
    /// rather than a [`Statement`]. Same category as the global index space.
    /// See [`canon::CanonDecl`].
    pub canon: canon::ComponentSection,
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
    Kotlin,
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
// Types and generics
// ════════════════════════════════════════════════════════════════════════════

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TypePath {
    pub segments: Vec<String>,
}

impl TypePath {
    pub fn new(segments: Vec<String>) -> Self {
        Self { segments }
    }

    pub fn from_dotted(path: &str) -> Self {
        Self {
            segments: path
                .split('.')
                .map(str::trim)
                .filter(|segment| !segment.is_empty())
                .map(str::to_string)
                .collect(),
        }
    }

    pub fn display_name(&self) -> String {
        self.segments.join(".")
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TypeRef {
    pub kind: TypeRefKind,
}

impl TypeRef {
    pub fn named(path: impl Into<String>) -> Self {
        Self {
            kind: TypeRefKind::Named {
                path: TypePath::from_dotted(&path.into()),
                args: Vec::new(),
            },
        }
    }

    pub fn generic_param(name: impl Into<String>) -> Self {
        Self {
            kind: TypeRefKind::GenericParam { name: name.into() },
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum TypeRefKind {
    Named {
        path: TypePath,
        args: Vec<GenericArg>,
    },
    GenericParam {
        name: String,
    },
    Array {
        element: Box<TypeRef>,
        rank: usize,
    },
    Tuple {
        elements: Vec<TypeRef>,
    },
    Function {
        params: Vec<TypeRef>,
        result: Box<TypeRef>,
    },
    Union {
        members: Vec<TypeRef>,
    },
    Intersection {
        members: Vec<TypeRef>,
    },
    Nullable {
        inner: Box<TypeRef>,
    },
    Pointer {
        inner: Box<TypeRef>,
    },
    Reference {
        inner: Box<TypeRef>,
    },
    Wildcard {
        bound: Option<GenericBound>,
    },
    SelfType,
    Infer,
    Error,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum GenericArg {
    Type(TypeRef),
    Const(String),
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum GenericBound {
    Extends(Box<TypeRef>),
    Super(Box<TypeRef>),
}

#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct GenericDecl {
    pub params: Vec<GenericParam>,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct GenericParam {
    pub name: String,
    pub constraints: Vec<GenericConstraint>,
    pub variance: GenericVariance,
    pub default: Option<TypeRef>,
    pub runtime: GenericRuntimeMode,
}

impl GenericParam {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            constraints: Vec::new(),
            variance: GenericVariance::Invariant,
            default: None,
            runtime: GenericRuntimeMode::Erased,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum GenericVariance {
    #[default]
    Invariant,
    Covariant,
    Contravariant,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum GenericRuntimeMode {
    #[default]
    Erased,
    Reified,
    Specialized,
    Shared,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum GenericConstraint {
    Any,
    Class,
    Struct,
    Record,
    Interface,
    Enum,
    Delegate,
    Constructor { argc: Option<usize> },
    Extends(TypeRef),
    Implements(TypeRef),
    Comparable,
    Numeric,
    Integer,
    Floating,
    NonNull,
    Nullable,
    Unmanaged,
    CopyLike,
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

    /// Does this statement DECLARE something rather than execute?
    ///
    /// A declaration is not a step in the flow, so a pass that reorders or
    /// partitions executable statements — `goto` lowering splitting a body at
    /// its labels, say — must leave declarations where every branch can still
    /// see them. This is a property of the NODE, not of the language: C's
    /// lowering hoisted only `VarDecl` because C has nothing else to put
    /// there, not because the others belong inside a numbered block.
    pub fn is_declaration(&self) -> bool {
        matches!(
            self.kind,
            StmtKind::VarDecl { .. }
                | StmtKind::FunctionDecl { .. }
                | StmtKind::ClassDecl { .. }
                | StmtKind::StructDecl { .. }
                | StmtKind::ModuleDecl { .. }
        )
    }
}

#[derive(Debug, Clone)]
pub enum StmtKind {
    /// Expression used as a statement.
    Expr(Expression),

    /// CSP `select` — readiness choice over channel communications
    /// (Go §Select statements). Structural, not an if-chain: readiness
    /// includes "closed" (always ready, yields zero/ok=false) and excludes
    /// nil channels, which no expression-level rewrite can express without
    /// re-evaluating operands. Lowered ONCE in `primitives/channels.rs`.
    Select {
        arms: Vec<SelectArm>,
        default: Option<Vec<Statement>>,
    },

    /// Block of statements.
    Block(Vec<Statement>),

    /// A change of declared policy, in force from this point on. The AST form
    /// of `{$R+}`, `declare(strict_types=1)`, `Option Explicit`. Emits no code
    /// — `scope` says how far the change reaches. See [`Directives`].
    Directive {
        set: Directives,
        scope: DirectiveScope,
    },

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
        /// Declared record semantics — storage, equality, layout, variant part.
        /// Defaults to a plain reference aggregate, so a walker that does not
        /// set it behaves exactly as before. See `recordprimitiveplan.md`.
        semantics: ValueSemantics,
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
        /// The target is bound BY REFERENCE — it aliases the value's storage,
        /// so later writes through either name are seen by both
        /// (php `$b = &$a`).
        ///
        /// Completes a vocabulary the AST already had everywhere else:
        /// `Param.pass_by`, `Argument.by_ref` and `ArrayElement.by_ref`.
        /// Assignment was the one binding form missing it, which is why the
        /// compiler had to ask `profile.name == "php"` to tell an ALIAS from a
        /// stored pointer VALUE — `p = &v` in c/go/c#/pascal keeps assigning a
        /// pointer and leaves this false, so rebinding `p` still rebinds rather
        /// than writing through.
        by_ref: bool,
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

    // ── Record files ─────────────────────────────────────────────────────
    //
    // These REPLACE the nine VB6 nodes above. Both models are live during the
    // migration (`recordfileplan.md` §5): the old nodes die when their last
    // emitter is gone, not before, and the `__vb_*` globals go after that.
    /// A file DECLARED: its own name, the record type it holds, and how it is
    /// organized. Everything a transfer needs to know, stated once.
    ///
    /// Replaces `OpenFile { file_number }`. A file's identity is its
    /// declaration — COBOL `SELECT`, Pascal `file of Rec`, Fortran's unit — and
    /// only VB6 identifies one by an integer the programmer invents. Six
    /// languages were faking a number to fit a model none of them has.
    ///
    /// Layout is NOT here: `record` names a `StructDecl`, whose members carry
    /// [`FieldStorage`]. One description of a record, in one place.
    FileDecl {
        /// The file's identity in source. Not a handle, not a number.
        name: String,
        path: Expression,
        /// The record type — resolves to a `StructDecl`.
        record: TypeRef,
        organization: FileOrganization,
        access: FileAccess,
        /// What opening it permits, and what it does to existing contents.
        mode: OpenMode,
        /// Indexed files only: which record fields are keys, in priority order.
        keys: Vec<RecordKey>,
    },

    /// ONE transfer over an addressing mode.
    ///
    /// COBOL `READ`/`WRITE`/`REWRITE`/`DELETE`, VB `Get #n`/`Put #n`, Pascal
    /// `Read(f,r)`/`Write(f,r)` and Fortran `READ(u, rec=n)` are the same
    /// operation with different `direction` and `at` — which is why they were
    /// four nodes that each knew a little and none knew the layout.
    ///
    /// Lowers onto `wasi:filesystem/types` `[method]descriptor.read-via-stream(offset)`
    /// / `write-via-stream(data, offset)` — the only byte-moving calls WASI
    /// 0.3.1 defines. `at` becomes the offset: record *n* of a fixed-width
    /// record sits at *n × width*, and the width comes from the record type's
    /// `FieldStorage`, never re-derived at emit time.
    RecordTransfer {
        /// The declared file — an `Ident` naming a [`StmtKind::FileDecl`].
        file: Expression,
        /// The record type this transfer moves, naming the same `StructDecl`
        /// as the file's declaration.
        ///
        /// Stated HERE as well as on [`StmtKind::FileDecl`] so the node is
        /// self-contained. Looking it up through the file instead makes a
        /// transfer depend on the declaration having been COMPILED first —
        /// and `READ F.` in a program that never opens `F`, or a paragraph
        /// compiled ahead of the one holding the `OPEN`, then has no type at
        /// all. This is a NAME, not a layout: the extents still live in one
        /// place, on the type's own members.
        record_type: TypeRef,
        direction: RecordDirection,
        at: RecordAddress,
        /// The struct read into, or written from. `None` for `Delete`, which
        /// addresses a record without moving one.
        record: Option<Expression>,
        /// Where the OUTCOME of the transfer is stored, if the program asked
        /// for one: COBOL `FILE STATUS`, Fortran `IOSTAT=`, VB `Err`.
        ///
        /// A transfer that reached the end of the file is not an error and
        /// not a silent no-op — it is a fact the program is entitled to read,
        /// and `AT END` / `EOF(n)` / `IOSTAT < 0` are three spellings of
        /// asking for it. Without this the emitter could position a read and
        /// then have nowhere to say it found nothing, which is the one shape
        /// that turns an empty file into a plausible-looking record.
        ///
        /// Values are the two-character COBOL status codes, because they are
        /// the only vocabulary among these that distinguishes the cases
        /// (`"00"` ok, `"10"` at end, `"23"` key not found); a language whose
        /// own spelling is coarser narrows at its walker.
        status: Option<Expression>,
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
    /// Terminate the program with a status, from any depth — PHP `exit`/`die`,
    /// Ruby `exit`, Lua `os.exit`, Go `os.Exit`, Pascal `Halt`, COBOL
    /// `STOP RUN`, Python `sys.exit`, JS `process.exit`, C `exit`.
    ///
    /// A non-local transfer like `Return`/`Break`/`Throw`, which is why it is a
    /// statement rather than a call: it unwinds every frame and never returns.
    /// `None` means status 0. Lowered once by
    /// `primitives/control_flow.rs::compile_exit_stmt`; per-language argument
    /// quirks (Lua's inverted boolean, PHP's string message) are normalized by
    /// the walker before they get here.
    Exit {
        status: Option<Expression>,
    },
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
        /// The INDEX type: `(memory i64 …)` (memory64) addresses with i64,
        /// the default with i32. It adds no opcodes — every load, store and
        /// `memory.*` op reads it to decide the width of its address and
        /// count operands, and of the value `memory.size`/`grow` answer.
        is_64: bool,
    },

    /// `(table $id? min max? funcref)` — a reference-table declaration.
    /// Declaration order is the table index.
    TableDecl {
        /// Minimum element count.
        min_size: u64,
        /// Maximum element count; `None` = unbounded.
        max_size: Option<u64>,
        /// The INDEX type: `(table i64 …)` addresses with i64. Same role as
        /// `MemoryDecl::is_64`, for `table.size`/`grow`/`fill`/`copy`/`init`
        /// and `call_indirect`.
        is_64: bool,
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

    // ── WASM exception handling (canonical `try_table`) ──────────────────────
    // The WASM text format declares exception *tags* and matches catch clauses
    // by TAG IDENTITY (not by exception class like the higher-level `Try`
    // node). These nodes model that directly; the wast frontend folds both the
    // canonical `try_table` form and the legacy `try/catch/delegate/rethrow`
    // sugar into them, and the compiler lowers them to the VM's `try_table` /
    // `throw` / `throw_ref` opcodes.
    /// `(tag $id? (param t*))` — a WASM exception-tag declaration. Declares a
    /// distinct tag entity, keyed by `name` so `throw $e`/`catch $e` resolve to
    /// the SAME entity. Lowered via `import_exception_tag`.
    WasmTagDecl {
        /// Tag name (the `$id`), used to key the imported tag entity.
        name: String,
        /// Payload arity (number of `param` values the tag carries).
        arity: u8,
    },

    // ── Call Tags proposal (proposals/call-tags) ─────────────────────────────
    // A CALL tag is a different entity from the exception tag above: that one
    // names an exception TYPE, this one names a calling CONVENTION over a
    // signature. Both are module entities, so both are declarations rather than
    // expressions, and both key by name so a declaration and its uses meet at
    // one entity.
    /// `(call_tag $id (param t*) (result t*) (fallback $f)?)`.
    ///
    /// Without a fallback this is `call_tag.canon` — the canonical tag for the
    /// signature, interned, and an unhandled call traps. With one it is
    /// `call_tag.new`: a FRESH identity over a signature that may already have
    /// a canonical tag, which is what lets two structurally identical functions
    /// stay distinguishable after GC type canonicalisation.
    WasmCallTagDecl {
        /// Tag name (the `$id`), keying the entity.
        name: String,
        /// Signature shape — parameter and result counts.
        params: u8,
        results: u8,
        /// The DECLARED functype, as its source spelling (`"i32->i32"`).
        ///
        /// ⛔ `params`/`results` above are a SHAPE, and a shape is not a
        /// functype. `call_tag.canon $functype` derives the canonical tag *of
        /// that functype*, so `[i32]->[i32]` and `[f64]->[f64]` are two tags;
        /// keyed on counts they intern to one, and an `i32`-shaped funcref
        /// answers the `f64` canonical tag. Carried as the source spelling
        /// because the VM erases runtime types — `Chunk` has no value types at
        /// all — so this is the only place the functype survives. Compared,
        /// never interpreted. Empty when the producer did not supply one, which
        /// keeps every non-wast frontend on the old shape-only behaviour.
        signature: String,
        /// `(canon)` — `call_tag.canon`, interned per signature. Otherwise this
        /// is `call_tag.new`: a fresh identity over that signature.
        canonical: bool,
        /// `(fallback $f)`: the handler called when a `funcref` does not handle
        /// this tag, receiving `[ti* funcref]`. `None` ⇒ canonical ⇒ trap.
        fallback: Option<String>,
    },

    /// `(func_switch $id (case $tag $func)* (forward $other)?)` — the
    /// Overview's "alternative to `func`". Has no type and cannot be called
    /// directly; a `funcref` to it dispatches on the tag it is called with,
    /// forwarding unmatched tags to `$other` when given.
    WasmFuncSwitchDecl {
        name: String,
        /// `($call_tag $func)*`, matched in declaration order.
        arms: Vec<(String, String)>,
        /// The trailing `$func_switch?`.
        forward: Option<String>,
    },

    /// `(func … (call_tag $t+))` — which call tags a func's `funcref` handles.
    /// Declaring any REPLACES the default of its own canonical tag, which is
    /// what gives the proposal its security property: a func that lists only
    /// non-exported tags cannot be reached indirectly from outside the module.
    WasmFuncCallTags {
        /// The func being declared.
        func: String,
        /// Tag names it handles.
        tags: Vec<String>,
    },

    /// `throw $tag` — raise the WASM exception `$tag` with `args` as its
    /// payload (already popped off the stack machine in push order). Lowers to
    /// each arg followed by `THROW <tagidx>`.
    WasmThrow {
        /// Tag name to raise (resolves to the same entity as `WasmTagDecl`).
        tag: String,
        /// Payload values, bottom-to-top.
        args: Vec<Expression>,
    },

    /// A WASM exception-handling block — canonical `try_table`, or the legacy
    /// `try/catch/catch_all/delegate` sugar the wast walker folds into it.
    /// Unlike the class-based [`StmtKind::Try`], catch clauses match by TAG
    /// IDENTITY and payloads are raw stack values. The compiler lowers this to
    /// a `TRY_TABLE` (one clause per catch, each a forward offset to its inline
    /// handler), the body, a structural `end`, then the handlers.
    WasmTryTable {
        body: Vec<Statement>,
        catches: Vec<WasmCatch>,
        /// Spec blocktype `bt` in `try_table bt vec(catch) instr* end` —
        /// `try_table` IS a block and may take and produce values. Dropping it
        /// made `(try_table (result i32) …)` unrepresentable: the VM pushed a
        /// zero result arity and discarded the value.
        params: u8,
        results: u8,
    },

    /// `rethrow N` (legacy) — re-raise the exception caught by the `N`th
    /// enclosing catch, whose `exnref` the handler captured (`capture_ref`).
    /// Lowers to `THROW_REF` of that captured reference.
    WasmRethrow {
        /// The `exnref` local name captured by the target catch handler.
        exnref_local: String,
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
        /// Declared byte extent — see [`FieldStorage`]. `None` unless the
        /// language declares a fixed width (COBOL `PIC`, VB `String * n`,
        /// Fortran `character(len=n)`, Pascal packed records).
        ///
        /// It is NOT part of `modifiers`: everything there — visibility,
        /// static, readonly, virtual — is about access and dispatch. A byte
        /// extent is neither, and filing it there to save edits would put the
        /// fact where nobody looks for it.
        storage: Option<FieldStorage>,
    },

    Method(Box<Statement>),

    Constructor {
        /// A *named* constructor's name (Dart `Point.origin()`, Pascal
        /// `constructor Create`). `None` is the ordinary unnamed constructor,
        /// which every other language uses — those dispatch by signature, not
        /// by name.
        name: Option<String>,
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

    /// The class draws members from another declared type: PHP `use T;`,
    /// Dart `with M`, Ruby `include`/`prepend`, Java interface `default`
    /// methods, Go field promotion.
    ///
    /// This is a DECLARATION, not a fold. Every language parsed it and then
    /// dropped it into a walker thread-local (`TRAIT_USAGES` / `TRAIT_ALIASES` /
    /// `TRAIT_PRECEDENCES` in PHP, `DART_CLASS_MIXINS` in Dart) because the AST
    /// had nowhere to put it — which is why each language then had to fold it
    /// itself. A thread-local is also cleared per `parse()`, so it cannot
    /// survive multi-file compilation.
    ///
    /// The walker records what the source said; `normalize_class` turns it into
    /// an `Augmentation`; the shared `class_augmentation` pass applies it once.
    /// See flexclassplan.md §4c-R.
    Augment(AugmentDecl),
}

/// One augmentation clause as the source wrote it.
#[derive(Debug, Clone, Default)]
pub struct AugmentDecl {
    /// The augmenting type's name, in the SOURCE spelling — the compiler
    /// resolves it against declared classes (a trait may be referenced short
    /// and declared fully qualified).
    pub from: String,
    /// Go field promotion only: the field the receiver rebinds to.
    pub via_field: Option<String>,
    /// Per-member adjustments (PHP `as` / `insteadof`). Empty for every other
    /// language's mechanism.
    pub adjustments: Vec<AugmentAdjustment>,
}

/// A per-member adjustment on an augmentation clause: PHP `A::run as protected
/// go;` (rename and/or change visibility) and `A::run insteadof B;` (exclude).
///
/// PHP's `as` is ADDITIVE — the member stays bound under its own name too — and
/// it composes with `insteadof`, so an excluded member is still reachable under
/// its alias.
#[derive(Debug, Clone, Default)]
pub struct AugmentAdjustment {
    /// Source member this applies to.
    pub member: String,
    /// Also bind under this name (PHP `as other`).
    pub rename_to: Option<String>,
    /// Override the member's visibility (PHP `as protected run`).
    pub visibility: Option<Visibility>,
    /// Drop this member from THIS augmentation (PHP `insteadof`).
    pub exclude: bool,
}

/// What kind of type a `ClassDecl` actually declares.
///
/// `class`, `interface`, `trait`, `mixin` and `module` all parse to
/// `StmtKind::ClassDecl` with nothing to tell them apart, so PHP kept a
/// `trait_names: HashSet` and Dart a `DART_MIXIN_NAMES` in walker thread-locals
/// to recover it. The compiler needs it to know a type is not instantiable and
/// to answer `trait_exists` / `kind_of?`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ClassKind {
    #[default]
    Class,
    Interface,
    /// PHP `trait` — an augmentation source, never instantiable.
    Trait,
    /// Dart `mixin` — an augmentation source with an optional `on` constraint.
    Mixin,
    /// Ruby `module` — an augmentation source, and a namespace.
    Module,
    Struct,
    /// Kotlin `data class`, Java `record`, C# `record` — a declaration whose
    /// members are DERIVED from its primary constructor's components:
    /// component accessors, a copy/`with` constructor, and structural
    /// `ToString` / `Eq` / `Hash`.
    ///
    /// Without this, a frontend has no way to tell the normalizer that a
    /// declaration is derived, so each one synthesized the members in its
    /// WALKER as hand-built AST — Kotlin in `walk_class_decl`, Java at
    /// `walker.rs:2534`. That is the duplication flexclassplan §4a-ter
    /// collapses: the shape is derived from `constructors[0].params`, which is
    /// the same computation in every language, and it belongs in normalization
    /// where the bindings can be `ProtocolSlot`s rather than spellings (§2a).
    Record,
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

impl PlaceExpr {
    /// Does `expr` denote STORAGE? `Some` for a place, `None` for an rvalue.
    ///
    /// This is the whole of `referenceplan.md` §4's routing rule: a place takes
    /// `RefOf` and aliases its storage; an rvalue takes `Unary{AddrOf}` and gets
    /// a fresh cell. Both arms are legitimate — go's `&Foo{...}` has no prior
    /// storage and must NOT resolve — so the converter is what keeps them apart,
    /// and it belongs here rather than inside one language's walker.
    ///
    /// Lifted verbatim from go's private `go_expr_to_place`, which had been
    /// proving the design in the one language that could reach it while c and
    /// pascal routed places through the rvalue node for want of a copy.
    ///
    /// The four arms ARE the closed set of `PlaceExpr` — nothing left over.
    pub fn from_expr(expr: &Expression) -> Option<PlaceExpr> {
        match &expr.kind {
            ExprKind::Ident(name) => Some(PlaceExpr::Ident(name.clone())),
            ExprKind::Member {
                object,
                field,
                null_safe,
            } => Some(PlaceExpr::Member {
                object: object.clone(),
                field: field.clone(),
                null_safe: *null_safe,
            }),
            ExprKind::Index {
                object,
                index,
                null_safe,
            } => Some(PlaceExpr::Index {
                object: object.clone(),
                index: index.clone(),
                null_safe: *null_safe,
            }),
            // A dereference is itself a place: `*p = v` and `&(*p)` are both
            // legal, and `&*p` is just `p`.
            ExprKind::RefLoad(expr) => Some(PlaceExpr::Deref(expr.clone())),
            ExprKind::Unary {
                op: UnaryOp::Deref,
                expr,
            } => Some(PlaceExpr::Deref(expr.clone())),
            _ => None,
        }
    }
}

impl TryFrom<&Expression> for PlaceExpr {
    type Error = ();

    fn try_from(expr: &Expression) -> Result<Self, Self::Error> {
        PlaceExpr::from_expr(expr).ok_or(())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ZipMode {
    First,
    Shortest,
    Longest,
}

/// Logical traversal order for ranked array transforms.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ArrayTraversalOrder {
    RowMajor,
    ColumnMajor,
}

/// Cross-language ranked-array transformations.
///
/// This is distinct from binary [`packing`](crate) semantics: Fortran
/// `PACK`/`UNPACK` compact or scatter array elements under a mask, while Ruby,
/// Python, PHP, and Lua packing encode scalar values into bytes. Languages with
/// array/tensor intrinsics can normalize to this node and share one lowering.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ArrayTransformOp {
    /// `PACK(array, mask[, vector])`: compact elements where `mask` is true.
    PackMask,
    /// `UNPACK(vector, mask, field)`: scatter vector elements into mask-true
    /// positions and fill mask-false positions from `field`.
    UnpackMask,
    /// `MERGE(tsource, fsource, mask)`: elementwise SELECT — the result has the
    /// mask's shape, and each element comes from `tsource` where the mask is
    /// true and `fsource` where it is false. Either source may be a scalar,
    /// which broadcasts.
    ///
    /// The fourth member of the mask-driven family beside [`PackMask`] and
    /// [`UnpackMask`], and general rather than Fortran-specific: NumPy
    /// `where`, MATLAB logical indexing, APL compress, SQL `CASE` over a
    /// column, Julia `ifelse.`, wasm SIMD `v128.bitselect`.
    ///
    /// It belongs here rather than in a walker because RANK is the whole
    /// difficulty: a rank-2 array is a NEST, so a plain `mask.map(...)` hands
    /// the callback a ROW — which is always truthy, so every element takes the
    /// true branch. Flattening and re-shaping is what the other three already
    /// do.
    ///
    /// [`PackMask`]: ArrayTransformOp::PackMask
    /// [`UnpackMask`]: ArrayTransformOp::UnpackMask
    MergeMask,
    /// `RESHAPE(source, shape[, pad])`: the source's elements rebuilt into an
    /// array of the given shape, cycling `pad` when the shape asks for more
    /// than the source holds and truncating when it asks for fewer. The node's
    /// [`ArrayTraversalOrder`] is which subscript runs fastest as the source is
    /// consumed — Fortran's `ORDER=` permutation, NumPy's `order=`.
    Reshape,
}

impl ArrayTransformOp {
    /// Which argument the result takes its SHAPE from, or `None` when the
    /// result is always a rank-1 vector whatever went in.
    ///
    /// The rank of a transform's result is a property of the OPERATION, so the
    /// node answers it once for every language. Without this, each consumer
    /// asking "is this expression an array?" has to special-case the op list
    /// itself — and the one that mattered is `MERGE`, whose result is a SCALAR
    /// when its mask is one. Treating every transform as an array made
    /// `merge(1, 0, scalar_condition) /= 1` lower elementwise, and an array is
    /// truthy, so the comparison took the wrong branch while still printing the
    /// right value.
    pub fn shape_source_arg(self) -> Option<usize> {
        match self {
            // PACK compacts to a vector — rank 1 regardless of the source.
            ArrayTransformOp::PackMask => None,
            // UNPACK scatters into the mask's shape.
            ArrayTransformOp::UnpackMask => Some(1),
            // RESHAPE is told its shape outright.
            ArrayTransformOp::Reshape => Some(1),
            // MERGE selects elementwise, so it wears the mask's shape.
            ArrayTransformOp::MergeMask => Some(2),
        }
    }
}

#[derive(Debug, Clone)]
pub enum ExprKind {
    Lit(Literal),
    Ident(String),
    This,
    Super,
    /// The module's GLOBAL NAMESPACE OBJECT — ECMA-262 §9.1.1.4's global
    /// environment record, the object half.
    ///
    /// Four languages spell it and mean the same thing: JS `globalThis`, Lua
    /// `_G`, PHP `$GLOBALS`, Python `globals()`. It sits beside `This` and
    /// `Super` because it is the same kind of fact — a language-neutral
    /// reference to a well-known object, not a name that happens to resolve.
    ///
    /// ⛔ It is a NODE rather than an `Ident` because a name is SPELLING.
    /// `profile.global_namespace` used to carry the four spellings and shared
    /// code compared the source text against it (`names_global_namespace`),
    /// with `global_namespace_is_call` for Python's call form — a per-language
    /// spelling table consulted by the shared compiler, which is
    /// `directives.md` §10.3's under-normalization exactly: "the walker passes
    /// its own syntax through and shared code grows a branch to cope. That
    /// branch is a language check."
    ///
    /// Walkers own the spelling; this is the vocabulary they normalize into.
    GlobalNamespace,

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
    /// A named tuple — array-backed like `Tuple`, but each element may carry a
    /// field name and the tuple an optional type name (Python `namedtuple`,
    /// C# named `ValueTuple` `(x: 1, y: 2)`). Lowered by the shared compiler to
    /// a tagged array plus field-name keys and hidden `__fields`/`__typename`,
    /// so a named tuple is one runtime value across languages.
    /// See `vybe_compiler::primitives::tuples`.
    NamedTuple {
        fields: Vec<(Option<String>, Expression)>,
        type_name: Option<String>,
    },
    /// Python `{1, 2, 3}` — unordered unique collection
    Set(Vec<Expression>),
    /// A property-bag literal: string keys, JS object semantics.
    Object(Vec<ObjectProperty>),
    /// An ORDERED, `Value`-keyed collection literal — Python's `dict`, and the
    /// natural home for PHP arrays / Ruby hashes / JS `new Map` when they move.
    ///
    /// Distinct from [`ExprKind::Object`] because the difference is semantic,
    /// not cosmetic: a `Map` keeps the key's TYPE (`{1: 'a'}` stays an int key,
    /// where an object literal stringifies it) and guarantees insertion order,
    /// which a JS object does not for integer-like keys.
    ///
    /// It is a separate NODE rather than a profile flag on purpose. This used to
    /// be `profile.dict_literals_as_map`, and a per-language boolean deciding
    /// what a shared node MEANS is exactly the thing the AST is supposed to
    /// remove: a primitive holding an `Object` could not tell which of two
    /// runtime shapes it had without consulting the front end's profile. The
    /// front end knows; it says so here.
    ///
    /// Entries only — a spread (`{**a, 'k': 1}`) is a different operation and
    /// stays an `Object`.
    Map(Vec<(Expression, Expression)>),
    /// Cross-language zip/transpose primitive. Languages choose the length
    /// policy in their walker: Python `zip` uses `Shortest`, PHP
    /// `array_map(null, ...)` uses `Longest`, Ruby-style receiver zip uses
    /// `First`. Lowered by the shared collections primitive.
    Zip {
        iterables: Vec<Expression>,
        mode: ZipMode,
        strict: bool,
    },
    /// Cross-language array/map transform. Frontends use this instead of
    /// spelling a host-specific `.map(...)` member call when the source
    /// language means elemental collection mapping.
    ArrayMap {
        array: Box<Expression>,
        params: Vec<Param>,
        body: Box<Expression>,
    },
    /// Ranked array transform, normalized from language-specific intrinsics
    /// such as Fortran `PACK`/`UNPACK`.
    ArrayTransform {
        op: ArrayTransformOp,
        args: Vec<Expression>,
        order: ArrayTraversalOrder,
    },
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
    /// A normalized async-model operation — see [`AsyncOp`]. One vocabulary
    /// for every language's spelling (`Promise.resolve` / `Task.FromResult` /
    /// `Future.value` …), produced by WALKERS at parse and lowered ONCE in the
    /// compiler onto the ECMA-262 §27.2 host surface + the JSPI suspend
    /// mechanism. Languages differ only in normalization; the tree is common.
    Async(AsyncOp),
    /// A normalized channel operation — see [`ChanOp`]. CSP is its OWN model,
    /// deliberately not shoehorned into [`AsyncOp`]: a channel is a value
    /// with buffer/closed state and blocking rendezvous semantics, not a
    /// one-shot settled result. Go is the first normalizer; Rust
    /// (`std::sync::mpsc`) and Kotlin (`Channel<T>`) share the vocabulary.
    Chan(ChanOp),
    /// A normalized atomic (shared-memory read-modify-write) operation — see
    /// [`AtomicOp`]. The third member of the concurrency family, alongside
    /// [`AsyncOp`] and [`ChanOp`], and for the same reason: five languages
    /// spell it, and before this node each had invented its own channel —
    /// C# `Interlocked` on the .NET tree, C `atomic_fetch_add` desugared to a
    /// NON-atomic `Sequence` in its walker, Go a `sync` prelude written in Go,
    /// Java/Kotlin `AtomicInteger` in walker tables, Pascal `TInterlocked`
    /// bound to nothing. Three of the five were not atomic at all.
    Atomic(AtomicOp),

    /// `call_with_tag $tag` / `call_return_with_tag $tag` — Call Tags proposal.
    ///
    /// An ordinary `Call` cannot express this: the tag is not an argument, it
    /// is which CONVENTION the call is made under, and the callee decides
    /// whether it handles it. Same operand layout as `call_ref` — arguments
    /// first, `funcref` on top (`[ti* funcref] -> [to*]`).
    ///
    /// `call_indirect_with_tag $table $tag` is not a separate node: the Overview
    /// defines it as shorthand for `(call_with_tag $tag (table.get $table))`, so
    /// the front end desugars it and there is one path to keep correct.
    WasmCallWithTag {
        /// The call tag's name, resolving to the same entity its declaration
        /// created.
        tag: String,
        /// The `funcref` being called.
        callee: Box<Expression>,
        /// `ti*`, bottom-to-top.
        args: Vec<Expression>,
        /// `call_return_with_tag` — the tail-call form.
        tail: bool,
        /// `call_indirect_with_tag $table $tag`: the table index, with `callee`
        /// holding the ELEMENT INDEX rather than a funcref.
        ///
        /// The Overview defines this as shorthand for
        /// `(call_with_tag $tag (table.get $table))`, but there is no
        /// funcref-yielding table expression to desugar into here — plain
        /// `call_indirect` is itself a single opcode with the table as an
        /// immediate. So the shorthand keeps its own opcode and this node
        /// carries the immediate.
        table: Option<u32>,
    },

    Await(Box<Expression>),
    Yield(Option<Box<Expression>>),
    YieldFrom(Box<Expression>),

    // ── Callable references ──────────────────────────────────────────────
    /// A reference to existing callable code (for example VB `AddressOf F`).
    /// This is the narrow bare-function spelling. Receiver-aware method
    /// groups and delegates use `CallableRef` below.
    /// This is not storage address-of; `UnaryOp::AddrOf` owns that axis.
    FuncRef(String),
    /// A first-class callable reference with optional receiver policy.
    /// This preserves source intent for method groups, delegates, Kotlin/Java
    /// method refs and similar constructs without collapsing them into
    /// lambda bodies or proxy interception.
    CallableRef {
        target: Box<Expression>,
        receiver: Option<Box<Expression>>,
        binding: CallableBinding,
        adapter: Option<CallableAdapter>,
    },
    // ── VB / .NET ────────────────────────────────────────────────────────
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
    /// Object-operation interception boundary such as JS `new Proxy(target, handler)`.
    Proxy {
        target: Box<Expression>,
        handler: Box<Expression>,
    },

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
    /// Binary data — PHP byte strings, Python `bytes`, Go `[]byte`.
    ///
    /// A distinct TYPE, not a differently-encoded `Str`
    /// (`unifiedstringplan.md` §3c): it is what lets `(Bytes, slot)` bindings
    /// resolve, which `builtin_slots.rs` records as the reason `Bytes` is
    /// unbound today. Deliberately NOT `Utf8Str`/`Utf16Str` — encoding is a
    /// property of a conversion, not of a literal.
    Bytes(Vec<u8>),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CallableBinding {
    /// A bare function or static callable reference. No receiver is bound.
    Static,
    /// A callable whose receiver is already bound.
    BoundReceiver,
    /// A method reference where the receiver is supplied by the eventual call.
    UnboundReceiver,
}

#[derive(Debug, Clone)]
pub enum CallableAdapter {
    /// Invoke this callable reference by compiling the supplied expression-body
    /// adapter. This keeps source intent as `CallableRef` while preserving
    /// language-specific arity/receiver/helper semantics that are not yet a
    /// direct callable value.
    Expr {
        params: Vec<Param>,
        body: Box<Expression>,
    },
}

// ── Variables ───────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct VarDeclarator {
    pub pattern: BindingPattern,
    pub type_hint: Option<TypeHint>,
    pub init: Option<Expression>,
    pub array_bounds: Option<Vec<Expression>>,
    pub with_events: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub enum VarDeclKind {
    Dim,
    Let,
    Const,
    /// An ordinary declaration, bound in the scope that encloses it — the
    /// block, if there is one. java/kotlin/dart/c `var`, and the default.
    Var,
    /// A declaration bound in the enclosing FUNCTION rather than the block it
    /// is written in — ECMA-262's VariableEnvironment (§9.1: the record that
    /// holds `VariableStatement` bindings, which a block does NOT push).
    ///
    /// ⛔ This exists because `Var` MEANT TWO THINGS. Ten languages emit it,
    /// and for nine of them it is block-scoped; only JS's `var` outlives its
    /// block. The difference was carried by a profile flag, `hoist_var`, read
    /// as `kind == Var && profile.hoist_var` — a language check with a nicer
    /// name (`directives.md` §1), and §5's "reusing a field as a marker": one
    /// variant answering two questions, so a change to either broke the other.
    ///
    /// The scoping of a declaration is something the AUTHOR DECLARED, so it
    /// belongs on the declaration. Any language with function-scoped
    /// declarations states it and gets the behaviour; nothing asks whose
    /// language it is.
    FunctionScoped,
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

// ── Declared types ──────────────────────────────────────────────────────────

/// Whether binding a value to a declared type CONVERTS it.
///
/// This is a property of the DECLARATION, not of the language, and it is set by
/// the walker that parsed it — the only component that legitimately knows its
/// own language's rules. Shared code reads the fact and never asks whose
/// language produced it.
///
/// # Why it cannot live on the profile
///
/// It was a per-compilation boolean (`coerces_value_to_type_hint`) defaulting
/// to `true`, which no language declared. Python inherited it by omission and
/// its annotations were being narrowed — `x: int = 3.7` gave `3` where real
/// python gives `3.7`. A per-language switch also cannot express a language
/// that has both, and it puts a language question in shared code.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TypeBinding {
    /// Binding CONVERTS the value to the declared type. C `unsigned char c =
    /// 300` is 44; PHP weak mode coerces `"5"` to `5`. The common case, so it
    /// is the default a walker gets by writing a bare spelling.
    #[default]
    Converting,
    /// The declaration DOCUMENTS the value and never mutates it. Python's
    /// PEP 484 annotations; a type hint a dynamic language INFERRED for
    /// dispatch, which must not change what is stored.
    Descriptive,
    /// The declaration never converts, but the LANGUAGE'S OWN COMPILER
    /// statically enforces it — Kotlin's `var i: Int`, go's `i := 0`: a
    /// program storing another type there does not compile, so the runtime
    /// value is GUARANTEED without any coercion at the store. Distinct from
    /// `Descriptive` precisely because a Python annotation carries no such
    /// guarantee; a provable-type consumer (the numeric operator fold) may
    /// trust `Checked` exactly as it trusts `Converting`.
    Checked,
}

/// A declared type: its source spelling plus whether binding converts.
///
/// Derefs to `str`, so `Option<TypeHint>::as_deref()` still yields
/// `Option<&str>` and the several hundred sites that only want the spelling
/// keep working untouched. Only sites that care about binding say so.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TypeHint {
    spelling: String,
    pub binding: TypeBinding,
}

impl TypeHint {
    /// A declared type that converts on binding.
    pub fn converting(spelling: impl Into<String>) -> Self {
        TypeHint {
            spelling: spelling.into(),
            binding: TypeBinding::Converting,
        }
    }

    /// A type that documents but never mutates — an annotation, or a hint a
    /// dynamic language inferred for dispatch.
    pub fn descriptive(spelling: impl Into<String>) -> Self {
        TypeHint {
            spelling: spelling.into(),
            binding: TypeBinding::Descriptive,
        }
    }

    /// A declared type the language's own compiler statically enforces —
    /// no conversion at the store, guaranteed value. See
    /// [`TypeBinding::Checked`].
    pub fn checked(spelling: impl Into<String>) -> Self {
        TypeHint {
            spelling: spelling.into(),
            binding: TypeBinding::Checked,
        }
    }

    pub fn spelling(&self) -> &str {
        &self.spelling
    }

    /// Rewrite the spelling, KEEPING the binding.
    ///
    /// Walkers normalize spellings in place — trimming whitespace, folding VB's
    /// `Char` onto `String`, appending `()` for an array. None of that changes
    /// whether the declaration converts, so the binding must survive.
    pub fn set_spelling(&mut self, spelling: impl Into<String>) {
        self.spelling = spelling.into();
    }

    /// Append to the spelling, keeping the binding — the `()` array suffix.
    pub fn push_str(&mut self, suffix: &str) {
        self.spelling.push_str(suffix);
    }

    /// Does binding a value to this declaration convert it?
    pub fn converts(&self) -> bool {
        self.binding == TypeBinding::Converting
    }
}

impl std::ops::Deref for TypeHint {
    type Target = str;
    fn deref(&self) -> &str {
        &self.spelling
    }
}

impl From<String> for TypeHint {
    fn from(spelling: String) -> Self {
        TypeHint::converting(spelling)
    }
}

impl From<&str> for TypeHint {
    fn from(spelling: &str) -> Self {
        TypeHint::converting(spelling)
    }
}

impl std::fmt::Display for TypeHint {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.spelling)
    }
}

// ── Parameters ──────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct Param {
    pub name: String,
    pub type_hint: Option<TypeHint>,
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
    /// Copy-in / copy-out: the argument is passed BY VALUE and the caller writes
    /// the final value back into the argument's place after the call returns.
    ///
    /// This is an observably different mechanism from [`PassBy::Alias`], not an
    /// implementation of it. Two ways to tell them apart, both of which real
    /// languages can see:
    /// - a mutation is NOT visible through another binding DURING the call, and
    /// - if the callee THROWS, the write-back never runs and the mutation is
    ///   silently lost — no write-back can execute on a path that never returns.
    ///
    /// Languages still on this: pascal `var`, C# `ref`, VB `ByRef`, cobol
    /// `BY REFERENCE`, fortran `intent(inout)`. Most of them want `Alias` and
    /// are simply not migrated yet — migrate one at a time, with a differential
    /// test for each, per §3 of `referenceplan.md`.
    Ref,
    /// True aliasing: the argument is passed AS A REFERENCE, and the parameter
    /// is bound to it, so reads auto-deref and writes go through to the caller's
    /// storage. Nothing is written back, because nothing was copied.
    ///
    /// This is what php's `&$x` means, and what most `Ref` languages actually
    /// mean too. The place kinds all work: a name gives a cell, `&$a[i]` and
    /// `&$o->p` give a `(base, key)` carray — see `primitives/references.rs`.
    Alias,
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

// ── WASM catch clause (tag-identity, for `StmtKind::WasmTryTable`) ────────────

/// One catch clause of a [`StmtKind::WasmTryTable`]. Matching is by TAG
/// IDENTITY: `tag == None` is `catch_all`. The delivered payload values are
/// bound, in push order, to fresh locals named in `payload_binds`, which the
/// handler `body` reads. `capture_ref` marks the `catch_ref`/`catch_all_ref`
/// forms that also bind the exception's `exnref` (for `rethrow`/`delegate`,
/// lowered to `throw_ref`) into `exnref_bind`.
#[derive(Debug, Clone)]
pub struct WasmCatch {
    /// Tag name to match; `None` for `catch_all`.
    pub tag: Option<String>,
    /// Payload-value locals, bottom-to-top, that the handler body reads.
    pub payload_binds: Vec<String>,
    /// Whether the VM also delivers an `exnref` (`catch_ref`/`catch_all_ref`).
    pub capture_ref: bool,
    /// Local the captured `exnref` is bound to when `capture_ref` is set.
    pub exnref_bind: Option<String>,
    /// Handler statements, run with the payload bound.
    pub body: Vec<Statement>,
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

/// True when a statement list contains `yield` / `yield from` in the current
/// function scope. Nested functions, lambdas, and class expressions are scope
/// boundaries.
pub fn statements_contain_yield_outside_nested_scopes(stmts: &[Statement]) -> bool {
    stmts.iter().any(stmt_contains_yield_outside_nested_scopes)
}

fn stmt_contains_yield_outside_nested_scopes(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::FunctionDecl { .. } | StmtKind::ClassDecl { .. } => false,
        StmtKind::Expr(expr) | StmtKind::Assert { test: expr, .. } => {
            expr_contains_yield_outside_nested_scopes(expr)
        }
        StmtKind::Block(stmts) | StmtKind::NamespaceDecl { body: stmts, .. } => {
            statements_contain_yield_outside_nested_scopes(stmts)
        }
        StmtKind::VarDecl { declarations, .. } => declarations.iter().any(|decl| {
            decl.init
                .as_ref()
                .is_some_and(expr_contains_yield_outside_nested_scopes)
        }),
        StmtKind::Return(expr) | StmtKind::CloseFile(expr) => expr
            .as_ref()
            .is_some_and(expr_contains_yield_outside_nested_scopes),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            expr_contains_yield_outside_nested_scopes(cond)
                || statements_contain_yield_outside_nested_scopes(then_body)
                || elifs.iter().any(|(cond, body)| {
                    expr_contains_yield_outside_nested_scopes(cond)
                        || statements_contain_yield_outside_nested_scopes(body)
                })
                || else_body
                    .as_ref()
                    .is_some_and(|body| statements_contain_yield_outside_nested_scopes(body))
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            expr_contains_yield_outside_nested_scopes(cond)
                || statements_contain_yield_outside_nested_scopes(body)
                || else_body
                    .as_ref()
                    .is_some_and(|body| statements_contain_yield_outside_nested_scopes(body))
        }
        StmtKind::DoWhile { body, cond, .. } => {
            statements_contain_yield_outside_nested_scopes(body)
                || expr_contains_yield_outside_nested_scopes(cond)
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            init.as_ref()
                .is_some_and(|stmt| stmt_contains_yield_outside_nested_scopes(stmt))
                || cond
                    .as_ref()
                    .is_some_and(expr_contains_yield_outside_nested_scopes)
                || update
                    .as_ref()
                    .is_some_and(expr_contains_yield_outside_nested_scopes)
                || statements_contain_yield_outside_nested_scopes(body)
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            expr_contains_yield_outside_nested_scopes(iter)
                || statements_contain_yield_outside_nested_scopes(body)
                || else_body
                    .as_ref()
                    .is_some_and(|body| statements_contain_yield_outside_nested_scopes(body))
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            expr_contains_yield_outside_nested_scopes(expr)
                || cases
                    .iter()
                    .any(|case| statements_contain_yield_outside_nested_scopes(&case.body))
                || default
                    .as_ref()
                    .is_some_and(|body| statements_contain_yield_outside_nested_scopes(body))
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            statements_contain_yield_outside_nested_scopes(body)
                || catches
                    .iter()
                    .any(|catch| statements_contain_yield_outside_nested_scopes(&catch.body))
                || else_body
                    .as_ref()
                    .is_some_and(|body| statements_contain_yield_outside_nested_scopes(body))
                || finally
                    .as_ref()
                    .is_some_and(|body| statements_contain_yield_outside_nested_scopes(body))
        }
        StmtKind::With { body, .. }
        | StmtKind::Using { body, .. }
        | StmtKind::Lock { body, .. } => statements_contain_yield_outside_nested_scopes(body),
        StmtKind::Assign { targets, value, .. } => {
            targets
                .iter()
                .any(expr_contains_yield_outside_nested_scopes)
                || expr_contains_yield_outside_nested_scopes(value)
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            expr_contains_yield_outside_nested_scopes(target)
                || expr_contains_yield_outside_nested_scopes(value)
        }
        StmtKind::Throw { expr, cause } => {
            expr.as_ref()
                .is_some_and(expr_contains_yield_outside_nested_scopes)
                || cause
                    .as_ref()
                    .is_some_and(expr_contains_yield_outside_nested_scopes)
        }
        StmtKind::Labeled { body, .. } => stmt_contains_yield_outside_nested_scopes(body),
        StmtKind::Echo(exprs)
        | StmtKind::Delete(exprs)
        | StmtKind::RaiseEvent { args: exprs, .. } => {
            exprs.iter().any(expr_contains_yield_outside_nested_scopes)
        }
        StmtKind::Export {
            declaration,
            default,
            ..
        } => {
            declaration
                .as_ref()
                .is_some_and(|stmt| stmt_contains_yield_outside_nested_scopes(stmt))
                || default
                    .as_ref()
                    .is_some_and(|expr| expr_contains_yield_outside_nested_scopes(expr))
        }
        StmtKind::WasmTryTable { body, catches, .. } => {
            statements_contain_yield_outside_nested_scopes(body)
                || catches
                    .iter()
                    .any(|catch| statements_contain_yield_outside_nested_scopes(&catch.body))
        }
        _ => false,
    }
}

fn expr_contains_yield_outside_nested_scopes(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Yield(_) | ExprKind::YieldFrom(_) => true,
        ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_) | ExprKind::ClassExpr { .. } => false,
        ExprKind::RefOf(place) => match place.as_ref() {
            PlaceExpr::Ident(_) => false,
            PlaceExpr::Member { object, .. } => expr_contains_yield_outside_nested_scopes(object),
            PlaceExpr::Index { object, index, .. } => {
                expr_contains_yield_outside_nested_scopes(object)
                    || expr_contains_yield_outside_nested_scopes(index)
            }
            PlaceExpr::Deref(expr) => expr_contains_yield_outside_nested_scopes(expr),
        },
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::IsType { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::TypeOf(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Await(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr) => expr_contains_yield_outside_nested_scopes(expr),
        ExprKind::Async(op) => op
            .children()
            .into_iter()
            .any(expr_contains_yield_outside_nested_scopes),
        ExprKind::Chan(op) => op
            .children()
            .into_iter()
            .any(expr_contains_yield_outside_nested_scopes),
        ExprKind::Atomic(op) => op
            .children()
            .into_iter()
            .any(expr_contains_yield_outside_nested_scopes),
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
        }
        | ExprKind::StaticAccess {
            class: left,
            member: right,
        }
        | ExprKind::Index {
            object: left,
            index: right,
            ..
        } => {
            expr_contains_yield_outside_nested_scopes(left)
                || expr_contains_yield_outside_nested_scopes(right)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            expr_contains_yield_outside_nested_scopes(cond)
                || expr_contains_yield_outside_nested_scopes(then)
                || expr_contains_yield_outside_nested_scopes(else_)
        }
        ExprKind::Member { object, .. } => expr_contains_yield_outside_nested_scopes(object),
        ExprKind::CallableRef {
            target,
            receiver,
            adapter,
            ..
        } => {
            expr_contains_yield_outside_nested_scopes(target)
                || receiver.as_ref().map_or(false, |expr| {
                    expr_contains_yield_outside_nested_scopes(expr)
                })
                || adapter.as_ref().map_or(false, |adapter| match adapter {
                    CallableAdapter::Expr { body, .. } => {
                        expr_contains_yield_outside_nested_scopes(body)
                    }
                })
        }
        ExprKind::Call { callee, args, .. } => {
            expr_contains_yield_outside_nested_scopes(callee)
                || args
                    .iter()
                    .any(|arg| expr_contains_yield_outside_nested_scopes(&arg.value))
        }
        ExprKind::New { class, args } => {
            expr_contains_yield_outside_nested_scopes(class)
                || args
                    .iter()
                    .any(|arg| expr_contains_yield_outside_nested_scopes(&arg.value))
        }
        ExprKind::SuperCall { args, .. } => args
            .iter()
            .any(|arg| expr_contains_yield_outside_nested_scopes(&arg.value)),
        ExprKind::Array(elems) => elems.iter().any(|elem| {
            expr_contains_yield_outside_nested_scopes(&elem.value)
                || elem
                    .key
                    .as_ref()
                    .is_some_and(expr_contains_yield_outside_nested_scopes)
        }),
        ExprKind::Tuple(exprs)
        | ExprKind::Set(exprs)
        | ExprKind::Sequence(exprs)
        | ExprKind::Zip {
            iterables: exprs, ..
        } => exprs.iter().any(expr_contains_yield_outside_nested_scopes),
        ExprKind::NamedTuple { fields, .. } => fields
            .iter()
            .any(|(_, expr)| expr_contains_yield_outside_nested_scopes(expr)),
        ExprKind::Object(props) => props.iter().any(|prop| match prop {
            ObjectProperty::KeyValue { key, value } | ObjectProperty::Computed { key, value } => {
                expr_contains_yield_outside_nested_scopes(key)
                    || expr_contains_yield_outside_nested_scopes(value)
            }
            ObjectProperty::Spread(expr) => expr_contains_yield_outside_nested_scopes(expr),
            ObjectProperty::Shorthand(_)
            | ObjectProperty::Method { .. }
            | ObjectProperty::Accessor { .. } => false,
        }),
        ExprKind::Interpolation(parts) => parts.iter().any(|part| match part {
            InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                expr_contains_yield_outside_nested_scopes(expr)
            }
            InterpolPart::Text(_) => false,
        }),
        ExprKind::Match { subject, arms } => {
            expr_contains_yield_outside_nested_scopes(subject)
                || arms.iter().any(|arm| {
                    arm.conditions.as_ref().is_some_and(|conditions| {
                        conditions
                            .iter()
                            .any(expr_contains_yield_outside_nested_scopes)
                    }) || expr_contains_yield_outside_nested_scopes(&arm.body)
                })
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            expr_contains_yield_outside_nested_scopes(element)
                || generators.iter().any(|generator| {
                    expr_contains_yield_outside_nested_scopes(&generator.iter)
                        || generator
                            .conditions
                            .iter()
                            .any(expr_contains_yield_outside_nested_scopes)
                })
        }
        ExprKind::Slice { lower, upper, step } => [lower, upper, step].iter().any(|expr| {
            expr.as_ref()
                .is_some_and(|expr| expr_contains_yield_outside_nested_scopes(expr))
        }),
        _ => false,
    }
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

/// Whether a spread argument may bind by NAME. See [`Directives::spread_arguments`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SpreadArguments {
    /// `f(...xs)` unpacks POSITIONALLY only. Every language but php.
    Positional,
    /// A string-keyed array binds by PARAMETER NAME; anything else is
    /// positional. php.
    PositionalOrNamed,
}

/// What the argument to a terminate builtin means. See [`Directives::exit_argument`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExitArgument {
    /// The argument is the exit STATUS. Every language but php.
    Status,
    /// A string argument is a farewell MESSAGE (printed, status 0); anything
    /// else is the STATUS. php `exit` / `die`.
    MessageOrStatus,
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

/// How a scope resolves a name that is not one of its own locals.
///
/// All three are statements about resolution, which is why they share a node:
/// `Closed` sets the scope's policy, `Global` and `Nonlocal` re-open individual
/// names within it.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ScopeDeclKind {
    /// This scope does NOT chain outward: a name that misses its locals
    /// resolves to nothing — it reads null and an assignment creates a local
    /// here — instead of falling through to the module globals. PHP function
    /// bodies, where a module `$x` is invisible without `global $x;`.
    ///
    /// `names` seeds the exceptions: names that stay open regardless of the
    /// policy. PHP's superglobals (`$_SERVER`, `$_GET`, …) are visible in every
    /// scope without being imported, and that list is PHP's to supply.
    Closed,
    /// Re-open these names to the module scope for the rest of this scope.
    /// PHP `global $x;`, Python `global x`.
    Global,
    /// Re-open these names to the enclosing function scope. Python `nonlocal`.
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

/// How records are arranged in the file — COBOL `ORGANIZATION`, VB `Open For`.
///
/// A property of the DECLARATION, not a directive: it is fixed where the file
/// is declared and is the same fact at every site that touches it, in any file
/// and any language (`directives.md` §3, question 3).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FileOrganization {
    /// Fixed-width records back to back. Record *n* is at *n × width*.
    #[default]
    Sequential,
    /// Newline-delimited text. Records are NOT fixed width, so positioned
    /// addressing does not apply — COBOL LINE SEQUENTIAL.
    Line,
    /// Fixed-width records addressed by 1-based record number.
    Relative,
    /// Fixed-width records addressed by key. See `keys` on the declaration.
    Indexed,
}

/// What the program may do with the file, and what opening it does to what is
/// already there — COBOL `OPEN INPUT/OUTPUT/I-O/EXTEND`, VB `Open For`, Pascal
/// `Reset`/`Rewrite`/`Append`.
///
/// Named for the intent, not for one language's spelling, and it maps onto
/// WASI's two flag words without interpretation: `descriptor-flags` says which
/// directions are permitted, `open-flags` says whether to create and whether
/// to truncate.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum OpenMode {
    /// Existing contents, read only. Opening a file that is not there is an
    /// error, not an empty file.
    #[default]
    Read,
    /// A fresh file. Creates it, and truncates whatever was there — COBOL
    /// `OPEN OUTPUT` starts an empty file even over a full one.
    Write,
    /// Both directions over existing contents, creating the file if absent
    /// but never truncating. COBOL `OPEN I-O`, which is what `REWRITE` needs.
    ReadWrite,
    /// Writes land after the existing contents. COBOL `OPEN EXTEND`.
    Append,
}

/// How the program intends to reach records — COBOL `ACCESS MODE`.
///
/// Distinct from [`FileOrganization`]: an INDEXED file may be read
/// sequentially. Organization is how the bytes lie; access is how this program
/// walks them.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FileAccess {
    #[default]
    Sequential,
    Random,
    /// Either, chosen per statement. COBOL DYNAMIC.
    Dynamic,
}

/// A key field of an INDEXED file, in priority order.
#[derive(Debug, Clone, PartialEq)]
pub struct RecordKey {
    /// Field name within the record type.
    pub field: String,
    /// A primary key rejects duplicates; an alternate key may allow them
    /// (COBOL `WITH DUPLICATES`).
    pub duplicates: bool,
}

/// Which way the bytes move.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RecordDirection {
    Read,
    Write,
    /// Replace the record already at `at` — COBOL `REWRITE`, VB `Put` over an
    /// existing record. Distinct from `Write`, which may extend the file.
    Rewrite,
    Delete,
}

/// WHICH record a transfer addresses.
///
/// No `PartialEq`: the payload is an `Expression`, which has none — two
/// addresses are compared by evaluating them, not by comparing their syntax.
#[derive(Debug, Clone)]
pub enum RecordAddress {
    /// The record most recently transferred — COBOL `REWRITE` after a `READ`,
    /// VB `Put #n` with no record number after a `Get`.
    ///
    /// Distinct from [`Self::Next`] rather than expressible as it: a
    /// sequential read leaves the position ON the following record, so
    /// "rewrite what I just read" is one record BEHIND where the next
    /// transfer would land. Spelling it `Next` would overwrite the record
    /// after the one the program meant, and the file would still look
    /// plausible.
    Current,
    /// The next record in sequence — the file's position advances.
    Next,
    /// A 1-based record number. COBOL RELATIVE, Fortran `rec=n`, VB `Get #f, n`.
    ///
    /// ⚠ 1-BASED in every source language that has it. The lowering subtracts
    /// one exactly once, when computing the byte offset.
    Number(Expression),
    /// By key value — COBOL `START`/`READ KEY`, indexed files only.
    Key {
        /// Index into the declaration's `keys`.
        key_index: usize,
        value: Expression,
        /// `START ... KEY IS >= X` and friends; `Equal` for a plain keyed read.
        relation: FileKeyRelation,
    },
}

#[derive(Debug, Clone)]
pub struct ExportName {
    pub name: String,
    pub alias: Option<String>,
}

// ════════════════════════════════════════════════════════════════════════════
// Operators
// ════════════════════════════════════════════════════════════════════════════

/// The integer lane a bit operation works in.
///
/// This is NOT a directive, and the distinction matters. The width of a bit
/// operation is a property of the OPERAND'S DECLARED TYPE — Fortran writes
/// `integer(kind=8)`, Go writes `uint32` vs `uint64`, Java encodes it in the
/// spelling (`Integer.bitCount` vs `Long.bitCount`). A language-wide default
/// would give that fact two homes and they would drift.
///
/// wasm agrees: `i32.popcnt` and `i64.popcnt` are different instructions, and
/// the answers genuinely differ — `LeadingZeros(1)` is 31 in `W32` and 63 in
/// `W64`. So the walker, which can read the declared kind from its own source,
/// states the lane; nothing downstream has to infer it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BitLane {
    W32,
    W64,
}

/// The float lane an operation works in — the storage width whose neighbours,
/// ULP and rounding are being asked about.
///
/// The same call as [`BitLane`], for the same reason: it is a property of the
/// operand's DECLARED type, not of a region. Fortran's default `real` is kind 4
/// and `double precision` is kind 8, C writes `float` vs `double`, .NET writes
/// `Single` vs `Double`. And the answers genuinely differ — `SPACING(1.0)` is
/// 2⁻²³ in `F32` and 2⁻⁵² in `F64`, so a shared implementation that picks one
/// lane is simply wrong for the other language half the time.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FloatLane {
    F32,
    F64,
}

/// Which way a value exactly halfway between two integers goes.
///
/// Not one operation and not a directive: it is what a given language's `round`
/// MEANS, so it is selected per spelling by the profile row, the way
/// [`BitLane`] is selected per operand. Languages genuinely disagree, and
/// verified against every installed toolchain they disagree three ways.
///
/// | policy | 2.5 | −2.5 | 0.5 | languages |
/// |---|---|---|---|---|
/// | `HalfEven` | 2 | −2 | 0 | Python, Pascal, Kotlin, C#, VB, wasm `f64.nearest` |
/// | `HalfAwayFromZero` | 3 | −3 | 1 | C, PHP, Go, Fortran `NINT`, Ruby, Dart |
/// | `HalfUp` | 3 | −2 | 1 | JS, Java |
///
/// ⛔ Distinct from the IEEE rounding DIRECTION of
/// `proposals/rounding-mode-control`. A direction applies to an inexact
/// ARITHMETIC result and is lexical (C `#pragma STDC FENV_ROUND`), so that one
/// is a directive. This applies to round-to-integral, and `HalfAwayFromZero` /
/// `HalfUp` are language conventions IEEE has no name for.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum MidpointPolicy {
    /// Ties to the even neighbour — banker's rounding, and what
    /// `f64.nearest` does, so it is the zero-instruction default.
    #[default]
    HalfEven,
    /// Ties away from zero. Sign-preserving: `round(-0.2)` is `-0.0`.
    HalfAwayFromZero,
    /// Ties toward +∞ — `floor(x + 0.5)`, which is Java's spec verbatim.
    HalfUp,
}

impl FloatLane {
    /// Bits of mantissa precision, including the implicit leading one.
    pub fn mantissa_bits(self) -> u32 {
        match self {
            FloatLane::F32 => 24,
            FloatLane::F64 => 53,
        }
    }
}

impl BitLane {
    /// Bits in the lane — Fortran's `BIT_SIZE`, Java's `SIZE`.
    pub fn bits(self) -> u32 {
        match self {
            BitLane::W32 => 32,
            BitLane::W64 => 64,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum BinOp {
    /// Rotate left — the bits shifted out re-enter at the other end, so
    /// nothing is lost. Fortran `ISHFTC`, Go `bits.RotateLeft*`, Java
    /// `Integer.rotateLeft`, C# `BitOperations.RotateLeft`, wasm `i32.rotl`.
    /// Distinct from [`BinOp::Shl`], which discards them.
    RotL(BitLane),
    /// Rotate right. Fortran `ISHFTC` with a negative count, Go
    /// `bits.RotateLeft*` with a negative count, Java `Integer.rotateRight`,
    /// C# `BitOperations.RotateRight`, wasm `i32.rotr`.
    RotR(BitLane),
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

/// The async model — one set of concepts behind seventeen spellings.
///
/// | op | JS | C# | Python | Dart |
/// |---|---|---|---|---|
/// | `Resolved` | `Promise.resolve` | `Task.FromResult`, `Task.CompletedTask` | `asyncio.ensure_future` / `create_task`¹ | `Future.value` |
/// | `Rejected` | `Promise.reject` | `Task.FromException` | — | `Future.error` |
/// | `Continue` | `.then` | `.ContinueWith` | — | `.then` |
/// | `Cleanup` | `.finally` | — | — | `.whenComplete` |
/// | `Join` | `all`/`allSettled`/`race`/`any` | `WhenAll` | `asyncio.gather` | `Future.wait` |
/// | `Spawn` | — | `Task.Run` | `loop.run_in_executor` | — |
/// | `Sleep` | — | `Task.Delay` | `asyncio.sleep` | `Future.delayed` |
/// | `BlockOn` | — | `.GetAwaiter().GetResult()` | `asyncio.run` | — |
/// | `AwaitEager` | — | `await` (continuation may be synchronous) | — | — |
///
/// ¹ `PromiseResolve` (§27.2.4.7) adopts a thenable, which is exactly
/// `ensure_future`'s contract for an already-started coroutine.
///
/// Deliberately NOT mapped: `Task.WhenAny` — it settles with the completed
/// TASK where `Promise.race` settles with its VALUE; a wrong mapping is worse
/// than a missing one. Channels (`go`/Rust/Kotlin) are the NEXT vocabulary —
/// CSP is its own model and is not forced into promises.
///
/// The lowering (`primitives/async_ops.rs`) targets `ecma:promise` (§27.2) and
/// the JSPI await imports — under the hood wasm/ecma. Where languages'
/// semantics genuinely differ (eager vs deferred await), the difference is a
/// DIFFERENT operation in this vocabulary, chosen at normalization — never a
/// runtime-consulted property.
#[derive(Debug, Clone)]
pub enum AsyncOp {
    /// §27.2.4.7 PromiseResolve: an already-settled (or adopted) async value.
    Resolved(Box<Expression>),
    /// §27.2.4.6 Promise.reject.
    Rejected(Box<Expression>),
    /// §27.2.5.4 then — chain a continuation (and optionally a handler).
    Continue {
        source: Box<Expression>,
        on_fulfilled: Option<Box<Expression>>,
        on_rejected: Option<Box<Expression>>,
    },
    /// §27.2.5.3 finally — runs on settle, passes the outcome through.
    Cleanup {
        source: Box<Expression>,
        on_settled: Box<Expression>,
    },
    /// §27.2.4.1-4.5: combine many async values into one.
    Join {
        mode: JoinMode,
        sources: Vec<Expression>,
    },
    /// Run a callable as scheduled work; the result is an async value.
    Spawn(Box<Expression>),
    /// An async value that settles after a duration (milliseconds) — one
    /// concept behind `Task.Delay`, `asyncio.sleep`, `Future.delayed`. Time
    /// itself comes from the HOST's timer surface at lowering; the vocabulary
    /// only says "later".
    Sleep(Box<Expression>),
    /// Await with EAGER continuation: if the antecedent is already settled,
    /// continue synchronously with its value (throw its rejection); suspend
    /// only when pending. This is .NET's contract — a completed Task may run
    /// its continuation on the completing thread — and it is a DIFFERENT
    /// OPERATION from `ExprKind::Await`, which is ECMA-262 §6.2.3.1 and
    /// always yields one turn. The distinction lives HERE, on the node,
    /// normalized by the walker — never in a runtime-consulted property —
    /// so any consumer of the tree (an exporter to Java, another backend)
    /// reads the semantics off the operation itself.
    AwaitEager(Box<Expression>),
    /// Yield one full turn of the ready queue: the fiber requeues at the
    /// BACK so every already-queued job runs first. Never continues
    /// synchronously — even under eager-await semantics. C# `Task.Yield`,
    /// the async channel surface's polling tick.
    Yield,
    /// Synchronously drive the loop until `source` settles; yield its value
    /// (throw its rejection). The sync↔async boundary: `GetAwaiter().GetResult()`,
    /// `asyncio.run`. Lowers to the JSPI suspend at an async-capable boundary.
    BlockOn(Box<Expression>),
}

impl AsyncOp {
    /// Whether this operation PRODUCES an async value rather than consuming
    /// one — i.e. whether the expression it forms is still a task/promise.
    ///
    /// ⛔ **The distinction is the whole point.** `Resolved`, `Spawn`, `Join`
    /// and friends yield a pending value; `AwaitEager`, `BlockOn` and `Yield`
    /// UNWRAP one, so their result is whatever the task carried. A frontend
    /// that rewrites `Task.Run(…)` into an `AsyncOp` erases the .NET type the
    /// declaration had, and its inference needs this back — otherwise
    /// `Dim t = Task.Run(...)` then `t.Result` falls through to an ordinary
    /// member read and answers `undefined` (measured: VB
    /// `vb_task_run_exception_capture` 13 -> 17 while this was missing).
    ///
    /// It lives here rather than in each walker because it is a fact about the
    /// NODE, and every typed frontend that normalises onto this vocabulary asks
    /// the same question — C# already carries a private copy
    /// (`csharp_task_valued_type`), which is one table in two places.
    pub fn yields_async_value(&self) -> bool {
        match self {
            AsyncOp::Resolved(_)
            | AsyncOp::Rejected(_)
            | AsyncOp::Spawn(_)
            | AsyncOp::Sleep(_)
            | AsyncOp::Join { .. }
            | AsyncOp::Continue { .. }
            | AsyncOp::Cleanup { .. } => true,
            // These CONSUME an async value; the result is what it carried.
            AsyncOp::AwaitEager(_) | AsyncOp::BlockOn(_) | AsyncOp::Yield => false,
        }
    }

    /// Every child expression, in evaluation order — for the structural
    /// traversals (yield detection, span walks) that must not skip async
    /// operands.
    pub fn children(&self) -> Vec<&Expression> {
        match self {
            AsyncOp::Resolved(e)
            | AsyncOp::Rejected(e)
            | AsyncOp::Spawn(e)
            | AsyncOp::Sleep(e)
            | AsyncOp::AwaitEager(e)
            | AsyncOp::BlockOn(e) => {
                vec![e]
            }
            AsyncOp::Continue {
                source,
                on_fulfilled,
                on_rejected,
            } => {
                let mut v: Vec<&Expression> = vec![source];
                v.extend(on_fulfilled.iter().map(|b| &**b));
                v.extend(on_rejected.iter().map(|b| &**b));
                v
            }
            AsyncOp::Cleanup { source, on_settled } => vec![source, on_settled],
            AsyncOp::Yield => Vec::new(),
            AsyncOp::Join { sources, .. } => sources.iter().collect(),
        }
    }

    /// Mutable [`AsyncOp::children`], for walker rewrite passes.
    pub fn children_mut(&mut self) -> Vec<&mut Expression> {
        match self {
            AsyncOp::Resolved(e)
            | AsyncOp::Rejected(e)
            | AsyncOp::Spawn(e)
            | AsyncOp::Sleep(e)
            | AsyncOp::AwaitEager(e)
            | AsyncOp::BlockOn(e) => {
                vec![e]
            }
            AsyncOp::Continue {
                source,
                on_fulfilled,
                on_rejected,
            } => {
                let mut v: Vec<&mut Expression> = vec![source];
                v.extend(on_fulfilled.iter_mut().map(|b| &mut **b));
                v.extend(on_rejected.iter_mut().map(|b| &mut **b));
                v
            }
            AsyncOp::Cleanup { source, on_settled } => vec![source, on_settled],
            AsyncOp::Yield => Vec::new(),
            AsyncOp::Join { sources, .. } => sources.iter_mut().collect(),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum JoinMode {
    /// First rejection rejects (§27.2.4.1 all).
    All,
    /// Never rejects; outcomes recorded (§27.2.4.2 allSettled).
    AllSettled,
    /// First SETTLED wins with its value (§27.2.4.5 race).
    Race,
    /// First FULFILLED wins; all-rejected → AggregateError (§27.2.4.3 any).
    Any,
}

/// The channel (CSP) vocabulary — one model behind every language's spelling.
///
/// | op | Go | Rust | Kotlin |
/// |----|----|------|--------|
/// | `New` | `make(chan T, n)` | `mpsc::channel` | `Channel<T>(n)` |
/// | `Send` | `ch <- v` | `tx.send(v)` | `ch.send(v)` |
/// | `Recv` | `<-ch` | `rx.recv()` | `ch.receive()` |
/// | `RecvOk` | `v, ok := <-ch` | `recv().ok()` | `receiveCatching` |
/// | `Len`/`Cap` | `len(ch)`/`cap(ch)` | — | — |
/// | `Close` | `close(ch)` | drop tx | `ch.close()` |
///
/// Semantics live in the ONE lowering (`primitives/channels.rs`), Go-spec
/// anchored: receive on a closed channel drains the buffer then yields the
/// element ZERO VALUE with `ok == false`; send/close on a closed channel
/// panic; a nil channel is never ready. The zero value is normalized onto
/// `New` by the walker (which knows the declared element type) and travels
/// WITH the channel — any consumer of the tree reads the semantics off the
/// operation, never off a runtime property.
///
/// Blocking `Send`/`Recv` (empty-buffer rendezvous) is fiber + scheduler
/// territory and lands on the `DeferredSource`/scheduler seam; until then
/// the lowering keeps the historical non-blocking shapes.
#[derive(Debug, Clone)]
pub enum ChanOp {
    /// `make(chan T, capacity?)` — `zero` is T's zero value, stored with the
    /// channel so closed-receive can produce it far from the declaration.
    New {
        capacity: Option<Box<Expression>>,
        zero: Box<Expression>,
    },
    Send {
        channel: Box<Expression>,
        value: Box<Expression>,
    },
    /// Receive the value alone (`<-ch`).
    Recv(Box<Expression>),
    /// Receive `(value, ok)` — `ok == false` iff the channel is closed AND
    /// drained (Go spec: a closed channel first yields its buffered values).
    RecvOk(Box<Expression>),
    Len(Box<Expression>),
    Cap(Box<Expression>),
    Close(Box<Expression>),
    /// Non-suspending send: `true` iff the value was accepted (room and not
    /// closed). .NET `Writer.TryWrite`, Kotlin `trySend`, Rust `try_send`;
    /// Go spells it `select { case ch <- v: ... default: ... }`.
    TrySend {
        channel: Box<Expression>,
        value: Box<Expression>,
    },
    /// Non-suspending receive of `(value, ok)` — `ok == false` when nothing
    /// is buffered (empty OR closed-and-drained; the value half is then the
    /// channel's zero). .NET `Reader.TryRead(out v)`, Kotlin `tryReceive`,
    /// Rust `try_recv`.
    TryRecv(Box<Expression>),
    /// Non-consuming read of `(value, ok)` — the head stays buffered.
    /// .NET `Reader.TryPeek(out v)`.
    TryPeek(Box<Expression>),
    /// `true` iff the channel is closed AND drained — the point where a
    /// consumer is definitively done. .NET `Reader.Completion.IsCompleted`,
    /// Kotlin `isClosedForReceive`.
    Drained(Box<Expression>),
    /// `true` iff the channel is closed for WRITING — buffered values may
    /// remain readable. .NET `Writer.TryWrite` returning false after
    /// `Complete()`, Kotlin `isClosedForSend`. Distinct from [`Drained`]:
    /// closed-with-backlog is Closed but not yet Drained.
    Closed(Box<Expression>),
    /// Blocking receive that THROWS `error` when the channel is closed and
    /// drained instead of yielding the zero value. The failure value is
    /// language policy, declared here: .NET `ReadAsync` →
    /// ChannelClosedException, Rust `recv()` → RecvError.
    RecvOrFail {
        channel: Box<Expression>,
        error: Box<Expression>,
    },
    /// Block until the channel is READABLE (buffered value present) or
    /// definitively done; yields the bool "a read will succeed". .NET
    /// `WaitToReadAsync`.
    WaitReadable(Box<Expression>),
}

/// The atomic vocabulary — one model behind every language's spelling, and the
/// WASM threads proposal is the substrate underneath all of them.
///
/// Every operand that names storage is a PLACE, not a value: an atomic acts on
/// a word in SHARED linear memory, so `place` must resolve to an address. That
/// is the whole reason the old per-language answers were wrong — C# handed the
/// atomic its variable's VALUE as an address, and C, Go and the JVM languages
/// gave up and emitted a plain read-modify-write, which is not atomic at all.
///
/// Quirks ride the node, per `documentation/directives.md` §3: none of them
/// governs a REGION of code (question 1), so none is a `Directives` entry.
/// They describe the operation being invoked (question 3) and therefore belong
/// to it — `getAndAdd` and `addAndGet` differ per CALL and appear in the same
/// file, which a lexical policy could never express.
#[derive(Debug, Clone)]
pub enum AtomicOp {
    /// Atomic read. C `atomic_load`, .NET `Interlocked.Read`, Java `get`.
    Load {
        place: Box<Expression>,
        ordering: MemoryOrder,
    },
    /// Atomic write. C `atomic_store`, Java `set`.
    Store {
        place: Box<Expression>,
        value: Box<Expression>,
        ordering: MemoryOrder,
    },
    /// Read-modify-write. C `atomic_fetch_*`, .NET `Interlocked.Add` /
    /// `Increment` / `Decrement` / `Exchange`, Go `atomic.AddInt32`, Java
    /// `getAndAdd` / `addAndGet`, Pascal `TInterlocked.Add`.
    Rmw {
        op: AtomicRmw,
        place: Box<Expression>,
        operand: Box<Expression>,
        result: RmwResult,
        ordering: MemoryOrder,
    },
    /// Compare-and-swap. The field NAMES settle the operand order that every
    /// language spells differently: .NET writes
    /// `CompareExchange(ref location, value, comparand)` — `comparand` is the
    /// EXPECTED and `value` the REPLACEMENT, the reverse of WASM's
    /// `cmpxchg(addr, expected, replacement)`. Each walker maps its own
    /// spelling onto these names once, and nothing downstream can get it
    /// backwards.
    CompareExchange {
        place: Box<Expression>,
        expected: Box<Expression>,
        replacement: Box<Expression>,
        result: RmwResult,
        ordering: MemoryOrder,
    },
    /// A standalone barrier. C `atomic_thread_fence`,
    /// .NET `Interlocked.MemoryBarrier`, Go `runtime.KeepAlive` ordering points.
    Fence { ordering: MemoryOrder },
}

/// The read-modify-write operations WASM provides directly.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AtomicRmw {
    Add,
    Sub,
    And,
    Or,
    Xor,
    /// Unconditional swap — .NET `Interlocked.Exchange`, Java `getAndSet`.
    Xchg,
}

/// Which value a read-modify-write yields.
///
/// WASM's `i32.atomic.rmw.*` always yields the OLD value, and so do C's
/// `atomic_fetch_*` and Java's `getAndAdd`. .NET's `Interlocked.Add` /
/// `Increment` / `Decrement`, Go's `atomic.AddInt32` and Java's `addAndGet`
/// yield the NEW one. One field, decided by the walker that knows which
/// function was called — not a second emitter per language.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RmwResult {
    Old,
    New,
}

/// Memory ordering. C names it per call (`memory_order_relaxed`); Go, .NET and
/// Pascal specify sequential consistency and nothing else, so their walkers
/// fill `SeqCst` and the distinction costs them nothing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MemoryOrder {
    Relaxed,
    Acquire,
    Release,
    AcqRel,
    SeqCst,
}

impl AtomicOp {
    /// Every child expression, in evaluation order — the same contract
    /// [`AsyncOp::children`] keeps, so the structural traversals cannot skip
    /// an atomic's operands.
    pub fn children(&self) -> Vec<&Expression> {
        match self {
            AtomicOp::Load { place, .. } => vec![place],
            AtomicOp::Store { place, value, .. } => vec![place, value],
            AtomicOp::Rmw { place, operand, .. } => vec![place, operand],
            AtomicOp::CompareExchange {
                place,
                expected,
                replacement,
                ..
            } => vec![place, expected, replacement],
            AtomicOp::Fence { .. } => Vec::new(),
        }
    }

    pub fn children_mut(&mut self) -> Vec<&mut Expression> {
        match self {
            AtomicOp::Load { place, .. } => vec![place.as_mut()],
            AtomicOp::Store { place, value, .. } => vec![place.as_mut(), value.as_mut()],
            AtomicOp::Rmw { place, operand, .. } => vec![place.as_mut(), operand.as_mut()],
            AtomicOp::CompareExchange {
                place,
                expected,
                replacement,
                ..
            } => vec![place.as_mut(), expected.as_mut(), replacement.as_mut()],
            AtomicOp::Fence { .. } => Vec::new(),
        }
    }
}

impl ChanOp {
    pub fn children(&self) -> Vec<&Expression> {
        match self {
            ChanOp::New { capacity, zero } => {
                let mut v: Vec<&Expression> = Vec::new();
                if let Some(c) = capacity {
                    v.push(c);
                }
                v.push(zero);
                v
            }
            ChanOp::Send { channel, value } | ChanOp::TrySend { channel, value } => {
                vec![channel, value]
            }
            ChanOp::RecvOrFail { channel, error } => vec![channel, error],
            ChanOp::Recv(e)
            | ChanOp::RecvOk(e)
            | ChanOp::Len(e)
            | ChanOp::Cap(e)
            | ChanOp::Close(e)
            | ChanOp::TryRecv(e)
            | ChanOp::TryPeek(e)
            | ChanOp::Drained(e)
            | ChanOp::Closed(e)
            | ChanOp::WaitReadable(e) => vec![e],
        }
    }

    pub fn children_mut(&mut self) -> Vec<&mut Expression> {
        match self {
            ChanOp::New { capacity, zero } => {
                let mut v: Vec<&mut Expression> = Vec::new();
                if let Some(c) = capacity {
                    v.push(c);
                }
                v.push(zero);
                v
            }
            ChanOp::Send { channel, value } | ChanOp::TrySend { channel, value } => {
                vec![channel, value]
            }
            ChanOp::RecvOrFail { channel, error } => vec![channel, error],
            ChanOp::Recv(e)
            | ChanOp::RecvOk(e)
            | ChanOp::Len(e)
            | ChanOp::Cap(e)
            | ChanOp::Close(e)
            | ChanOp::TryRecv(e)
            | ChanOp::TryPeek(e)
            | ChanOp::Drained(e)
            | ChanOp::Closed(e)
            | ChanOp::WaitReadable(e) => vec![e],
        }
    }
}

/// One arm of a `select` — the communication (used for the READINESS test)
/// and the body. The body's first statement performs the communication and
/// binds its results (`v, ok := ChanOp::RecvOk(ch)` as a plain declaration),
/// so binding, scoping and destructuring ride the ordinary statement
/// machinery instead of a parallel surface.
#[derive(Debug, Clone)]
pub struct SelectArm {
    pub comm: ChanOp,
    pub body: Vec<Statement>,
}

/// A numeric storage representation — the four wasm value types a bit cast can
/// read. Reinterpretation only ever swaps a float for an integer of the SAME
/// width, so naming the target implies the source and an invalid pairing
/// cannot be spelled.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumericRepr {
    I32,
    F32,
    I64,
    F64,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum UnaryOp {
    /// Read this value's STORAGE as another type of the same width — a bit
    /// cast, not a conversion. `Reinterpret(I32)` of `1.0f32` is `1065353216`,
    /// not `1`. Fortran `TRANSFER`, Go `math.Float32bits`, Java
    /// `Float.floatToIntBits`, C#/VB `BitConverter.SingleToInt32Bits`, C
    /// `union`, C++ `std::bit_cast`, Rust `transmute`, JS `DataView`.
    Reinterpret(NumericRepr),
    /// The number of one bits. Fortran `POPCNT`, Go `bits.OnesCount*`, Java
    /// `Integer/Long.bitCount`, C# `BitOperations.PopCount`, C
    /// `__builtin_popcount`, Python `int.bit_count`, Rust `count_ones`.
    PopCount(BitLane),
    /// Zero bits above the most significant one bit. Fortran `LEADZ`, Go
    /// `bits.LeadingZeros*`, Java `numberOfLeadingZeros`, JS `Math.clz32`,
    /// C# `LeadingZeroCount`.
    LeadingZeros(BitLane),
    /// Zero bits below the least significant one bit. Fortran `TRAILZ`, Go
    /// `bits.TrailingZeros*`, Java `numberOfTrailingZeros`, C#
    /// `TrailingZeroCount`.
    TrailingZeros(BitLane),
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

/// Cross-language operator + protocol method identity — the PROTOCOL SLOT.
/// Every language that defines one of these concepts under its own name
/// resolves to the same variant: Python `__str__`, Ruby `to_s`, PHP
/// `__toString`, C# `ToString` are one slot, not five spellings.
///
/// A slot is reached by its [`slot_id`](ProtocolSlot::slot_id) — an
/// integer — never by a name. That is the whole point: a shared *string* like
/// `"tostring"` lives in the identifier namespace, so it collides with a user
/// method of that name in one direction and cannot carry a per-language
/// signature in the other (flexclassplan.md §1e, §2a, §2g). Construction
/// already works this way — `ExprKind::New` names no constructor in any of the
/// 261 walker sites that emit it — and this is that mechanism for the rest.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ProtocolSlot {
    // ── Lifecycle ───────────────────────────────────────────────────
    /// The class's destructor / finaliser. PHP `__destruct`, Python
    /// `__del__`, C# `~Foo()`, Pascal `destructor Destroy`, VB `Finalize`.
    ///
    /// Every one of those was previously re-derived by a bespoke string test
    /// inside the language's own normalizer — four different checks in four
    /// crates for one concept. It resolves here like every other cross-language
    /// method identity, so a language declares its spelling once and the
    /// normalizer routes on the KIND.
    Destructor,

    // ── Coercion / representation ───────────────────────────────────
    ToString,    // JS toString, C# ToString, Python __str__, Ruby to_s, PHP __toString
    Repr,        // Python __repr__, Ruby inspect, PHP __debugInfo
    ValueOf,     // JS valueOf
    ToPrimitive, // JS Symbol.toPrimitive
    /// Truthiness — Python `__bool__`, C# `operator true`, Ruby `truthy?`.
    /// Distinct from [`ValueOf`](ProtocolSlot::ValueOf): a class can define
    /// both, and `if (obj)` must reach this one.
    Bool,
    Int,   // Python __int__
    Float, // Python __float__
    /// Coerce to a CHARACTER — C# `(char)n`, Pascal `Chr(n)`. The mirror of
    /// [`Int`](ProtocolSlot::Int) on a char, which reads a code point out;
    /// this one puts one back. Named by `builtinslotplan.md` as the
    /// replacement for `arrays.rs`'s `name == "pascal"` char coercion.
    Char,
    Bytes,  // Python __bytes__
    Format, // Python __format__
    /// Serialization hooks — PHP `__serialize` / `__sleep`, `jsonSerialize`.
    Serialize,
    /// The reverse — PHP `__unserialize` / `__wakeup`.
    Deserialize,
    /// Explicit copy — PHP `__clone`, Java `clone`, C# `Clone`, Python
    /// `__copy__`.
    Clone,

    // ── Iteration ───────────────────────────────────────────────────
    Iterator,      // JS Symbol.iterator, Python __iter__, Ruby each, C# GetEnumerator
    AsyncIterator, // JS Symbol.asyncIterator, Python __aiter__
    Next,          // JS iterator.next, Python __next__, Dart moveNext
    AsyncNext,     // Python __anext__
    Reversed,      // Python __reversed__, Ruby reverse_each

    // ── Arithmetic operators ────────────────────────────────────────
    Add,
    Sub,
    Mul,
    Div,
    /// Truncating division — Python `__floordiv__`, Dart `operator ~/`.
    /// Its own slot rather than sharing [`Div`](ProtocolSlot::Div): Python
    /// classes routinely define both, and one slot cannot hold two methods.
    FloorDiv,
    Mod,
    Pow,
    MatMul, // Python __matmul__ (@)
    Neg,
    Pos, // unary + — Python __pos__
    Abs, // Python __abs__

    // ── Numeric rounding protocol ───────────────────────────────────
    Round, // Python __round__
    Floor, // Python __floor__
    Ceil,  // Python __ceil__
    Trunc, // Python __trunc__
    Index, // Python __index__ — lossless conversion to an integer index

    // ── Comparison ──────────────────────────────────────────────────
    Eq,      // ==
    Ne,      // != / <> — C# `operator !=`, VB `Operator <>`, Python __ne__
    Compare, // <=> (Ruby) / __cmp__ (Python legacy) / CompareTo (C#)
    Lt,
    Le,
    Gt,
    Ge,

    // ── Bitwise ─────────────────────────────────────────────────────
    And,
    Or,
    Xor,
    Not,
    LShift,
    RShift,

    // ── Container protocol ──────────────────────────────────────────
    Len, // len() / length / size / Count
    /// Emptiness as its OWN question — Dart `isEmpty`, Java/Kotlin `isEmpty()`,
    /// Ruby `empty?`, PHP `empty()`, C# `Any()`/`IsEmpty`, Pascal `IsEmpty`.
    ///
    /// Not derived from [`Len`](ProtocolSlot::Len) even though `len == 0`
    /// answers it for most receivers: a lazy sequence can know it is non-empty
    /// without counting, and a language that spells only `isEmpty` on a type
    /// (Dart's `StringBuffer`) would otherwise have to publish a length it does
    /// not mean. One slot per QUESTION is what keeps the spelling in the
    /// language and the dispatch on the receiver.
    IsEmpty,
    GetItem,  // Python __getitem__, Ruby [], Dart operator [], PHP offsetGet
    SetItem,  // Python __setitem__, Ruby []=, Dart operator []=, PHP offsetSet
    DelItem,  // Python __delitem__, PHP offsetUnset
    HasItem,  // PHP offsetExists
    Missing,  // Python __missing__ — key absent from a mapping subclass
    Contains, // Python __contains__, Ruby include?, Dart contains

    // ── Callable / reflection ───────────────────────────────────────
    Call, // Python __call__, PHP __invoke, Dart call, C# ()
    /// The missing-method interceptor — PHP `__call`, Ruby `method_missing`,
    /// Dart `noSuchMethod`. NOT [`Call`](ProtocolSlot::Call): a PHP class may
    /// define `__invoke` and `__call` at once, and folding both onto one slot
    /// means the second install silently evicts the first.
    CallMissing,
    /// The static-side missing-method interceptor — PHP `__callStatic`.
    CallStatic,
    HasInstance, // JS Symbol.hasInstance, Python __instancecheck__

    // ── Property access interception ────────────────────────────────
    GetAttr, // Python __getattr__, PHP __get, JS Proxy get
    SetAttr, // Python __setattr__, PHP __set, JS Proxy set
    DelAttr, // Python __delattr__, PHP __unset
    HasAttr, // PHP __isset, JS Proxy has

    // ── Context managers ────────────────────────────────────────────
    Enter,      // Python __enter__, C#/Java using-block acquire
    Exit,       // Python __exit__, Java AutoCloseable.close, C# Dispose
    AsyncEnter, // Python __aenter__
    AsyncExit,  // Python __aexit__

    // ── Hash ────────────────────────────────────────────────────────
    Hash, // Python __hash__, Ruby hash, C# GetHashCode, Java hashCode

    // ── In-place (augmented-assignment) operators ───────────────────
    //
    // `x += y` is a DISTINCT method from `x + y` wherever a language lets a
    // class mutate in place (Python's `__iadd__` family). Without their own
    // slots these fall back to the binary op, which silently turns a mutation
    // into a rebind.
    IAdd,
    ISub,
    IMul,
    IDiv,
    IFloorDiv,
    IMod,
    IPow,
    IMatMul,
    IAnd,
    IOr,
    IXor,
    ILShift,
    IRShift,

    // ── Step operators ──────────────────────────────────────────────
    //
    // `x++` / `x--` are their OWN operations, not sugar for `+ 1`, wherever a
    // language steps non-numbers: PHP's alphanumeric increment ("a"++ is "b",
    // "2026-03-25"++ carries the date). A type binds them per
    // `[builtin_slots.<type>] inc/dec`; with no binding the shared step is
    // numeric (ECMA §13.4 ToNumeric).
    Inc,
    Dec,

    // ── Reflected (right-hand) operators ────────────────────────────
    //
    // `2 + vec` — the LEFT operand's type has no rule for the right one, so
    // dispatch reflects onto the right operand's method (Python `__radd__`).
    // A separate slot per operator because the parameter order differs.
    RAdd,
    RSub,
    RMul,
    RDiv,
    RFloorDiv,
    RMod,
    RPow,
    RMatMul,
    RAnd,
    ROr,
    RXor,
    RLShift,
    RRShift,

    // ── Character-class predicates ──────────────────────────────────
    //
    // `isdigit`/`isalpha`/… — non-standard behaviour (no ECMA-262 string
    // surface defines them), so the platform default rows point at tier-3
    // adapter primitives (`common:str_is_*`), never at a host fn. A language
    // whose classes differ (PHP ctype is C-locale ASCII) overrides per
    // `[builtin_slots.string] is_*`.
    IsDigit,
    IsAlpha,
    IsAlnum,
    IsSpace,
    IsUpper,
    IsLower,
}

/// The reserved property holding a class's protocol slot table.
///
/// ONE hidden key per object, whose value would map `slot_id` → the bound
/// method. The per-key form ([`protocol_slot_key`]) is what shipped; this name
/// is reserved so the two cannot both be claimed.
///
/// What both replace: a synonym table that stamped every cross-language
/// SPELLING of a method as its own property, so a Python class declaring
/// `__str__` also published `toString`, `tostring`, `ToString`, `to_s` and
/// `__toString` — five extra names in the same namespace user members live in.
/// That is what let a synonym set capture an unrelated user method (Dart's
/// `add`, `contains`, `length`). Deleted 2026-07-28.
pub const PROTOCOL_SLOT_TABLE: &str = "__vybe_slots";

/// The reserved member key a slot's implementation is published under.
///
/// Derived from the slot's NUMBER, never from any language's spelling — that is
/// the whole difference from the synonym stamping it replaces. `ToString` is
/// `__vybe_slot_1` whether the source wrote `__str__`, `to_s`, `toString` or
/// `__toString`, so a caller in any language reaches the same member, and a
/// user method genuinely named `toString` stays an ordinary member that nothing
/// else can capture.
pub fn protocol_slot_key(slot: ProtocolSlot) -> String {
    format!("__vybe_slot_{}", slot.slot_id())
}

impl ProtocolSlot {
    /// The slot's stable numeric identity — what dispatch keys on.
    ///
    /// Stable because it is written out, not derived from declaration order: a
    /// variant inserted in the middle must not renumber the others, since the
    /// ids are emitted into bytecode. Add new slots at the END of this match.
    /// Every slot, for exhaustive iteration.
    ///
    /// Generated from the same exhaustive `slot_id` match as
    /// [`Self::as_key`], so it cannot drift from the enum without `slot_id`
    /// failing to compile first.
    pub const ALL: [ProtocolSlot; 103] = [
        ProtocolSlot::Destructor,
        ProtocolSlot::ToString,
        ProtocolSlot::Repr,
        ProtocolSlot::ValueOf,
        ProtocolSlot::ToPrimitive,
        ProtocolSlot::Iterator,
        ProtocolSlot::AsyncIterator,
        ProtocolSlot::Next,
        ProtocolSlot::Add,
        ProtocolSlot::Sub,
        ProtocolSlot::Mul,
        ProtocolSlot::Div,
        ProtocolSlot::Mod,
        ProtocolSlot::Pow,
        ProtocolSlot::Neg,
        ProtocolSlot::Eq,
        ProtocolSlot::Compare,
        ProtocolSlot::Lt,
        ProtocolSlot::Le,
        ProtocolSlot::Gt,
        ProtocolSlot::Ge,
        ProtocolSlot::And,
        ProtocolSlot::Or,
        ProtocolSlot::Xor,
        ProtocolSlot::Not,
        ProtocolSlot::LShift,
        ProtocolSlot::RShift,
        ProtocolSlot::Len,
        ProtocolSlot::GetItem,
        ProtocolSlot::SetItem,
        ProtocolSlot::DelItem,
        ProtocolSlot::Contains,
        ProtocolSlot::Call,
        ProtocolSlot::HasInstance,
        ProtocolSlot::GetAttr,
        ProtocolSlot::SetAttr,
        ProtocolSlot::DelAttr,
        ProtocolSlot::Enter,
        ProtocolSlot::Exit,
        ProtocolSlot::Hash,
        ProtocolSlot::FloorDiv,
        ProtocolSlot::Bool,
        ProtocolSlot::Int,
        ProtocolSlot::Float,
        ProtocolSlot::Bytes,
        ProtocolSlot::Format,
        ProtocolSlot::Serialize,
        ProtocolSlot::Deserialize,
        ProtocolSlot::Clone,
        ProtocolSlot::AsyncNext,
        ProtocolSlot::Reversed,
        ProtocolSlot::MatMul,
        ProtocolSlot::Pos,
        ProtocolSlot::Abs,
        ProtocolSlot::Round,
        ProtocolSlot::Floor,
        ProtocolSlot::Ceil,
        ProtocolSlot::Trunc,
        ProtocolSlot::Index,
        ProtocolSlot::Ne,
        ProtocolSlot::HasItem,
        ProtocolSlot::Missing,
        ProtocolSlot::CallStatic,
        ProtocolSlot::HasAttr,
        ProtocolSlot::AsyncEnter,
        ProtocolSlot::AsyncExit,
        ProtocolSlot::IAdd,
        ProtocolSlot::ISub,
        ProtocolSlot::IMul,
        ProtocolSlot::IDiv,
        ProtocolSlot::IFloorDiv,
        ProtocolSlot::IMod,
        ProtocolSlot::IPow,
        ProtocolSlot::IMatMul,
        ProtocolSlot::IAnd,
        ProtocolSlot::IOr,
        ProtocolSlot::IXor,
        ProtocolSlot::ILShift,
        ProtocolSlot::IRShift,
        ProtocolSlot::RAdd,
        ProtocolSlot::RSub,
        ProtocolSlot::RMul,
        ProtocolSlot::RDiv,
        ProtocolSlot::RFloorDiv,
        ProtocolSlot::RMod,
        ProtocolSlot::RPow,
        ProtocolSlot::RMatMul,
        ProtocolSlot::RAnd,
        ProtocolSlot::ROr,
        ProtocolSlot::RXor,
        ProtocolSlot::RLShift,
        ProtocolSlot::RRShift,
        ProtocolSlot::CallMissing,
        ProtocolSlot::Inc,
        ProtocolSlot::Dec,
        ProtocolSlot::IsDigit,
        ProtocolSlot::IsAlpha,
        ProtocolSlot::IsAlnum,
        ProtocolSlot::IsSpace,
        ProtocolSlot::IsUpper,
        ProtocolSlot::IsLower,
        ProtocolSlot::Char,
        ProtocolSlot::IsEmpty,
    ];

    /// The slot's stable STRING key, for profile declarations.
    ///
    /// `builtinslotplan.md` step 4b. A profile writes
    /// `[value_methods] length = { emit = "...", slot = "len" }` to say which
    /// protocol slot its method implements — so the language owns the mapping
    /// from ITS spelling to the shared slot, and no method-name table ever
    /// appears in shared code.
    ///
    /// Keys are the variant name in snake_case, generated mechanically from the
    /// exhaustive [`ProtocolSlot::slot_id`] match rather than transcribed. All
    /// 93 are present and all are distinct — [`Self::every_slot_key_roundtrips`]
    /// pins both.
    pub fn as_key(self) -> &'static str {
        use ProtocolSlot::*;
        match self {
            Destructor => "destructor",
            ToString => "to_string",
            Repr => "repr",
            ValueOf => "value_of",
            ToPrimitive => "to_primitive",
            Iterator => "iterator",
            AsyncIterator => "async_iterator",
            Next => "next",
            Add => "add",
            Sub => "sub",
            Mul => "mul",
            Div => "div",
            Mod => "mod",
            Pow => "pow",
            Neg => "neg",
            Eq => "eq",
            Compare => "compare",
            Lt => "lt",
            Le => "le",
            Gt => "gt",
            Ge => "ge",
            And => "and",
            Or => "or",
            Xor => "xor",
            Not => "not",
            LShift => "l_shift",
            RShift => "r_shift",
            Len => "len",
            GetItem => "get_item",
            SetItem => "set_item",
            DelItem => "del_item",
            Contains => "contains",
            Call => "call",
            HasInstance => "has_instance",
            GetAttr => "get_attr",
            SetAttr => "set_attr",
            DelAttr => "del_attr",
            Enter => "enter",
            Exit => "exit",
            Hash => "hash",
            FloorDiv => "floor_div",
            Bool => "bool",
            Int => "int",
            Float => "float",
            Bytes => "bytes",
            Format => "format",
            Serialize => "serialize",
            Deserialize => "deserialize",
            Clone => "clone",
            AsyncNext => "async_next",
            Reversed => "reversed",
            MatMul => "mat_mul",
            Pos => "pos",
            Abs => "abs",
            Round => "round",
            Floor => "floor",
            Ceil => "ceil",
            Trunc => "trunc",
            Index => "index",
            Ne => "ne",
            HasItem => "has_item",
            Missing => "missing",
            CallStatic => "call_static",
            HasAttr => "has_attr",
            AsyncEnter => "async_enter",
            AsyncExit => "async_exit",
            IAdd => "i_add",
            ISub => "i_sub",
            IMul => "i_mul",
            IDiv => "i_div",
            IFloorDiv => "i_floor_div",
            IMod => "i_mod",
            IPow => "i_pow",
            IMatMul => "i_mat_mul",
            IAnd => "i_and",
            IOr => "i_or",
            IXor => "i_xor",
            ILShift => "i_l_shift",
            IRShift => "i_r_shift",
            RAdd => "r_add",
            RSub => "r_sub",
            RMul => "r_mul",
            RDiv => "r_div",
            RFloorDiv => "r_floor_div",
            RMod => "r_mod",
            RPow => "r_pow",
            RMatMul => "r_mat_mul",
            RAnd => "r_and",
            ROr => "r_or",
            RXor => "r_xor",
            RLShift => "r_l_shift",
            RRShift => "r_r_shift",
            CallMissing => "call_missing",
            Inc => "inc",
            Dec => "dec",
            IsDigit => "is_digit",
            IsAlpha => "is_alpha",
            IsAlnum => "is_alnum",
            IsSpace => "is_space",
            IsUpper => "is_upper",
            IsLower => "is_lower",
            Char => "char",
            IsEmpty => "is_empty",
        }
    }

    /// The slot a profile's `slot = "..."` names, or `None` if unrecognised.
    ///
    /// An unknown key is NOT an error: it means the profile names a slot this
    /// build does not have, and the declaration is ignored so the method keeps
    /// its existing emit. A profile from a newer toolchain must still load.
    pub fn from_key(key: &str) -> Option<Self> {
        use ProtocolSlot::*;
        Some(match key {
            "destructor" => Destructor,
            "to_string" => ToString,
            "repr" => Repr,
            "value_of" => ValueOf,
            "to_primitive" => ToPrimitive,
            "iterator" => Iterator,
            "async_iterator" => AsyncIterator,
            "next" => Next,
            "add" => Add,
            "sub" => Sub,
            "mul" => Mul,
            "div" => Div,
            "mod" => Mod,
            "pow" => Pow,
            "neg" => Neg,
            "eq" => Eq,
            "compare" => Compare,
            "lt" => Lt,
            "le" => Le,
            "gt" => Gt,
            "ge" => Ge,
            "and" => And,
            "or" => Or,
            "xor" => Xor,
            "not" => Not,
            "l_shift" => LShift,
            "r_shift" => RShift,
            "len" => Len,
            "get_item" => GetItem,
            "set_item" => SetItem,
            "del_item" => DelItem,
            "contains" => Contains,
            "call" => Call,
            "has_instance" => HasInstance,
            "get_attr" => GetAttr,
            "set_attr" => SetAttr,
            "del_attr" => DelAttr,
            "enter" => Enter,
            "exit" => Exit,
            "hash" => Hash,
            "floor_div" => FloorDiv,
            "bool" => Bool,
            "int" => Int,
            "float" => Float,
            "bytes" => Bytes,
            "format" => Format,
            "serialize" => Serialize,
            "deserialize" => Deserialize,
            "clone" => Clone,
            "async_next" => AsyncNext,
            "reversed" => Reversed,
            "mat_mul" => MatMul,
            "pos" => Pos,
            "abs" => Abs,
            "round" => Round,
            "floor" => Floor,
            "ceil" => Ceil,
            "trunc" => Trunc,
            "index" => Index,
            "ne" => Ne,
            "has_item" => HasItem,
            "missing" => Missing,
            "call_static" => CallStatic,
            "has_attr" => HasAttr,
            "async_enter" => AsyncEnter,
            "async_exit" => AsyncExit,
            "i_add" => IAdd,
            "i_sub" => ISub,
            "i_mul" => IMul,
            "i_div" => IDiv,
            "i_floor_div" => IFloorDiv,
            "i_mod" => IMod,
            "i_pow" => IPow,
            "i_mat_mul" => IMatMul,
            "i_and" => IAnd,
            "i_or" => IOr,
            "i_xor" => IXor,
            "i_l_shift" => ILShift,
            "i_r_shift" => IRShift,
            "r_add" => RAdd,
            "r_sub" => RSub,
            "r_mul" => RMul,
            "r_div" => RDiv,
            "r_floor_div" => RFloorDiv,
            "r_mod" => RMod,
            "r_pow" => RPow,
            "r_mat_mul" => RMatMul,
            "r_and" => RAnd,
            "r_or" => ROr,
            "r_xor" => RXor,
            "r_l_shift" => RLShift,
            "r_r_shift" => RRShift,
            "call_missing" => CallMissing,
            "inc" => Inc,
            "dec" => Dec,
            "is_digit" => IsDigit,
            "is_alpha" => IsAlpha,
            "is_alnum" => IsAlnum,
            "is_space" => IsSpace,
            "is_upper" => IsUpper,
            "is_lower" => IsLower,
            "char" => Char,
            "is_empty" => IsEmpty,
            _ => return None,
        })
    }

    pub fn slot_id(self) -> u16 {
        use ProtocolSlot::*;
        match self {
            Destructor => 0,
            ToString => 1,
            Repr => 2,
            ValueOf => 3,
            ToPrimitive => 4,
            Iterator => 5,
            AsyncIterator => 6,
            Next => 7,
            Add => 8,
            Sub => 9,
            Mul => 10,
            Div => 11,
            Mod => 12,
            Pow => 13,
            Neg => 14,
            Eq => 15,
            Compare => 16,
            Lt => 17,
            Le => 18,
            Gt => 19,
            Ge => 20,
            And => 21,
            Or => 22,
            Xor => 23,
            Not => 24,
            LShift => 25,
            RShift => 26,
            Len => 27,
            GetItem => 28,
            SetItem => 29,
            DelItem => 30,
            Contains => 31,
            Call => 32,
            HasInstance => 33,
            GetAttr => 34,
            SetAttr => 35,
            DelAttr => 36,
            Enter => 37,
            Exit => 38,
            Hash => 39,
            // Appended 2026-07-28 — ids continue from 39, existing ones unmoved.
            FloorDiv => 40,
            Bool => 41,
            Int => 42,
            Float => 43,
            Bytes => 44,
            Format => 45,
            Serialize => 46,
            Deserialize => 47,
            Clone => 48,
            AsyncNext => 49,
            Reversed => 50,
            MatMul => 51,
            Pos => 52,
            Abs => 53,
            Round => 54,
            Floor => 55,
            Ceil => 56,
            Trunc => 57,
            Index => 58,
            Ne => 59,
            HasItem => 60,
            Missing => 61,
            CallStatic => 62,
            HasAttr => 63,
            AsyncEnter => 64,
            AsyncExit => 65,
            IAdd => 66,
            ISub => 67,
            IMul => 68,
            IDiv => 69,
            IFloorDiv => 70,
            IMod => 71,
            IPow => 72,
            IMatMul => 73,
            IAnd => 74,
            IOr => 75,
            IXor => 76,
            ILShift => 77,
            IRShift => 78,
            RAdd => 79,
            RSub => 80,
            RMul => 81,
            RDiv => 82,
            RFloorDiv => 83,
            RMod => 84,
            RPow => 85,
            RMatMul => 86,
            RAnd => 87,
            ROr => 88,
            RXor => 89,
            RLShift => 90,
            RRShift => 91,
            CallMissing => 92,
            Inc => 93,
            Dec => 94,
            // Appended 2026-08-07 — ids continue from 94, existing ones unmoved.
            IsDigit => 95,
            IsAlpha => 96,
            IsAlnum => 97,
            IsSpace => 98,
            IsUpper => 99,
            IsLower => 100,
            // Appended 2026-08-07 — the coerce-to-char slot.
            Char => 101,
            // Appended 2026-08-07 — emptiness, its own question (see the
            // variant's doc). Ids continue from 101; existing ones unmoved.
            IsEmpty => 102,
        }
    }
}

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
    /// FINAL: no subclass may override this member. PHP `final`, Java `final`,
    /// VB `NotOverridable`, Fortran `non_overridable`.
    ///
    /// This is about the FUTURE — what descendants may do — and says nothing
    /// about an ancestor. Read by the compiler only to decide that a member can
    /// never be virtual.
    pub is_not_overridable: bool,
    /// HIDING: this member deliberately shadows an ancestor's member of the
    /// same name rather than overriding it, so the two occupy DISTINCT slots
    /// and a call resolves by the receiver's static type. C# `new`, VB
    /// `Shadows`, Pascal `reintroduce`.
    ///
    /// The exact opposite direction from [`is_not_overridable`]: hiding looks
    /// BACKWARD at an ancestor, final looks FORWARD at descendants. They were
    /// one field until it became clear the compiler was reading it for both
    /// meanings at once — as the hiding trigger and as "can never be virtual" —
    /// so a PHP `final` method whose name collided with an ancestor's was
    /// storage-renamed and its override silently broke, while a VB `Shadows`
    /// method was wrongly devirtualized. Five languages rode that ambiguity.
    ///
    /// Only ever a HINT that the language marked it: whether a member actually
    /// hides anything is answered by the compiler against `pending_classes`,
    /// because the ancestor may be declared later or in another file — which is
    /// why the walker cannot decide it alone.
    pub is_hiding: bool,
    /// This member is the class's DESTRUCTOR / finaliser, not an ordinary
    /// method. Set by the walker, which is the only place that knows how its
    /// language spells one — Pascal and C# mark it syntactically (`destructor
    /// Destroy;`, `~Foo()`), PHP and Python by a reserved name (`__destruct`,
    /// `__del__`).
    ///
    /// Without this the AST could not say "this class has a destructor" at
    /// all: it arrived as an ordinary `ClassMember::Method` and each normalizer
    /// re-derived the fact with its own name check — four different string
    /// tests in four crates, the same shape as the hardcoded `isInstance` name
    /// list that §4e retired.
    ///
    /// Subsumed by `protocol_slot == Some(ProtocolSlot::Destructor)`; kept
    /// until every walker sets the slot instead.
    pub is_destructor: bool,
    /// The cross-language ROLE this member fills, if any — `ToString`, `Add`,
    /// `Iterator`, `Destructor`, … See [`ProtocolSlot`].
    ///
    /// Set by the WALKER, which is the only place that knows how its language
    /// marks a role: Python by a reserved name (`__str__`), Dart and C# by
    /// syntax (`operator ==`, `implicit operator`), Java by conformance
    /// (`Comparable.compareTo`). Recovering it later from the method's name —
    /// which is what `canonicalize_method` does today — throws away what the
    /// frontend already knew and reintroduces the spelling as identity.
    ///
    /// `None` means an ordinary method, and that is the common case: a Dart
    /// `add` or a PHP `tostring` is an ordinary member unless its language
    /// says otherwise.
    pub protocol_slot: Option<ProtocolSlot>,
    /// WHERE this member's implementation lives.
    ///
    /// `None` means the ordinary case — the implementation is the source body
    /// in `FunctionDecl::body`, which is what every walker produces and what
    /// every reader assumes today.
    ///
    /// It exists because the model could not previously describe a member that
    /// has no source body but is nonetheless implemented. `StringBuilder.Append`
    /// is a method of a class by every definition this AST uses — a name, an
    /// arity, a receiver, visibility, a parent chain — and the only thing it
    /// lacks is text in a file. Unable to say that, the platform BCL tables went
    /// to the one model that could (`vybe_runtime::component_model::ClassType`),
    /// which put a class model in the VM. See flexclassplan §4a-octies.
    ///
    /// ⚠ [`Modifiers::is_abstract`] is the degenerate case of this field, and
    /// the two are meant to collapse: an abstract member is one whose
    /// implementation is declared absent. They coexist only until every walker
    /// sets the richer form, exactly as `is_destructor` coexists with
    /// `protocol_slot` above.
    pub implementation: Option<Implementation>,
    pub decorators: Vec<Expression>,
}

/// Where a member's implementation comes from.
///
/// **A method is a name, a signature, and an implementation.** The AST used to
/// hardcode one of these three cases — the source body — and could spell
/// neither of the others, which is why platform classes could not be expressed
/// as ordinary classes.
///
/// These are the same three cases `vybe_runtime`'s `MethodBody` enumerates
/// (`UserChunk` / `HostCall` / `Common`). That is the evidence the vocabulary is
/// right rather than invented: the VM had already discovered the distinction and
/// had to define a type to hold it *because this crate could not*. Naming them
/// here moves an existing abstraction to the layer that owns it.
#[derive(Debug, Clone, PartialEq)]
pub enum Implementation {
    /// A host import — `(module, name)`, called directly.
    Host { module: String, name: String },
    /// A shared compiler emit, named by string, that lowers to instructions.
    Intrinsic(String),
    /// Declared to have NO implementation here; a subclass must provide one.
    /// What `is_abstract` says today.
    Abstract,
}

#[derive(Debug, Clone, Default)]
pub struct ClassModifiers {
    pub visibility: Visibility,
    pub is_partial: bool,
    pub is_abstract: bool,
    pub is_sealed: bool,
    pub is_static: bool,
    /// `class` / `interface` / `trait` / `mixin` / `module` / `struct` — all of
    /// which parse to `StmtKind::ClassDecl`. Defaults to `Class`, so a walker
    /// that does not set it is unchanged.
    pub kind: ClassKind,
    /// Declared record semantics, for the declarations that parse to a
    /// `ClassDecl` rather than a `StructDecl` — a C# `record`, a Java `record`,
    /// a Kotlin `data class`, a Python `@dataclass`.
    ///
    /// It lives HERE rather than as a direct field because `ClassDecl` carries
    /// `parents` and `StructDecl` does not: a C# `record B : A` inherits, so
    /// normalizing records onto `StructDecl` would silently drop the base type.
    /// `ClassModifiers` already answers "what flavour of declaration is this",
    /// derives `Default`, and is built with `..default()` at every site but one
    /// — so this reaches all 28 `ClassDecl` constructions at no cost.
    ///
    /// The two declaration nodes converging is the real endgame; see
    /// `recordprimitiveplan.md`.
    pub semantics: ValueSemantics,
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
// Records
// ════════════════════════════════════════════════════════════════════════════
//
// One flexible record concept. The grammar is language-specific, with every
// quirk; after the walker normalizes, it is the same thing. A language declares
// POLICY here and the shared compiler owns the BEHAVIOUR — so a new language
// gets records by setting three properties, not by writing a lowering.
//
// The two semantic axes are INDEPENDENT, which is why one boolean cannot
// express them: `storage` is whether `b = a` copies, `equality` is whether
// `a == b` compares fields. Most languages pick one. C# needs all three
// combinations (`struct`, `record`, `record struct`).
//
// Every field defaults to what the tree did before this type existed, so a
// walker that sets nothing is unchanged. See `recordprimitiveplan.md`.

/// Does assignment copy, or alias?
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ValueStorage {
    /// `b = a` aliases — a class, a C# `record`, a Python `@dataclass`.
    #[default]
    Reference,
    /// `b = a` produces an independent value — a Pascal `record`, a C/Go
    /// `struct`, a C# `struct`, a VB `Structure`. Bites at THREE sites:
    /// assignment, argument passing and return.
    Value,
}

/// Does `==` compare fields, or identity?
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ValueEquality {
    /// Reference identity.
    #[default]
    Identity,
    /// Field-wise. The instance stamp `__value_eq` is this policy's runtime
    /// channel, already read by the language equality paths.
    Structural,
}

/// A field's DECLARED EXTENT — how many bytes it occupies when the record is
/// laid out as bytes, on disk or in a fixed-layout aggregate.
///
/// Distinct from [`FieldLayout`], which is memory ALIGNMENT. `PIC X(10)` is ten
/// bytes on disk whatever the alignment, and `RELATIVE`/`INDEXED` files put
/// record *n* at offset *n × width* — so a record cannot be read or written at
/// all without this. It is `None` for every language that does not declare
/// fixed widths, which is most of them.
///
/// Why it is on the DECLARATION rather than on the transfer node: it describes
/// a declared thing, `directives.md` §3 question 3 — the same call as
/// `Param.pass_by`. It is the same fact at every site that touches the record,
/// in any file, in any language. `RecordFieldFormat { decimal_places }`, which
/// hung one integer of this off `RewriteRecordFile` and nothing else, is what
/// it replaces.
///
/// ⚠ The declared width and the blank-padding rule are ONE fact, not two.
/// Fortran blank-pads `character(len=N)` comparison (F2018 §10.1.5.5.2) exactly
/// as COBOL pads alphanumeric; both walkers implement it separately today. A
/// comparison against a declared-width field pads to that width — reading this
/// is what lets that stop being per-language.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct FieldStorage {
    /// Declared width in bytes. `PIC X(10)` → 10; VB `String * 10` → 10;
    /// Fortran `character(len=10)` → 10.
    pub bytes: u32,
    /// Digits after an implied decimal point — COBOL `PIC 9(5)V99` → 2. The
    /// point is not stored, which is why the scale has to be declared.
    pub decimal_places: u8,
    /// How the sign is stored, when the field is signed at all.
    pub sign: Option<SignFormat>,
    /// Which end is padded when the value is shorter than `bytes`.
    pub justify: Justify,
}

/// Where a signed field keeps its sign, and whether it costs a byte.
///
/// COBOL `SIGN IS LEADING/TRAILING SEPARATE CHARACTER` is currently dropped on
/// the floor — the clause parses and nothing carries it, so a signed DISPLAY
/// field round-trips wrong. `SEPARATE` also changes the byte count, which is
/// why it belongs beside `bytes` rather than in a walker table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SignFormat {
    /// Sign overpunched onto the first digit — no extra byte.
    LeadingOverpunch,
    /// Sign overpunched onto the last digit — no extra byte. COBOL's default
    /// for a signed DISPLAY field.
    TrailingOverpunch,
    /// A `+`/`-` byte before the digits. Costs one byte.
    LeadingSeparate,
    /// A `+`/`-` byte after the digits. Costs one byte.
    TrailingSeparate,
}

/// Which end of a fixed-width field is padded.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Justify {
    /// Value at the left, padding on the right. COBOL alphanumeric default,
    /// Fortran `character`.
    #[default]
    Left,
    /// Value at the right, padding on the left. COBOL `JUSTIFIED RIGHT`, and
    /// how numeric DISPLAY fields are stored.
    Right,
}

/// Storage layout. Only meaningful where bytes are observable — C, COBOL
/// groups, Pascal `packed record`.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum FieldLayout {
    #[default]
    Auto,
    Packed,
    Explicit {
        align: u32,
    },
}

/// Overlapping storage: a Pascal variant part, a C `union`, a COBOL
/// `REDEFINES`. **One feature, currently implemented three times and two of
/// them are wrong** — Pascal flattens the arms into sibling fields and loses
/// the overlap, and a COBOL write through a REDEFINES alias does not propagate
/// back because it copies rather than sharing storage.
///
/// Modelled as TRUE overlapping storage rather than discriminated alternatives:
/// C's type-punning tests need a byte view, and "tag plus one active arm" can
/// be expressed on top of overlap but not the other way round.
#[derive(Debug, Clone, Default)]
pub struct VariantPart {
    /// The discriminant field, when the language has one (Pascal `case tag:`).
    /// `None` is a plain overlap — a C `union`, a COBOL `REDEFINES`.
    pub tag: Option<String>,
    /// Each arm is a set of members sharing the same region.
    pub arms: Vec<VariantArm>,
}

#[derive(Debug, Clone, Default)]
pub struct VariantArm {
    /// Tag values selecting this arm; empty for an untagged overlap.
    pub labels: Vec<Expression>,
    pub members: Vec<ClassMember>,
}

/// The declared semantics of a record. Defaults reproduce a plain reference
/// aggregate, which is what every `StructDecl` was before this existed.
#[derive(Debug, Clone, Default)]
pub struct ValueSemantics {
    pub storage: ValueStorage,
    pub equality: ValueEquality,
    pub layout: FieldLayout,
    pub variant: Option<VariantPart>,
}

// ════════════════════════════════════════════════════════════════════════════
// Sets
// ════════════════════════════════════════════════════════════════════════════
//
// One runtime storage, multiple language contracts. The backing store is the
// shared Set primitive, but languages disagree about the public surface:
// ECMA `add` returns the receiver, .NET/Kotlin `Add/add` returns changed?,
// Python `add` returns None, Python `remove` throws on a missing value while
// `discard` does not, and Python's algebra methods accept many operands.
//
// These are semantics, not host functions. A walker/profile chooses one mode
// and the shared set primitive lowers it. This mirrors record/enum
// normalization: frontends state policy, primitives own behavior.

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SetMutationResult {
    /// ECMA: return the receiver.
    #[default]
    Receiver,
    /// Kotlin/.NET: return whether membership changed.
    ChangedBool,
    /// Python mutators: return None/null.
    Void,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SetMissingDelete {
    /// `delete`/`discard`: missing is fine.
    #[default]
    Ignore,
    /// `remove`: missing is an error.
    ThrowKeyError,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SetAlgebraArity {
    /// Pascal/ECMA operators: one right operand.
    #[default]
    Binary,
    /// Python methods: receiver plus zero or more operands.
    Variadic,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SetMembership {
    /// ECMA SameValueZero — object elements compare by identity.
    #[default]
    SameValueZero,
    /// JDK `hashCode`/`equals` model: membership identity is a STRUCTURAL
    /// snapshot of the element, rendered at INSERTION time. Structurally
    /// equal values collide (a set deduplicates equivalent data keys), and
    /// mutating an element after insertion breaks its lookup — Java's
    /// hash-at-insert behaviour, which `java.util.HashSet` and Kotlin's
    /// data-class sets both contract. The backing store remains the ECMA
    /// Set; the snapshot keys ride in a sidecar the primitive owns.
    SnapshotKey,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SetSemantics {
    /// Contract for adding an element.
    pub mutation_result: SetMutationResult,
    /// Contract for deleting/removing an element when no error is raised.
    pub delete_result: SetMutationResult,
    pub missing_delete: SetMissingDelete,
    pub algebra_arity: SetAlgebraArity,
    /// Convert raw predicate results to a language-visible bool object.
    pub predicate_bool_object: bool,
    /// How membership identity is decided.
    pub membership: SetMembership,
}

impl Default for SetSemantics {
    fn default() -> Self {
        Self {
            mutation_result: SetMutationResult::Receiver,
            delete_result: SetMutationResult::ChangedBool,
            missing_delete: SetMissingDelete::Ignore,
            algebra_arity: SetAlgebraArity::Binary,
            predicate_bool_object: false,
            membership: SetMembership::SameValueZero,
        }
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Directives — declared policy over CODE
// ════════════════════════════════════════════════════════════════════════════
//
// `ValueSemantics` above declares policy for one DECLARATION and travels on the
// instance. A directive declares policy for a REGION OF CODE and does not: the
// two answer different questions and the assign path consults both. A Pascal
// record passed into PHP still copies, because its stamp came with it; a PHP
// array copies because the code assigning it is governed by PHP's directive.
//
// Real languages state these in the source and change them mid-file —
// `Option Explicit`, `declare(strict_types=1)`, `{$R+}`/`{$R-}`, `"use strict"`.
// A profile flag would work but is invisible in the program and cannot be
// overridden by it, so policy that a program is allowed to state lives here.
//
// Isolation across languages is structural, not enforced here: each unit of a
// multi-language program is compiled on its own terms and nothing is
// concatenated across languages (`DynamicRuntime::run_program_unit`). Pascal
// never sees PHP's directives because it never shares PHP's compile.

/// How a method call obtains its receiver — see [`Directives::method_receiver`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MethodReceiver {
    /// The call site passes the receiver as an explicit leading argument,
    /// because the callable is the raw function off the class struct and
    /// carries no receiver of its own. php.
    CallSite,
    /// The callable rides an ambient receiver plus a bound-receiver marker.
    /// JS, Dart. Currently still declared as
    /// `class_method_dispatch = "prototype"`.
    Prototype,
    /// Reading the method produces a fresh callable with the receiver already
    /// bound. Python. Currently still declared as `methods_bind_on_access`.
    BindOnAccess,
}

/// A statement of policy. Every field is an `Option`: `None` means "not stated
/// here", so the same type serves as a module's declared defaults and as an
/// in-source delta that changes one thing and leaves the rest alone.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Directives {
    /// Does assigning the language's builtin ARRAY copy it? PHP says yes —
    /// `$b = $a` on an array is a value copy, and objects inside stay shared.
    /// Every other language in the tree aliases, which is the `None` default.
    ///
    /// Distinct from `ValueSemantics::storage`, which is about a declared
    /// aggregate. This is about the builtin type, which has no declaration to
    /// hang policy on and no constructor to stamp.
    pub array_storage: Option<ValueStorage>,

    /// When an assignment's VALUE is a reference — [`ExprKind::RefOf`] or
    /// `Unary{AddrOf}` — does the target BIND to that reference, or STORE it as
    /// an ordinary pointer value?
    ///
    /// This is Axis B vs Axis A of `referenceplan.md` §3, and it is the one
    /// distinction the AST could not state. php `$b = &$a` BINDS: `$b = 42`
    /// afterwards writes `$a`'s storage. c/go/pascal `p = &v` STORES: `p = &w`
    /// rebinds `p` and leaves `v` alone. Same node shape, opposite meanings.
    ///
    /// It belongs here rather than on the node because it is uniform within a
    /// language — checked across every walker that builds a reference: only php
    /// aliases on assignment, and no language does both. `ExprKind::Assign`
    /// carries just a target and a value, so without this the shared compiler
    /// had nothing to read and the two meanings collapsed into one path.
    ///
    /// Only [`PassBy::Alias`] is meaningful; `None` and everything else are the
    /// STORE behaviour, which is what every language but php wants.
    pub reference_binding: Option<PassBy>,

    /// What does the argument to the language's TERMINATE builtin mean?
    ///
    /// Every language binds its own spelling to `common:control_flow.exit` on
    /// its OWN profile row — `os.Exit` (go), `halt` (pascal), `sys.exit`
    /// (python), `process.exit` (js), `os.exit` (lua), `__process_exit` (java).
    /// For all of them the argument IS the status, which is the `None` default.
    ///
    /// php is the one language where the same syntax carries two things:
    /// `die("bye")` prints a farewell MESSAGE and exits 0, while `exit(3)`
    /// exits with STATUS 3 and prints nothing. Which one it is can only be
    /// decided at runtime, so it cannot be a lowering — the emitter has to
    /// branch on the value's type.
    ///
    /// It is stated here rather than as a profile flag because it is a
    /// SEMANTIC property of the language, and it replaced two
    /// `profile.name == "php"` checks in shared code (`calls.rs`,
    /// `statements.rs`) that recognised the SPELLINGS `exit`/`die` directly.
    pub exit_argument: Option<ExitArgument>,

    /// Can a SPREAD argument carry NAMES as well as positions?
    ///
    /// Every language's `f(...xs)` unpacks positionally, which is the `None`
    /// default. php also accepts a STRING-KEYED array — `f(...['b' => 2])`
    /// binds by PARAMETER NAME, and which one a given call means depends on
    /// the array's keys at RUNTIME, so it cannot be decided by a lowering.
    ///
    /// Stated here because it is a property of the language's call semantics,
    /// replacing a `profile.name == "php"` check in `calls.rs` that guarded
    /// the named-unpack probe.
    pub spread_arguments: Option<SpreadArguments>,

    /// How a method CALL obtains its receiver.
    ///
    /// Three models exist and shared code has to pick one per unit:
    /// - **prototype dispatch** (JS/Dart) — the callable rides `__js_this` and
    ///   a bound-receiver marker;
    /// - **bind-on-access** (Python) — reading the method burns the receiver
    ///   into a fresh bound method;
    /// - [`MethodReceiver::CallSite`] — the callable is the raw function off
    ///   the class struct and carries no receiver, so the CALL supplies one.
    ///
    /// `None` means "not the call-site model", which is every language but php
    /// today, and is the safe answer for a language that has not stated one.
    ///
    /// ⛔ This replaces a `profile.name == "php"` in shared code — a language
    /// NAME standing in for a property of the language's dispatch model, which
    /// is the arrangement the tree is removing. A directive is stated by the
    /// walker on `Module.directives`, so it travels with the UNIT that declared
    /// it and a multi-language bundle gets the right answer per unit; a profile
    /// is installed once per compilation and cannot.
    ///
    /// The other two models are still profile-declared
    /// (`class_method_dispatch = "prototype"`, `methods_bind_on_access`) and
    /// belong in this enum too — they are named here so that migration does not
    /// have to change its shape.
    pub method_receiver: Option<MethodReceiver>,


    /// Is a parameter with no supplied argument bound to `undefined`, rather
    /// than being an error or a language-specific sentinel? ECMA-262 §10.2.1.1.
    pub missing_arg_is_undefined: Option<bool>,

    /// Are a class's STATIC fields own properties of the class object —
    /// enumerable through the ordinary object surface — or a separate storage
    /// that reflection does not see?
    pub static_fields_are_own_properties: Option<bool>,

    /// Is a declared function a first-class OBJECT carrying properties —
    /// `name`, `length`, `prototype`, a `__nonenum` set — or just code?
    ///
    /// ECMA-262 §10.2.5 / §10.2.9 / §10.2.10 say yes, and js, php, python and
    /// ruby all want that. **wast, c, cobol, fortran and pascal do not have
    /// function objects at all**, and stamping the metadata onto their
    /// functions is not a harmless extra: for wast it emits a `struct.set` of
    /// `name` / `length` / `prototype` into a module that declares no such
    /// fields, so the module does not load on a spec engine.
    ///
    /// ⛔ This is NOT the "flag with one consumer" this document warns about
    /// (§1, `args_pass_by_reference`). Half the languages in the tree answer
    /// `false`; it reads `None` as the ECMA behaviour only because that is what
    /// the majority of the CURRENTLY-WIRED walkers want, and every language
    /// that states it states a fact about its own semantics.
    ///
    /// Why it cannot be a property of the declaration: nothing in a
    /// `FunctionDecl` distinguishes "a wasm function" from "a JS function" —
    /// the node is identical. The fact belongs to the UNIT being compiled,
    /// which is exactly what `Module.directives` is.
    pub functions_are_objects: Option<bool>,

    /// Inside a SUBPROGRAM BODY, do local declarations compile before nested
    /// procedure declarations, whatever order the flattened statement list
    /// puts them in?
    ///
    /// Fortran writes a procedure's locals above `contains` and its internal
    /// procedures below, so an internal procedure may reference a local that
    /// follows it once the body is flattened. Compiling in list order leaves
    /// that local undefined at the point the inner procedure binds.
    ///
    /// ⛔ The flag this replaces was called
    /// `class_body_declarations_before_procedures` and the "class" in that name
    /// was WRONG — its only reader is in `compile_function_decl`, on a FUNCTION
    /// body. A class body never reaches it. The name is corrected here rather
    /// than carried, because the old one sent me to the class path looking for
    /// an effect that could not be there.
    ///
    /// `directives.md` §3 question 1: it governs how a REGION OF CODE is
    /// compiled, rather than describing a declaration or travelling with a
    /// value.
    ///
    /// ⚠ UNPROVEN. Ablating it moves ZERO fortran tests (537 either way, 0 by
    /// name), and no dump differs on a derived type. Preserved verbatim because
    /// a refactor is not the place to delete a behaviour nobody has shown to be
    /// dead — but it is a candidate for M8's "or cease to exist" branch, and
    /// the next person to touch it should try to KILL it rather than assume it
    /// works.
    pub body_declarations_first: Option<bool>,

    /// The same question for a class's declared INSTANCE fields: are they own
    /// properties of the instance — what `Object.keys`, `in`, `for…in` and
    /// `JSON.stringify` enumerate — or a separate typed storage?
    ///
    /// ECMA-262 §10.2.11 settles it for js: `InitializeInstanceElements`
    /// performs `CreateDataPropertyOrThrow`, so a declared field IS an own
    /// enumerable property and nothing else will do. A language whose fields
    /// are a fixed record (java, C#, pascal) states nothing and keeps the
    /// indexed storage.
    ///
    /// ⛔ It is stated here rather than inferred because inferring it is how
    /// the bug arose: `seam3_indexable` grants indexed GC-struct storage on a
    /// fact about ALLOCATION (`published && no parent`), which has nothing to
    /// say about whether the result must be enumerable. A parentless js class
    /// took the licence and wrote `struct.set`, while the host's key walk
    /// reads the `properties` map — two storage locations on one object, with
    /// `d.a` answering 1 while `Object.keys(d)` answered `[]`. Adding a parent
    /// withheld the licence and the same field became visible, which is the
    /// signature of a storage split rather than a field bug.
    pub instance_fields_are_own_properties: Option<bool>,













    /// How LOCAL and PARAMETER names compare in this region.
    ///
    /// `None` means [`CaseMatch::Exact`]: eleven of the seventeen languages
    /// here, and the safe answer for a language that has not stated one.
    /// [`CaseMatch::Folded`] is vb, pascal, cobol, fortran and powershell.
    ///
    /// ⛔ **Folding a variable name is a COMPILE-TIME normalisation, not a
    /// runtime behaviour.** `Scope::resolve` answers with a SLOT, so a VB
    /// reference to `MYVAR` is resolved once against the local declared
    /// `myVar` and the runtime never sees either spelling. That is the whole
    /// design: fix the name where the declaration is known, and there is no
    /// case problem left downstream. A directive that made the RUNTIME fold
    /// would be the opposite of this.
    pub variable_case: Option<CaseMatch>,

    /// How FUNCTION and CLASS names compare in this region.
    ///
    /// Separate from [`Self::variable_case`] because PHP splits them —
    /// `$Foo` and `$foo` are different variables while `strlen` and `StrLen`
    /// are the same function. A language whose variables fold necessarily
    /// folds its callables too; the reverse does not hold, and that asymmetry
    /// is the only reason this is a second field rather than one.
    pub callable_case: Option<CaseMatch>,

    /// How a NAMED TUPLE's field names compare in this region.
    ///
    /// ⛔ **A tuple field is not a namespace member and not a local.** It is a
    /// key on a VALUE, resolved at run time against the shape the literal
    /// built — so unlike [`Self::variable_case`], which `Scope::resolve`
    /// settles at compile time into a slot, this one has to survive into the
    /// emitted value. It is stated separately because a language can fold its
    /// identifiers and still want the DECLARED spelling reported back:
    /// `.ToString()` and `_asdict` read `__fields`, which keeps what the
    /// source wrote.
    ///
    /// `None` means [`CaseMatch::Exact`] — C#, Dart and Python, whose tuple
    /// fields are case-sensitive. VB states [`CaseMatch::Folded`]: `(Id:=2).id`
    /// and `.ID` are the same field, because VB identifiers are.
    pub tuple_field_case: Option<CaseMatch>,

    /// Which alphabet [`CaseMatch::Folded`] folds. `None` means
    /// [`CaseAlphabet::Ascii`], which is what every folding language in this
    /// tree specifies and what `Scope` already implements.
    ///
    /// Stated separately from the two `CaseMatch` fields, rather than as extra
    /// enum variants on them, so that "do we fold?" and "how wide is the fold?"
    /// stay independent questions. A language answering the first differently
    /// for variables and callables still answers the second once.
    pub case_alphabet: Option<CaseAlphabet>,

    /// Declared contract for builtin sets in this region. The storage remains
    /// the common Set primitive; this records the source-language surface.
    pub set_semantics: Option<SetSemantics>,

    /// How a method's RECEIVER reaches its body in this region.
    ///
    /// Where `this` is ambient (ECMA §10.2.1.1 — JS, Dart, Lua's `:` sugar
    /// aside) the callee reads it from the call's own binding and its declared
    /// parameters are untouched. Where the receiver is an explicit leading
    /// parameter — Pascal's `Self`, C#/VB's `this`, Python's `self` — it must be
    /// PASSED, and anything that constructs or binds a callable has to supply
    /// it.
    ///
    /// It belongs here for the same reason [`Self::reference_binding`] does: it
    /// is uniform within a language and no single node can state it. A handler
    /// value carries no record of how its receiver arrives, so
    /// `ecma:function.bind` at an `addEventListener` site, a lambda capture and
    /// a delegate construction each had to know — and each read a profile flag,
    /// which the program itself could never see.
    ///
    /// `None` means the receiver is an explicit parameter, which is what all but
    /// a handful of languages want.
    pub receiver_binding: Option<ReceiverBinding>,

    /// How this region decides whether a value is TRUE.
    ///
    /// Two genuinely different questions, and no node can tell them apart:
    /// `if x` is the same shape in every language. ECMA §7.1.2 ToBoolean asks
    /// the VALUE — `null`/`false`/`0`/`""` are false and every object is true.
    /// CPython §3.3.1 asks the OBJECT — [`ProtocolSlot::Bool`] first, then
    /// [`ProtocolSlot::Len`], and only a value that answers neither is true.
    /// The two disagree on `[]`, `{}`, `set()` and on any class defining
    /// `__bool__`/`__len__`.
    ///
    /// It belongs here for the reason [`Self::receiver_binding`] does: it is
    /// uniform within a language and the program itself could never see it.
    /// It replaced `truthiness_via_dunder_or_length`, a PROFILE property that
    /// one language declared and three shared sites read — and because those
    /// sites each decided separately, they drifted: `if`, `bool()` and
    /// `not not` applied the protocol while `emit_dyn_not` did not, so a
    /// hand-built `Unary{Not}` (which is what `assert` desugars to) answered
    /// `[]` truthy and `assert []` silently passed.
    ///
    /// `None` inherits [`Truthiness::Value`] — ECMA's rule, which is what every
    /// language but the protocol ones wants.
    pub truthiness: Option<Truthiness>,

    /// What a shift or rotate count outside `[0, width)` does in this region.
    ///
    /// Genuinely lexical policy: the operand's declared type does NOT
    /// distinguish these, the language does. wasm — and therefore JS, Java and
    /// C# — MASK the count, so `1 << 32` is `1`. Fortran's `ISHFT` yields
    /// ZERO whenever `|shift| >= BIT_SIZE`, and `gfortran` proves it:
    /// `ishft(1, 32)` prints `0` where a masking lane prints `1`.
    ///
    /// `None` inherits [`ShiftOverflow::Mask`], which is what every language
    /// but Fortran wants and what the compiler already emitted.
    pub shift_overflow: Option<ShiftOverflow>,

    /// Which `pow` contract this region wants for the two cases where the two
    /// standards genuinely disagree.
    ///
    /// Sibling of [`Self::shift_overflow`], and the same kind of fact: a
    /// numeric edge case the operand's TYPE cannot distinguish, only the
    /// language can. ECMA-262 §6.1.6.1.3 answers `NaN` for `1 ** ±∞` and
    /// `1 ** NaN`; IEEE 754-2019 answers `1`, and the standard says so in its
    /// own note. Neither is a bug: JS and Java want the first, C, Python,
    /// Fortran, Go and Lua want the second.
    ///
    /// `None` inherits [`PowSemantics::Ieee`] — what `f64::powf` already gave
    /// every language, so declaring nothing keeps today's behaviour.
    pub pow_semantics: Option<PowSemantics>,

    /// Does this program present a user interface?
    ///
    /// A WHOLE-PROGRAM property in the same sense case sensitivity is: the
    /// project kind states it — a Delphi `.dproj` whose `.dpr` calls
    /// `Application.Run`, a WinForms `.vbproj`/`.csproj`, a Flutter `runApp` —
    /// and no single statement can. It is declared rather than inferred because
    /// the alternative tests are all wrong in one direction: an entry point that
    /// merely LINKS GUI code is not a GUI program, and a document that happens
    /// to be empty when `main` returns is not a console one.
    ///
    /// It answers a question the runtime asks after the program has run —
    /// whether to present a window and wait, or exit — which used to be a host
    /// call setting `GuiState.should_run` from inside the
    /// guest. That made "is this a UI program" a side effect of calling a host
    /// function, invisible to anything that did not.
    ///
    /// `None` states nothing, and the document answers instead: a document with
    /// content is a running one. [`AppShell::Headless`] is the case nothing else
    /// can express — a program that builds controls but must not open a window.
    pub app_shell: Option<AppShell>,

    /// When a NAME stops referring to a value in this region, does the value's
    /// finaliser run?
    ///
    /// Question 1 of `directives.md` §3: it governs a region of CODE. Python
    /// says `del x` and `x = None` both drop a reference and may finalise;
    /// C#, Java and JS drop the name and leave collection to the runtime. No
    /// property of the VALUE distinguishes those — the language does, and the
    /// same object reached from two languages must obey whichever one is
    /// executing. That is the definition of lexical.
    ///
    /// It states only WHETHER. WHICH method runs is
    /// [`ProtocolSlot::Destructor`], a property of the class (question 3), so
    /// the two facts keep one home each. Before this, python's walker
    /// synthesised `typeof x.__del__ == "function"` into the tree at every drop
    /// site — a spelling carried across a layer, re-tested at runtime, for a
    /// slot `NormalClass::destructor` had already filled by construction.
    ///
    /// `None` inherits "dropping a name finalises nothing", which is what every
    /// language but python wants and what the compiler already emitted.
    pub name_drop: Option<NameDrop>,

    /// The TEXT of a Boolean when a string needs it.
    ///
    /// Question 1 of `directives.md` §3 — it governs a region of code, and the
    /// answers genuinely differ: ECMA writes `true`/`false`, .NET and python
    /// write `True`/`False`. No property of the VALUE distinguishes them, and
    /// the same `true` reached from two languages must read the way whichever
    /// one is executing spells it.
    ///
    /// ⛔ IT WAS A WALKER PASS. VB carried ~120 lines of
    /// `normalize_vb_concat_bool_text` walking every statement to rewrite a
    /// concatenated Boolean into a literal, C# carried nothing at all and
    /// printed `1`/`0` for every `bool` in a string, and python got it right
    /// only inside `str()`. One fact, three homes, two of them wrong.
    ///
    /// `None` inherits ECMA's `true`/`false`, which is what the compiler
    /// already emitted and what JS, java, kotlin, ruby and lua all want.
    pub bool_text: Option<BoolText>,

    /// When a field declaration repeats a name an ancestor already declared,
    /// does it get its OWN storage or write the ancestor's?
    ///
    /// Question 1 of `directives.md` §3: it governs a region of CODE. The two
    /// declarations are identical in every language — what differs is the
    /// language executing them. Java, C# and VB give the shadowing field a
    /// declaring-class-keyed slot so both survive on the object and a read
    /// resolves by the reference's STATIC type; python, js and ruby have one
    /// slot per name on the instance. No property of the field distinguishes
    /// those, which is what makes it lexical.
    ///
    /// It states only the language-wide RULE. Whether a particular declaration
    /// *intends* to shadow is [`Modifiers::is_hiding`] (C# `new`, VB `Shadows`,
    /// Pascal `reintroduce`) — a property of the DECLARATION, question 3 — so
    /// the two facts keep one home each.
    ///
    /// A **private** field takes its own slot whatever this says: an ancestor's
    /// private field is not visible here, so a same-named field is a different
    /// field. That follows from `Access::Private` on the declaration and is not
    /// the language's shadowing rule.
    ///
    /// `None` inherits [`FieldShadowing::Share`], which is what every language
    /// but the statically-typed three wants.
    pub field_shadowing: Option<FieldShadowing>,
}

/// Whether a field declaration that repeats an ancestor's name gets its own
/// storage — see [`Directives::field_shadowing`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FieldShadowing {
    /// One slot per field NAME on the instance: a subclass declaration writes
    /// the ancestor's storage.
    #[default]
    Share,
    /// The shadowing declaration gets a declaring-class-keyed slot. Both
    /// survive on the object and a read resolves by the declared type.
    Hide,
}

/// What happens to the referent when a name stops referring to it — see
/// [`Directives::name_drop`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum NameDrop {
    /// The name stops referring and nothing else happens.
    #[default]
    Ignore,
    /// Run the referent's [`ProtocolSlot::Destructor`] first, if it binds one.
    ///
    /// **Known imprecision, stated rather than hidden.** CPython finalises when
    /// the last REFERENCE goes away; this finalises when the NAME does. With an
    /// alias live (`y = x; del x`) CPython runs nothing and this runs the
    /// finaliser early. Getting that exact needs refcounting in the VM. The
    /// same trade was already made for PHP's `unset`.
    Finalise,
    /// The value becomes ELIGIBLE for finalisation; the finaliser runs at the
    /// next collection point, not at the drop.
    ///
    /// This is .NET's answer, and it is a third answer rather than a shade of
    /// [`Finalise`]: a C# or VB finaliser is emphatically NOT run when the name
    /// goes away — it runs when the collector gets to it, which a program
    /// forces with `GC.Collect()` / `GC.WaitForPendingFinalizers()`. Running it
    /// at the drop would make `Dispose` and `Finalize` indistinguishable, and
    /// the whole .NET convention is that they are not.
    ///
    /// Same stated imprecision as [`Finalise`] about WHICH values become
    /// eligible — a dropped NAME, not a dropped last REFERENCE — but the timing
    /// it expresses is exact.
    Defer,
}

/// How a Boolean spells itself — see [`Directives::bool_text`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum BoolText {
    /// `true` / `false` — ECMA, java, kotlin, ruby, lua, go.
    #[default]
    Lowercase,
    /// `True` / `False` — .NET (`Boolean.ToString`) and python (`str`).
    TitleCase,
}

impl BoolText {
    /// The two spellings, true first.
    pub fn texts(self) -> (&'static str, &'static str) {
        match self {
            BoolText::Lowercase => ("true", "false"),
            BoolText::TitleCase => ("True", "False"),
        }
    }
}

/// Whether a program presents a user interface — see [`Directives::app_shell`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AppShell {
    /// The program presents a UI and the runtime waits for it.
    #[default]
    Windowed,
    /// The program must NOT present one, whatever it builds. A designer or a
    /// test harness that constructs controls to inspect them wants this.
    Headless,
}

/// How a method's receiver reaches its body — see
/// [`Directives::receiver_binding`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ReceiverBinding {
    /// An explicit leading parameter: Pascal `Self`, Python `self`, C# `this`.
    #[default]
    ExplicitParameter,
    /// Ambient, read from the call's own `this` binding: JS, Dart.
    Ambient,
    /// **Every** callable takes a leading receiver parameter, not just methods
    /// — ECMA-262 §10.2.1 `[[Call]](thisArgument, argumentsList)`, where the
    /// receiver is an argument of the call itself rather than a property of
    /// being a method.
    ///
    /// ⛔ This is a THIRD fact, not a spelling of [`Self::ExplicitParameter`],
    /// and the difference is load-bearing for M5. Under `ExplicitParameter`
    /// (python, pascal, C#) a plain `def f(x)` takes NO receiver, so a caller
    /// must know whether the callee is a method to know the argument count.
    /// That is answerable in those languages because their method calls are
    /// resolved statically. It is NOT answerable in JS: `o.m(1)` compiles to a
    /// dynamic `call_ref` after a runtime property lookup, and `const f = o.m;
    /// f()` reaches the identical instruction — so a receiver that only
    /// METHODS take is a receiver the call site cannot count.
    ///
    /// ⇒ Uniform arity is what makes the receiver expressible as a parameter
    /// at all in an ambient-dispatch language, which is why M5 moves js and
    /// dart HERE rather than to `ExplicitParameter`. A plain call passes the
    /// undefined receiver explicitly (§10.2.1.1 OrdinaryCallBindThis with a
    /// non-object thisArgument), which is also the ECMA answer rather than a
    /// filler value.
    UniversalParameter,
}

/// Which `pow` contract a region wants — see [`Directives::pow_semantics`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PowSemantics {
    /// IEEE 754-2019: `pow(1, ±∞)` and `pow(1, NaN)` are `1`. C, Python,
    /// Fortran, Go, Lua — and what `f64::powf` does natively.
    #[default]
    Ieee,
    /// ECMA-262 §6.1.6.1.3: those same cases are `NaN`. Kept by the standard
    /// "for compatibility reasons" with the first edition. JS, Java.
    Ecma,
}

/// How a value's truth is decided — see [`Directives::truthiness`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Truthiness {
    /// Ask the VALUE. ECMA §7.1.2 ToBoolean: `null`, `false`, `0`, `NaN` and
    /// `""` are false; every object — an empty array included — is true.
    #[default]
    Value,
    /// Ask the OBJECT, then fall back to the value. CPython §3.3.1: try
    /// [`ProtocolSlot::Bool`], then [`ProtocolSlot::Len`] (`!= 0`), and only
    /// then the value itself.
    ///
    /// Stated as a protocol rather than as "empty collections are falsy"
    /// deliberately: a builtin `[]` is falsy for exactly the reason a user
    /// class with `__len__` returning 0 is, so one rule covers both and a
    /// class in ANY language gets it by binding the slot.
    Protocol,
}

/// How two identifiers are compared for equality in this region.
///
/// ⛔ **This is a property of the NAME KIND, not of the language.** PHP is the
/// proof: its *variables* are case-sensitive while its *function and class
/// names* are not. A single per-language boolean cannot state that, which is
/// why [`Directives`] carries `variable_case` and `callable_case` separately.
/// The previous carrier was a `self.name == "php"` check inside the VM crate's
/// `lookup_builtin` — a language-name gate, which is what this replaces.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CaseMatch {
    /// Byte-exact. Every language that does not say otherwise.
    Exact,
    /// Case-insensitive, folding the alphabet [`CaseAlphabet`] names.
    Folded,
}

/// Which alphabet a [`CaseMatch::Folded`] comparison folds.
///
/// ⛔ **The tree already disagrees with itself about this**, which is why it is
/// stated rather than assumed. `primitives/scope.rs` folds with
/// `eq_ignore_ascii_case`; `vybe_runtime::namespaces` folds with
/// `to_lowercase()`, which is Unicode. So a Turkish dotted `I` resolves one way
/// as a local and another way as a namespace path *in the same program*, and
/// nothing reports the disagreement because the two are never compared.
///
/// Every case-insensitive language in this tree — vb, pascal, cobol, fortran,
/// powershell, and PHP's callables — specifies ASCII folding for identifiers,
/// so `Ascii` is the answer for all of them today. `Unicode` exists because the
/// AST has to be able to STATE the other answer: a language whose identifiers
/// fold beyond ASCII is expressible without another carrier being invented for
/// it later.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CaseAlphabet {
    /// `A`-`Z` only. Nothing outside ASCII folds.
    Ascii,
    /// Unicode simple case folding.
    Unicode,
}

impl Directives {
    /// Whether a LOCAL/PARAMETER reference folds here, and over which alphabet.
    ///
    /// One accessor rather than two `Option` reads, because the defaults have
    /// to be applied together and applying them at each call site is how
    /// `case_sensitive` went wrong: it was a compiler flag, so all 33 sites had
    /// to remember `!self.case_sensitive &&` and **23 did not**, silently
    /// breaking Go. Anything that needs the policy reads it from here.
    pub fn variable_fold(&self) -> Option<CaseAlphabet> {
        match self.variable_case.unwrap_or(CaseMatch::Exact) {
            CaseMatch::Exact => None,
            CaseMatch::Folded => Some(self.case_alphabet.unwrap_or(CaseAlphabet::Ascii)),
        }
    }

    /// Whether a NAMED TUPLE field reference folds here, and over which
    /// alphabet. Same contract as [`Self::variable_fold`].
    ///
    /// ⛔ Read where the tuple is BUILT, not where it is read. `variable_case`
    /// can be settled at compile time because `Scope::resolve` answers with a
    /// slot; a tuple field is a key on a value, so the folded spelling has to
    /// be on the value before it travels anywhere.
    pub fn tuple_field_fold(&self) -> Option<CaseAlphabet> {
        match self.tuple_field_case.unwrap_or(CaseMatch::Exact) {
            CaseMatch::Exact => None,
            CaseMatch::Folded => Some(self.case_alphabet.unwrap_or(CaseAlphabet::Ascii)),
        }
    }

    /// Whether a FUNCTION/CLASS reference folds here, and over which alphabet.
    /// Same contract as [`Self::variable_fold`].
    pub fn callable_fold(&self) -> Option<CaseAlphabet> {
        match self.callable_case.unwrap_or(CaseMatch::Exact) {
            CaseMatch::Exact => None,
            CaseMatch::Folded => Some(self.case_alphabet.unwrap_or(CaseAlphabet::Ascii)),
        }
    }

    /// Nothing stated.
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }

    /// Apply `other` on top of `self`: a field `other` states wins, a field it
    /// leaves `None` is inherited. This is the only way directives combine.
    pub fn overlay(&mut self, other: &Directives) {
        if other.array_storage.is_some() {
            self.array_storage = other.array_storage;
        }
        if other.reference_binding.is_some() {
            self.reference_binding = other.reference_binding;
        }
        if other.variable_case.is_some() {
            self.variable_case = other.variable_case;
        }
        if other.callable_case.is_some() {
            self.callable_case = other.callable_case;
        }
        if other.case_alphabet.is_some() {
            self.case_alphabet = other.case_alphabet;
        }
        if other.set_semantics.is_some() {
            self.set_semantics = other.set_semantics;
        }
        if other.receiver_binding.is_some() {
            self.receiver_binding = other.receiver_binding;
        }
        if other.shift_overflow.is_some() {
            self.shift_overflow = other.shift_overflow;
        }
        if other.app_shell.is_some() {
            self.app_shell = other.app_shell;
        }
        if other.name_drop.is_some() {
            self.name_drop = other.name_drop;
        }
        if other.bool_text.is_some() {
            self.bool_text = other.bool_text;
        }
        if other.field_shadowing.is_some() {
            self.field_shadowing = other.field_shadowing;
        }
        // ⛔ §7 step 4: every field needs a rule here, or `Block` scope cannot
        // inherit it and the field is a live bug rather than an omission. This
        // one was added without one.
        //
        // "May a spread argument bind by NAME here" governs the code doing the
        // CALLING — §3 question 1, the same shape as php's historical
        // `allow_call_time_pass_reference` — so a nested region may state it
        // and a `Some` wins over the enclosing frame like every other field.
        if other.spread_arguments.is_some() {
            self.spread_arguments = other.spread_arguments;
        }
        // A unit-level fact about the object model, not something a nested
        // region redefines — but it gets a rule all the same, because the note
        // above is right: a field with no rule here is dropped by any merge,
        // and "it could never differ" is exactly the reasoning that leaves a
        // live bug looking like an omission. Its STATIC twin
        // (`static_fields_are_own_properties`) still has no rule, along with
        // `method_receiver`, `missing_arg_is_undefined` and `exit_argument`;
        // left visible rather than quietly copied.
        if other.instance_fields_are_own_properties.is_some() {
            self.instance_fields_are_own_properties = other.instance_fields_are_own_properties;
        }
        // Question 1 through and through — it governs a region of code — so a
        // nested region stating it wins over the enclosing frame, like every
        // other §3-Q1 field here.
        if other.body_declarations_first.is_some() {
            self.body_declarations_first = other.body_declarations_first;
        }
        // A unit-level fact about the object model, like
        // `instance_fields_are_own_properties` above — a nested region has no
        // business redefining whether functions are objects. It gets a rule
        // regardless, for the reason stated there: a field with no rule here is
        // silently dropped by any merge.
        if other.functions_are_objects.is_some() {
            self.functions_are_objects = other.functions_are_objects;
        }
    }
}

/// What a shift or rotate count outside `[0, width)` does — see
/// [`Directives::shift_overflow`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ShiftOverflow {
    /// The count is taken modulo the lane width, so `1 << 32 == 1`. wasm's
    /// own rule (`i32.shl` masks to 5 bits), and therefore JS §13.9, Java
    /// §15.19 and C#.
    #[default]
    Mask,
    /// The result is zero once the count reaches the lane width. Fortran
    /// `ISHFT`/`SHIFTL`/`SHIFTR`: every bit has been shifted out, so nothing
    /// remains.
    Zero,
}

/// How far a directive statement's effect reaches — itself a language quirk,
/// so it is declared rather than assumed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DirectiveScope {
    /// Until the end of the enclosing block, then restored. A JS
    /// `"use strict"` at the head of a function body works this way.
    #[default]
    Block,
    /// Until the end of the module, surviving every intervening block end.
    /// Pascal's `{$R+}`/`{$R-}` and a C `#pragma` are positional like this —
    /// switched on halfway down a procedure, they stay on afterwards.
    Module,
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
        StmtKind::Assign { targets, value, .. } => {
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

impl Expression {
    /// Post-order mutable visitor: children first, then `f` on this node.
    ///
    /// THE traversal facility for normalization passes. Before this, every
    /// rewrite pass hand-rolled its own full `ExprKind` recursion inside a
    /// walker (the reason those files run to tens of thousands of lines), and
    /// each new `ExprKind` variant had to be added to every copy. Passes match
    /// on the shapes they care about inside `f` and ignore the rest.
    ///
    /// Coverage mirrors `expr_has_yield` — the best-exercised container map in
    /// the tree. Contract: a NEW container variant must be added here (and
    /// there); statement-carrying bodies (`Lambda`, `FunctionExpr`,
    /// `ClassExpr`) are NOT descended into — statement slots are the driving
    /// pass's job, exactly as they are for every existing pass.
    pub fn walk_exprs_mut(&mut self, f: &mut dyn FnMut(&mut Expression)) {
        match &mut self.kind {
            ExprKind::Unary { expr, .. }
            | ExprKind::IsType { expr, .. }
            | ExprKind::Cast { expr, .. }
            | ExprKind::TypeOf(expr)
            | ExprKind::Spread(expr)
            | ExprKind::Await(expr)
            | ExprKind::YieldFrom(expr)
            | ExprKind::RefLoad(expr)
            | ExprKind::Void(expr)
            | ExprKind::Delete(expr) => expr.walk_exprs_mut(f),
            ExprKind::Yield(inner) => {
                if let Some(expr) = inner {
                    expr.walk_exprs_mut(f);
                }
            }
            ExprKind::Async(op) => {
                for child in op.children_mut() {
                    child.walk_exprs_mut(f);
                }
            }
            ExprKind::Chan(op) => {
                for child in op.children_mut() {
                    child.walk_exprs_mut(f);
                }
            }
            ExprKind::Atomic(op) => {
                for child in op.children_mut() {
                    child.walk_exprs_mut(f);
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
                left.walk_exprs_mut(f);
                right.walk_exprs_mut(f);
            }
            ExprKind::StaticAccess { class, member } => {
                class.walk_exprs_mut(f);
                member.walk_exprs_mut(f);
            }
            ExprKind::Ternary { cond, then, else_ } => {
                cond.walk_exprs_mut(f);
                then.walk_exprs_mut(f);
                else_.walk_exprs_mut(f);
            }
            ExprKind::Member { object, .. } => object.walk_exprs_mut(f),
            ExprKind::CallableRef {
                target,
                receiver,
                adapter,
                ..
            } => {
                target.walk_exprs_mut(f);
                if let Some(receiver) = receiver {
                    receiver.walk_exprs_mut(f);
                }
                if let Some(adapter) = adapter {
                    match adapter {
                        CallableAdapter::Expr { body, .. } => body.walk_exprs_mut(f),
                    }
                }
            }
            ExprKind::Index { object, index, .. } => {
                object.walk_exprs_mut(f);
                index.walk_exprs_mut(f);
            }
            ExprKind::Call { callee, args, .. } => {
                callee.walk_exprs_mut(f);
                for arg in args {
                    arg.value.walk_exprs_mut(f);
                }
            }
            ExprKind::New { class, args } => {
                class.walk_exprs_mut(f);
                for arg in args {
                    arg.value.walk_exprs_mut(f);
                }
            }
            ExprKind::Proxy { target, handler } => {
                target.walk_exprs_mut(f);
                handler.walk_exprs_mut(f);
            }
            ExprKind::SuperCall { args, .. } => {
                for arg in args {
                    arg.value.walk_exprs_mut(f);
                }
            }
            ExprKind::Array(items) => {
                for item in items {
                    if let Some(key) = &mut item.key {
                        key.walk_exprs_mut(f);
                    }
                    item.value.walk_exprs_mut(f);
                }
            }
            ExprKind::Tuple(items)
            | ExprKind::Set(items)
            | ExprKind::Sequence(items)
            | ExprKind::Zip {
                iterables: items, ..
            } => {
                for item in items {
                    item.walk_exprs_mut(f);
                }
            }
            ExprKind::NamedTuple { fields, .. } => {
                for (_, value) in fields {
                    value.walk_exprs_mut(f);
                }
            }
            ExprKind::Object(props) => {
                for prop in props {
                    match prop {
                        ObjectProperty::KeyValue { key, value }
                        | ObjectProperty::Computed { key, value } => {
                            key.walk_exprs_mut(f);
                            value.walk_exprs_mut(f);
                        }
                        ObjectProperty::Spread(expr) => expr.walk_exprs_mut(f),
                        _ => {}
                    }
                }
            }
            ExprKind::Interpolation(parts) => {
                for part in parts {
                    match part {
                        InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                            expr.walk_exprs_mut(f)
                        }
                        InterpolPart::Text(_) => {}
                    }
                }
            }
            ExprKind::Match { subject, arms } => {
                subject.walk_exprs_mut(f);
                for arm in arms {
                    if let Some(conditions) = &mut arm.conditions {
                        for c in conditions {
                            c.walk_exprs_mut(f);
                        }
                    }
                    arm.body.walk_exprs_mut(f);
                }
            }
            ExprKind::Comprehension {
                element,
                generators,
                ..
            } => {
                element.walk_exprs_mut(f);
                for g in generators {
                    g.target.walk_exprs_mut(f);
                    g.iter.walk_exprs_mut(f);
                    for c in &mut g.conditions {
                        c.walk_exprs_mut(f);
                    }
                }
            }
            ExprKind::Slice { lower, upper, step } => {
                for slot in [lower, upper, step] {
                    if let Some(expr) = slot {
                        expr.walk_exprs_mut(f);
                    }
                }
            }
            // Unlike the yield traversals (which stop at function boundaries
            // because yield SCOPES to its function), a normalization pass must
            // reach nested bodies: the shapes being normalized appear inside
            // `async () => …` and function expressions.
            ExprKind::Lambda { body, .. } => match body {
                LambdaBody::Expr(expr) => expr.walk_exprs_mut(f),
                LambdaBody::Block(stmts) => {
                    for stmt in stmts {
                        stmt.walk_exprs_mut(f);
                    }
                }
            },
            ExprKind::FunctionExpr(decl) => decl.walk_exprs_mut(f),
            _ => {}
        }
        f(self);
    }
}

impl Statement {
    /// Statement half of [`Expression::walk_exprs_mut`]: visit every
    /// expression slot in this statement, recursing through nested statement
    /// bodies INCLUDING function and class declarations. Coverage mirrors
    /// `statement_has_yield` plus the declaration bodies it deliberately
    /// skips. Together the pair make a normalization pass one callback,
    /// instead of a hand-rolled recursion per pass per walker.
    pub fn walk_exprs_mut(&mut self, f: &mut dyn FnMut(&mut Expression)) {
        fn body(stmts: &mut [Statement], f: &mut dyn FnMut(&mut Expression)) {
            for stmt in stmts {
                stmt.walk_exprs_mut(f);
            }
        }
        match &mut self.kind {
            StmtKind::Expr(expr) => expr.walk_exprs_mut(f),
            StmtKind::Block(stmts) => body(stmts, f),
            StmtKind::Select { arms, default } => {
                for arm in arms {
                    for child in arm.comm.children_mut() {
                        child.walk_exprs_mut(f);
                    }
                    body(&mut arm.body, f);
                }
                if let Some(default) = default {
                    body(default, f);
                }
            }
            StmtKind::FunctionDecl {
                body: b, params, ..
            } => {
                for param in params {
                    if let Some(default) = &mut param.default {
                        default.walk_exprs_mut(f);
                    }
                }
                body(b, f);
            }
            StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    match member {
                        ClassMember::Field { init, .. } => {
                            if let Some(expr) = init {
                                expr.walk_exprs_mut(f);
                            }
                        }
                        ClassMember::Method(decl) | ClassMember::NestedType(decl) => {
                            decl.walk_exprs_mut(f)
                        }
                        ClassMember::Constructor { body: b, .. } => body(b, f),
                        ClassMember::Property { getter, setter, .. } => {
                            if let Some(stmts) = getter {
                                body(stmts, f);
                            }
                            if let Some(prop_setter) = setter {
                                body(&mut prop_setter.body, f);
                            }
                        }
                        ClassMember::Const { value, .. } => value.walk_exprs_mut(f),
                        _ => {}
                    }
                }
            }
            // InterfaceDecl members are SIGNATURES — no bodies to walk.
            StmtKind::NamespaceDecl { body: b, .. } => body(b, f),
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &mut decl.init {
                        init.walk_exprs_mut(f);
                    }
                }
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                cond.walk_exprs_mut(f);
                body(then_body, f);
                for (c, b) in elifs {
                    c.walk_exprs_mut(f);
                    body(b, f);
                }
                if let Some(b) = else_body {
                    body(b, f);
                }
            }
            StmtKind::For {
                init,
                cond,
                update,
                body: b,
            } => {
                if let Some(stmt) = init {
                    stmt.walk_exprs_mut(f);
                }
                for slot in [cond, update] {
                    if let Some(expr) = slot {
                        expr.walk_exprs_mut(f);
                    }
                }
                body(b, f);
            }
            StmtKind::ForIn {
                iter,
                body: b,
                else_body,
                ..
            } => {
                iter.walk_exprs_mut(f);
                body(b, f);
                if let Some(eb) = else_body {
                    body(eb, f);
                }
            }
            StmtKind::While {
                cond,
                body: b,
                else_body,
            } => {
                cond.walk_exprs_mut(f);
                body(b, f);
                if let Some(eb) = else_body {
                    body(eb, f);
                }
            }
            StmtKind::DoWhile { body: b, cond, .. } => {
                body(b, f);
                cond.walk_exprs_mut(f);
            }
            StmtKind::Switch {
                expr,
                cases,
                default,
            } => {
                expr.walk_exprs_mut(f);
                for case in cases {
                    body(&mut case.body, f);
                }
                if let Some(b) = default {
                    body(b, f);
                }
            }
            StmtKind::Try {
                body: b,
                catches,
                else_body,
                finally,
            } => {
                body(b, f);
                for catch in catches {
                    if let Some(when) = &mut catch.when_clause {
                        when.walk_exprs_mut(f);
                    }
                    body(&mut catch.body, f);
                }
                for slot in [else_body, finally] {
                    if let Some(stmts) = slot {
                        body(stmts, f);
                    }
                }
            }
            StmtKind::With { items, body: b, .. } => {
                for item in items {
                    item.expr.walk_exprs_mut(f);
                }
                body(b, f);
            }
            StmtKind::Using {
                resource, body: b, ..
            } => {
                resource.walk_exprs_mut(f);
                body(b, f);
            }
            StmtKind::Lock { expr, body: b } => {
                expr.walk_exprs_mut(f);
                body(b, f);
            }
            StmtKind::Return(expr) => {
                if let Some(expr) = expr {
                    expr.walk_exprs_mut(f);
                }
            }
            StmtKind::Throw { expr, cause } => {
                for slot in [expr, cause] {
                    if let Some(expr) = slot {
                        expr.walk_exprs_mut(f);
                    }
                }
            }
            StmtKind::Assign { targets, value, .. } => {
                for target in targets {
                    target.walk_exprs_mut(f);
                }
                value.walk_exprs_mut(f);
            }
            StmtKind::CompoundAssign { target, value, .. } => {
                target.walk_exprs_mut(f);
                value.walk_exprs_mut(f);
            }
            StmtKind::RaiseEvent { args, .. } | StmtKind::Delete(args) | StmtKind::Echo(args) => {
                for arg in args {
                    arg.walk_exprs_mut(f);
                }
            }
            _ => {}
        }
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
        ExprKind::Async(op) => op.children().into_iter().any(expr_has_yield),
        ExprKind::Chan(op) => op.children().into_iter().any(expr_has_yield),
        ExprKind::Atomic(op) => op.children().into_iter().any(expr_has_yield),
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
        ExprKind::CallableRef {
            target,
            receiver,
            adapter,
            ..
        } => {
            expr_has_yield(target)
                || receiver.as_ref().map_or(false, |expr| expr_has_yield(expr))
                || adapter.as_ref().map_or(false, |adapter| match adapter {
                    CallableAdapter::Expr { body, .. } => expr_has_yield(body),
                })
        }
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
        ExprKind::Tuple(items)
        | ExprKind::Set(items)
        | ExprKind::Sequence(items)
        | ExprKind::Zip {
            iterables: items, ..
        } => items.iter().any(expr_has_yield),
        ExprKind::NamedTuple { fields, .. } => {
            fields.iter().any(|(_, value)| expr_has_yield(value))
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
        StmtKind::Assign { targets, value, .. } => {
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
        ExprKind::CallableRef {
            target,
            receiver,
            adapter,
            ..
        } => {
            collect_rest_in_expr(target, out);
            if let Some(receiver) = receiver {
                collect_rest_in_expr(receiver, out);
            }
            if let Some(adapter) = adapter {
                match adapter {
                    CallableAdapter::Expr { body, .. } => collect_rest_in_expr(body, out),
                }
            }
        }
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
        ExprKind::Tuple(items)
        | ExprKind::Set(items)
        | ExprKind::Sequence(items)
        | ExprKind::Zip {
            iterables: items, ..
        } => {
            for item in items {
                collect_rest_in_expr(item, out);
            }
        }
        ExprKind::NamedTuple { fields, .. } => {
            for (_, value) in fields {
                collect_rest_in_expr(value, out);
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

#[cfg(test)]
mod protocol_slot_key_tests {
    use super::ProtocolSlot;

    /// Every slot in the exhaustive `slot_id` match must round-trip through its
    /// string key, and no two slots may share one.
    ///
    /// `as_key`/`from_key` were GENERATED from `slot_id`'s arms rather than
    /// transcribed, precisely because 93 hand-written pairs is where a slip
    /// hides — and a collision would silently alias two protocols, so a profile
    /// declaring `slot = "eq"` could bind `Ne`. This is the check that makes
    /// the generated pair trustworthy; it walks slot ids rather than a
    /// `ProtocolSlot::ALL` rather than a hand-listed set, so a slot added later
    /// is covered without editing it.
    #[test]
    fn every_slot_key_roundtrips_and_is_unique() {
        let mut seen: std::collections::HashMap<&'static str, ProtocolSlot> =
            std::collections::HashMap::new();
        let mut count = 0;
        for slot in ProtocolSlot::ALL {
            count += 1;
            let key = slot.as_key();
            assert_eq!(
                ProtocolSlot::from_key(key),
                Some(slot),
                "{slot:?} does not round-trip through key {key:?}"
            );
            if let Some(other) = seen.insert(key, slot) {
                panic!("key {key:?} is shared by {other:?} and {slot:?}");
            }
        }
        assert_eq!(count, 102, "ProtocolSlot::ALL lost a slot");
    }

    /// An unrecognised key is ignored, not an error: a profile written against
    /// a newer toolchain must still load, with the unknown declaration simply
    /// having no effect.
    #[test]
    fn an_unknown_slot_key_is_none() {
        assert_eq!(ProtocolSlot::from_key("no_such_slot"), None);
        assert_eq!(ProtocolSlot::from_key(""), None);
    }
}
