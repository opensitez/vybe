//! `NormalClass` and its component types — the compile-time IR for a
//! normalised class declaration. Produced by each language's walker
//! and consumed by `emit::emit_class`.
//!
//! Design rules:
//! - Every user-visible name the walker observed is preserved on
//!   `source_name` fields for diagnostics.
//! - Canonical names (`canonical_name` fields) are the cross-language
//!   identity. Dispatch at runtime keys on canonical names.
//! - Every statement / expression reference is a raw AST node —
//!   `emit::emit_class` compiles bodies to chunks as a separate step.
//!
//! See `classnormalization.md` for the full shape rationale.

use crate::ast::{Argument, Expression, Modifiers, Param, Span, Statement};

/// The normalised class declaration — single source of truth for every
/// per-language class idiom after the walker has flattened it.
#[derive(Debug, Clone)]
pub struct NormalClass {
    pub span: Span,
    pub name: String,

    /// Single superclass name. Multiple-inheritance / mixins / traits
    /// are flattened at walker time — see `interfaces` for instanceof
    /// identity and the method list for flattened dispatch.
    pub parent: Option<String>,

    /// Implemented / mixed-in / included interfaces + mixins. Used only
    /// for `instanceof` / `isinstance` / `is` / `kind_of?` identity
    /// checks; method dispatch never walks this list.
    pub interfaces: Vec<String>,

    pub is_abstract: bool,
    pub is_sealed: bool,
    /// Walker merged all partial parts before producing this value.
    /// Flag is informational / diagnostic — `emit_class` ignores it.
    pub is_partial: bool,

    pub instance_fields: Vec<NormalField>,
    pub static_fields: Vec<NormalField>,
    pub instance_methods: Vec<NormalMethod>,
    pub static_methods: Vec<NormalMethod>,
    pub properties: Vec<NormalProperty>,
    pub constructor: Option<NormalConstructor>,
    pub destructor: Option<NormalMethod>,

    /// Methods the compiler auto-calls at the start of each constructor
    /// body — e.g. `["InitializeComponent"]` for .NET forms. Walker
    /// populates; `emit_class` emits the call sequence unconditionally.
    pub auto_init_methods: Vec<String>,

    /// Operator / protocol methods, cross-referenced from the normal
    /// method list. `kind` is the canonical cross-language concept
    /// (`ToString`, `Add`, `Iterator`, …); `canonical_name` matches
    /// the method's `canonical_name` in `instance_methods`.
    pub special_methods: Vec<SpecialMethod>,

    /// VB `Handles ctrl.Event` bindings. Walker extracts; `emit_class`
    /// emits the corresponding `vybe:gui.bindEvent` calls during
    /// constructor compilation.
    pub event_bindings: Vec<EventBinding>,

    /// ClassMembers the normalizer doesn't explicitly model yet
    /// (`ClassMember::Event`, `::Const`, `::NestedType`). Shim
    /// reconstruction appends these back verbatim so the legacy
    /// `compile_class` path still sees them. Each language's walker
    /// copies its unhandled members here. Phase 2b.2 eventually
    /// replaces this pass-through by adding first-class fields to
    /// `NormalClass` for every member kind actually in use.
    pub raw_extra_members: Vec<crate::ast::ClassMember>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Access {
    Public,
    Protected,
    Internal,   // package / assembly visibility
    Private,
}

#[derive(Debug, Clone)]
pub struct NormalField {
    pub span: Span,
    pub name: String,
    pub init: Option<Expression>,
    pub access: Access,
    pub readonly: bool,
}

#[derive(Debug, Clone)]
pub struct NormalMethod {
    pub span: Span,
    /// Cross-language canonical name (the vtable key at runtime).
    /// E.g. Python `__str__` → `"tostring"`.
    pub canonical_name: String,
    /// Name as it appeared in the source file — preserved for error
    /// messages and populated into `ClassType.method_aliases` so
    /// callers from any language find the method.
    pub source_name: String,
    /// Additional alias names the walker computed (e.g. a VB method
    /// tagged `Implements IDisposable.Dispose` gets both `Dispose`
    /// and `dispose` aliased).
    pub aliases: Vec<String>,
    pub params: Vec<Param>,
    pub return_type: Option<String>,
    pub body: Vec<Statement>,
    pub access: Access,
    pub is_virtual: bool,
    pub is_override: bool,
    pub is_async: bool,
    pub is_generator: bool,
    pub is_abstract: bool,
    /// Whether the source declared this as a `Sub` (no return) vs a
    /// function. VB / Pascal care; other languages leave as `false`.
    pub is_sub: bool,
    /// Full original `Modifiers` the walker saw, preserved verbatim so
    /// reconstruction into `ClassMember::Method` is lossless for
    /// language-specific flags (`is_readonly`, `is_shared`,
    /// `is_extension`, `is_overloads`, `is_not_overridable`,
    /// `decorators`) that the per-language compile path reads but
    /// that aren't first-class in `NormalMethod`. The canonical
    /// fields (`is_virtual`, `is_override`, `is_abstract`, `access`)
    /// remain authoritative; `raw_modifiers` is just a carrier.
    pub raw_modifiers: Modifiers,
}

#[derive(Debug, Clone)]
pub struct NormalConstructor {
    pub span: Span,
    pub params: Vec<Param>,
    pub body: Vec<Statement>,
    /// Whether / how the parent constructor is called at the start of
    /// this ctor body. Resolved by the walker from source syntax +
    /// profile defaults — no profile branching in `emit_class`.
    pub base_call: BaseCall,
    /// Dart named constructors: `ClassName.named(args)` — carry the
    /// name suffix so `emit_class` can emit it as a named factory.
    /// `None` for the unnamed / primary ctor.
    pub named_name: Option<String>,
}

#[derive(Debug, Clone)]
pub enum BaseCall {
    /// User wrote `super(args)` / `MyBase.New(args)` / `: base(args)`.
    Explicit(Vec<Argument>),
    /// Profile says auto-call parent ctor with no args (C# default,
    /// VB when no `MyBase.New` in source). Walker promotes to this;
    /// compiler emits `super()` preamble.
    Auto,
    /// JS root class (no `extends`) or explicit no-op.
    None,
}

#[derive(Debug, Clone)]
pub struct NormalProperty {
    pub span: Span,
    pub canonical_name: String,
    pub source_name: String,
    pub getter: Option<NormalMethod>,
    pub setter: Option<NormalMethod>,
    /// For C# `{ get; set; }` auto-properties: the backing field name.
    /// `None` for fully-implemented properties.
    pub auto_field: Option<String>,
}

#[derive(Debug, Clone)]
pub struct SpecialMethod {
    pub kind: SpecialMethodKind,
    /// Matches the `canonical_name` of a method in the same class.
    pub canonical_name: String,
    /// Original name in source (for diagnostics).
    pub source_name: String,
}

/// Cross-language operator + protocol method identity. Every language
/// that defines any of these concepts under its own name resolves to
/// the same `SpecialMethodKind`. Consumers of the class access the
/// behaviour via the corresponding canonical method (e.g. `ToString`
/// → `canonical_name = "tostring"`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SpecialMethodKind {
    // ── Coercion / representation ───────────────────────────────────
    ToString,    // JS toString, C# ToString, Python __str__, Ruby to_s, PHP __toString
    Repr,        // Python __repr__, Ruby inspect
    ValueOf,     // JS valueOf, Python __int__/__float__
    ToPrimitive, // JS Symbol.toPrimitive

    // ── Iteration ───────────────────────────────────────────────────
    Iterator,      // JS Symbol.iterator, Python __iter__, Ruby each, C# GetEnumerator
    AsyncIterator, // JS Symbol.asyncIterator
    Next,          // JS iterator.next, Python __next__

    // ── Arithmetic operators ────────────────────────────────────────
    Add, Sub, Mul, Div, Mod, Pow, Neg,

    // ── Comparison ──────────────────────────────────────────────────
    Eq,       // ==
    Compare,  // <=> (Ruby) / __cmp__ (Python legacy) / CompareTo (C#)
    Lt, Le, Gt, Ge,

    // ── Bitwise ─────────────────────────────────────────────────────
    And, Or, Xor, Not,
    LShift, RShift,

    // ── Container protocol ──────────────────────────────────────────
    Len,       // len() / length / size / Count
    GetItem,   // Python __getitem__, Ruby [], Dart operator []
    SetItem,   // Python __setitem__, Ruby []=, Dart operator []=
    DelItem,   // Python __delitem__
    Contains,  // Python __contains__, Ruby include?

    // ── Callable / reflection ───────────────────────────────────────
    Call,         // Python __call__, PHP __invoke, Dart call, C# ()
    HasInstance,  // JS Symbol.hasInstance, Python __instancecheck__

    // ── Property access interception ────────────────────────────────
    GetAttr,  // Python __getattr__, PHP __get, JS Proxy get
    SetAttr,  // Python __setattr__, PHP __set, JS Proxy set
    DelAttr,  // Python __delattr__, PHP __unset

    // ── Context managers ────────────────────────────────────────────
    Enter,  // Python __enter__
    Exit,   // Python __exit__

    // ── Hash ────────────────────────────────────────────────────────
    Hash,  // Python __hash__, Ruby hash, C# GetHashCode, Java hashCode
}

#[derive(Debug, Clone)]
pub struct EventBinding {
    pub control: String, // "btn1"
    pub event: String,   // "Click"
    pub handler: String, // method name on this class
}
