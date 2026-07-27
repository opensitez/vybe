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

use vybe_ast::{Argument, Expression, Modifiers, Param, Span, Statement};

/// The normalised class declaration — single source of truth for every
/// per-language class idiom after the walker has flattened it.
#[derive(Debug, Clone)]
pub struct NormalClass {
    pub span: Span,
    pub name: String,

    /// Primary superclass name (`bases.first()`). Single-inheritance
    /// languages use only this; it drives the constructor chain and the
    /// static `super`/`__base_<name>` snapshot. Mixins / traits that a
    /// language chooses to flatten still land in the method list + `interfaces`.
    pub parent: Option<String>,

    /// ALL declared direct bases, in source order (`parent == bases.first()`).
    /// Filled centrally in `emit_class_from_ast` from the AST `parents`, so
    /// per-language normalizers don't populate it. Only consumed when the
    /// profile opts into `class_multiple_inheritance` (Python C3 MI); every
    /// other language ignores `bases[1..]`, keeping bytecode byte-identical.
    pub bases: Vec<String>,

    /// Implemented / mixed-in / included interfaces + mixins. Used only
    /// for `instanceof` / `isinstance` / `is` / `kind_of?` identity
    /// checks; method dispatch never walks this list.
    pub interfaces: Vec<String>,

    pub is_abstract: bool,
    pub is_sealed: bool,
    /// Walker merged all partial parts before producing this value.
    /// Flag is informational / diagnostic — `emit_class` ignores it.
    pub is_partial: bool,
    /// Value types (for example C# structs) still use the shared class
    /// pipeline, but they need different default-value semantics.
    pub is_value_type: bool,

    /// Every instance method (including the constructor) has `self` as
    /// the first positional parameter, as in Python (`def f(self, …)`).
    /// `emit_class` skips the first param when binding user parameters.
    /// For languages where `self`/`this` is an implicit slot (JS / VB /
    /// C# / Ruby / PHP / Dart / Pascal), this stays `false`.
    pub explicit_self_param: bool,

    /// Bare identifiers inside instance methods resolve to fields on
    /// `self` before falling through to locals / globals, as in Python
    /// and VB.NET. Walker-set; threaded through the compiler via
    /// `current_class_implicit_self` so `calls.rs` / expression
    /// resolution can honour it without consulting the profile.
    pub implicit_self_fields: bool,

    pub instance_fields: Vec<NormalField>,
    pub static_fields: Vec<NormalField>,
    pub instance_methods: Vec<NormalMethod>,
    pub static_methods: Vec<NormalMethod>,
    pub properties: Vec<NormalProperty>,
    pub constructors: Vec<NormalConstructor>,
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
    pub raw_extra_members: Vec<vybe_ast::ClassMember>,

    /// Declared AUGMENTATIONS — where this class's members come from besides
    /// its own body (PHP traits, Dart mixins, Ruby include/prepend, Java
    /// interface defaults, Go field promotion, Dart `extension on MyClass`).
    ///
    /// A language's normalizer declares these as DATA; the compiler's
    /// `class_augmentation` pass applies them ONCE, before member
    /// registration. Empty for a language that has not been migrated yet —
    /// those still fold in their own walker, which is exactly the duplication
    /// this replaces. See flexclassplan.md §4c.
    pub augmentations: Vec<Augmentation>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Access {
    Public,
    Protected,
    Internal, // package / assembly visibility
    Private,
}

#[derive(Debug, Clone)]
pub struct NormalField {
    pub span: Span,
    pub name: String,
    pub type_hint: Option<String>,
    pub init: Option<Expression>,
    pub array_bounds: Option<Vec<Expression>>,
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
    /// User wrote `: this(args)` / `this(args)` to chain to a sibling
    /// constructor on the same class.
    This(Vec<Argument>),
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
    pub is_static: bool,
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
    Add,
    Sub,
    Mul,
    Div,
    Mod,
    Pow,
    Neg,

    // ── Comparison ──────────────────────────────────────────────────
    Eq,      // ==
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
    Len,      // len() / length / size / Count
    GetItem,  // Python __getitem__, Ruby [], Dart operator []
    SetItem,  // Python __setitem__, Ruby []=, Dart operator []=
    DelItem,  // Python __delitem__
    Contains, // Python __contains__, Ruby include?

    // ── Callable / reflection ───────────────────────────────────────
    Call,        // Python __call__, PHP __invoke, Dart call, C# ()
    HasInstance, // JS Symbol.hasInstance, Python __instancecheck__

    // ── Property access interception ────────────────────────────────
    GetAttr, // Python __getattr__, PHP __get, JS Proxy get
    SetAttr, // Python __setattr__, PHP __set, JS Proxy set
    DelAttr, // Python __delattr__, PHP __unset

    // ── Context managers ────────────────────────────────────────────
    Enter, // Python __enter__
    Exit,  // Python __exit__

    // ── Hash ────────────────────────────────────────────────────────
    Hash, // Python __hash__, Ruby hash, C# GetHashCode, Java hashCode
}

#[derive(Debug, Clone)]
pub struct EventBinding {
    pub control: String, // "btn1"
    pub event: String,   // "Click"
    pub handler: String, // method name on this class
}

// ── Class augmentation ──────────────────────────────────────────────────
//
// Where a class's members come from BESIDES its own body: PHP traits, Dart
// mixins, Ruby `include`/`prepend`, Java interface defaults, Go field
// promotion, Dart `extension E on MyClass`.
//
// One vocabulary, per-language declared data — NOT one algorithm. These
// mechanisms differ in KIND, so a single fold would be wrong for most of them.
// See flexclassplan.md §4c.
//
// "Augmentation" is reserved for this concept. The primitive prototype
// FALLBACK (§4d) is a different mechanism and never uses the word.

/// How an augmenting type's members reach the class.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AugmentationMode {
    /// Members are duplicated into the class. PHP traits, Dart mixins,
    /// Dart `extension E on MyClass`.
    Copy,
    /// Members are inserted into the LOOKUP ORDER, not copied. Ruby
    /// `include`/`prepend`; Java default methods resolve at dispatch.
    Chain,
    /// Members are PROMOTED from an inner value and the receiver rebinds to
    /// it. Go field promotion — and Go's own spec word, chosen because
    /// "delegate" already means a first-class function type in this codebase
    /// (C# `delegate_declaration`, `vybe_compiler::emitter/src/delegates.rs`).
    Promote,
}

/// Where the augmenting type sits relative to the class's own members.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AugmentationPosition {
    /// The class's own members win. PHP traits, Dart mixins, Ruby `include`.
    AfterOwn,
    /// The augmenting type wins over the class's own members. Ruby `prepend`.
    BeforeOwn,
}

/// What happens when two augmenting types supply the same member name.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AugmentationConflict {
    /// Later augmentation overrides earlier. Dart mixin linearization.
    LastWins,
    /// Earlier wins; later is ignored.
    FirstWins,
    /// A diagnosable error. Go promotion at EQUAL depth; Java default-method
    /// diamonds. Silently picking one is a bug, not a policy.
    Error,
    /// An error unless the class explicitly resolves it (PHP `insteadof`,
    /// Java overriding the diamond, `X.super.m()`).
    RequireExplicit,
}

/// What `super` means inside an augmenting type's member.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AugmentationSuper {
    /// The augmented class's own parent. PHP traits.
    OwnParent,
    /// The next entry in the resolution order — NOT the augmenting type's
    /// own parent. Dart mixins, Ruby modules.
    NextInOrder,
}

/// A per-member adjustment applied while augmenting: PHP `as` (rename and/or
/// change visibility) and `insteadof` (exclude).
#[derive(Debug, Clone, Default)]
pub struct AugmentationAdjustment {
    /// Source member name this applies to.
    pub member: String,
    /// Bind under this name instead (PHP `as other`).
    pub rename_to: Option<String>,
    /// Override the member's visibility (PHP `as protected foo`). Uses the
    /// normalized model's `Access`, not the AST `Visibility`, so the record
    /// stays in the same vocabulary as `NormalMethod.access`.
    pub visibility: Option<Access>,
    /// Drop this member from THIS augmentation (PHP `insteadof`).
    pub exclude: bool,
}

/// Which member kinds may cross from the augmenting type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AugmentationContributes {
    pub methods: bool,
    pub fields: bool,
    /// PHP quirk: a trait's static property gives each using class its OWN
    /// copy, not a shared one.
    pub statics: bool,
    /// Dart mixins declare no constructors.
    pub constructors: bool,
    pub abstract_members: bool,
}

impl Default for AugmentationContributes {
    fn default() -> Self {
        Self {
            methods: true,
            fields: true,
            statics: false,
            constructors: false,
            abstract_members: true,
        }
    }
}

/// One declared augmentation of a class. A language's normalizer produces
/// these as DATA; the compiler's `class_augmentation` pass applies them once.
#[derive(Debug, Clone)]
pub struct Augmentation {
    /// The augmenting type's name. For `Promote` this is the field's type,
    /// and `via_field` names the field.
    pub from: String,
    /// `Promote` only: the field the receiver rebinds to (Go promotion).
    pub via_field: Option<String>,
    pub mode: AugmentationMode,
    pub position: AugmentationPosition,
    pub conflict: AugmentationConflict,
    pub super_target: AugmentationSuper,
    pub adjustments: Vec<AugmentationAdjustment>,
    pub contributes: AugmentationContributes,
    /// `Promote` only: promotion depth. Shallower wins; EQUAL depth with the
    /// same name is an `Error` (Go).
    pub depth: u8,
}
