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

use crate::{Argument, Expression, Modifiers, Param, Span, Statement};

/// The protocol-slot vocabulary lives in the AST, because the WALKER is what
/// knows a method's role — Python from a reserved name, Dart from `operator`
/// syntax, Java from conformance. Recovering it afterwards from the method's
/// spelling is what made the name the identity. Re-exported under its original
/// name so existing consumers are unaffected.
pub use crate::{PROTOCOL_SLOT_TABLE, ProtocolSlot, ProtocolSlot as SpecialMethodKind};

/// Platform/type-construction metadata attached to a normalized class whose
/// parent resolves to a registered platform type rather than a user class.
///
/// This mirrors the semantic subset the compiler needs from the namespace
/// registry without making normalized class IR depend on the bytecode crate.
#[derive(Debug, Clone, Default)]
pub struct PlatformBaseSpec {
    pub params: Vec<String>,
    pub fields: Vec<String>,
    pub ancestry: Vec<String>,
    pub control_fn: Option<String>,
    pub field_gui: Vec<PlatformFieldGui>,
    pub value_equality: bool,
}

/// How a platform constructor arg maps onto a GUI/control-backed field.
#[derive(Debug, Clone)]
pub enum PlatformFieldGui {
    NestOrProp(String),
    Children,
    Event(String),
    Caption,
}

impl Default for PlatformFieldGui {
    fn default() -> Self {
        PlatformFieldGui::NestOrProp(String::new())
    }
}

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

    /// ClassMembers the normalizer doesn't explicitly model yet
    /// (`ClassMember::Event`, `::Const`, `::NestedType`). Shim
    /// reconstruction appends these back verbatim so the legacy
    /// `compile_class` path still sees them. Each language's walker
    /// copies its unhandled members here. Phase 2b.2 eventually
    /// replaces this pass-through by adding first-class fields to
    /// `NormalClass` for every member kind actually in use.
    pub raw_extra_members: Vec<crate::ClassMember>,

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

    /// LINK classes synthesized for this class's `NextInOrder` augmentations,
    /// parent-first, each already present in the compiler's class map.
    ///
    /// A Dart mixin's `super` means "the next entry in the linearization", and
    /// a flat member list has no "next" — so those members are not copied into
    /// the class at all. They are given real classes spliced into the parent
    /// chain (`Base ← D&A ← D&A&B ← D`), which is how the Dart VM names them
    /// (`_D&Base&A`), how Ruby's `ancestors` reads, and how a C3 MRO is built.
    /// `super` then resolves through the ordinary prototype hop with no
    /// change to super emission at all — flexclassplan.md §4c-R.
    ///
    /// Emission must emit these BEFORE the class itself: the class's parent is
    /// the last of them. Empty for every class with no such augmentation, so a
    /// language that declares `OwnParent` (PHP traits, Go embedding) is
    /// untouched.
    pub synthesized_bases: Vec<String>,

    /// The class's parent is a registered PLATFORM type (a `Type` node in the
    /// namespace tree), not a user class — so this carries that type's
    /// construction spec.
    ///
    /// Recorded by the declaration pass, which is the only place that can tell
    /// the two apart: `class TForm1 = class(TForm)` and `class B extends A`
    /// are the same syntax, and only the tree knows `TForm` is data while `A`
    /// is a compiled class. Without it the emitter reaches for a constructor
    /// global that never existed (`global.get (tform)` → undefined) and the
    /// only way to satisfy it is to manufacture one — which is the
    /// compiler-side registration pass this design removes.
    ///
    /// `Some` means: this class gains the spec's fields, constructs its backing
    /// control through `control_fn`, and stamps the spec's `ancestry` — the
    /// same contribution a mixin makes, from a source that is already data.
    pub platform_base: Option<PlatformBaseSpec>,
}

/// The member buckets a normalizer fills while walking a class body.
///
/// Every normalizer declared these as 5–9 separate `let mut` locals (87 across
/// the twelve), filled them in a `match member { … }` loop, then spelled every
/// one back out in the `NormalClass` literal. This is that accumulation, once.
///
/// It is deliberately an ACCUMULATOR and not a partitioner: the caller says
/// which bucket a member belongs in. Static-vs-instance and name
/// canonicalisation are language RULES — Ruby's trailing `private`, VB's
/// `Shared`, Python's `@staticmethod`, Pascal's `class procedure` — and a
/// shared `partition_members(members, policy)` would drag them into common
/// code. The walker keeps the decision; this only holds the result.
#[derive(Debug, Clone, Default)]
pub struct NormalMembers {
    pub instance_fields: Vec<NormalField>,
    pub static_fields: Vec<NormalField>,
    pub instance_methods: Vec<NormalMethod>,
    pub static_methods: Vec<NormalMethod>,
    pub properties: Vec<NormalProperty>,
    pub constructors: Vec<NormalConstructor>,
    pub constructor: Option<NormalConstructor>,
    pub destructor: Option<NormalMethod>,
    pub auto_init_methods: Vec<String>,
    pub special_methods: Vec<SpecialMethod>,
    pub raw_extra_members: Vec<crate::ClassMember>,
    pub augmentations: Vec<Augmentation>,
}

impl NormalMembers {
    /// Route a field to the static or instance bucket. The caller decides
    /// `is_static` — that is the language's rule, not this type's.
    pub fn push_field(&mut self, is_static: bool, field: NormalField) {
        if is_static {
            self.static_fields.push(field);
        } else {
            self.instance_fields.push(field);
        }
    }

    /// Route a method to the static or instance bucket.
    pub fn push_method(&mut self, is_static: bool, method: NormalMethod) {
        if is_static {
            self.static_methods.push(method);
        } else {
            self.instance_methods.push(method);
        }
    }

    /// Record a constructor — ONE path, whether the language overloads or not.
    ///
    /// `constructors` (the list) and `constructor` (the single view) are two
    /// representations of one concept, and `classes.rs` BRANCHES on which is
    /// populated (`if !class.constructors.is_empty()` → per-variant dispatch).
    /// Languages written at different times picked different ones, so the
    /// shared compiler carries two emit paths for the same thing — the
    /// duplication this file exists to remove, sitting in the common compiler
    /// rather than in a language.
    ///
    /// This always fills both, putting every migrated language on the
    /// per-variant path. Measured: cobol 88/9 unchanged, and fortran identical
    /// either way — the paths agree for a single constructor. Once every
    /// normalizer is migrated, `constructor` becomes a derived view and the
    /// second emit path in `classes.rs` deletes.
    ///
    /// The single view is the PRIMARY constructor: an UNNAMED one when the
    /// language has named constructors (Dart's `Point(this.x)` beside
    /// `Point.origin()`, Lua's named factories), otherwise simply the first.
    /// `named_name` already records the distinction, so no caller has to
    /// restate it — a language with no named constructors leaves it `None`
    /// everywhere and gets first-wins unchanged.
    pub fn push_constructor(&mut self, ctor: NormalConstructor) {
        let takes_primary_slot = match &self.constructor {
            None => true,
            // An unnamed constructor displaces a named one already in the
            // slot; it never displaces another unnamed one (first wins).
            Some(held) => held.named_name.is_some() && ctor.named_name.is_none(),
        };
        if takes_primary_slot {
            self.constructor = Some(ctor.clone());
        }
        self.constructors.push(ctor);
    }

    /// Record an augmentation the class body declared (`ClassMember::Augment`),
    /// under this language's policy.
    ///
    /// The language states its POLICY once — what mode, whose members win, what
    /// `super` means — and the AST supplies the per-clause data (which type,
    /// which adjustments). Neither the walker nor this type folds anything; the
    /// shared `class_augmentation` pass does that once, for every language.
    /// See flexclassplan.md §4c-R.
    pub fn push_augment_decl(&mut self, decl: &crate::AugmentDecl, policy: AugmentationPolicy) {
        self.augmentations.push(policy.applied_to(decl));
    }

    /// Re-derive the single view from `constructors` under the same primary
    /// rule `push_constructor` applies.
    ///
    /// For a language that REWRITES constructor bodies after collecting them —
    /// Pascal rewrites static-value members, implicit `Self.` qualification and
    /// GCL property accessors across the whole list — the view taken at push
    /// time is a clone of the ORIGINAL, so it would silently carry
    /// un-rewritten code. Calling this after the rewrites lands makes the view
    /// agree with the list again.
    pub fn resync_constructor_view(&mut self) {
        self.constructor = self
            .constructors
            .iter()
            .find(|c| c.named_name.is_none())
            .or_else(|| self.constructors.first())
            .cloned();
    }

    /// Route a field the language declares STATIC but which must ALSO be
    /// readable through an instance — Python's class attributes (`A.kind` and
    /// `a.kind` both resolve), and Java's `instance.staticField`.
    ///
    /// `instance_init` is what the instance copy initialises FROM — a read of
    /// the class attribute, not a re-evaluation of the original initialiser, so
    /// a mutable class attribute stays ONE shared object across instances
    /// (`A.items` is the classic Python gotcha) and `a.kind = x` shadows
    /// without touching `A.kind`. The expression is built by the language,
    /// because only it knows how to name the class attribute.
    pub fn push_static_field_readable_on_instances(
        &mut self,
        field: NormalField,
        instance_init: Expression,
    ) {
        self.instance_fields.push(NormalField {
            init: Some(instance_init),
            ..field.clone()
        });
        self.static_fields.push(field);
    }
}

impl NormalClass {
    /// Attach accumulated members. Pairs with `Default` so a normalizer states
    /// only its class-level facts:
    ///
    /// ```ignore
    /// NormalClass { span, name, parent, ..Default::default() }.with_members(members)
    /// ```
    pub fn with_members(mut self, m: NormalMembers) -> Self {
        self.instance_fields = m.instance_fields;
        self.static_fields = m.static_fields;
        self.instance_methods = m.instance_methods;
        self.static_methods = m.static_methods;
        self.properties = m.properties;
        self.constructors = m.constructors;
        self.constructor = m.constructor;
        self.destructor = m.destructor;
        self.auto_init_methods = m.auto_init_methods;
        self.special_methods = m.special_methods;
        self.raw_extra_members = m.raw_extra_members;
        // Appended, not assigned: a language may declare augmentations in the
        // `NormalClass` literal too (Dart reads its `with` clause from the class
        // header, which is not a member), and those must not be dropped when the
        // body also contributes `Augment` members.
        self.augmentations.extend(m.augmentations);
        self
    }

    /// Fold another PARTIAL declaration of the same type into this one.
    ///
    /// A type's members do not always arrive in one syntactic declaration. Go
    /// writes methods outside the type — `func (t Tag) String() string` — so
    /// its walker emits one `StructDecl` per method, each a partial view of the
    /// same struct. Replacing on the second one leaves the type holding only
    /// its LAST method; this appends instead.
    ///
    /// Own members win: a name already present is not overwritten, so folding
    /// is order-independent for anything the type declares once (which is all
    /// a source language permits).
    pub fn merge_partial(&mut self, other: NormalClass) {
        for field in other.instance_fields {
            if !self.instance_fields.iter().any(|f| f.name == field.name) {
                self.instance_fields.push(field);
            }
        }
        for field in other.static_fields {
            if !self.static_fields.iter().any(|f| f.name == field.name) {
                self.static_fields.push(field);
            }
        }
        for method in other.instance_methods {
            if !self
                .instance_methods
                .iter()
                .any(|m| m.canonical_name == method.canonical_name)
            {
                self.instance_methods.push(method);
            }
        }
        for method in other.static_methods {
            if !self
                .static_methods
                .iter()
                .any(|m| m.canonical_name == method.canonical_name)
            {
                self.static_methods.push(method);
            }
        }
        for property in other.properties {
            if !self
                .properties
                .iter()
                .any(|p| p.canonical_name == property.canonical_name)
            {
                self.properties.push(property);
            }
        }
        for special in other.special_methods {
            if !self
                .special_methods
                .iter()
                .any(|s| s.canonical_name == special.canonical_name)
            {
                self.special_methods.push(special);
            }
        }
        for aug in other.augmentations {
            if !self
                .augmentations
                .iter()
                .any(|a| a.from == aug.from && a.via_field == aug.via_field)
            {
                self.augmentations.push(aug);
            }
        }
        self.constructors.extend(other.constructors);
        if self.constructor.is_none() {
            self.constructor = other.constructor;
        }
        if self.destructor.is_none() {
            self.destructor = other.destructor;
        }
        self.raw_extra_members.extend(other.raw_extra_members);
    }
}

/// The neutral class: no bases, no members, no language quirks enabled.
///
/// Written out rather than derived so each choice is a stated decision, not an
/// accident of field order — a normalizer that omits a field is relying on
/// these values, and a wrong one changes a language's semantics silently.
///
/// Every `bool` here is "this language does NOT do that": `explicit_self_param`
/// false means `self`/`this` is an implicit slot (JS/VB/C#/Ruby/PHP/Dart/
/// Pascal), `implicit_self_fields` false means bare identifiers do not resolve
/// to fields first (everything except Python and VB). The collections start
/// empty and the options `None` because "not declared" is the absence of a
/// member, never a placeholder one.
///
/// The point is that a normalizer states only what its language actually says.
/// Before this, every one of the twelve spelled out 34–90 fields exhaustively,
/// so adding a single field to this struct meant editing twelve crates and
/// writing twelve lines that carried no information. Now it costs nothing.
impl Default for NormalClass {
    fn default() -> Self {
        NormalClass {
            span: Span::default(),
            // A real class always names itself; this is a placeholder that
            // every construction path overwrites.
            name: String::new(),
            parent: None,
            bases: Vec::new(),
            interfaces: Vec::new(),
            is_abstract: false,
            is_sealed: false,
            is_partial: false,
            is_value_type: false,
            explicit_self_param: false,
            implicit_self_fields: false,
            instance_fields: Vec::new(),
            static_fields: Vec::new(),
            instance_methods: Vec::new(),
            static_methods: Vec::new(),
            properties: Vec::new(),
            constructors: Vec::new(),
            constructor: None,
            destructor: None,
            auto_init_methods: Vec::new(),
            special_methods: Vec::new(),
            raw_extra_members: Vec::new(),
            augmentations: Vec::new(),
            // Filled only by the compiler's augmentation pass; a normalizer
            // declares augmentations, never their lowering.
            synthesized_bases: Vec::new(),
            platform_base: None,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Access {
    Public,
    Protected,
    Internal, // package / assembly visibility
    Private,
}

/// The AST's declared visibility IS the normalized access level — the two
/// vocabularies match one-for-one, and every language that has visibility maps
/// them the same way.
///
/// Five normalizers (java, vb, ruby, php, csharp) each carried a byte-identical
/// `fn access_from_visibility` doing exactly this. A language with a genuinely
/// different rule (Ruby's `private` applying to everything after it, Dart's
/// leading-underscore convention) still decides that in its own walker and
/// passes the `Access` it means — this only removes the copies of the mapping
/// that never differed.
impl From<crate::Visibility> for Access {
    fn from(v: crate::Visibility) -> Self {
        match v {
            crate::Visibility::Public => Access::Public,
            crate::Visibility::Protected => Access::Protected,
            crate::Visibility::Private => Access::Private,
            crate::Visibility::Internal => Access::Internal,
        }
    }
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
    /// Name as it appeared in the source file. This is what the emitter binds
    /// the member under (`canon(source_name)`), so it is the member's callable
    /// identity, not just a diagnostic.
    ///
    /// A method reachable under a SECOND name is not an extra spelling in a
    /// list — it is a distinct binding with its own visibility (PHP
    /// `A::run as protected go;` leaves `run` public). Build it with
    /// [`NormalMethod::bound_as`], and mark cross-language ROLES with
    /// `SpecialMethodKind` rather than by name.
    pub source_name: String,
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

impl NormalMethod {
    /// The same implementation, bound under a DIFFERENT name.
    ///
    /// PHP `use A { run as go; }`, and anything else that gives one body a
    /// second entry point. A rebound member is its own member — it can carry a
    /// different visibility from the original — so this is a new binding, not
    /// an entry in a list of spellings.
    ///
    /// The point of having it here is that the identity has to move as a UNIT.
    /// `canonical_name` is the vtable key and `source_name` is what the emitter
    /// binds under (`canon(source_name)`), so a producer that sets one and not
    /// the other publishes a member under the very name it was meant to differ
    /// from, and two members end up claiming one key. Stating the rule once
    /// means no caller can get half of it right.
    ///
    /// This is NOT how a cross-language role is expressed. `__toString`,
    /// `__str__` and `to_s` are the same ROLE, marked with
    /// `SpecialMethodKind::ToString` and resolved by that — never by rebinding
    /// them all to one hardcoded spelling.
    pub fn bound_as(&self, name: &str) -> NormalMethod {
        NormalMethod {
            canonical_name: name.to_string(),
            source_name: name.to_string(),
            ..self.clone()
        }
    }
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
    /// A diagnosable error. Java default-method diamonds. Silently picking one
    /// is a bug, not a policy.
    Error,
    /// NEITHER candidate is contributed, and the clash itself is not an error.
    ///
    /// Go promotion at EQUAL depth: `type c struct { a; b }` where both supply
    /// `f` is a legal type — `x.a.f()` and `x.b.f()` both compile. Only an
    /// UNQUALIFIED `x.f()` is illegal, because the spec resolves a selector to
    /// the unique member at the shallowest depth and there is no unique one.
    /// So the name is simply absent from the promoted set, and the diagnosis
    /// belongs to the use site rather than the declaration.
    Ambiguous,
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

/// A language's augmentation RULES, stated once.
///
/// `Augmentation` mixes two things: what the source clause said (which type,
/// which adjustments — per clause) and what the language's mechanism means
/// (mode, who wins, what `super` reaches — the same for every clause in that
/// language). This is the second half, so a normalizer declares its mechanism
/// as one constant instead of restating five fields per `use`/`with`/`include`.
///
/// PHP traits, Dart mixins and Ruby `include` are then three constants, and the
/// difference between them is readable at a glance.
#[derive(Debug, Clone, Copy)]
pub struct AugmentationPolicy {
    pub mode: AugmentationMode,
    pub position: AugmentationPosition,
    pub conflict: AugmentationConflict,
    pub super_target: AugmentationSuper,
    pub contributes: AugmentationContributes,
}

impl AugmentationPolicy {
    /// Combine this language's rules with one source clause.
    pub fn applied_to(&self, decl: &crate::AugmentDecl) -> Augmentation {
        Augmentation {
            from: decl.from.clone(),
            via_field: decl.via_field.clone(),
            mode: self.mode,
            position: self.position,
            conflict: self.conflict,
            super_target: self.super_target,
            adjustments: decl
                .adjustments
                .iter()
                .map(|adj| AugmentationAdjustment {
                    member: adj.member.clone(),
                    rename_to: adj.rename_to.clone(),
                    // The AST records source `Visibility`; the normalized model
                    // speaks `Access`, so the conversion happens once, here,
                    // rather than in each language.
                    visibility: adj.visibility.map(Access::from),
                    exclude: adj.exclude,
                })
                .collect(),
            contributes: self.contributes,
            depth: 0,
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
