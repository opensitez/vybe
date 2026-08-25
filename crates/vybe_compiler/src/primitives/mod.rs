//! Compilation pipeline: AST (from any language walker) → bytecode Chunks.
//!
//! This `mod.rs` owns the `Compiler` struct, its constructor + top-level
//! `compile` entry point, small helpers (emit, scope, canon, var_get/set),
//! and `compile_stmt` / `compile_var_declarator` / `compile_assign_target`.
//! The three largest concerns live in sibling modules:
//!
//!   * `classes`     — function + class + constructor chunk compilation
//!   * `expressions` — `compile_expr` dispatch over every `ExprKind`
//!   * `calls`       — `compile_call` + `compile_lambda`
//!
//! Each sibling is a separate `impl Compiler { ... }` block; the public
//! surface is unchanged (`Compiler::with_profile`, `Compiler::compile`).

// ── Chunk-level emit surface ────────────────────────────────────────────
// The former `emitter` module. Its free functions over `&mut Chunk` and the
// `impl Compiler` walkers alongside them are one layer now — this crate IS
// the emitter, so there is no second module to route through.
pub mod addressable_storage;
pub mod array_transforms;
pub mod atomic_ops;
pub mod base64;
pub mod bigint;
pub mod bits;
pub mod builtin_slots;
pub mod canon_marshal;
pub mod bundle;
pub mod callable;
pub mod clone; // what it means to COPY a value — records, collections, arguments
pub mod codepoints;
pub mod collections;
pub mod complex;
pub mod convert;
pub mod config;
pub mod csv;
pub mod datetime;
pub mod delegates;
pub mod dict;
pub mod dispatch;
pub mod dynamic_symbols;
pub mod enum_lowering;
pub mod errors;
pub mod fs_path;
pub mod functions;
pub mod generators;
pub mod generics;
pub mod canon;
pub mod globals;
pub mod gui;
pub mod heap;
pub mod instructions;
pub mod invoke;
pub mod io;
pub mod json;
pub mod loops;
pub mod math;
pub mod memory;
pub mod multivalue;
pub mod paths;
pub mod object;
pub mod ops;
pub mod packing;
pub mod platforms;
pub mod pointers;
pub mod polyfills;
pub mod prelude; // parse cache + splice for language preludes — one place, not four
pub mod proxy;
pub mod random;
pub mod record_files; // FileDecl / RecordTransfer → wasi:filesystem 0.3.1
pub mod regex;
pub mod sets;
pub mod sorted_collection;
pub mod sprintf;
pub mod string_encoding;
pub mod string_similarity;
pub mod strings;
pub mod target;
pub mod threading;
pub mod tuples;
pub mod type_registry;
pub mod url;
pub mod xml;
pub use polyfills::RuntimeHelpers;
pub use target::Target;
pub use type_registry::CompileTimeTypes;

macro_rules! inst {
    ($self:expr, $($path:ident)::+ $(, $arg:expr)*) => {{
        crate::primitives::instructions::$($path)::+(&mut $self.chunks[$self.current], $self.line $(, $arg)*)
    }};
}

macro_rules! fn_call {
    ($self:expr, $module:literal, $name:literal, $argc:expr) => {{
        crate::primitives::instructions::host::CapabilityContext::get()
            .functions
            .emit(
                &mut $self.chunks[$self.current],
                $module,
                $name,
                $argc,
                $self.line,
            )
    }};
}

mod arrays;
pub mod async_ops;
mod bindings;
mod builtins;
mod calls;
pub mod canonical;
mod case_insensitive_collections;
pub mod channels;
mod class_augmentation;
mod class_context;
pub mod class_normalize; // cross-language class normalisation (was crate::common::classes)
pub mod classes;
pub mod closures;
pub mod components;
// `pub` for `lower_gotos`: `goto`/label lowering is shared machinery every
// language with a goto calls (C, PHP), so it has to be reachable from the
// language crates. It previously lived in the C walker, which is why nothing
// outside this crate needed the module before.
pub mod control_flow;
mod emit_helpers;
pub mod enums;
pub mod events;
pub mod expressions;
pub mod http_cookie;
pub mod http_form;
pub mod http_request_env;
pub mod http_session;
pub mod imports;
mod lambdas;
mod link;
mod metadata;
pub mod namespaces;
mod operators;
mod overloads;
pub mod prototypes;
pub mod records;
pub mod references;
pub mod reflection;
mod resolver;
mod scope;
pub mod slices;
mod statements;
mod type_inference;

use crate::ast::*;
use crate::primitives as common;
#[allow(unused_imports)]
use crate::primitives::instructions as inst;
use crate::primitives::loops::LoopState;
use crate::primitives::scope::Scope;
use crate::profile::*;
use std::collections::{BTreeSet, HashMap, HashSet};
use std::sync::Arc;
use vybe_runtime::chunk::Import as BytecodeImport;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

// ════════════════════════════════════════════════════════════════════════════
// Loop context for break/continue patching
// ════════════════════════════════════════════════════════════════════════════

struct LoopCtx {
    label: Option<String>,
    /// Label stack depth at the BLOCK that wraps this loop (for break).
    /// break = current_label_depth - break_label_depth
    break_label_depth: u32,
    /// Label stack depth at the LOOP (for continue).
    /// continue = current_label_depth - continue_label_depth
    continue_label_depth: u32,
    /// Local slot tracking whether `break` fired in this loop.
    /// Set to true at every `break` site; checked after the loop to
    /// decide whether the Python/Ruby `else` clause runs.
    /// `None` for loops without an `else` clause — no slot allocated.
    did_break_slot: Option<u16>,
    /// JS iterator object to close with `return()` when a `break` exits
    /// a lazy custom for-of loop early.
    iterator_close_slot: Option<u16>,
    /// True for actual loops (while/for/for-in); false for switch blocks
    /// and labeled non-loop blocks. `continue` skips non-continuable contexts.
    is_continuable: bool,
    /// Length of `active_finally_blocks` when this loop was entered. A
    /// `break`/`continue` targeting this loop must execute the finally
    /// blocks pushed *inside* it (those at indices >= this value) before
    /// branching out — ECMA-262 §14.2: abrupt loop completion still runs
    /// pending `finally` bodies.
    finally_depth: usize,
}

// ════════════════════════════════════════════════════════════════════════════
// Pending class bookkeeping
// ════════════════════════════════════════════════════════════════════════════
/// What a class declared about one of its fields: the type SPELLING, plus the
/// answers already resolved from it. One record of one fact — a second map
/// keyed the same way is two places to disagree, and the disagreement is
/// silent.
#[derive(Clone)]
struct FieldType {
    /// The declared type hint, as `Compiler::normalize_type_hint` left it.
    hint: String,
    /// Canonical class name when `hint` names a type that stores BY VALUE,
    /// else `None`. Read by the record deep-copy to recurse into nested
    /// records.
    ///
    /// Carried from `NormalField::value_type`, resolved once in the
    /// declaration pass against `normalized_classes`. Never derived from
    /// `hint` at a read site: that spelling match is exactly what resolution
    /// exists to remove, and a miss is a SILENT shallow copy.
    ///
    /// This merge is only lossless because `resolve_field_value_types`
    /// (`link.rs`) requires a hint before it resolves anything
    /// (`field.type_hint.as_deref()?`), so `value_type: Some` never occurs
    /// with `type_hint: None`. If that ever stops holding, entries land in
    /// neither map and nested records alias again.
    value_type: Option<String>,
}

struct PendingClass {
    parent: Option<String>,
    /// ALL declared direct bases (raw names), for C3 linearization under
    /// `class_multiple_inheritance`. `parent` remains `bases.first()`;
    /// single-inheritance classes have 0 or 1 entry here.
    bases: Vec<String>,
    enclosing_class: Option<String>,
    fields: Vec<String>,
    field_storage_names: HashMap<String, String>,
    is_value_type: bool,
    instance_member_names: Vec<String>,
    instance_pointer_method_names: Vec<String>,
    /// Declared types of instance fields, keyed by canonical field name.
    /// Used when implicit-self resolution turns a bare field name into
    /// `this.<field>` so member access keeps the original receiver type.
    instance_field_types: HashMap<String, FieldType>,
    /// Static field names (declared `static T name`). Looked up from
    /// inside instance methods so a bare `Name` resolves to
    /// `<ClassName>.Name` (struct_get on the class global) rather than
    /// falling through to a non-existent module global.
    static_fields: Vec<String>,
    /// Type hints for static fields, keyed by canonical field name.
    static_field_types: HashMap<String, String>,
    /// Declared static method names on this class. Used for bare
    /// in-class resolution (`Double(x)` inside `Converter`) so the
    /// call resolves through the class object.
    static_method_names: Vec<String>,
    /// Compiled instance-method overloads keyed by canonical source
    /// name. Used by the call compiler to choose the right overload
    /// when the receiver type and argument types are known.
    instance_method_overloads: HashMap<String, Vec<PendingMethodOverload>>,
    /// Compiled static-method overloads keyed by canonical source
    /// name. Kept alongside instance overloads for future shared
    /// overload resolution paths.
    static_method_overloads: HashMap<String, Vec<PendingMethodOverload>>,
    /// Nested type names attached to this class constructor object.
    nested_types: Vec<String>,
    /// Static methods: (name, chunk_idx) — tracked for inheritance
    statics: Vec<(String, usize)>,
}

#[derive(Debug, Clone)]
struct PendingMethodOverload {
    param_types: Vec<String>,
    chunk_idx: usize,
    return_type: Option<String>,
    signature: CallSignature,
    /// Whether this method dispatches on the receiver's RUNTIME type. Resolved
    /// once at registration from the normalized `is_virtual`/`is_override`/
    /// `is_abstract` marks plus `profile.methods_virtual_by_default`, so the
    /// call path never has to re-derive per-language virtuality rules.
    /// `chunk_idx` names the DECLARED type's body, which is the wrong target
    /// for a virtual call — see `resolve_instance_method_overload_chunk`.
    is_virtual: bool,
}

#[derive(Debug, Clone)]
pub(crate) struct CallSignature {
    param_names: Vec<String>,
    /// The declared default expression for each param (positionally aligned
    /// with `param_names`), so named-arg reordering can fill an OMITTED
    /// optional param with its real default instead of `null`.
    param_defaults: Vec<Option<Expression>>,
    min_arity: usize,
    has_rest: bool,
    /// The last param is a `**kwargs`-style collector: named args that match no
    /// declared param are gathered into a dict bound to it (data-driven, so any
    /// language whose frontend emits an `is_kwargs` param opts in).
    has_kwargs: bool,
    /// Position of the variadic (`*args`) collector, wherever it sits. Distinct
    /// from `has_rest` (which is "the LAST param is rest", the shape the runtime
    /// rest-packing handles): with a trailing `**kwargs` the rest is NOT last, so
    /// only the named-arg reorder collects positionals into it.
    rest_index: Option<usize>,
}

impl CallSignature {
    pub(crate) fn from_params(params: &[Param]) -> Self {
        Self {
            param_names: params
                .iter()
                .map(|param| param.name.trim_start_matches('$').to_string())
                .collect(),
            param_defaults: params.iter().map(|param| param.default.clone()).collect(),
            min_arity: params
                .iter()
                .take_while(|param| param.default.is_none() && !param.is_rest && !param.is_kwargs)
                .count(),
            has_rest: params.last().is_some_and(|param| param.is_rest),
            has_kwargs: params.last().is_some_and(|param| param.is_kwargs),
            rest_index: params.iter().position(|param| param.is_rest),
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct AttributeUsageMetadata {
    pub allow_multiple: bool,
    pub inherited: bool,
}

impl Default for AttributeUsageMetadata {
    fn default() -> Self {
        Self {
            allow_multiple: false,
            inherited: true,
        }
    }
}

#[derive(Debug, Clone, Default)]
pub(crate) struct ReflectionParamMetadata {
    pub name: String,
    pub decorators: Vec<Expression>,
    pub type_name: Option<String>,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct ReflectionMethodMetadata {
    pub decorators: Vec<Expression>,
    pub params: Vec<ReflectionParamMetadata>,
    pub is_static: bool,
    pub return_type: Option<String>,
    #[allow(dead_code)]
    pub visibility: Visibility,
    pub is_abstract: bool,
    pub is_virtual: bool,
    pub generic_params: Vec<String>,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct ReflectionMemberMetadata {
    pub decorators: Vec<Expression>,
    pub is_static: bool,
    pub can_write: bool,
    pub type_name: Option<String>,
    pub params: Vec<ReflectionParamMetadata>,
    pub visibility: Visibility,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct ReflectionConstructorMetadata {
    pub param_types: Vec<String>,
    pub params: Vec<ReflectionParamMetadata>,
    pub decorators: Vec<Expression>,
    pub visibility: Visibility,
    pub is_static: bool,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct ReflectionTypeMetadata {
    pub parents: Vec<String>,
    pub decorators: Vec<Expression>,
    pub interfaces: Vec<String>,
    pub nested_types: Vec<String>,
    pub constructors: Vec<ReflectionConstructorMetadata>,
    pub is_value_type: bool,
    pub is_sealed: bool,
    pub methods: HashMap<String, ReflectionMethodMetadata>,
    pub properties: HashMap<String, ReflectionMemberMetadata>,
    pub fields: HashMap<String, ReflectionMemberMetadata>,
    /// The type's declared generic parameters, as the generics primitive's own
    /// `GenericParam` — parsed with `generics::parse_generic_params_hint`, not
    /// a local spelling.
    ///
    /// Members store their declared type as a string, so a field of type `T`
    /// records `"T"`. Without the parameter list there is nothing to bind that
    /// against, so `FieldType.Name` answered `T` instead of the type argument.
    /// The metadata KEY stays the erased name — the declaration lives on the
    /// open type — and this is the other half `GenericSignature::bind_args`
    /// needs to build a `GenericContext` for a closed use.
    pub generic_params: Vec<vybe_ast::GenericParam>,
}

#[derive(Debug, Clone)]
pub(crate) enum ReflectionBinding {
    Type(String),
    Constructor {
        type_name: String,
        param_types: Vec<String>,
    },
    Method {
        type_name: String,
        method_name: String,
        generic_args: Vec<String>,
    },
    Property {
        type_name: String,
        property_name: String,
    },
    Field {
        type_name: String,
        field_name: String,
    },
    Assembly,
    AssemblyName,
    Parameter {
        type_name: String,
        method_name: String,
        index: usize,
    },
}

#[derive(Debug, Clone)]
struct StaticLocalBinding {
    global_name: String,
    init_flag_name: String,
    type_hint: Option<String>,
}

#[derive(Debug, Clone)]
struct PascalArrayDimensionMetadata {
    first_index: i64,
    length: usize,
    uses_char_ordinal: bool,
}

#[derive(Debug, Clone)]
struct PascalArrayBoundsMetadata {
    is_fixed: bool,
    dimensions: Vec<PascalArrayDimensionMetadata>,
}

#[derive(Debug, Clone)]
struct ArrayBindingMetadata {
    is_fixed: bool,
    type_hint: Option<String>,
    pascal_bounds: Option<PascalArrayBoundsMetadata>,
}

#[derive(Debug, Clone)]
struct FortranInterfaceOverload {
    target_name: String,
    min_arity: usize,
    param_types: Vec<Option<String>>,
}

#[derive(Debug, Clone)]
struct JsArgumentsBinding {
    args_slot: u16,
    aliased_params: HashMap<String, (u16, usize)>,
    aliased_indices: HashMap<usize, u16>,
}

// ════════════════════════════════════════════════════════════════════════════
// Compiler
pub struct Compiler {
    pub(crate) chunks: Vec<Chunk>,
    scopes: Vec<Scope>,
    pub(crate) current: usize,
    /// Chunk index of the most recently compiled lambda BODY (not its factory).
    /// The only handle a caller has on an anonymous lambda's chunk.
    pub(crate) last_lambda_body_chunk: Option<usize>,
    loops: Vec<LoopCtx>,
    loop_states: Vec<LoopState>,
    label_depth: u32,
    function_label_base: u32,
    pub(crate) line: u32,
    pub(crate) defined_globals: HashSet<String>,
    /// Module-level VARIABLE names, as written. Separate from
    /// `defined_globals` (classes/functions/modules) because a language whose
    /// script-scope variables are implicit — PHP's `$x = 5` — records them
    /// nowhere else: they emit as VM globals with no compile-time trace.
    /// Read by `primitives/globals.rs` to answer "what does `$GLOBALS`
    /// contain"; nothing else consults it.
    pub(crate) module_variable_names: HashSet<String>,
    /// Serial for the default NAME every created control gets. A control has a
    /// name whether the program assigns one or not — see
    /// `gui::emit_control_type_stamp`.
    pub(crate) gui_auto_name_counter: u32,
    const_globals: HashSet<String>,
    /// Compile-time integer values of immutable globals whose initializer is a
    /// constant expression (WASM). Feeds extended-const evaluation of data/elem
    /// segment offsets (`(offset (i32.add (global.get $g) (i32.const N)))`).
    pub(crate) global_const_values: std::collections::HashMap<String, i64>,
    in_strict: bool,
    /// Declared policy in force, innermost last — never empty. Frame 0 is the
    /// module's own declaration; a block pushes a frame only when its body
    /// actually states something, so ordinary code pays nothing.
    ///
    /// A `DirectiveScope::Module` statement writes through EVERY frame, which
    /// is what makes Pascal's `{$R+}` outlive the procedure it appeared in
    /// while JS's `"use strict"` does not. See `vybe_ast::Directives`.
    directives: Vec<vybe_ast::Directives>,
    /// The module's canon section, lowered from `Module::canon` at the start of
    /// the compile and published to every chunk at emission. Empty for every
    /// language that is not a Component Model front end.
    canon_section: Vec<vybe_runtime::canon_def::CanonDef>,
    /// The DECLARED rows, kept alongside the lowered ones because a
    /// `CoreExport` callee can only be resolved once every chunk exists.
    canon_decls: Vec<vybe_ast::canon::CanonDecl>,
    canon_functypes: Vec<Option<vybe_runtime::canon_def::CanonFuncType>>,
    canon_valtypes: Vec<Option<vybe_runtime::component::ValType>>,
    component_funcs: Vec<Option<u32>>,
    /// True while compiling the operand of a `typeof`. `typeof undeclaredName`
    /// must evaluate to `"undefined"`, never throw — so the unresolvable-binding
    /// ReferenceError in `emit_var_get` is suppressed in this context.
    in_typeof_operand: bool,
    /// A CONDITION is being compiled and would rather have a raw i32 than a
    /// boxed `Bool`. Set only by `compile_condition_to_i32`, and **taken** at
    /// the top of `compile_expr` — so it reaches the outermost operator and
    /// can never leak into a nested operand, where the boxed value is what the
    /// surrounding expression is entitled to.
    want_i32_condition: bool,
    /// The request above was honoured: the value on the stack is an i32.
    /// Written ONLY by `emit_i32_to_bool_or_report`, where reporting and
    /// skipping the boxing are the same statement — so "the boxing was
    /// skipped" and "the stack holds an i32" cannot disagree. The dangerous
    /// direction is a report without a skip: `BR_IF` accepts a `Bool` happily,
    /// so the loop would key on the wrong thing in silence.
    gave_i32_condition: bool,
    /// Every name the program lexically declares (`let`/`const`/`var`/params/
    /// etc.) as a local, across all scopes — populated only for languages with
    /// `unresolved_reference_throws`. A name in this set that is unresolvable in
    /// the current scope is provably an out-of-scope user binding (never an
    /// untracked host global), so reading it is a ReferenceError even in sloppy
    /// mode (§9.1.1.4.6 applies in both strict and sloppy). See `emit_var_get`.
    program_lexical_names: HashSet<String>,

    /// Fx-hashed, not SipHash: `resolve_namespaced_function_identity` probes
    /// this set once or more per identifier compiled, and a warm-job profile
    /// put std's default hasher among the largest single costs. The keys are
    /// function names from the program being compiled, in the compiler's own
    /// table — SipHash's resistance to attacker-chosen keys buys nothing here.
    /// See [`vybe_runtime::chunk::FxBuildHasher`].
    pub(crate) defined_functions: HashSet<String, vybe_runtime::chunk::FxBuildHasher>,
    function_param_modes: HashMap<String, Vec<PassBy>>,
    function_param_types: HashMap<String, Vec<Option<String>>>,
    function_min_arity: HashMap<String, usize>,
    function_signatures: HashMap<String, Vec<CallSignature>>,
    rest_fixed_arities: BTreeSet<u8>,
    function_return_types: HashMap<String, String>,
    fortran_interface_overloads: HashMap<String, Vec<FortranInterfaceOverload>>,
    fortran_operator_overloads: HashMap<String, Vec<FortranInterfaceOverload>>,
    constructor_signatures: HashMap<String, Vec<CallSignature>>,
    pub(crate) defined_classes: HashSet<String>,
    pub(crate) abstract_classes: HashSet<String>,
    /// Names of methods defined on any user class — used to avoid value method
    /// hijacking (e.g. user class `Calc.Add()` shouldn't match array `add`).
    defined_class_methods: HashSet<String>,
    /// Classes whose STATIC methods take the called class as a receiver —
    /// `NormalClass::late_static_binding`, recorded when the class is compiled
    /// so a CALL SITE can read the same declaration the method chunk's arity
    /// was built from.
    ///
    /// Both ends used to ask `profile.name == "php"` independently. That is one
    /// fact answered twice from a proxy; a call site that disagreed with the
    /// declaration would push a receiver the callee never bound.
    pub(crate) classes_with_late_static_binding: HashSet<String>,
    /// Classes declaring an index operator (`operator []` / `__getitem__`).
    /// Indexing one of these is a method call, not a key lookup — resolved
    /// from the receiver's static type so arrays, dicts and strings keep the
    /// plain index path with no runtime probe.
    pub(crate) classes_with_indexer: HashSet<String>,
    /// Classes declaring an index *setter* (`operator []=` / `__setitem__`).
    /// Kept apart from `classes_with_indexer` — a class may define either
    /// half on its own.
    pub(crate) classes_with_index_setter: HashSet<String>,
    /// Any class in this program binds the `GetAttr` role (Python
    /// `__getattr__`, PHP `__get`, JS Proxy get) — the attribute-miss
    /// interceptor.
    ///
    /// A program-level flag, not a per-class set like `classes_with_indexer`,
    /// because the receiver's type is usually unknown exactly where this
    /// matters: `f = FlexObj(); f.undefined` has no static hint in a
    /// dynamically-typed language, so a hint-keyed gate would miss the case it
    /// exists for. Programs that bind the role nowhere pay nothing.
    pub(crate) program_has_getattr: bool,
    /// The WRITE half of the same role: any class binds `ProtocolSlot::SetAttr`
    /// (PHP `__set`, Python `__setattr__`).
    ///
    /// `SetAttr` was in the slot vocabulary with NO reader anywhere — the
    /// frontends declared it and nothing consumed it, so a language wanting a
    /// catch-all property write had to synthesise a direct `__set` member call
    /// in its walker instead. Same program-level gating as its read twin, for
    /// the same reason: the receiver's type is unknown exactly where this
    /// matters, and programs binding the role nowhere pay nothing.
    pub(crate) program_has_setattr: bool,
    /// Any class in this program exposes its index role as an ACCESSOR pair
    /// (`__get___index__` / `__set___index__`) rather than as a slot-bound
    /// method. A program-level flag for the same reason as
    /// `program_has_getattr`: the receiver's type is unknown exactly where it
    /// matters — an interface-typed reference, or a local whose initializer's
    /// type never reached its binding.
    ///
    /// Deliberately narrower than `classes_with_indexer`. That set answers
    /// "does some type fill `GetItem`", which is true wherever a Python
    /// `__getitem__` or a Dart `operator []` exists — and a runtime probe on
    /// THEIR receivers takes the emit away from what those languages declared
    /// for `get_item`, so an ordinary `d[k]` on a dict stops maintaining what
    /// its own path maintains. This flag answers the narrower question the
    /// probe below actually asks, so producer and consumer agree by
    /// construction.
    pub(crate) program_has_index_accessor: bool,
    /// The declared type of each global, WITH its `TypeBinding`.
    ///
    /// Was a bare `String`, which meant a global's declaration could say what
    /// it was but not whether that was enforced — and a top-level `var i:
    /// Integer` is how most Pascal is written, so the whole language's globals
    /// were unusable to any caller that needs the guarantee.
    global_type_hints: HashMap<String, vybe_ast::TypeHint>,
    /// Map from member name → containing namespace name.
    /// Used for bare-name resolution within modules/namespaces/enums.
    /// E.g. `Main` inside `Module Program` resolves to `Program.Main`.
    /// `Green` inside `enum TColor` resolves to `TColor.Green`.
    /// Models the WASM Component Model's namespace-scoped imports.
    enum_members: HashMap<String, String>,
    /// Every module that contributes a given BARE member name, in declaration
    /// order.
    ///
    /// [`Self::enum_members`] answers "which owner does this bare name belong
    /// to" with ONE owner, so a second contributor silently displaces the
    /// first — the name resolves to whichever module was declared last. That
    /// is not what the languages say. A module contributes its members to the
    /// enclosing scope, and when two modules contribute the SAME name the
    /// unqualified reference is an ERROR, not a silent pick:
    ///
    /// | language | diagnostic |
    /// |---|---|
    /// | VB.NET | BC30562 `'X' is ambiguous between declarations in Modules 'A, B'` |
    /// | C# | CS0104 `'X' is an ambiguous reference` |
    ///
    /// Keeping every contributor rather than the last is what makes that
    /// question answerable at all: a lookup returning `Option` cannot say "two
    /// answers", so the ambiguity had nowhere to be represented and the
    /// overwrite was the only reachable outcome.
    ///
    /// Only names a module actually PUBLISHES are recorded — a module-private
    /// member is not contributed and so cannot collide (see the visibility
    /// rule where module members are compiled). Qualified access is unaffected:
    /// `A.X` and `B.X` name different members and stay legal, which is why the
    /// error belongs at the REFERENCE and not at the declaration.
    module_member_contributors: HashMap<String, Vec<String>>,
    /// Reverse enum lookup: enum type -> underlying integer -> member name.
    enum_value_names: HashMap<String, HashMap<i64, String>>,
    enum_flags: HashSet<String>,
    pub(crate) reflection_types: HashMap<String, ReflectionTypeMetadata>,
    pub(crate) attribute_usage: HashMap<String, AttributeUsageMetadata>,
    pub(crate) reflection_bindings: HashMap<String, ReflectionBinding>,
    case_sensitive: bool,
    /// Resolved ONCE at construction, like `case_sensitive` — `registry::hooks`
    /// takes a mutex and linear-scans, and `canon` is called on nearly every
    /// name the compiler touches.
    variable_namespace: Option<&'static vybe_runtime::registry::VariableNamespace>,
    pub(crate) profile: LanguageProfile,
    current_func_name: Option<String>,
    current_result_slot: Option<u16>,
    current_ref_out_params: Option<Vec<u16>>,
    pending_classes: HashMap<String, PendingClass>,
    /// Every class NORMALIZED ONCE, during the declaration pass, keyed by
    /// canonical name. This is the single class model: computed before any
    /// body compiles, so a class's member set is knowable regardless of
    /// compilation order, and augmenting types (traits / mixins / promoted
    /// fields) can be resolved by name. See flexclassplan.md §3a, §4c.
    normalized_classes: HashMap<String, vybe_ast::class_normalize::NormalClass>,
    /// For the class being emitted: bound method name → the PROTOCOL SLOT key
    /// that method also publishes under (`__vybe_slot_<id>`).
    ///
    /// Built in `compile_class`, where the `NormalMethod` and the class's
    /// `special_methods` are both still in hand. By the time the bind sites run
    /// a method is only a `(name, chunk_idx)` pair, so the role has to be
    /// carried here or it is lost — which is exactly how the compiler ended up
    /// re-deriving roles from spellings. See flexclassplan.md §2g.
    current_class_slot_keys: HashMap<String, String>,
    current_class: Option<String>,
    current_namespace: Option<String>,
    /// Mirrors `NormalClass.implicit_self_fields` for the class the
    /// compiler is currently inside. Saved/restored by `compile_class`
    /// alongside `current_class`. Expression + call-site resolution
    /// consults this instead of `profile.implicit_self_fields`, so the
    /// walker stays the single source of truth for per-language class
    /// semantics.
    pub(crate) current_class_implicit_self: bool,
    pub(crate) current_member_is_static: bool,
    static_local_bindings: Vec<HashMap<String, StaticLocalBinding>>,
    array_bindings: HashMap<String, ArrayBindingMetadata>,
    /// Label for the next loop to be pushed (set by StmtKind::Labeled).
    pending_label: Option<String>,
    /// The BINDING NAME of each open `with` target, innermost last.
    ///
    /// A name, not a slot, because a bare name inside the block is compiled as
    /// an ordinary member access on that binding — see `emit_with_target_get`.
    /// A slot could only be read with `LOCAL_GET`, which is what forced the old
    /// raw `STRUCT_SET` and made every declared property inside a `with` block
    /// write a field nobody reads.
    with_targets: Vec<String>,
    capture_by_value_vars: Vec<String>,
    capture_locals: HashMap<u8, u16>,
    closure_env_names: Vec<String>,
    /// Shared env array for the current outer function: holds locals that
    /// are captured by inner closures. All reads/writes of these locals
    /// go through array.get/array.set so mutations are visible to closures.
    shared_env_slot: Option<u16>,
    shared_env_names: Vec<String>,
    /// Names promoted to a module-level pointer cell. A GLOBAL being a cell is
    /// a module-wide fact; a local or parameter holding a reference is a
    /// per-binding one and lives on `scope::Local::holds_reference`. One map
    /// used to answer both — see `binding_uses_pointer_cell`.
    promoted_global_cells: HashSet<String>,
    /// Names taken by address ANYWHERE in the module, collected before any body
    /// is compiled. A "readers must deref" hint, deliberately separate from
    /// `promoted_global_cells`, which records that a wrap actually happened —
    /// mixing them makes promotion see itself as already done and skip.
    module_addr_taken_globals: HashSet<String>,
    /// Module-level names used as an ATOMIC place anywhere in the module —
    /// promoted to a shared-memory word at their top-level declaration.
    module_atomic_word_globals: HashSet<String>,

    /// Names of locals/params in the function currently being compiled whose
    /// address is taken somewhere in the body (`&v`). Populated by a pre-scan
    /// at function entry. Such bindings are promoted to a pointer cell *once*
    /// at their declaration (and params at entry) rather than lazily at the
    /// first `&v` use — taking the address inside a loop would otherwise
    /// re-wrap the cell every iteration and orphan prior mutations.
    current_addr_taken_locals: HashSet<String>,
    /// Locals of the CURRENT function used as an atomic place — promoted to a
    /// shared word at declaration, exactly where the pointer-cell promotion
    /// would have run.
    current_atomic_word_locals: HashSet<String>,
    current_closure_captured_locals: HashSet<String>,

    /// Functions whose every explicit `Return` carries an `ExprKind::Tuple`
    /// of the same arity. Populated by a pre-pass before any function is
    /// compiled so both callee (set `chunk.result_arity`, push N values
    /// without packing) and caller (destructure directly off the stack)
    /// can agree on the multi-value ABI at emit time.
    multi_return_functions: HashMap<String, u8>,
    /// Synchronous functions compiled with `chunk.is_generator = true`
    /// — tracked by canonical name so `for v in gen()` call-site
    /// emission knows to use the stack-switching iterator protocol
    /// rather than the array-index protocol. Async generators are
    /// deliberately excluded; they route through the async iterator
    /// path so await points stay JSPI-compliant.
    generator_functions: HashSet<String>,
    /// Source-level (is_async, is_generator) per method chunk index. Walker
    /// lowering (wrap_generator) can leave the CHUNK flags false while the
    /// source method was a generator — prototype-kind stamping at the class
    /// attach sites reads this instead (§27.3/§27.4/§27.7 intrinsic stamps).
    pub(crate) method_fn_kinds: HashMap<usize, (bool, bool)>,
    /// §9.1.1.3.4 (JS only): set to `(chunk_idx, this_slot)` while the body
    /// of a DERIVED class constructor is being compiled. `this` reads and
    /// `super()` calls in that chunk emit TDZ guards against this_slot
    /// (null until super() initializes it). Chunk index is checked so
    /// nested functions/classes compiled mid-body don't inherit the guard.
    pub(crate) js_derived_ctor_ctx: Option<(usize, u16)>,
    /// Number of user-visible parameters (excluding the hidden control
    /// slot) for each synchronous generator function. Used at call
    /// sites to pad missing optional args with `undefined` so the
    /// null resume value in GEN_NEXT never lands in an optional
    /// parameter slot and prevents default-parameter application.
    generator_param_counts: HashMap<String, usize>,
    /// ESM host-module import bindings: canon(local) → (module, func).
    /// Populated from user `import { X } from "wasi:foo"` statements.
    /// Calls and reads both go through the installed runtime binding so
    /// function exports stay callable while value exports preserve
    /// normal JS non-callable semantics.
    host_import_bindings: HashMap<String, (String, String)>,
    /// ESM named imports that resolve to `ExportEntry::Value` — constant
    /// values provided by the host at registration time (e.g. `ecma:math::PI`).
    /// canon(local_name) → Value. These are inlined as constants at the
    /// use-site rather than routed through `CALL_IMPORT`, which only handles
    /// callable function exports.
    host_const_bindings: HashMap<String, vybe_runtime::Value>,
    /// Named imports resolved through the common namespace tree rather than an
    /// ESM host module. This covers language/module surfaces registered as
    /// namespace data (`from math import sqrt`, `use function Lib\fn`) whose
    /// leaves may be `CommonEmit`, `HostCall`, or `Const`.
    namespace_import_bindings: HashMap<String, crate::primitives::namespaces::ResolutionTarget>,
    /// ESM wildcard namespace aliases: canon(alias) → module specifier.
    /// `import * as cli from "wasi:cli"` records `cli` → `"wasi:cli"`.
    /// Bare-value access and calls both route through the runtime Module
    /// Namespace object built by `host_imports::install`.
    host_namespace_aliases: HashMap<String, String>,
    /// Component-Model package roots: canon(prefix) → module_root.
    /// Populated by the Linker from profile `PackageRoot` defaults
    /// (e.g. `{"vybe": "vybe:", "wasi": "wasi:", "wasm": "wasm:"}`).
    /// Phase 3 will wire `calls.rs`'s qualified-chain path to consume
    /// this map instead of `profile.namespaces.host_packages`.
    host_package_roots: HashMap<String, String>,
    /// Namespace-tree mounts: canon(prefix) → tree path. Populated by the
    /// Linker from profile `TreeMount` defaults (VB/C# `system` →
    /// `dotnet.system`); the resolver rebases a matching qualified chain
    /// onto the tree path before walking the global namespace tree.
    tree_mounts: HashMap<String, String>,
    /// Ambient namespace-tree roots: tree paths bare qualified chains
    /// additionally search under — .NET `Imports`/`using` context as data.
    /// Profile `TreeAmbient` defaults + user import statements (rebased
    /// through `tree_mounts` at link time). Order matters: first hit wins.
    ambient_tree_roots: Vec<String>,
    /// Source-language type aliases: canon(alias) -> target type path.
    /// Shared across languages so `Imports X = System.Text.StringBuilder`
    /// and `using X = System.Text.StringBuilder` normalize below the walker.
    source_type_aliases: HashMap<String, String>,
    /// Source-language namespace imports (`Imports Demo.Core`,
    /// `using Demo.Core`) that expose declarations under that namespace to
    /// unqualified source lookup.
    source_namespace_imports: Vec<String>,
    /// The `user.<unit>.*` root of namespaceplan.md, held per COMPILATION UNIT
    /// instead of in `vybe_runtime::namespaces`.
    ///
    /// User declarations are a *mount*, not tree data. The global registry is
    /// process-global, merges on registration and never unregisters, while
    /// `--serve` compiles many units in one process — registering
    /// `myapp.models.customer` there would leave it resolvable while compiling
    /// the NEXT program, and silently, as a stale resolution rather than an
    /// error. The plan's own gotcha: "mounts are per-VM; tree registration is
    /// process-global data (immutable after startup) — keep them distinct."
    ///
    /// It is walked by the SAME `resolve_segments` as every platform root, so
    /// this is one more root in one resolver, not a second resolver.
    ///
    /// Written at the point of DECLARATION by
    /// `declare_user_namespace_member`, not derived from `defined_classes`
    /// afterwards. That ordering is forced: predeclaration resolves while it
    /// declares — normalizing a class asks for its parent's identity — so a
    /// root built only after that pass answers nothing for the pass's own
    /// queries. Being the storage rather than a projection also removes the
    /// staleness question a derived tree could only ever guess at.
    ///
    /// `RefCell` because resolution takes `&self` and must be able to read it
    /// from the middle of a query.
    user_namespace_tree: std::cell::RefCell<crate::primitives::namespaces::Subtree>,
    /// Snapshot of the current module's source imports.
    ///
    /// Used for narrow source-shape decisions that depend on the ambient
    /// framework surface, such as WinForms form inference for bare VB/C#
    /// classes inside a module that explicitly imports System.Windows.Forms.
    current_module_imports: Vec<Import>,
    /// Activation set for gated namespace roots, built from the module's
    /// imports at compile start. `None` until then; only consulted when the
    /// profile declares `gated_namespace_roots`.
    active_namespaces: Option<std::collections::HashSet<String>>,
    /// JS-only: set when the module references `new Proxy(...)`. Member /
    /// Index reads + writes route through `emitter::js::proxy_adapter`
    /// for runtime trap dispatch. Off → direct `STRUCT_GET` / `ARRAY_GET`
    /// (zero overhead for non-Proxy code paths).
    pub(crate) uses_proxy: bool,
    /// BigInt values are POSSIBLE in this compile — derived at
    /// `compile_with_imports` from the three declarations that can produce
    /// one: a `[builtin_types] bigint` spelling (Kotlin `Long`), a builtin
    /// whose emit target reaches the `ecma:bigint` host (JS's `BigInt`,
    /// Java's BigInteger surface), or a `Literal::BigInt` anywhere in the
    /// module (PHP's big literals, JS `1n`). Replaces the `has_ecma_bigint`
    /// profile bool: off → the `++` path emits no runtime type test and the
    /// bigint routing arms are skipped (they would be unreachable anyway).
    pub(crate) bigint_enabled: bool,
    /// Read-only snapshot of `vm.modules` keyed by specifier. Lets the
    /// Linker resolve `import { X } from "node:http"` against Adapter
    /// modules (Phase 6) — walking the `Indirect` re-export chain to
    /// the ultimate Synthetic export so `X` binds directly to that
    /// `(module, func)` pair, same as a direct host import.
    ///
    /// Stored by specifier → per-module name → (final module, final
    /// name). Empty map when the caller didn't supply one (Bundle's
    /// legacy compile path, tests that don't use adapters).
    module_exports: HashMap<String, HashMap<String, (String, String)>>,
    /// Constant-value snapshot of host module exports — `ExportEntry::Value`
    /// entries from `flatten_module_value_exports`. Populated at compile time
    /// and used during `collect_host_imports` to route value exports into
    /// `host_const_bindings` instead of `host_import_bindings`.
    module_value_exports: HashMap<String, HashMap<String, vybe_runtime::Value>>,
    /// Active finally blocks for the current control-flow path.
    ///
    /// Used to make early returns execute structured `finally` bodies
    /// even though the VM's TRY_START handler currently ignores the
    /// reserved finally offset operand.
    active_finally_blocks: Vec<FinallyAction>,
    /// Indices into `active_finally_blocks` whose try's runtime handler has
    /// FIRED (we are compiling that try's catch-arms section). A `throw`
    /// inside a catch arm inlines exactly these — the runtime can no longer
    /// run them — and leaves live-handler finallys to the runtime.
    fired_finally_indices: Vec<usize>,
    /// Stack of enclosing `try`-with-`finally` join points. A `break` /
    /// `continue` / `return` inside a protected body cannot run its `finally`
    /// under the `try_table` handler (a throwing finally would be self-caught
    /// — non-spec). Instead it sets the top join's completion code and `br`s
    /// to the join, which runs `finally` OUTSIDE the handler, then dispatches
    /// the pending exit onward. This is the whole reason `finally` needs no VM
    /// opcode: it is lowered here into standard wasm. Innermost is last.
    finally_joins: Vec<FinallyJoin>,
    /// Nesting depth of the catch body currently being compiled.
    catch_depth: usize,
    /// JS async-function wrapper try depth currently active for the
    /// function body being compiled. Explicit returns inside that body
    /// must emit matching TRY_END opcodes before RETURN so the VM does
    /// not retain stale handlers from the callee frame.
    active_async_try_depth: usize,
    js_arguments_bindings: Vec<Option<JsArgumentsBinding>>,
}

/// §16.2.1.3 wildcard — `import * as alias from "module"`.
#[derive(Debug, Clone)]
pub struct HostWildcardImport {
    pub alias: String,
    pub module: String,
}

#[derive(Debug, Clone)]
enum FinallyAction {
    Statements(Vec<Statement>),
    ResourceDispose {
        slot: u16,
        method: String,
        line: u32,
    },
}

/// Completion codes stored in a join's completion local. Written before the
/// `br` to the join; read by the dispatch emitted after the `finally` body.
/// `NORMAL` is the zero value (fall through after `finally`).
mod completion {
    pub const NORMAL: f64 = 0.0;
    pub const BREAK: f64 = 1.0;
    pub const CONTINUE: f64 = 2.0;
    pub const RETURN: f64 = 3.0;
}

/// A `try`-with-`finally` join point (see [`Compiler::finally_joins`]).
struct FinallyJoin {
    /// `self.label_depth` captured where the try's wrapping block is the
    /// innermost label, so `label_depth - join_label_depth` is the `br` depth
    /// to the join from any point inside the protected body.
    join_label_depth: u32,
    /// Local holding the completion code (see [`completion`]).
    completion_slot: u16,
    /// Local holding the pending `return` value while `finally` runs.
    ret_slot: u16,
}

/// §16.2.1 named — `import { name as local } from "module"`.
#[derive(Debug, Clone)]
pub struct HostImportNamed {
    pub local: String,
    pub module: String,
    pub func: String,
}

/// The number-path opcode for `emit_js_dynamic_arith` (the non-BigInt
/// branch of a dynamically-dispatched `-`/`*`/`/`/`%`).
#[derive(Clone, Copy)]
enum NumberArith {
    Sub,
    Mul,
    Div,
    Mod,
}

/// Map a compound-assignment operator to its plain binary operator, for
/// desugaring `t OP= v` → `t = t OP v`. Returns `None` for the logical /
/// null-coalescing / `+` forms, which have their own short-circuit or
/// string-concat handling and are left on the direct compound path.
/// JS builtin constructors that the host exposes a canonical `__ctor_<Name>`
/// anchor for (see `ecma_globals::register`). Bare reads of these as VALUES
/// resolve through the anchor so `constructor`/`prototype` identity survives
/// the user-facing global being re-bound by later compile/link passes.
fn is_js_builtin_ctor_value(name: &str) -> bool {
    matches!(
        name,
        "Object"
            | "Array"
            | "Function"
            | "Number"
            | "String"
            | "Boolean"
            | "Symbol"
            | "BigInt"
            | "Date"
            | "RegExp"
    )
}

fn compound_op_to_binop(op: &CompoundOp) -> Option<BinOp> {
    Some(match op {
        CompoundOp::Sub => BinOp::Sub,
        CompoundOp::Mul => BinOp::Mul,
        CompoundOp::Div => BinOp::Div,
        CompoundOp::IDiv => BinOp::IDiv,
        CompoundOp::Mod => BinOp::Mod,
        CompoundOp::Pow => BinOp::Pow,
        CompoundOp::BitAnd => BinOp::BitAnd,
        CompoundOp::BitOr => BinOp::BitOr,
        CompoundOp::BitXor => BinOp::BitXor,
        CompoundOp::Shl => BinOp::Shl,
        CompoundOp::Shr => BinOp::Shr,
        CompoundOp::UShr => BinOp::UShr,
        _ => return None,
    })
}

/// AST scan: returns true if the statement (or anything nested within
/// it) constructs a `Proxy` (i.e. contains `new Proxy(...)`). Used to
/// gate the Member / Index proxy dispatcher emit so non-Proxy code
/// keeps the zero-overhead direct-opcode path.
/// Collect the names of variables whose address is taken (`&name`) anywhere in
/// `stmts`. Does NOT descend into nested function/lambda/class bodies — those
/// open their own scopes and are pre-scanned when compiled. Used to promote
/// address-taken bindings to a pointer cell at declaration time (once), instead
/// of re-wrapping at every `&name` use (which corrupts the cell inside loops).
fn collect_addr_taken_idents(stmts: &[Statement], out: &mut HashSet<String>) {
    for stmt in stmts {
        collect_addr_taken_in_stmt(stmt, out);
    }
}

fn collect_addr_taken_in_stmt(stmt: &Statement, out: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::Expr(e) => collect_addr_taken_in_expr(e, out),
        StmtKind::Return(opt) => {
            if let Some(e) = opt {
                collect_addr_taken_in_expr(e, out);
            }
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(e) = expr {
                collect_addr_taken_in_expr(e, out);
            }
            if let Some(e) = cause {
                collect_addr_taken_in_expr(e, out);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations {
                if let Some(e) = &d.init {
                    collect_addr_taken_in_expr(e, out);
                }
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            for t in targets {
                collect_addr_taken_in_expr(t, out);
            }
            collect_addr_taken_in_expr(value, out);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            collect_addr_taken_in_expr(target, out);
            collect_addr_taken_in_expr(value, out);
        }
        // A namespace is not a scope for VARIABLES — php's `$n` under
        // `namespace App;` is the same module global as at top level — so the
        // scan must descend. Skipping it left every statement in a namespaced
        // file invisible here, and the reader compiled before the promotion
        // (the whole reason this pre-pass exists) went back to reading the cell
        // raw.
        StmtKind::Block(stmts) | StmtKind::NamespaceDecl { body: stmts, .. } => {
            collect_addr_taken_idents(stmts, out)
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            collect_addr_taken_in_expr(cond, out);
            collect_addr_taken_idents(then_body, out);
            for (c, b) in elifs {
                collect_addr_taken_in_expr(c, out);
                collect_addr_taken_idents(b, out);
            }
            if let Some(b) = else_body {
                collect_addr_taken_idents(b, out);
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            collect_addr_taken_in_expr(cond, out);
            collect_addr_taken_idents(body, out);
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
            ..
        } => {
            if let Some(i) = init {
                collect_addr_taken_in_stmt(i, out);
            }
            if let Some(c) = cond {
                collect_addr_taken_in_expr(c, out);
            }
            if let Some(u) = update {
                collect_addr_taken_in_expr(u, out);
            }
            collect_addr_taken_idents(body, out);
        }
        StmtKind::ForIn { iter, body, .. } => {
            collect_addr_taken_in_expr(iter, out);
            collect_addr_taken_idents(body, out);
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            collect_addr_taken_idents(body, out);
            for c in catches {
                collect_addr_taken_idents(&c.body, out);
            }
            if let Some(b) = else_body {
                collect_addr_taken_idents(b, out);
            }
            if let Some(b) = finally {
                collect_addr_taken_idents(b, out);
            }
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            collect_addr_taken_in_expr(expr, out);
            for case in cases {
                collect_addr_taken_idents(&case.body, out);
            }
            if let Some(b) = default {
                collect_addr_taken_idents(b, out);
            }
        }
        StmtKind::Labeled { body, .. } => collect_addr_taken_in_stmt(body, out),
        // Nested function/class declarations open their own scope — skip.
        _ => {}
    }
}

fn collect_addr_taken_in_expr(expr: &Expression, out: &mut HashSet<String>) {
    match &expr.kind {
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr: inner,
        } => {
            if let ExprKind::Ident(name) = &inner.kind {
                out.insert(name.clone());
            }
            collect_addr_taken_in_expr(inner, out);
        }
        ExprKind::Unary { expr, .. } => collect_addr_taken_in_expr(expr, out),
        ExprKind::Binary { left, right, .. } => {
            collect_addr_taken_in_expr(left, out);
            collect_addr_taken_in_expr(right, out);
        }
        ExprKind::Call { callee, args, .. } => {
            collect_addr_taken_in_expr(callee, out);
            for a in args {
                // Address-taking is not always SPELLED at the call site: php
                // writes `w($g)` and the `&` lives on the callee's PARAMETER,
                // so no `AddrOf` node reaches this scan even though `$g`'s
                // storage is aliased into the call. The walker resolves the
                // signature and declares the fact on the argument, which makes
                // it the same fact as `&$g` — so read it the same way.
                if a.by_ref {
                    if let ExprKind::Ident(name) = &a.value.kind {
                        out.insert(name.clone());
                    }
                }
                collect_addr_taken_in_expr(&a.value, out);
            }
        }
        ExprKind::Member { object, .. } => collect_addr_taken_in_expr(object, out),
        ExprKind::Index { object, index, .. } => {
            collect_addr_taken_in_expr(object, out);
            collect_addr_taken_in_expr(index, out);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            collect_addr_taken_in_expr(cond, out);
            collect_addr_taken_in_expr(then, out);
            collect_addr_taken_in_expr(else_, out);
        }
        ExprKind::Assign { target, value } => {
            collect_addr_taken_in_expr(target, out);
            collect_addr_taken_in_expr(value, out);
        }
        ExprKind::Array(elems) => {
            for e in elems {
                collect_addr_taken_in_expr(&e.value, out);
            }
        }
        ExprKind::Object(props) => {
            for p in props {
                match p {
                    ObjectProperty::KeyValue { value, .. } => {
                        collect_addr_taken_in_expr(value, out)
                    }
                    ObjectProperty::Computed { key, value } => {
                        collect_addr_taken_in_expr(key, out);
                        collect_addr_taken_in_expr(value, out);
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Sequence(exprs) => {
            for e in exprs {
                collect_addr_taken_in_expr(e, out);
            }
        }
        ExprKind::Cast { expr, .. } => collect_addr_taken_in_expr(expr, out),
        ExprKind::RefLoad(inner) => collect_addr_taken_in_expr(inner, out),
        _ => {}
    }
}

/// Pre-scan: collect names used as the PLACE of an `ExprKind::Atomic`
/// operation, so their bindings can be promoted to a SHARED-MEMORY WORD at
/// declaration (`references::emit_shared_word_new`) — the storage a WASM
/// atomic acts on. A pointer cell is the wrong kind here: an atomic on a cell
/// object is an atomic on a copy.
///
/// Unlike `collect_addr_taken_idents`, this DOES descend into nested function,
/// lambda and class-member bodies: whether a binding is a shared word is a
/// whole-module property (an `Interlocked.Add` inside a `Task.Run` lambda
/// makes the TOP-LEVEL variable shared), the same forward-pass argument
/// `module_addr_taken_globals` documents. Over-approximation is safe for the
/// same reason stated there: the deref dispatchers pass a non-reference
/// through untouched.
///
/// A place this scan misses is LOUD, not wrong: the binding stays unpromoted
/// and `emit_atomic` refuses it at compile time.
fn collect_atomic_place_idents(stmts: &[Statement], out: &mut HashSet<String>) {
    for stmt in stmts {
        collect_atomic_place_in_stmt(stmt, out);
    }
}

fn collect_atomic_place_in_stmt(stmt: &Statement, out: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::Expr(e) => collect_atomic_place_in_expr(e, out),
        StmtKind::Return(opt) => {
            if let Some(e) = opt {
                collect_atomic_place_in_expr(e, out);
            }
        }
        StmtKind::Throw { expr, cause } => {
            for e in [expr, cause].into_iter().flatten() {
                collect_atomic_place_in_expr(e, out);
            }
        }
        StmtKind::Echo(exprs) => {
            for e in exprs {
                collect_atomic_place_in_expr(e, out);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations {
                if let Some(e) = &d.init {
                    collect_atomic_place_in_expr(e, out);
                }
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            for t in targets {
                collect_atomic_place_in_expr(t, out);
            }
            collect_atomic_place_in_expr(value, out);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            collect_atomic_place_in_expr(target, out);
            collect_atomic_place_in_expr(value, out);
        }
        StmtKind::Block(stmts) | StmtKind::NamespaceDecl { body: stmts, .. } => {
            collect_atomic_place_idents(stmts, out)
        }
        StmtKind::FunctionDecl { body, .. } => collect_atomic_place_idents(body, out),
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for m in members {
                match m {
                    vybe_ast::ClassMember::Method(inner)
                    | vybe_ast::ClassMember::NestedType(inner) => {
                        collect_atomic_place_in_stmt(inner, out)
                    }
                    vybe_ast::ClassMember::Constructor { body, .. } => {
                        collect_atomic_place_idents(body, out)
                    }
                    vybe_ast::ClassMember::Property { getter, setter, .. } => {
                        if let Some(g) = getter {
                            collect_atomic_place_idents(g, out);
                        }
                        if let Some(s) = setter {
                            collect_atomic_place_idents(&s.body, out);
                        }
                    }
                    vybe_ast::ClassMember::Field { init: Some(e), .. } => {
                        collect_atomic_place_in_expr(e, out)
                    }
                    _ => {}
                }
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            collect_atomic_place_in_expr(cond, out);
            collect_atomic_place_idents(then_body, out);
            for (c, b) in elifs {
                collect_atomic_place_in_expr(c, out);
                collect_atomic_place_idents(b, out);
            }
            if let Some(b) = else_body {
                collect_atomic_place_idents(b, out);
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            collect_atomic_place_in_expr(cond, out);
            collect_atomic_place_idents(body, out);
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
            ..
        } => {
            if let Some(i) = init {
                collect_atomic_place_in_stmt(i, out);
            }
            for e in [cond, update].into_iter().flatten() {
                collect_atomic_place_in_expr(e, out);
            }
            collect_atomic_place_idents(body, out);
        }
        StmtKind::ForIn { iter, body, .. } => {
            collect_atomic_place_in_expr(iter, out);
            collect_atomic_place_idents(body, out);
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            collect_atomic_place_idents(body, out);
            for c in catches {
                collect_atomic_place_idents(&c.body, out);
            }
            for b in [else_body, finally].into_iter().flatten() {
                collect_atomic_place_idents(b, out);
            }
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            collect_atomic_place_in_expr(expr, out);
            for case in cases {
                collect_atomic_place_idents(&case.body, out);
            }
            if let Some(b) = default {
                collect_atomic_place_idents(b, out);
            }
        }
        StmtKind::Using { body, .. } | StmtKind::Lock { body, .. } => {
            collect_atomic_place_idents(body, out)
        }
        StmtKind::Labeled { body, .. } => collect_atomic_place_in_stmt(body, out),
        _ => {}
    }
}

fn collect_atomic_place_in_expr(expr: &Expression, out: &mut HashSet<String>) {
    if let ExprKind::Atomic(op) = &expr.kind {
        // The point of the scan: record the place's root name.
        for child in op.children() {
            collect_atomic_place_in_expr(child, out);
        }
        let place = match op {
            vybe_ast::AtomicOp::Load { place, .. }
            | vybe_ast::AtomicOp::Store { place, .. }
            | vybe_ast::AtomicOp::Rmw { place, .. }
            | vybe_ast::AtomicOp::CompareExchange { place, .. } => Some(place),
            vybe_ast::AtomicOp::Fence { .. } => None,
        };
        if let Some(place) = place {
            if let ExprKind::Ident(name) = &place.kind {
                out.insert(name.clone());
            }
        }
        return;
    }
    match &expr.kind {
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::Spread(expr) => collect_atomic_place_in_expr(expr, out),
        ExprKind::Binary { left, right, .. } | ExprKind::NullCoalesce { left, right } => {
            collect_atomic_place_in_expr(left, out);
            collect_atomic_place_in_expr(right, out);
        }
        ExprKind::Assign { target, value } => {
            collect_atomic_place_in_expr(target, out);
            collect_atomic_place_in_expr(value, out);
        }
        ExprKind::Call { callee, args, .. } => {
            collect_atomic_place_in_expr(callee, out);
            for a in args {
                collect_atomic_place_in_expr(&a.value, out);
            }
        }
        ExprKind::New { class, args } => {
            collect_atomic_place_in_expr(class, out);
            for a in args {
                collect_atomic_place_in_expr(&a.value, out);
            }
        }
        ExprKind::Member { object, .. } => collect_atomic_place_in_expr(object, out),
        ExprKind::Index { object, index, .. } => {
            collect_atomic_place_in_expr(object, out);
            collect_atomic_place_in_expr(index, out);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            for e in [cond, then, else_] {
                collect_atomic_place_in_expr(e, out);
            }
        }
        ExprKind::Array(elems) => {
            for e in elems {
                collect_atomic_place_in_expr(&e.value, out);
            }
        }
        ExprKind::Sequence(exprs) => {
            for e in exprs {
                collect_atomic_place_in_expr(e, out);
            }
        }
        ExprKind::Async(op) => {
            for child in op.children() {
                collect_atomic_place_in_expr(child, out);
            }
        }
        ExprKind::Chan(op) => {
            for child in op.children() {
                collect_atomic_place_in_expr(child, out);
            }
        }
        // The whole-module property: descend into nested callables.
        ExprKind::Lambda { body, .. } => match body {
            vybe_ast::LambdaBody::Expr(e) => collect_atomic_place_in_expr(e, out),
            vybe_ast::LambdaBody::Block(stmts) => collect_atomic_place_idents(stmts, out),
        },
        ExprKind::FunctionExpr(decl) => collect_atomic_place_in_stmt(decl, out),
        _ => {}
    }
}

/// Pre-scan: collect identifiers referenced inside nested function/lambda
/// bodies. These are potential closure captures — the declaring function
/// wraps them in a cell object so mutations propagate through shared refs.
fn collect_closure_captured_idents(stmts: &[Statement], out: &mut HashSet<String>) {
    for stmt in stmts {
        collect_closure_captured_in_stmt(stmt, out);
    }
}

fn collect_declared_names(stmts: &[Statement], out: &mut HashSet<String>) {
    for stmt in stmts {
        collect_declared_names_in_stmt(stmt, out);
    }
}

fn collect_declared_names_in_stmt(stmt: &Statement, out: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations {
                collect_binding_pattern_names(&d.pattern, out);
            }
        }
        StmtKind::FunctionDecl { name, .. } => {
            out.insert(name.clone());
        }
        StmtKind::ClassDecl { name, .. } => {
            out.insert(name.clone());
        }
        StmtKind::Block(stmts) => collect_declared_names(stmts, out),
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            collect_declared_names(then_body, out);
            for (_, b) in elifs {
                collect_declared_names(b, out);
            }
            if let Some(b) = else_body {
                collect_declared_names(b, out);
            }
        }
        StmtKind::While { body, .. } | StmtKind::DoWhile { body, .. } => {
            collect_declared_names(body, out);
        }
        StmtKind::For { init, body, .. } => {
            if let Some(i) = init {
                collect_declared_names_in_stmt(i, out);
            }
            collect_declared_names(body, out);
        }
        StmtKind::ForIn { var, body, .. } => {
            out.insert(var.clone());
            collect_declared_names(body, out);
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            collect_declared_names(body, out);
            for c in catches {
                if let Some(name) = &c.var_name {
                    out.insert(name.clone());
                }
                collect_declared_names(&c.body, out);
            }
            if let Some(b) = else_body {
                collect_declared_names(b, out);
            }
            if let Some(b) = finally {
                collect_declared_names(b, out);
            }
        }
        StmtKind::Switch { cases, default, .. } => {
            for case in cases {
                collect_declared_names(&case.body, out);
            }
            if let Some(b) = default {
                collect_declared_names(b, out);
            }
        }
        StmtKind::Labeled { body, .. } => collect_declared_names_in_stmt(body, out),
        _ => {}
    }
}

/// Public wrapper for tooling (vybex eval's §19.2.1.1 var-name harvest).
pub fn collect_binding_pattern_names_pub(
    pat: &crate::ast::BindingPattern,
    out: &mut HashSet<String>,
) {
    collect_binding_pattern_names(pat, out);
}

fn collect_binding_pattern_names(pat: &crate::ast::BindingPattern, out: &mut HashSet<String>) {
    match pat {
        crate::ast::BindingPattern::Ident(name) => {
            out.insert(name.clone());
        }
        crate::ast::BindingPattern::Array(elems) => {
            for elem in elems {
                match elem {
                    crate::ast::ArrayPatternElem::Pattern(p, _) => {
                        collect_binding_pattern_names(p, out)
                    }
                    crate::ast::ArrayPatternElem::Rest(name) => {
                        out.insert(name.clone());
                    }
                    crate::ast::ArrayPatternElem::Hole => {}
                }
            }
        }
        crate::ast::BindingPattern::Object(fields) => {
            for f in fields {
                if let Some(p) = &f.value {
                    collect_binding_pattern_names(p, out);
                } else {
                    out.insert(f.key.clone());
                }
            }
        }
    }
}

/// Free identifiers of a class body — the names it closes over from the
/// enclosing frame. Each member is scoped independently: a method's own params
/// and locals are not captures. Methods are `FunctionDecl` statements, so they
/// reuse the nested-function arm verbatim.
fn collect_closure_captured_in_class_members(members: &[ClassMember], out: &mut HashSet<String>) {
    for member in members {
        match member {
            ClassMember::Method(stmt) => collect_closure_captured_in_stmt(stmt, out),
            ClassMember::Constructor { params, body, .. } => {
                let mut all_idents = HashSet::new();
                collect_all_idents_in_stmts(body, &mut all_idents);
                let mut local_names: HashSet<String> =
                    params.iter().map(|p| p.name.clone()).collect();
                collect_declared_names(body, &mut local_names);
                for name in all_idents {
                    if !local_names.contains(&name) {
                        out.insert(name);
                    }
                }
            }
            ClassMember::Field {
                init: Some(init), ..
            } => collect_closure_captured_in_expr(init, out),
            _ => {}
        }
    }
}

fn collect_closure_captured_in_stmt(stmt: &Statement, out: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut all_idents = HashSet::new();
            collect_all_idents_in_stmts(body, &mut all_idents);
            let mut local_names: HashSet<String> = params.iter().map(|p| p.name.clone()).collect();
            collect_declared_names(body, &mut local_names);
            for name in all_idents {
                if !local_names.contains(&name) {
                    out.insert(name);
                }
            }
        }
        // A class body closes over the enclosing frame exactly like a nested
        // function does: its methods can read an enclosing local, so those
        // names must be boxed into the shared env here. Without this the
        // capture is resolved but never boxed, and the method reads through an
        // unboxed value (`function mk(msg){ return class { greet(){ return msg; } }; }`).
        StmtKind::ClassDecl { members, .. } => {
            collect_closure_captured_in_class_members(members, out);
        }
        StmtKind::Expr(e) => collect_closure_captured_in_expr(e, out),
        StmtKind::Return(opt) => {
            if let Some(e) = opt {
                collect_closure_captured_in_expr(e, out);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations {
                if let Some(e) = &d.init {
                    collect_closure_captured_in_expr(e, out);
                }
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            for t in targets {
                collect_closure_captured_in_expr(t, out);
            }
            collect_closure_captured_in_expr(value, out);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            collect_closure_captured_in_expr(target, out);
            collect_closure_captured_in_expr(value, out);
        }
        StmtKind::Block(stmts) => collect_closure_captured_idents(stmts, out),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            collect_closure_captured_in_expr(cond, out);
            collect_closure_captured_idents(then_body, out);
            for (c, b) in elifs {
                collect_closure_captured_in_expr(c, out);
                collect_closure_captured_idents(b, out);
            }
            if let Some(b) = else_body {
                collect_closure_captured_idents(b, out);
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            collect_closure_captured_in_expr(cond, out);
            collect_closure_captured_idents(body, out);
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
            ..
        } => {
            if let Some(i) = init {
                collect_closure_captured_in_stmt(i, out);
            }
            if let Some(c) = cond {
                collect_closure_captured_in_expr(c, out);
            }
            if let Some(u) = update {
                collect_closure_captured_in_expr(u, out);
            }
            collect_closure_captured_idents(body, out);
        }
        StmtKind::ForIn { iter, body, .. } => {
            collect_closure_captured_in_expr(iter, out);
            collect_closure_captured_idents(body, out);
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            collect_closure_captured_idents(body, out);
            for c in catches {
                collect_closure_captured_idents(&c.body, out);
            }
            if let Some(b) = else_body {
                collect_closure_captured_idents(b, out);
            }
            if let Some(b) = finally {
                collect_closure_captured_idents(b, out);
            }
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            collect_closure_captured_in_expr(expr, out);
            for case in cases {
                collect_closure_captured_idents(&case.body, out);
            }
            if let Some(b) = default {
                collect_closure_captured_idents(b, out);
            }
        }
        StmtKind::Labeled { body, .. } => collect_closure_captured_in_stmt(body, out),
        StmtKind::Throw { expr, .. } => {
            if let Some(e) = expr {
                collect_closure_captured_in_expr(e, out);
            }
        }
        _ => {}
    }
}

pub(crate) fn body_contains_this(stmts: &[Statement]) -> bool {
    stmts.iter().any(|s| stmt_contains_this(s))
}

fn stmt_contains_this(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Expr(e) | StmtKind::Return(Some(e)) | StmtKind::Throw { expr: Some(e), .. } => {
            expr_contains_this(e)
        }
        StmtKind::VarDecl { declarations, .. } => declarations
            .iter()
            .any(|d| d.init.as_ref().is_some_and(expr_contains_this)),
        StmtKind::Assign { value, .. } | StmtKind::CompoundAssign { value, .. } => {
            expr_contains_this(value)
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
            ..
        } => {
            expr_contains_this(cond)
                || body_contains_this(then_body)
                || elifs
                    .iter()
                    .any(|(c, b)| expr_contains_this(c) || body_contains_this(b))
                || else_body.as_ref().is_some_and(|b| body_contains_this(b))
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            expr_contains_this(cond) || body_contains_this(body)
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
            ..
        } => {
            init.as_ref().is_some_and(|s| stmt_contains_this(s))
                || cond.as_ref().is_some_and(expr_contains_this)
                || update.as_ref().is_some_and(expr_contains_this)
                || body_contains_this(body)
        }
        StmtKind::ForIn { iter, body, .. } => expr_contains_this(iter) || body_contains_this(body),
        StmtKind::Block(stmts) => body_contains_this(stmts),
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            body_contains_this(body)
                || catches.iter().any(|c| body_contains_this(&c.body))
                || finally.as_ref().is_some_and(|b| body_contains_this(b))
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
            ..
        } => {
            expr_contains_this(expr)
                || cases.iter().any(|c| body_contains_this(&c.body))
                || default.as_ref().is_some_and(|b| body_contains_this(b))
        }
        StmtKind::Labeled { body, .. } => stmt_contains_this(body),
        _ => false,
    }
}

pub(crate) fn expr_contains_this(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::This | ExprKind::Super => true,
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Block(stmts) => body_contains_this(stmts),
            LambdaBody::Expr(e) => expr_contains_this(e),
        },
        ExprKind::FunctionExpr(_) => false,
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::Spread(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Delete(expr) => expr_contains_this(expr),
        ExprKind::Async(op) => op.children().into_iter().any(expr_contains_this),
        ExprKind::Chan(op) => op.children().into_iter().any(expr_contains_this),
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Assign {
            target: left,
            value: right,
        } => expr_contains_this(left) || expr_contains_this(right),
        ExprKind::Ternary { cond, then, else_ } => {
            expr_contains_this(cond) || expr_contains_this(then) || expr_contains_this(else_)
        }
        ExprKind::Call { callee, args, .. } => {
            expr_contains_this(callee) || args.iter().any(|a| expr_contains_this(&a.value))
        }
        ExprKind::Member { object, .. } | ExprKind::Index { object, .. } => {
            expr_contains_this(object)
        }
        ExprKind::Array(elems) => elems.iter().any(|e| expr_contains_this(&e.value)),
        ExprKind::Object(props) => props.iter().any(|p| match p {
            ObjectProperty::KeyValue { value, .. } => expr_contains_this(value),
            _ => false,
        }),
        ExprKind::Interpolation(parts) => parts.iter().any(|p| match p {
            crate::ast::InterpolPart::Expr(e) | crate::ast::InterpolPart::Formatted(e, _) => {
                expr_contains_this(e)
            }
            _ => false,
        }),
        ExprKind::Sequence(exprs) => exprs.iter().any(expr_contains_this),
        ExprKind::New { class, args } => {
            expr_contains_this(class) || args.iter().any(|a| expr_contains_this(&a.value))
        }
        ExprKind::Yield(Some(inner)) => expr_contains_this(inner),
        _ => false,
    }
}

pub(crate) fn closures_in_body_reference_this(stmts: &[Statement]) -> bool {
    stmts.iter().any(|s| stmt_has_closure_with_this(s))
}

fn stmt_has_closure_with_this(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Expr(e) | StmtKind::Return(Some(e)) | StmtKind::Throw { expr: Some(e), .. } => {
            expr_has_closure_with_this(e)
        }
        StmtKind::VarDecl { declarations, .. } => declarations
            .iter()
            .any(|d| d.init.as_ref().is_some_and(expr_has_closure_with_this)),
        StmtKind::Assign { value, .. } | StmtKind::CompoundAssign { value, .. } => {
            expr_has_closure_with_this(value)
        }
        StmtKind::Block(stmts) => closures_in_body_reference_this(stmts),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
            ..
        } => {
            expr_has_closure_with_this(cond)
                || closures_in_body_reference_this(then_body)
                || elifs.iter().any(|(c, b)| {
                    expr_has_closure_with_this(c) || closures_in_body_reference_this(b)
                })
                || else_body
                    .as_ref()
                    .is_some_and(|b| closures_in_body_reference_this(b))
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            expr_has_closure_with_this(cond) || closures_in_body_reference_this(body)
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
            ..
        } => {
            init.as_ref().is_some_and(|s| stmt_has_closure_with_this(s))
                || cond.as_ref().is_some_and(expr_has_closure_with_this)
                || update.as_ref().is_some_and(expr_has_closure_with_this)
                || closures_in_body_reference_this(body)
        }
        StmtKind::ForIn { iter, body, .. } => {
            expr_has_closure_with_this(iter) || closures_in_body_reference_this(body)
        }
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            closures_in_body_reference_this(body)
                || catches
                    .iter()
                    .any(|c| closures_in_body_reference_this(&c.body))
                || finally
                    .as_ref()
                    .is_some_and(|b| closures_in_body_reference_this(b))
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
            ..
        } => {
            expr_has_closure_with_this(expr)
                || cases
                    .iter()
                    .any(|c| closures_in_body_reference_this(&c.body))
                || default
                    .as_ref()
                    .is_some_and(|b| closures_in_body_reference_this(b))
        }
        StmtKind::Labeled { body, .. } => stmt_has_closure_with_this(body),
        _ => false,
    }
}

fn expr_has_closure_with_this(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Block(stmts) => body_contains_this(stmts),
            LambdaBody::Expr(e) => expr_contains_this(e),
        },
        ExprKind::FunctionExpr(_) => false,
        ExprKind::Unary { expr, .. } | ExprKind::Await(expr) | ExprKind::Spread(expr) => {
            expr_has_closure_with_this(expr)
        }
        ExprKind::Async(op) => op.children().into_iter().any(expr_has_closure_with_this),
        ExprKind::Chan(op) => op.children().into_iter().any(expr_has_closure_with_this),
        ExprKind::Binary { left, right, .. }
        | ExprKind::Assign {
            target: left,
            value: right,
        } => expr_has_closure_with_this(left) || expr_has_closure_with_this(right),
        ExprKind::Ternary { cond, then, else_ } => {
            expr_has_closure_with_this(cond)
                || expr_has_closure_with_this(then)
                || expr_has_closure_with_this(else_)
        }
        ExprKind::Call { callee, args, .. } => {
            expr_has_closure_with_this(callee)
                || args.iter().any(|a| expr_has_closure_with_this(&a.value))
        }
        ExprKind::Member { object, .. } | ExprKind::Index { object, .. } => {
            expr_has_closure_with_this(object)
        }
        ExprKind::Array(elems) => elems.iter().any(|e| expr_has_closure_with_this(&e.value)),
        ExprKind::Object(props) => props.iter().any(|p| match p {
            ObjectProperty::KeyValue { value, .. } => expr_has_closure_with_this(value),
            _ => false,
        }),
        ExprKind::New { class, args } => {
            expr_has_closure_with_this(class)
                || args.iter().any(|a| expr_has_closure_with_this(&a.value))
        }
        ExprKind::Sequence(exprs) => exprs.iter().any(expr_has_closure_with_this),
        _ => false,
    }
}

pub(crate) fn collect_closure_captured_in_expr(expr: &Expression, out: &mut HashSet<String>) {
    match &expr.kind {
        // `const K = class { m(){ return msg; } }` closes over the enclosing
        // frame just as the declaration form does — same members, same rule.
        ExprKind::ClassExpr { members, .. } => {
            collect_closure_captured_in_class_members(members, out);
        }
        ExprKind::Lambda { params, body, .. } => {
            let mut all_idents = HashSet::new();
            match body {
                LambdaBody::Block(stmts) => collect_all_idents_in_stmts(stmts, &mut all_idents),
                LambdaBody::Expr(e) => collect_all_idents_in_expr(e, &mut all_idents),
            }
            let mut local_names: HashSet<String> = params.iter().map(|p| p.name.clone()).collect();
            if let LambdaBody::Block(stmts) = body {
                collect_declared_names(stmts, &mut local_names);
            }
            for name in all_idents {
                if !local_names.contains(&name) {
                    out.insert(name);
                }
            }
        }
        ExprKind::FunctionExpr(stmt) => {
            if let StmtKind::FunctionDecl { params, body, .. } = &stmt.kind {
                let mut all_idents = HashSet::new();
                collect_all_idents_in_stmts(body, &mut all_idents);
                let mut local_names: HashSet<String> =
                    params.iter().map(|p| p.name.clone()).collect();
                collect_declared_names(body, &mut local_names);
                for name in all_idents {
                    if !local_names.contains(&name) {
                        out.insert(name);
                    }
                }
            }
        }
        ExprKind::Unary { expr, .. } => collect_closure_captured_in_expr(expr, out),
        ExprKind::Binary { left, right, .. } => {
            collect_closure_captured_in_expr(left, out);
            collect_closure_captured_in_expr(right, out);
        }
        ExprKind::Call { callee, args, .. } => {
            collect_closure_captured_in_expr(callee, out);
            for a in args {
                collect_closure_captured_in_expr(&a.value, out);
            }
        }
        ExprKind::Member { object, .. } => collect_closure_captured_in_expr(object, out),
        ExprKind::Index { object, index, .. } => {
            collect_closure_captured_in_expr(object, out);
            collect_closure_captured_in_expr(index, out);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            collect_closure_captured_in_expr(cond, out);
            collect_closure_captured_in_expr(then, out);
            collect_closure_captured_in_expr(else_, out);
        }
        ExprKind::Assign { target, value } => {
            collect_closure_captured_in_expr(target, out);
            collect_closure_captured_in_expr(value, out);
        }
        ExprKind::Array(elems) => {
            for e in elems {
                collect_closure_captured_in_expr(&e.value, out);
            }
        }
        ExprKind::Object(props) => {
            for p in props {
                match p {
                    ObjectProperty::KeyValue { value, .. } => {
                        collect_closure_captured_in_expr(value, out)
                    }
                    ObjectProperty::Computed { key, value } => {
                        collect_closure_captured_in_expr(key, out);
                        collect_closure_captured_in_expr(value, out);
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        if let StmtKind::FunctionDecl { params, body, .. } = &value.kind {
                            let mut all_idents = HashSet::new();
                            collect_all_idents_in_stmts(body, &mut all_idents);
                            let mut local_names: HashSet<String> =
                                params.iter().map(|p| p.name.clone()).collect();
                            collect_declared_names(body, &mut local_names);
                            for name in all_idents {
                                if !local_names.contains(&name) {
                                    out.insert(name);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Sequence(exprs) => {
            for e in exprs {
                collect_closure_captured_in_expr(e, out);
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let crate::ast::InterpolPart::Expr(e)
                | crate::ast::InterpolPart::Formatted(e, _) = part
                {
                    collect_closure_captured_in_expr(e, out);
                }
            }
        }
        ExprKind::Await(inner) | ExprKind::Spread(inner) | ExprKind::Yield(Some(inner)) => {
            collect_closure_captured_in_expr(inner, out);
        }
        // The async/channel vocabularies carry expressions too — a goroutine
        // body's `ch <- v` is `Chan(Send{Ident(ch)})`, and missing it here
        // left `ch` out of the enclosing function's env array while the
        // lambda still read it there (measured: every goroutine channel
        // capture arrived undefined).
        ExprKind::Async(op) => {
            for child in op.children() {
                collect_closure_captured_in_expr(child, out);
            }
        }
        ExprKind::Chan(op) => {
            for child in op.children() {
                collect_closure_captured_in_expr(child, out);
            }
        }
        ExprKind::New { class, args } => {
            collect_closure_captured_in_expr(class, out);
            for a in args {
                collect_closure_captured_in_expr(&a.value, out);
            }
        }
        ExprKind::Lit(_) | ExprKind::This | ExprKind::Super | ExprKind::Yield(None) => {}
        _ => {}
    }
}

fn collect_all_idents_in_stmts(stmts: &[Statement], out: &mut HashSet<String>) {
    for stmt in stmts {
        collect_all_idents_in_stmt(stmt, out);
    }
}

fn collect_all_idents_in_stmt(stmt: &Statement, out: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::Expr(e) => collect_all_idents_in_expr(e, out),
        StmtKind::Return(opt) => {
            if let Some(e) = opt {
                collect_all_idents_in_expr(e, out);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations {
                if let Some(e) = &d.init {
                    collect_all_idents_in_expr(e, out);
                }
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            for t in targets {
                collect_all_idents_in_expr(t, out);
            }
            collect_all_idents_in_expr(value, out);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            collect_all_idents_in_expr(target, out);
            collect_all_idents_in_expr(value, out);
        }
        StmtKind::Block(stmts) => collect_all_idents_in_stmts(stmts, out),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            collect_all_idents_in_expr(cond, out);
            collect_all_idents_in_stmts(then_body, out);
            for (c, b) in elifs {
                collect_all_idents_in_expr(c, out);
                collect_all_idents_in_stmts(b, out);
            }
            if let Some(b) = else_body {
                collect_all_idents_in_stmts(b, out);
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            collect_all_idents_in_expr(cond, out);
            collect_all_idents_in_stmts(body, out);
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
            ..
        } => {
            if let Some(i) = init {
                collect_all_idents_in_stmt(i, out);
            }
            if let Some(c) = cond {
                collect_all_idents_in_expr(c, out);
            }
            if let Some(u) = update {
                collect_all_idents_in_expr(u, out);
            }
            collect_all_idents_in_stmts(body, out);
        }
        StmtKind::ForIn { iter, body, .. } => {
            collect_all_idents_in_expr(iter, out);
            collect_all_idents_in_stmts(body, out);
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            collect_all_idents_in_stmts(body, out);
            for c in catches {
                collect_all_idents_in_stmts(&c.body, out);
            }
            if let Some(b) = else_body {
                collect_all_idents_in_stmts(b, out);
            }
            if let Some(b) = finally {
                collect_all_idents_in_stmts(b, out);
            }
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            collect_all_idents_in_expr(expr, out);
            for case in cases {
                collect_all_idents_in_stmts(&case.body, out);
            }
            if let Some(b) = default {
                collect_all_idents_in_stmts(b, out);
            }
        }
        StmtKind::FunctionDecl { body, .. } => collect_all_idents_in_stmts(body, out),
        StmtKind::Labeled { body, .. } => collect_all_idents_in_stmt(body, out),
        StmtKind::Throw { expr, .. } => {
            if let Some(e) = expr {
                collect_all_idents_in_expr(e, out);
            }
        }
        _ => {}
    }
}

fn collect_all_idents_in_expr(expr: &Expression, out: &mut HashSet<String>) {
    match &expr.kind {
        ExprKind::Ident(name) => {
            out.insert(name.clone());
        }
        ExprKind::Unary { expr, .. } => collect_all_idents_in_expr(expr, out),
        ExprKind::Binary { left, right, .. } => {
            collect_all_idents_in_expr(left, out);
            collect_all_idents_in_expr(right, out);
        }
        ExprKind::Call { callee, args, .. } => {
            collect_all_idents_in_expr(callee, out);
            for a in args {
                collect_all_idents_in_expr(&a.value, out);
            }
        }
        ExprKind::Member { object, .. } => collect_all_idents_in_expr(object, out),
        ExprKind::Index { object, index, .. } => {
            collect_all_idents_in_expr(object, out);
            collect_all_idents_in_expr(index, out);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            collect_all_idents_in_expr(cond, out);
            collect_all_idents_in_expr(then, out);
            collect_all_idents_in_expr(else_, out);
        }
        ExprKind::Assign { target, value } => {
            collect_all_idents_in_expr(target, out);
            collect_all_idents_in_expr(value, out);
        }
        ExprKind::Array(elems) => {
            for e in elems {
                collect_all_idents_in_expr(&e.value, out);
            }
        }
        ExprKind::Object(props) => {
            for p in props {
                match p {
                    ObjectProperty::KeyValue { value, .. } => {
                        collect_all_idents_in_expr(value, out)
                    }
                    ObjectProperty::Computed { key, value } => {
                        collect_all_idents_in_expr(key, out);
                        collect_all_idents_in_expr(value, out);
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        if let StmtKind::FunctionDecl { params, body, .. } = &value.kind {
                            let mut all_idents = HashSet::new();
                            collect_all_idents_in_stmts(body, &mut all_idents);
                            let mut local_names: HashSet<String> =
                                params.iter().map(|p| p.name.clone()).collect();
                            collect_declared_names(body, &mut local_names);
                            for name in all_idents {
                                if !local_names.contains(&name) {
                                    out.insert(name);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Sequence(exprs) => {
            for e in exprs {
                collect_all_idents_in_expr(e, out);
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let crate::ast::InterpolPart::Expr(e)
                | crate::ast::InterpolPart::Formatted(e, _) = part
                {
                    collect_all_idents_in_expr(e, out);
                }
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Block(stmts) => collect_all_idents_in_stmts(stmts, out),
            LambdaBody::Expr(e) => collect_all_idents_in_expr(e, out),
        },
        ExprKind::FunctionExpr(stmt) => {
            if let StmtKind::FunctionDecl { body, .. } = &stmt.kind {
                collect_all_idents_in_stmts(body, out);
            }
        }
        ExprKind::New { class, args } => {
            collect_all_idents_in_expr(class, out);
            for a in args {
                collect_all_idents_in_expr(&a.value, out);
            }
        }
        ExprKind::Await(inner) | ExprKind::Spread(inner) | ExprKind::Yield(Some(inner)) => {
            collect_all_idents_in_expr(inner, out);
        }
        ExprKind::Async(op) => {
            for child in op.children() {
                collect_all_idents_in_expr(child, out);
            }
        }
        ExprKind::Chan(op) => {
            for child in op.children() {
                collect_all_idents_in_expr(child, out);
            }
        }
        _ => {}
    }
}

fn stmt_uses_proxy(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Expr(e) => expr_uses_proxy(e),
        StmtKind::Return(opt) => opt.as_ref().map_or(false, expr_uses_proxy),
        StmtKind::Throw { expr, cause } => {
            expr.as_ref().map_or(false, expr_uses_proxy)
                || cause.as_ref().map_or(false, expr_uses_proxy)
        }
        StmtKind::VarDecl { declarations, .. } => declarations
            .iter()
            .any(|d| d.init.as_ref().map_or(false, expr_uses_proxy)),
        StmtKind::Assign { value, .. } | StmtKind::CompoundAssign { value, .. } => {
            expr_uses_proxy(value)
        }
        StmtKind::Block(stmts) => stmts.iter().any(stmt_uses_proxy),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            expr_uses_proxy(cond)
                || then_body.iter().any(stmt_uses_proxy)
                || elifs
                    .iter()
                    .any(|(c, b)| expr_uses_proxy(c) || b.iter().any(stmt_uses_proxy))
                || else_body
                    .as_ref()
                    .map_or(false, |b| b.iter().any(stmt_uses_proxy))
        }
        StmtKind::While { cond, body, .. } => {
            expr_uses_proxy(cond) || body.iter().any(stmt_uses_proxy)
        }
        StmtKind::DoWhile { cond, body, .. } => {
            expr_uses_proxy(cond) || body.iter().any(stmt_uses_proxy)
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
            ..
        } => {
            init.as_ref().map_or(false, |i| stmt_uses_proxy(i))
                || cond.as_ref().map_or(false, expr_uses_proxy)
                || update.as_ref().map_or(false, expr_uses_proxy)
                || body.iter().any(stmt_uses_proxy)
        }
        StmtKind::ForIn { iter, body, .. } => {
            expr_uses_proxy(iter) || body.iter().any(stmt_uses_proxy)
        }
        StmtKind::FunctionDecl { body, .. } => body.iter().any(stmt_uses_proxy),
        StmtKind::ClassDecl { members, .. } => members.iter().any(|m| match m {
            ClassMember::Method(s) => stmt_uses_proxy(s),
            ClassMember::Constructor { body, .. } => body.iter().any(stmt_uses_proxy),
            ClassMember::Field { init, .. } => init.as_ref().map_or(false, expr_uses_proxy),
            ClassMember::Property { getter, setter, .. } => {
                getter
                    .as_ref()
                    .map_or(false, |b| b.iter().any(stmt_uses_proxy))
                    || setter
                        .as_ref()
                        .map_or(false, |s| s.body.iter().any(stmt_uses_proxy))
            }
            _ => false,
        }),
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            body.iter().any(stmt_uses_proxy)
                || catches.iter().any(|c| c.body.iter().any(stmt_uses_proxy))
                || finally
                    .as_ref()
                    .map_or(false, |b| b.iter().any(stmt_uses_proxy))
        }
        _ => false,
    }
}

fn expr_uses_proxy(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Proxy { .. } => true,
        ExprKind::New { args, .. } => args.iter().any(|a| expr_uses_proxy(&a.value)),
        ExprKind::Call { callee, args, .. } => {
            // Proxy.revocable(...) creates a proxy without `new Proxy`.
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if field == "revocable"
                    && matches!(&object.kind, ExprKind::Ident(n) if n == "Proxy")
                {
                    return true;
                }
            }
            expr_uses_proxy(callee) || args.iter().any(|a| expr_uses_proxy(&a.value))
        }
        ExprKind::Binary { left, right, .. } => expr_uses_proxy(left) || expr_uses_proxy(right),
        ExprKind::Unary { expr, .. } => expr_uses_proxy(expr),
        ExprKind::Member { object, .. } => expr_uses_proxy(object),
        ExprKind::Index { object, index, .. } => expr_uses_proxy(object) || expr_uses_proxy(index),
        ExprKind::Ternary { cond, then, else_ } => {
            expr_uses_proxy(cond) || expr_uses_proxy(then) || expr_uses_proxy(else_)
        }
        ExprKind::Array(elems) => elems.iter().any(|e| expr_uses_proxy(&e.value)),
        ExprKind::Object(props) => props.iter().any(|p| match p {
            ObjectProperty::KeyValue { value, .. } => expr_uses_proxy(value),
            ObjectProperty::Computed { key, value } => {
                expr_uses_proxy(key) || expr_uses_proxy(value)
            }
            _ => false,
        }),
        ExprKind::Assign { value, .. } => expr_uses_proxy(value),
        ExprKind::Lambda { body, .. } => match body {
            crate::ast::LambdaBody::Expr(e) => expr_uses_proxy(e),
            crate::ast::LambdaBody::Block(b) => b.iter().any(stmt_uses_proxy),
        },
        _ => false,
    }
}

fn stmt_uses_js_arguments(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Expr(expr) => expr_uses_js_arguments(expr),
        StmtKind::Return(expr) => expr.as_ref().is_some_and(expr_uses_js_arguments),
        StmtKind::Throw { expr, cause } => {
            expr.as_ref().is_some_and(expr_uses_js_arguments)
                || cause.as_ref().is_some_and(expr_uses_js_arguments)
        }
        StmtKind::VarDecl { declarations, .. } => declarations
            .iter()
            .any(|decl| decl.init.as_ref().is_some_and(expr_uses_js_arguments)),
        StmtKind::Assign { targets, value, .. } => {
            targets.iter().any(expr_uses_js_arguments) || expr_uses_js_arguments(value)
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            expr_uses_js_arguments(target) || expr_uses_js_arguments(value)
        }
        StmtKind::Block(body) | StmtKind::Using { body, .. } | StmtKind::Lock { body, .. } => {
            body.iter().any(stmt_uses_js_arguments)
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            expr_uses_js_arguments(cond)
                || then_body.iter().any(stmt_uses_js_arguments)
                || elifs.iter().any(|(cond, body)| {
                    expr_uses_js_arguments(cond) || body.iter().any(stmt_uses_js_arguments)
                })
                || else_body
                    .as_ref()
                    .is_some_and(|body| body.iter().any(stmt_uses_js_arguments))
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            expr_uses_js_arguments(cond) || body.iter().any(stmt_uses_js_arguments)
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
            ..
        } => {
            init.as_ref()
                .is_some_and(|stmt| stmt_uses_js_arguments(stmt))
                || cond.as_ref().is_some_and(expr_uses_js_arguments)
                || update.as_ref().is_some_and(expr_uses_js_arguments)
                || body.iter().any(stmt_uses_js_arguments)
        }
        StmtKind::ForIn { iter, body, .. } => {
            expr_uses_js_arguments(iter) || body.iter().any(stmt_uses_js_arguments)
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            expr_uses_js_arguments(expr)
                || cases.iter().any(|case| {
                    case.conditions.iter().any(|condition| match condition {
                        CaseCondition::Value(expr) => expr_uses_js_arguments(expr),
                        CaseCondition::Range { from, to } => {
                            expr_uses_js_arguments(from) || expr_uses_js_arguments(to)
                        }
                        CaseCondition::Comparison { expr, .. } => expr_uses_js_arguments(expr),
                    }) || case.body.iter().any(stmt_uses_js_arguments)
                })
                || default
                    .as_ref()
                    .is_some_and(|body| body.iter().any(stmt_uses_js_arguments))
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            body.iter().any(stmt_uses_js_arguments)
                || catches
                    .iter()
                    .any(|catch| catch.body.iter().any(stmt_uses_js_arguments))
                || else_body
                    .as_ref()
                    .is_some_and(|body| body.iter().any(stmt_uses_js_arguments))
                || finally
                    .as_ref()
                    .is_some_and(|body| body.iter().any(stmt_uses_js_arguments))
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            params
                .iter()
                .any(|param| param.default.as_ref().is_some_and(expr_uses_js_arguments))
                || body.iter().any(stmt_uses_js_arguments)
        }
        StmtKind::ClassDecl { members, .. } => members.iter().any(|member| match member {
            ClassMember::Method(stmt) => stmt_uses_js_arguments(stmt),
            ClassMember::Constructor { params, body, .. } => {
                params
                    .iter()
                    .any(|param| param.default.as_ref().is_some_and(expr_uses_js_arguments))
                    || body.iter().any(stmt_uses_js_arguments)
            }
            ClassMember::Field { init, .. } => init.as_ref().is_some_and(expr_uses_js_arguments),
            ClassMember::Property { getter, setter, .. } => {
                getter
                    .as_ref()
                    .is_some_and(|body| body.iter().any(stmt_uses_js_arguments))
                    || setter
                        .as_ref()
                        .is_some_and(|setter| setter.body.iter().any(stmt_uses_js_arguments))
            }
            _ => false,
        }),
        _ => false,
    }
}

fn expr_uses_js_arguments(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => name == "arguments",
        ExprKind::Binary { left, right, .. } => {
            expr_uses_js_arguments(left) || expr_uses_js_arguments(right)
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::Spread(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Delete(expr)
        | ExprKind::Void(expr) => expr_uses_js_arguments(expr),
        ExprKind::Async(op) => op.children().into_iter().any(expr_uses_js_arguments),
        ExprKind::Chan(op) => op.children().into_iter().any(expr_uses_js_arguments),
        ExprKind::Yield(expr) => expr
            .as_ref()
            .is_some_and(|expr| expr_uses_js_arguments(expr)),
        ExprKind::YieldFrom(expr) => expr_uses_js_arguments(expr),
        ExprKind::Call { callee, args, .. } => {
            expr_uses_js_arguments(callee)
                || args.iter().any(|arg| expr_uses_js_arguments(&arg.value))
        }
        ExprKind::New { class, args } => {
            expr_uses_js_arguments(class)
                || args.iter().any(|arg| expr_uses_js_arguments(&arg.value))
        }
        ExprKind::Member { object, .. } => expr_uses_js_arguments(object),
        ExprKind::Index { object, index, .. } => {
            expr_uses_js_arguments(object) || expr_uses_js_arguments(index)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            expr_uses_js_arguments(cond)
                || expr_uses_js_arguments(then)
                || expr_uses_js_arguments(else_)
        }
        ExprKind::Array(elements) => elements.iter().any(|element| {
            element.key.as_ref().is_some_and(expr_uses_js_arguments)
                || expr_uses_js_arguments(&element.value)
        }),
        ExprKind::Object(props) => props.iter().any(|prop| match prop {
            ObjectProperty::KeyValue { value, .. } => expr_uses_js_arguments(value),
            ObjectProperty::Computed { key, value } => {
                expr_uses_js_arguments(key) || expr_uses_js_arguments(value)
            }
            ObjectProperty::Spread(expr) => expr_uses_js_arguments(expr),
            ObjectProperty::Method { value, .. } => stmt_uses_js_arguments(value),
            ObjectProperty::Accessor { value, .. } => stmt_uses_js_arguments(value),
            _ => false,
        }),
        ExprKind::Assign { target, value } => {
            expr_uses_js_arguments(target) || expr_uses_js_arguments(value)
        }
        ExprKind::Lambda { params, body, .. } => {
            params
                .iter()
                .any(|param| param.default.as_ref().is_some_and(expr_uses_js_arguments))
                || match body {
                    LambdaBody::Expr(expr) => expr_uses_js_arguments(expr),
                    LambdaBody::Block(body) => body.iter().any(stmt_uses_js_arguments),
                }
        }
        _ => false,
    }
}

/// Named + wildcard ESM imports a compiled module binds against host
/// Component Model namespaces.
#[derive(Debug, Default, Clone)]
pub struct HostImportMetadata {
    pub named: Vec<HostImportNamed>,
    pub wildcard: Vec<HostWildcardImport>,
}

/// Result of `Compiler::compile_with_imports` — chunks + ESM host-import
/// metadata the VM setup uses to install runtime globals for
/// `read-as-value` and reflective namespace access.
#[derive(Debug, Default)]
pub struct CompileResult {
    pub chunks: Vec<Chunk>,
    pub host_imports: HostImportMetadata,
    /// What the module DECLARED about presenting a UI
    /// ([`vybe_ast::Directives::app_shell`]), carried out so the embedder can
    /// read it without re-walking the tree. `None` states nothing — the
    /// document answers instead.
    pub app_shell: Option<vybe_ast::AppShell>,
}

fn is_php_builtin_constant_name(name: &str) -> bool {
    matches!(
        name,
        "PHP_VERSION"
            | "PHP_VERSION_ID"
            | "PHP_MAJOR_VERSION"
            | "PHP_MINOR_VERSION"
            | "PHP_RELEASE_VERSION"
            | "PHP_OS"
            | "PHP_OS_FAMILY"
            | "PHP_EOL"
            | "PHP_MAXPATHLEN"
            | "PHP_INT_MAX"
            | "PHP_INT_MIN"
            | "PHP_INT_SIZE"
            | "PHP_FLOAT_MAX"
            | "PHP_FLOAT_MIN"
            | "PHP_FLOAT_EPSILON"
            | "PHP_FLOAT_DIG"
            | "M_PI"
            | "M_E"
            | "M_LN2"
            | "M_LN10"
            | "M_LOG2E"
            | "M_LOG10E"
            | "M_SQRT2"
            | "M_SQRT1_2"
            | "INF"
            | "NAN"
            | "STDIN"
            | "STDOUT"
            | "STDERR"
            | "SORT_REGULAR"
            | "SORT_NUMERIC"
            | "SORT_STRING"
            | "SORT_NATURAL"
            | "SORT_ASC"
            | "SORT_DESC"
            | "SORT_FLAG_CASE"
            | "ARRAY_FILTER_USE_KEY"
            | "ARRAY_FILTER_USE_BOTH"
            | "JSON_PRETTY_PRINT"
            | "JSON_UNESCAPED_UNICODE"
            | "JSON_THROW_ON_ERROR"
            | "DIRECTORY_SEPARATOR"
            | "PATH_SEPARATOR"
    )
}

impl Compiler {
    pub(crate) fn strip_global_namespace_prefix(name: &str) -> String {
        let trimmed = name.trim();
        let lower = trimmed.to_ascii_lowercase();
        if lower.starts_with("global::") {
            return trimmed[8..].trim().to_string();
        }
        if lower.starts_with("global.") {
            return trimmed[7..].trim().to_string();
        }
        trimmed.to_string()
    }

    pub fn with_profile(profile: LanguageProfile) -> Self {
        Self {
            last_lambda_body_chunk: None,
            chunks: vec![Chunk::new("<script>")],
            scopes: vec![Scope::new(scope::profile_variable_fold(&profile))],
            current: 0,
            loops: Vec::new(),
            loop_states: Vec::new(),
            label_depth: 0,
            function_label_base: 0,
            line: 1,
            defined_globals: HashSet::new(),
            module_variable_names: HashSet::new(),
            gui_auto_name_counter: 0,
            const_globals: HashSet::new(),
            global_const_values: std::collections::HashMap::new(),
            in_strict: false,
            directives: vec![vybe_ast::Directives::default()],
            canon_section: Vec::new(),
            canon_decls: Vec::new(),
            canon_functypes: Vec::new(),
            canon_valtypes: Vec::new(),
            component_funcs: Vec::new(),
            in_typeof_operand: false,
            want_i32_condition: false,
            gave_i32_condition: false,
            program_lexical_names: HashSet::new(),
            defined_functions: HashSet::default(),
            function_param_modes: HashMap::new(),
            function_param_types: HashMap::new(),
            function_min_arity: HashMap::new(),
            function_signatures: HashMap::new(),
            rest_fixed_arities: BTreeSet::new(),
            function_return_types: HashMap::new(),
            fortran_interface_overloads: HashMap::new(),
            fortran_operator_overloads: HashMap::new(),
            constructor_signatures: HashMap::new(),
            defined_classes: HashSet::new(),
            abstract_classes: HashSet::new(),
            defined_class_methods: HashSet::new(),
            classes_with_late_static_binding: HashSet::new(),
            classes_with_indexer: HashSet::new(),
            classes_with_index_setter: HashSet::new(),
            program_has_getattr: false,
            program_has_setattr: false,
            program_has_index_accessor: false,
            global_type_hints: HashMap::new(),
            enum_members: HashMap::new(),
            module_member_contributors: HashMap::new(),
            enum_value_names: HashMap::new(),
            enum_flags: HashSet::new(),
            reflection_types: HashMap::new(),
            attribute_usage: HashMap::new(),
            reflection_bindings: HashMap::new(),
            case_sensitive: profile.case_sensitive,
            variable_namespace: vybe_runtime::registry::hooks(&profile.name).variable_namespace,
            profile,
            current_func_name: None,
            current_result_slot: None,
            current_ref_out_params: None,
            pending_classes: HashMap::new(),
            normalized_classes: HashMap::new(),
            current_class_slot_keys: HashMap::new(),
            current_class: None,
            current_namespace: None,
            current_class_implicit_self: false,
            current_member_is_static: false,
            static_local_bindings: Vec::new(),
            array_bindings: HashMap::new(),
            pending_label: None,
            with_targets: Vec::new(),
            capture_by_value_vars: Vec::new(),
            capture_locals: HashMap::new(),
            closure_env_names: Vec::new(),
            shared_env_slot: None,
            shared_env_names: Vec::new(),
            promoted_global_cells: HashSet::new(),
            module_addr_taken_globals: HashSet::new(),
            current_addr_taken_locals: HashSet::new(),
            current_atomic_word_locals: HashSet::new(),
            module_atomic_word_globals: HashSet::new(),
            current_closure_captured_locals: HashSet::new(),

            multi_return_functions: HashMap::new(),
            generator_functions: HashSet::new(),
            method_fn_kinds: HashMap::new(),
            js_derived_ctor_ctx: None,
            generator_param_counts: HashMap::new(),
            host_import_bindings: HashMap::new(),
            host_const_bindings: HashMap::new(),
            namespace_import_bindings: HashMap::new(),
            host_namespace_aliases: HashMap::new(),
            host_package_roots: HashMap::new(),
            tree_mounts: HashMap::new(),
            ambient_tree_roots: Vec::new(),
            source_type_aliases: HashMap::new(),
            source_namespace_imports: Vec::new(),
            user_namespace_tree: Default::default(),
            current_module_imports: Vec::new(),
            active_namespaces: None,
            module_exports: HashMap::new(),
            module_value_exports: HashMap::new(),
            active_finally_blocks: Vec::new(),
            finally_joins: Vec::new(),
            fired_finally_indices: Vec::new(),
            catch_depth: 0,
            active_async_try_depth: 0,
            uses_proxy: false,
            bigint_enabled: false,
            js_arguments_bindings: Vec::new(),
        }
    }

    fn current_js_arguments_binding(&self) -> Option<&JsArgumentsBinding> {
        self.js_arguments_bindings
            .last()
            .and_then(|binding| binding.as_ref())
    }

    fn js_arguments_alias_for_name(&self, name: &str) -> Option<(u16, usize)> {
        let binding = self.current_js_arguments_binding()?;
        binding
            .aliased_params
            .get(name)
            .map(|(_, index)| (binding.args_slot, *index))
            .or_else(|| {
                binding
                    .aliased_params
                    .get(&self.canon(name))
                    .map(|(_, index)| (binding.args_slot, *index))
            })
    }

    fn js_arguments_alias_for_index_target(
        &self,
        object: &Expression,
        index: &Expression,
    ) -> Option<(u16, u16, usize)> {
        let binding = self.current_js_arguments_binding()?;
        let ExprKind::Ident(name) = &object.kind else {
            return None;
        };
        if name != "arguments" {
            return None;
        }
        let index = match &index.kind {
            ExprKind::Lit(Literal::Int(value)) if *value >= 0 => *value as usize,
            ExprKind::Lit(Literal::Float(value)) if *value >= 0.0 && value.fract() == 0.0 => {
                *value as usize
            }
            _ => return None,
        };
        let slot = *binding.aliased_indices.get(&index)?;
        Some((binding.args_slot, slot, index))
    }








    /// Pre-populate the module-exports snapshot. Called by the Bundle
    /// before `compile_with_imports` so the Linker can resolve
    /// Adapter-module re-exports during Phase A.
    pub fn with_module_exports(
        mut self,
        module_exports: HashMap<String, HashMap<String, (String, String)>>,
    ) -> Self {
        self.module_exports = module_exports;
        self
    }

    /// Pre-populate the host constant-value snapshot. Called alongside
    /// `with_module_exports` so that `ExportEntry::Value` exports (e.g.
    /// `ecma:math::PI`) are inlined as compile-time constants rather than
    /// routed through `CALL_IMPORT` (which only resolves callable exports).
    ///
    /// The map has the same shape as `flatten_module_exports` but only
    /// includes Value exports: `module → (export_name → Value)`.
    pub fn with_module_value_exports(
        mut self,
        value_exports: HashMap<String, HashMap<String, vybe_runtime::Value>>,
    ) -> Self {
        self.module_value_exports = value_exports;
        self
    }

    /// Compile a module to bytecode chunks. Legacy API — returns just
    /// chunks; import bindings discarded. Callers that need bindings
    /// should use [`Self::compile_with_imports`].
    pub fn compile(self, module: &Module) -> Result<Vec<Chunk>, String> {
        self.compile_with_imports(module).map(|r| r.chunks)
    }

    /// Compile a module, returning chunks plus ESM host-module import
    /// metadata. The caller (typically the VM setup) uses the metadata
    /// to install runtime globals for imported names so `import { X }`
    /// followed by `const f = X` works, and to synthesize Module
    /// Namespace Objects for `import * as ns` reflective access.
    pub fn compile_with_imports(mut self, module: &Module) -> Result<CompileResult, String> {
        // ⛔ THE MODULE'S DECLARED POLICY, NOT THE PROFILE'S. `variable_case`
        // is stated by the walker on `Module.directives`, so a multi-language
        // bundle gets the right answer per UNIT rather than per whichever
        // profile happens to be installed. `variable_fold()` applies the
        // default in ONE place, which is the whole lesson of the flag this
        // replaces: 33 sites had to write `!self.case_sensitive &&` and 23
        // forgot.
        let fold = module.directives.variable_fold();
        self.case_sensitive = fold.is_none();
        // The profile can be swapped after construction (multi-language
        // bundles), so the ROOT scope's folding policy has to follow it — it
        // was built from whatever profile `with_profile` saw.
        for scope in &mut self.scopes {
            scope.fold = fold;
        }
        // This module's declared policy is in force from its first statement.
        // Nothing carries over from a previously compiled module: each unit of
        // a multi-language program is compiled on its own terms, so Pascal
        // never inherits what PHP declared.
        self.directives = vec![module.directives.clone()];
        // The canon section is module-level DATA, not policy and not code: it
        // has no execution position, so it is captured here rather than
        // emitted, and published to the chunks once compilation is done.
        self.canon_section = common::canon::lower_section(&module.canon.defs)?;
        self.canon_decls = module.canon.defs.clone();
        let (fts, vts) = common::canon::lower_types(&module.canon.types)?;
        self.canon_functypes = fts;
        self.canon_valtypes = vts;
        self.component_funcs = module.canon.funcs.clone();
        self.current_module_imports = module.imports.clone();
        // Whether a global holds a pointer cell is a WHOLE-MODULE property, but
        // compilation is one forward pass: `function r(){ global $g; echo $g; }`
        // is emitted before `$g = 1; w($g);` promotes `$g`, so the read would
        // correctly see "not a cell" and emit no autoderef — and then read the
        // cell object raw at runtime.
        //
        // Establish it up front, the same way `collect_addr_taken_idents`
        // already establishes address-taken LOCALS before a body is compiled.
        // Deliberately an OVER-approximation: a name taken by address anywhere
        // marks the global too, even if that site was a local. That is safe
        // because a local binding is consulted first (so it shadows correctly)
        // and `emit_autoderef_pointer_cell` passes a non-reference through
        // untouched — the cost of over-marking is one runtime shape check, and
        // the cost of under-marking is a wrong answer.
        collect_addr_taken_idents(&module.body, &mut self.module_addr_taken_globals);
        // Same forward-pass argument, for ATOMIC places: a module-level name
        // any statement (including one inside a nested function or lambda)
        // targets with an atomic op becomes a shared-memory word at its
        // declaration. It also enters `module_addr_taken_globals` so readers
        // compiled BEFORE the promotion emit the deref, which passes through
        // untouched until the wrap exists.
        collect_atomic_place_idents(&module.body, &mut self.module_atomic_word_globals);
        for name in &self.module_atomic_word_globals {
            self.module_addr_taken_globals.insert(name.clone());
        }
        // The module body IS a scope: locals declared at top level in a
        // script-shaped language resolve as locals, so the local set must be
        // populated here too.
        self.current_atomic_word_locals = self.module_atomic_word_globals.clone();
        // Gated-namespace activation: every import path activates its
        // namespace for builtin resolution (C includes lower to these).
        let mut active = std::collections::HashSet::new();
        for imp in &module.imports {
            let path = match &imp.kind {
                vybe_ast::ImportKind::Simple { path, .. }
                | vybe_ast::ImportKind::Wildcard { path, .. }
                | vybe_ast::ImportKind::Named { path, .. }
                | vybe_ast::ImportKind::Default { path, .. } => path.clone(),
            };
            active.insert(path);
        }
        self.active_namespaces = Some(active);

        // Pre-scan: detect `new Proxy(...)` anywhere in the module so the
        // Member / Index emit sites can route through the proxy dispatcher
        // even when the access appears before the construction in source
        // order. Only for a profile that actually binds `Proxy` to the ECMA
        // proxy surface — otherwise `new Proxy()` is a user class of that
        // name and must not enable proxy dispatch.
        let binds_ecma_proxy = self.profile.esm_defaults.iter().any(|entry| {
            matches!(
                entry,
                vybe_runtime::profile::EsmDefault::Namespace { alias, module }
                    if alias == "Proxy" && module == "ecma:proxy"
            )
        });
        // Pre-scan: can this compile produce a BigInt VALUE at all? Three
        // declarations can — a `[builtin_types] bigint` spelling, a builtin
        // whose emit target reaches the `ecma:bigint` host, or a
        // `Literal::BigInt` in the module itself (a walker only emits one
        // when its language means it). Nothing declared → the bigint arms
        // are unreachable and the `++` path emits no runtime type test.
        self.bigint_enabled = self.bigint_widens_mixes()
            || self.profile.builtins.values().any(|def| {
                matches!(&def.emit,
                    vybe_runtime::profile::BuiltinEmit::HostCall(module, _)
                        if module == "ecma:bigint")
            })
            || {
                let mut probe = module.body.clone();
                probe.iter_mut().any(|stmt| {
                    let mut found = false;
                    stmt.walk_exprs_mut(&mut |e| {
                        found |= matches!(e.kind, ExprKind::Lit(Literal::BigInt(_)));
                    });
                    found
                })
            };

        if binds_ecma_proxy {
            for stmt in &module.body {
                if stmt_uses_proxy(stmt) {
                    self.uses_proxy = true;
                    break;
                }
            }
        }

        // ── Phase A: Link ──────────────────────────────────────────────
        //
        // ECMA-262 §16.2.1.5 adapted for Vybe. Populates the three
        // resolver maps from profile defaults + user imports so every
        // downstream emit site consults a single source of truth.
        // Profile defaults seed the ambient namespaces (e.g. JS
        // `console`, VB `System`); user `import { X } from "wasi:foo"`
        // statements shadow them on key collision, per §16.2
        // lexical-over-module-scope.
        //
        // Host synthetic modules (`wasi:*`, `wasm:*`, `vybe:*`) are
        // linked immediately — they're leaves with no code. User
        // `.wasm` / source-file imports continue to resolve at Bundle
        // load time in a separate step.
        self.link(module);

        // .NET BCL classes no longer emit a per-class constructor prelude:
        // control/value/drawing types resolve through the component descriptor
        // (properties/methods) and either the GUI-direct path or a
        // descriptor constructor (construction), and user `class Form1 : Form`
        // base construction lowers via `try_emit_framework_control_base`. See
        // the retired `registry::dotnet` (keep-set went empty once the drawing
        // Body methods migrated to `MethodBody::Common`).

        // Ambient platform roots — the SAME `type_scopes` the language
        // declares for member resolution. Mounting them makes unqualified
        // platform names (`Scaffold(...)`, `TForm`, `Button`) resolve to their
        // tree `Type` and construct through the one common-resolver path. One
        // declaration, not a second list and not a per-platform hook.
        for root in self.profile.namespaces.type_scopes.clone() {
            self.mount_ambient_root(&root);
        }

        // Pre-pass: merge `Partial Class` declarations sharing the same name.
        // Walker-driven: only runs when at least one ClassDecl in the module
        // is flagged `modifiers.is_partial = true` (VB/C# walkers set this
        // on `Partial Class`; other languages leave it false and skip the
        // merge entirely). After merging, the body has exactly one ClassDecl
        // per class name with all fields/methods pooled together.
        let has_partial = module.body.iter().any(
            |s| matches!(&s.kind, StmtKind::ClassDecl { modifiers, .. } if modifiers.is_partial),
        );
        let mut merged_body = if has_partial {
            merge_partial_classes(&module.body, self.case_sensitive)
        } else {
            module.body.clone()
        };

        // Builtins declared with `pass_by = [...]` in the profile enter the tree
        // here, as an ordinary interface — `referenceplan.md` §10j.1. The profile
        // row is source text; this declaration is the fact. An interface is the
        // right carrier precisely because it can never become a call target
        // (`StmtKind::InterfaceDecl` compiles to a no-op), so `bindParam` still
        // routes through `[value_methods]` to its adapter — this DESCRIBES the
        // callee, it does not implement it.
        //
        // Appended LAST so a same-named declaration written in source registers
        // first and wins: `register_interface_method_signatures` is `or_insert`,
        // so the builtin fills in only where the program itself said nothing.
        if !self.profile.builtin_signatures.is_empty() {
            merged_body.push(Statement::new(StmtKind::InterfaceDecl {
                // Reserved spelling — a user type of this name is not expressible
                // in any surface language here, so registering it as a type name
                // cannot shadow a real class (php's own `PDOStatement` would).
                name: "__vybe_builtin_signatures".to_string(),
                parents: Vec::new(),
                members: self.profile.builtin_signatures.clone(),
                decorators: Vec::new(),
            }));
        }
        let merged_body = merged_body;

        self.predeclare_type_names(&merged_body, None);
        self.collect_module_variable_names(&merged_body);
        self.collect_reflection_metadata(&merged_body);

        // Multi-value pre-scan: any function whose every explicit `Return`
        // is a same-arity tuple literal is a candidate for the WASM
        // multi-value ABI. We only opt in when the language profile
        // requests it — other languages keep tuple-as-heap-object semantics.
        if self.profile.multi_value_tuple_returns {
            self.collect_multi_return_functions(&merged_body);
        }

        // Declaration pass, phases 2 and 3: fold declared augmentations across
        // ALL normalized classes, THEN register member surfaces. Order matters
        // — a contributed member missing from registration reintroduces the
        // order-dependence bug (flexclassplan.md §3a, §4c).
        self.record_platform_bases();
        self.apply_class_augmentations()?;
        // Every declaration is known here — including the ones augmentation
        // just contributed — so this is the earliest point a field's declared
        // type can be resolved, and the only one where it need happen once.
        self.resolve_field_value_types();
        self.predeclare_class_surfaces();
        self.predeclare_function_names(&merged_body);
        self.predeclare_interface_signatures_in_body(&merged_body);

        // Pre-collect every rest-parameter arity in the program so call-site
        // rest packing is emitted even when the callee (e.g. a `const f =
        // (...xs) => ...` arrow) is declared after a hoisted function body that
        // calls it. Without this, those calls silently drop arguments.
        {
            let mut rest_arities = Vec::new();
            crate::ast::collect_rest_param_arities(&merged_body, &mut rest_arities);
            for arity in rest_arities {
                self.rest_fixed_arities.insert(arity);
            }
        }

        for stmt in &merged_body {
            if matches!(&stmt.kind, StmtKind::FunctionDecl { .. }) {
                self.compile_stmt(stmt)?;
            }
        }

        // ECMA-262 §11.2.1: Detect top-level "use strict" directive prologue
        // so strict mode rules apply to module-level code
        if self.profile.ecma_strict_mode && Self::stmts_have_use_strict_directive(&merged_body) {
            self.in_strict = true;
        }

        for stmt in &merged_body {
            if matches!(&stmt.kind, StmtKind::FunctionDecl { .. }) {
                continue;
            }
            self.compile_stmt(stmt)?;
        }

        // Auto-call entry point if defined
        if let Some(ref ep) = self.profile.entry_point.clone() {
            let has_ep = self.defined_globals.contains(ep)
                || (!self.case_sensitive
                    && self
                        .defined_globals
                        .iter()
                        .any(|g| g.eq_ignore_ascii_case(ep)));
            if has_ep {
                self.emit_var_get(ep);
                self.emit_direct_callable_invoke(0);
                self.emit(Op::DROP);
            } else {
                // C#-style entry: `static void Main()` lives as a static
                // method stamped on the class object (class normalization),
                // not as a bare global. Find the class declaring it and
                // invoke `<Class>.<ep>()`.
                let ep_canon = self.canon(ep);
                let host_class = self
                    .pending_classes
                    .iter()
                    .find(|(_, pc)| pc.static_method_names.iter().any(|m| *m == ep_canon))
                    .map(|(name, _)| name.clone());
                if let Some(class_name) = host_class {
                    self.emit_var_get(&class_name);
                    let key = self.str_const(&ep_canon);
                    self.emit_struct_field_op(Op::STRUCT_GET, 0, key);
                    self.emit_direct_callable_invoke(0);
                    self.emit(Op::DROP);
                }
            }
        }

        // Any output buffer still open when the script ends is FLUSHED, not
        // discarded — a template that forgets its `ob_end_flush()` still
        // renders. Costs one null check on `__vybe_ob_stack` for a program that
        // never buffered, since the stack global is only created by the first
        // `ob_start`.
        // Shutdown callbacks run BEFORE the final flush, exactly as php does:
        // a handler that echoes must still land in an open buffer. Ungated —
        // the runner is a null check on `__php_shutdown_fns`, which only exists
        // once something registered one, so a program in any other language
        // pays the same single check the buffer flush above already costs.
        let saved = self.current;
        self.current = 0;
        self.emit_php_run_shutdown_fns();
        self.current = saved;

        let line = self.line;
        common::io::emit_ob_flush_all(&mut self.chunks, 0, line);

        // A module that recorded an exit status hands it to the process HERE —
        // after the entry point has run and after the flush, so no output is
        // lost. Placing it with the module BODY instead ended the run before
        // the entry point was even called: `main` never executed and every
        // test passed vacuously. Gated on the module actually declaring the
        // global, which is a property of the module rather than a language
        // name, so it costs other languages nothing.
        if self.defined_globals.contains("__c_exit_status") {
            self.emit_var_get("__c_exit_status");
            let exit_idx = self.import("wasi:cli/exit", "exit-with-code");
            self.emit_host_call(exit_idx, 1);
        }

        self.emit_null();
        self.emit(Op::RETURN);
        // Take the max of the scope's highest slot and whatever raw local
        // slots compiler_common helpers (e.g. `invoke::emit_invoke_method`)
        // reserved directly on the chunk — those bypass `Scope` but still
        // need the VM to reserve slots at call-frame entry.
        let ns = self.scope().next_slot;
        self.chunks[0].finalize_local_count(ns);
        self.chunks[0].local_names = self.scope().defined_names.clone();
        // Skip helper linking when compiling runtime helper source.
        // Re-running finalization here would call back into helper
        // compilation and recurse through `Compiler::compile`.
        // and recurse forever. Cheap thread-local guard since polyfill
        // compilation is single-threaded at vybex build time.
        if !crate::primitives::polyfills::is_compiling_runtime_helper() {
            // The exclusion list is profile data — a language that supplies
            // its own implementation of a shared helper names it there, rather
            // than the shared crate carrying a per-language table.
            let excluded: Vec<String> = self.profile.excluded_runtime_helpers.clone();
            if excluded.is_empty() {
                common::bundle::finalize_with_runtime_helpers(&mut self.chunks);
            } else {
                let excluded_refs: Vec<&str> = excluded.iter().map(String::as_str).collect();
                common::bundle::finalize_with_runtime_helpers_excluding(
                    &mut self.chunks,
                    &excluded_refs,
                );
            }
        }
        Self::normalize_import_table(&mut self.chunks);
        common::globals::declare_free_globals(&mut self.chunks);
        // Assign every global a real index over `global_imports ++ defined`
        // and rewrite the operands into it. Must follow the line above,
        // which decides the import half.
        common::globals::normalize_global_table(&mut self.chunks);
        // The canon section, published to every chunk on the same principle as
        // the global index space above: a chunk carrying a canonidx must be
        // able to say what that index means.
        let mut canon = std::mem::take(&mut self.canon_section);
        let decls = std::mem::take(&mut self.canon_decls);
        common::canon::resolve_core_export_callees(&mut self.chunks, &decls, &mut canon)?;
        common::canon::install_canon_section(&mut self.chunks, &canon);
        let fts = std::mem::take(&mut self.canon_functypes);
        let vts = std::mem::take(&mut self.canon_valtypes);
        let cfuncs = std::mem::take(&mut self.component_funcs);
        common::canon::install_type_space(&mut self.chunks, &fts, &vts, &cfuncs);
        let host_imports = self.collected_host_imports();
        // Frame 0 is the module's own declaration (installed above from
        // `module.directives`); an in-source `Directive` with `Module` scope
        // writes through to it, so reading the base frame answers for the whole
        // unit however it was stated.
        let app_shell = self.directives.first().and_then(|d| d.app_shell);
        Ok(CompileResult {
            chunks: self.chunks,
            host_imports,
            app_shell,
        })
    }
}

/// Merge `Partial Class` declarations sharing the same name into one.
/// Used by VB and C# (any language whose profile sets `partial_classes = true`).
///
/// Detect statements of the form `Me.__<identifier> = <literal>` so the
/// child-class JS/VB/Pascal-style ctor flow can include them in the
/// "preamble" that runs before method binding. The walker normalization
/// for VB injects `Me.__control_name = "<lower class name>"` immediately
/// after the implicit `MyBase.New()` so the canonical control name is
/// stamped before any property writes mirror to gui state.
/// Return `Some(N)` if every explicit `Return` in `body` carries an
/// `ExprKind::Tuple` literal of the same N elements. Returns `None` if
/// the body has no explicit returns, a return with no value, a return
/// with a non-tuple value, or tuples of mismatched arity. Recurses into
/// nested control-flow bodies (if/loops/try) but **not** into nested
/// function declarations — those are separate scopes.
fn uniform_tuple_return_arity(body: &[Statement]) -> Option<u8> {
    let mut arity: Option<u8> = None;
    let mut saw_any = false;
    fn walk(stmts: &[Statement], arity: &mut Option<u8>, saw_any: &mut bool) -> bool {
        for s in stmts {
            match &s.kind {
                StmtKind::Return(Some(expr)) => {
                    *saw_any = true;
                    if let ExprKind::Tuple(elems) = &expr.kind {
                        let n = elems.len();
                        if n < 2 || n > 255 {
                            return false;
                        }
                        match arity {
                            None => *arity = Some(n as u8),
                            Some(a) if *a as usize == n => {}
                            _ => return false,
                        }
                    } else {
                        return false;
                    }
                }
                StmtKind::Return(None) => {
                    *saw_any = true;
                    return false;
                }
                StmtKind::If {
                    then_body,
                    elifs,
                    else_body,
                    ..
                } => {
                    if !walk(then_body, arity, saw_any) {
                        return false;
                    }
                    for (_, b) in elifs {
                        if !walk(b, arity, saw_any) {
                            return false;
                        }
                    }
                    if let Some(b) = else_body {
                        if !walk(b, arity, saw_any) {
                            return false;
                        }
                    }
                }
                StmtKind::While {
                    body, else_body, ..
                }
                | StmtKind::ForIn {
                    body, else_body, ..
                } => {
                    if !walk(body, arity, saw_any) {
                        return false;
                    }
                    if let Some(b) = else_body {
                        if !walk(b, arity, saw_any) {
                            return false;
                        }
                    }
                }
                StmtKind::For { body, .. }
                | StmtKind::DoWhile { body, .. }
                | StmtKind::With { body, .. }
                | StmtKind::Using { body, .. } => {
                    if !walk(body, arity, saw_any) {
                        return false;
                    }
                }
                StmtKind::Try {
                    body,
                    catches,
                    else_body,
                    finally,
                } => {
                    if !walk(body, arity, saw_any) {
                        return false;
                    }
                    for c in catches {
                        if !walk(&c.body, arity, saw_any) {
                            return false;
                        }
                    }
                    if let Some(b) = else_body {
                        if !walk(b, arity, saw_any) {
                            return false;
                        }
                    }
                    if let Some(b) = finally {
                        if !walk(b, arity, saw_any) {
                            return false;
                        }
                    }
                }
                StmtKind::Block(b) => {
                    if !walk(b, arity, saw_any) {
                        return false;
                    }
                }
                StmtKind::Labeled { body, .. } => {
                    if !walk(std::slice::from_ref(body.as_ref()), arity, saw_any) {
                        return false;
                    }
                }
                // Nested function / class declarations are their own scopes.
                StmtKind::FunctionDecl { .. } | StmtKind::ClassDecl { .. } => {}
                _ => {}
            }
        }
        true
    }
    if !walk(body, &mut arity, &mut saw_any) {
        return None;
    }
    if saw_any { arity } else { None }
}

fn is_hoisted_deconstruction_block(stmts: &[Statement]) -> bool {
    let call_stmt = match stmts {
        [
            Statement {
                kind: StmtKind::Expr(expr),
                ..
            },
        ] => expr,
        [
            Statement {
                kind: StmtKind::VarDecl { .. },
                ..
            },
            Statement {
                kind: StmtKind::Expr(expr),
                ..
            },
        ] => expr,
        _ => return false,
    };

    let ExprKind::Call { callee, args, .. } = &call_stmt.kind else {
        return false;
    };
    let ExprKind::Member { field, .. } = &callee.kind else {
        return false;
    };

    field.eq_ignore_ascii_case("Deconstruct") && args.iter().all(|arg| arg.by_ref)
}

fn is_identity_stamp(stmt: &Statement) -> bool {
    if let StmtKind::Assign { targets, .. } = &stmt.kind {
        if targets.len() == 1 {
            if let ExprKind::Member { object, field, .. } = &targets[0].kind {
                let obj_is_self = matches!(&object.kind, ExprKind::This | ExprKind::Super)
                    || matches!(
                        &object.kind,
                        ExprKind::Ident(n) if matches!(n.to_lowercase().as_str(), "me" | "this" | "self")
                    );
                if obj_is_self && field.starts_with("__") {
                    return true;
                }
            }
        }
    }
    false
}

/// Check whether a constructor body already contains a call to the given
/// method (case-insensitive).  Matches `Me.Method()`, `this.Method()`,
/// and bare `Method()` call shapes — which is how all walkers emit it.
fn body_calls_method(body: &[Statement], method_name: &str) -> bool {
    body.iter().any(|s| {
        if let StmtKind::Expr(expr) = &s.kind {
            if let ExprKind::Call { callee, .. } = &expr.kind {
                if let ExprKind::Member { field, .. } = &callee.kind {
                    return field.eq_ignore_ascii_case(method_name);
                }
                if let ExprKind::Ident(name) = &callee.kind {
                    return name.eq_ignore_ascii_case(method_name);
                }
            }
        }
        false
    })
}

/// Check whether a constructor body contains a super/base call.
/// Matches `SuperCall { .. }` (VB/Pascal) and `Call { callee: Super }` (JS).
fn body_has_super_call(body: &[Statement]) -> bool {
    body.iter().any(|s| {
        if let StmtKind::Expr(e) = &s.kind {
            match &e.kind {
                ExprKind::SuperCall { .. } => true,
                ExprKind::Call { callee, .. } => matches!(callee.kind, ExprKind::Super),
                _ => false,
            }
        } else {
            false
        }
    })
}

fn body_has_identity_stamp(body: &[Statement]) -> bool {
    body.iter().any(is_identity_stamp)
}

/// The first declaration of a class name keeps its position in the body and
/// receives all subsequent partials' members appended in source order.
/// Subsequent partials are removed. Non-class statements pass through
/// unchanged.
///
/// This is intentionally a pure AST transform — no compiler state, no
/// language-specific quirks. The merged class compiles via the standard
/// compile_class path.
fn merge_partial_classes(body: &[Statement], case_sensitive: bool) -> Vec<Statement> {
    let key = |name: &str| -> String {
        if case_sensitive {
            name.to_string()
        } else {
            name.to_lowercase()
        }
    };

    // First pass: collect (name → first_index)
    let mut first_index: std::collections::HashMap<String, usize> =
        std::collections::HashMap::new();
    for (i, stmt) in body.iter().enumerate() {
        if let StmtKind::ClassDecl { name, .. } = &stmt.kind {
            first_index.entry(key(name)).or_insert(i);
        }
    }

    // Second pass: build merged body. For each class, only keep the first
    // declaration but append later partials' members into it.
    let mut result: Vec<Statement> = Vec::with_capacity(body.len());
    let mut emitted: std::collections::HashSet<String> = std::collections::HashSet::new();

    for (i, stmt) in body.iter().enumerate() {
        match &stmt.kind {
            StmtKind::ClassDecl { name, .. } => {
                let k = key(name);
                if first_index.get(&k) != Some(&i) {
                    // Not the first declaration — skip; its members were
                    // (or will be) merged into the first one.
                    continue;
                }
                if emitted.contains(&k) {
                    continue;
                }
                emitted.insert(k.clone());

                // Clone the first declaration; we'll mutate its members.
                let mut merged = stmt.clone();
                if let StmtKind::ClassDecl {
                    members: m,
                    parents: p,
                    interfaces: iface,
                    ..
                } = &mut merged.kind
                {
                    // Append members from every later declaration of this name.
                    // Skip duplicate Constructors — the VB walker's
                    // `inject_implicit_mybase_new` synthesizes a 2-stmt
                    // Constructor for any `Partial Class` with an
                    // `Inherits` clause that doesn't declare its own
                    // `Sub New`. If another partial DOES declare a real
                    // Sub New (with injected AddHandlers, user body, etc.),
                    // appending the synthesized stub would duplicate the
                    // ClassMember::Constructor entry. `normalize_class`
                    // then iterates and the last one wins, silently
                    // dropping every real ctor statement (including the
                    // injected Handles → AddHandler bindings). Keep the
                    // first Constructor; discard any later partial's
                    // Constructor clone.
                    let mut has_ctor = m
                        .iter()
                        .any(|mb| matches!(mb, ClassMember::Constructor { .. }));
                    for later in body.iter().skip(i + 1) {
                        if let StmtKind::ClassDecl {
                            name: ln,
                            members: lm,
                            parents: lp,
                            interfaces: li,
                            ..
                        } = &later.kind
                        {
                            if key(ln) == k {
                                for lmem in lm {
                                    if matches!(lmem, ClassMember::Constructor { .. }) {
                                        if has_ctor {
                                            continue;
                                        }
                                        has_ctor = true;
                                    }
                                    m.push(lmem.clone());
                                }
                                // Merge unique parents / interfaces
                                for parent in lp {
                                    if !p.iter().any(|existing| key(existing) == key(parent)) {
                                        p.push(parent.clone());
                                    }
                                }
                                for it in li {
                                    if !iface.iter().any(|e| key(e) == key(it)) {
                                        iface.push(it.clone());
                                    }
                                }
                            }
                        }
                    }
                }
                result.push(merged);
            }
            _ => result.push(stmt.clone()),
        }
    }

    result
}

/// `true` when an ESM import specifier points at a host Component Model
/// namespace rather than a source file on disk. The Linker treats these
/// as Synthetic Module Record imports — no filesystem resolution.
fn is_host_specifier(path: &str) -> bool {
    path.starts_with("wasi:")
        || path.starts_with("wasm:")
        || path.starts_with("vybe:")
        // Node.js built-ins under the `node:` prefix. Only enumerate
        // modules with a real host implementation in
        // `crates/vybe_host/src/node/*.rs`. `node:http` is still
        // served by a JS adapter in `crates/vybex/src/adapters/node/`
        // and must fall through so the linker walks the adapter chain.
        // Add an entry as each new `node:*` host module lands; remove
        // its matching adapter at the same time.
        || matches!(path,
            "node:fs"
            | "node:os"
            | "node:path"
            | "node:process"
            | "node:child_process"
            | "node:crypto")
}

/// Resolve a shared *platform* emit dispatcher by its `common:<prefix>.*`
/// prefix, through the plugin registry.
///
/// This was a hardcoded `match prefix { "dotnet" => …, "libc" => … }` — the
/// same name-check antipattern as `profile.name == "<lang>"`, one layer up, and
/// it forced `vybe_compiler` to depend on the platform crates at COMPILE time.
/// Platforms are plugins: they register themselves (see each platform's
/// `register_platform()`), so the compiler never names one and a platform can
/// become a dylib.
pub fn platform_emit_dispatch(prefix: &str) -> Option<crate::languages::EmitDispatch> {
    vybe_runtime::registry::platform_emit_dispatch_for(prefix)
}
