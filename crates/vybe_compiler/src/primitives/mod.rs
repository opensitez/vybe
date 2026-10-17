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

#[cfg(test)]
std::thread_local! {
    static COPY_MODULE_BODY_BASELINE: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
}

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
pub mod bundle;
pub mod callable;
pub mod canon;
pub mod canon_marshal;
pub mod clone; // what it means to COPY a value — records, collections, arguments
pub mod codepoints;
pub mod collections;
pub mod complex;
pub mod config;
pub mod convert;
pub mod csv;
pub mod datetime;
pub mod delegates;
pub mod dict;
pub mod dispatch;
pub mod dynamic_invoke;
pub mod dynamic_symbols;
pub mod enum_lowering;
pub mod errors;
pub mod exact_integer;
pub mod fs_path;
pub mod functions;
pub mod generators;
pub mod generics;
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
pub mod memory_compare;
pub mod memory_find;
pub mod memory_scan;
pub mod memory_span;
pub mod multivalue;
pub mod object;
pub mod ops;
pub mod packing;
pub mod paths;
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
pub mod helper_templates;
pub mod argument_emission;
mod calls;
mod immutable_calls;
mod body_symbols;
pub mod canonical;
mod case_insensitive_collections;
pub mod channels;
mod class_augmentation;
// ⛔ PUBLIC SO A LANGUAGE EMITTER CAN BIND A RECEIVER THE NORMALIZED WAY.
// The shared bind has to be reachable from where receivers are actually
// bound, or a frontend re-invents a channel of its own.
pub mod class_context;
pub mod deferred_class_bodies;
pub mod class_normalize; // cross-language class normalisation (was crate::common::classes)
pub mod class_slots;
pub mod classes;
mod closure_analysis;
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
mod user_tree_cache;
mod operators;
mod overloads;
pub mod prototypes;
pub mod records;
mod reference_analysis;
mod atomic_preparation;
mod source_feature_preparation;
mod declaration_preparation;
mod function_preparation;
mod unit_preparation;
pub mod references;
pub mod reflection;
mod resolver;
mod scope;
pub mod slices;
mod statements;
mod statement_stream;
pub use statement_stream::{StatementEmission, StatementProgress};
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
use std::time::{Duration, Instant};
use vybe_runtime::opcode::Op;
use vybe_runtime::registry::emitter_results::{EmitOutcome, EmitRequest, ResultRepresentation};
use vybe_runtime::{Chunk, Value};

fn compiler_timing_path() -> Option<String> {
    let value = std::env::var("VYBE_COMPILER_DEBUG_TIMINGS").ok()?;
    if value.is_empty() {
        None
    } else if value == "1" {
        Some("/private/tmp/vybe_compiler_timings.log".to_string())
    } else {
        Some(value)
    }
}

fn write_compiler_timing(path: &str, label: &str, elapsed: Duration) {
    use std::io::Write;

    if let Ok(mut file) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(path)
    {
        // Keep concurrent worker records whole: formatting into File can issue
        // separate writes for the label and duration, corrupting phase totals.
        let record = format!("{label} {elapsed:?}\n");
        let _ = file.write_all(record.as_bytes());
    }
}

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
    /// The declared type hint, IN THE SPELLING THE SOURCE WROTE.
    ///
    /// ⛔ Stored folded until 2026-08-27, which destroyed the only copy of the
    /// declared case. The namespace tree is keyed by the registrar's spelling
    /// and matches EXACTLY for a language that does not fold, so a receiver
    /// typed from this map asked the tree for `arraylist` and missed
    /// `ArrayList` — invisible in every folding language, fatal in the others,
    /// exactly as `tree_type_key` warns. Folding here also cannot be undone;
    /// folding at a READ site can.
    ///
    /// Every consumer that compares against lowercase literals folds on its
    /// own (`normalize_type_hint`, `builtin_types::classify_with`), so they are
    /// unaffected. See `documentation/casesensitivityplan.md` — this is the
    /// field-hint leg of the lowercase-canonical shortcut.
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

// Signature names are immutable declaration metadata, shared across overload
// selection/call staging copies. The empty declaration shape needs no fresh
// allocation at each registration.
fn empty_signature_param_names() -> std::sync::Arc<[String]> {
    static EMPTY: std::sync::OnceLock<std::sync::Arc<[String]>> = std::sync::OnceLock::new();
    EMPTY.get_or_init(|| std::sync::Arc::from([])).clone()
}

#[derive(Debug, Clone)]
pub(crate) struct CallSignature {
    param_names: std::sync::Arc<[String]>,
    /// The declared default expression for each param (positionally aligned
    /// with `param_names`), so named-arg reordering can fill an OMITTED
    /// optional param with its real default instead of `null`.
    param_defaults: std::sync::Arc<[Option<Expression>]>,
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
            param_names: if params.is_empty() {
                empty_signature_param_names()
            } else {
                params.iter().map(|param| param.name.trim_start_matches('$').to_string()).collect()
            },
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

    /// The signature a TREE LEAF declares, when it declares parameter names.
    /// `None` for a leaf with any unnamed parameter: such a leaf binds
    /// positionally, exactly as it did before leaves carried signatures, so
    /// a staging registration never rejects a call that worked. An optional
    /// parameter is the Component Model's `option<t>`; there is no default
    /// expression to fill, so an omitted one lands as `null` — its `none`.
    pub(crate) fn from_leaf(sig: &vybe_runtime::component::FuncSig) -> Option<Self> {
        if !sig.declares_names() {
            return None;
        }
        Some(Self {
            param_names: if sig.params.is_empty() {
                empty_signature_param_names()
            } else {
                sig.params.iter().map(|param| param.name.clone()).collect()
            },
            param_defaults: vec![None; sig.params.len()].into(),
            min_arity: sig.min_arity(),
            has_rest: false,
            has_kwargs: false,
            rest_index: None,
        })
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
    /// The declared spelling; the runtime key is canonicalised, which a
    /// case-insensitive language lowercases.
    pub display_name: String,
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
    declaration_scope_depth: usize,
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
/// `--dump-classes`. A process-wide switch rather than a `Compiler` field
/// because `Bundle` is built by struct literal at 43 sites and a diagnostic
/// must not churn them; the CLI sets it once, before any compile.
static DUMP_CLASSES: std::sync::atomic::AtomicBool = std::sync::atomic::AtomicBool::new(false);

pub fn set_dump_classes(on: bool) {
    DUMP_CLASSES.store(on, std::sync::atomic::Ordering::Relaxed);
}

pub(crate) fn dump_classes_enabled() -> bool {
    DUMP_CLASSES.load(std::sync::atomic::Ordering::Relaxed)
}

/// `(none)` rather than an empty line, so an EMPTY table is visibly empty —
/// the distinction the class dump exists to make.
fn join_or_none(items: &[String]) -> String {
    if items.is_empty() {
        "(none)".to_string()
    } else {
        items.join(", ")
    }
}

/// Declaration storage shared only while read-only publication runs. Ordinary
/// declaration updates have a unique owner; copy-on-write protects snapshots if
/// a later phase ever mutates the registry while one is retained.
#[derive(Default)]
struct NormalizedClasses(
    std::sync::Arc<HashMap<String, vybe_ast::class_normalize::NormalClass>>,
    std::sync::Mutex<Option<PrivateMemberIndex>>,
);

#[derive(Clone)]
struct PrivateMemberIndex {
    case_sensitive: bool,
    namespace_body: Option<usize>,
    members: std::sync::Arc<HashSet<String>>,
}

impl Clone for NormalizedClasses {
    fn clone(&self) -> Self {
        Self(self.0.clone(), std::sync::Mutex::new(
            self.1.lock().unwrap_or_else(|poison| poison.into_inner()).clone(),
        ))
    }
}

impl std::ops::Deref for NormalizedClasses {
    type Target = HashMap<String, vybe_ast::class_normalize::NormalClass>;
    fn deref(&self) -> &Self::Target { &self.0 }
}
impl std::ops::DerefMut for NormalizedClasses {
    fn deref_mut(&mut self) -> &mut Self::Target {
        // All declaration writes, including in-place member/augmentation
        // updates, pass through this boundary. Snapshots retain their own
        // cache and declaration view; only this owner's index is invalidated.
        self.1.get_mut().unwrap_or_else(|poison| poison.into_inner()).take();
        std::sync::Arc::make_mut(&mut self.0)
    }
}

#[cfg(test)]
mod normalized_class_index_tests {
    use super::{NormalizedClasses, PrivateMemberIndex};
    use std::sync::Arc;

    #[test]
    fn private_index_snapshots_invalidate_on_each_declaration_write() {
        fn send_sync<T: Send + Sync>() {}
        send_sync::<NormalizedClasses>();
        let mut classes = NormalizedClasses::default();
        classes.insert("owner".into(), vybe_ast::class_normalize::NormalClass::default());
        *classes.1.lock().unwrap() = Some(PrivateMemberIndex {
            case_sensitive: true,
            namespace_body: None,
            members: Arc::new(["secret".to_string()].into_iter().collect()),
        });
        let snapshot = classes.clone();
        assert!(Arc::ptr_eq(&classes.0, &snapshot.0));
        classes.get_mut("owner").unwrap().name = "Changed".into();
        assert!(classes.1.lock().unwrap().is_none());
        assert!(snapshot.1.lock().unwrap().as_ref().unwrap().members.contains("secret"));
        assert_ne!(classes.get("owner").unwrap().name, snapshot.get("owner").unwrap().name);
        assert!(!Arc::ptr_eq(&classes.0, &snapshot.0));
        *classes.1.lock().unwrap() = snapshot.1.lock().unwrap().clone();
        classes.remove("owner");
        assert!(classes.1.lock().unwrap().is_none());
        assert!(snapshot.contains_key("owner"));
    }
}

/// Diagnostic environment is captured per consumed module compilation.
/// Changes between compilations remain visible; no process-global cache.
struct StatementTraceConfig { path: String, function: String }

pub struct Compiler {
    argument_emitters: argument_emission::Selection,
    helper_parser: Option<vybe_runtime::registry::ParseFn>,
    helper_templates: helper_templates::Owner,
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
    pub(crate) defined_globals: body_symbols::BodySymbols,
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
    /// Named core callees kept until their emitted chunks have indices.
    core_export_callees: Vec<common::canon::CoreExportCallee>,
    canon_functypes: Vec<Option<vybe_runtime::canon_def::CanonFuncType>>,
    canon_valtypes: Vec<Option<vybe_runtime::component::ValType>>,
    component_funcs: Vec<Option<u32>>,
    component_func_links: Vec<vybe_runtime::canon_def::ComponentFuncLink>,
    component_instance_links: Option<Arc<vybe_runtime::component_instance_link::InstanceLinkMetadata>>,
    /// True while compiling the operand of a `typeof`. `typeof undeclaredName`
    /// must evaluate to `"undefined"`, never throw — so the unresolvable-binding
    /// ReferenceError in `emit_var_get` is suppressed in this context.
    in_typeof_operand: bool,
    /// A CONDITION is being compiled and would rather have a raw i32 than a
    /// boxed `Bool`. Set only by `compile_condition_to_i32`, and **taken** at
    /// the top of `compile_expr` — so it reaches the outermost operator and
    /// can never leak into a nested operand, where the boxed value is what the
    /// surrounding expression is entitled to.
    has_discarded_builtin_emits: bool,
    want_i32_condition: bool,
    /// The request above was honoured: the value on the stack is an i32.
    /// Written by `emit_i32_to_bool_or_report` or a terminal typed builtin
    /// dispatch reporting `RawI32Boolean`. Expression boundaries isolate
    /// nested reports, and condition boundaries restore their caller's state.
    /// In the comparison path, reporting and
    /// skipping the boxing are the same statement — so "the boxing was
    /// skipped" and "the stack holds an i32" cannot disagree. The dangerous
    /// direction is a report without a skip: `BR_IF` accepts a `Bool` happily,
    /// so the loop would key on the wrong thing in silence.
    gave_i32_condition: bool,
    /// Set only by a resolved terminal emitter for an explicit word request.
    gave_i32_word: bool,
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
    pub(crate) defined_functions: body_symbols::FunctionSymbols,
    immutable_function_bindings: immutable_calls::Bindings,
    /// Whether this compiled module owns whole-program end-of-module hooks.
    ///
    /// Normal entry modules do. Dynamic fragments/includes do not: they run in
    /// the caller's request and must not fire request-end hooks or flush every
    /// output buffer just because one included file reached EOF.
    run_module_end_hooks: bool,
    // These private symbol tables are queried by key, never emitted in hash
    // iteration order. Share the fast hasher already used for function identity.
    function_param_modes: HashMap<String, Vec<PassBy>, vybe_runtime::chunk::FxBuildHasher>,
    function_param_types: HashMap<String, Vec<Option<String>>, vybe_runtime::chunk::FxBuildHasher>,
    function_min_arity: HashMap<String, usize, vybe_runtime::chunk::FxBuildHasher>,
    function_signatures: HashMap<String, Vec<CallSignature>, vybe_runtime::chunk::FxBuildHasher>,
    rest_fixed_arities: BTreeSet<u8>,
    function_return_types: HashMap<String, String, vybe_runtime::chunk::FxBuildHasher>,
    fortran_interface_overloads: HashMap<String, Vec<FortranInterfaceOverload>>,
    fortran_operator_overloads: HashMap<String, Vec<FortranInterfaceOverload>>,
    constructor_signatures: HashMap<String, Vec<CallSignature>>,
    pub(crate) defined_classes: body_symbols::BodySymbols,
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
    /// `NormalClass::cooperative_super`, recorded when the class is compiled,
    /// for the sites that hold only the class NAME. Same two-part shape as
    /// `classes_with_late_static_binding` above: read the declaration directly
    /// where the class is in scope, and this set where it is not.
    pub(crate) classes_with_cooperative_super: HashSet<String>,
    /// Classes declaring an index operator (`operator []` / `__getitem__`).
    /// Indexing one of these is a method call, not a key lookup — resolved
    /// from the receiver's static type so arrays, dicts and strings keep the
    /// plain index path with no runtime probe.
    pub(crate) classes_with_indexer: HashSet<String>,
    /// Classes declaring an index *setter* (`operator []=` / `__setitem__`).
    /// Kept apart from `classes_with_indexer` — a class may define either
    /// half on its own.
    pub(crate) classes_with_index_setter: HashSet<String>,
    /// Protocol slots published by any class/struct normalized in this program.
    ///
    /// Rich operator dispatch is not a language-name property and it is not the
    /// default for a static program. If no compiled type publishes `Eq`, `Lt`,
    /// `Add`, etc., probing every unknown operator site for those roles is dead
    /// code. Recording the roles during the declaration pass lets operator
    /// lowering ask the program's actual surface.
    pub(crate) program_protocol_slots: HashSet<vybe_ast::ProtocolSlot>,
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
    /// Any class in this program binds `ProtocolSlot::CallMissing` (PHP
    /// `__call`, Ruby `method_missing`, Dart `noSuchMethod`) — the
    /// method-miss interceptor.
    ///
    /// The slot was in the vocabulary with NO reader anywhere: four frontends
    /// declared it and nothing consumed it, so each one synthesised its own
    /// `typeof obj.__call === "function" ? … : …` rewrite in its walker, once
    /// per call site. The reader is the Call Tags proposal's "Dynamic
    /// Invocation" application — a tag per name/arity pair whose fall-back
    /// handler runs the miss — so the decision is the CALLEE's and the call
    /// site is one instruction.
    ///
    /// Program-level for the same reason as `program_has_getattr`: the
    /// receiver's type is unknown exactly where this matters. Programs binding
    /// the role nowhere pay nothing.
    pub(crate) program_has_callmissing: bool,
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
    pub(crate) profile: std::sync::Arc<LanguageProfile>,
    current_func_name: Option<String>,
    statement_trace: Option<std::sync::Arc<StatementTraceConfig>>,
    current_result_slot: Option<u16>,
    current_ref_out_params: Option<Vec<u16>>,
    pending_classes: HashMap<String, PendingClass>,
    /// Canonical names of classes whose declared field ORDER this compiler
    /// published to the reserved type slot, so seam 3 may index them.
    ///
    /// ⛔ A type registered by any other path — a language emitter, a platform
    /// catalogue — has a field list this compiler did not author and cannot
    /// index against: python's classes arrive carrying
    /// `__class__`/`__dataclass_fields__`/`__match_args__` where the emitter
    /// would have indexed `x`/`y`/`label`. Indexing those wrote past the arity
    /// and trapped. Membership here is the ONLY licence to emit an indexed
    /// access.
    pub(crate) seam3_indexable: std::collections::HashSet<String>,
    /// Local derived constructors with a proven fixed parent initializer ABI.
    pub(crate) receiver_preserving_classes: std::collections::HashSet<String>,
    #[cfg(feature = "primitive-benchmarks")]
    pub(crate) inherited_gc_fields: bool,
    #[cfg(feature = "primitive-benchmarks")]
    pub(crate) compact_parent_forwarding: bool,
    #[cfg(feature = "primitive-benchmarks")]
    pub(crate) static_class_member_writes: bool,
    #[cfg(feature = "primitive-benchmarks")]
    pub(crate) console_string_producers: bool,
    #[cfg(feature = "primitive-benchmarks")]
    pub(crate) console_fixed_numeric: bool,
    #[cfg(feature = "primitive-benchmarks")]
    pub(crate) compound_numeric_proof: bool,
    #[cfg(feature = "primitive-benchmarks")]
    pub(crate) comparison_bool_reuse: bool,
    #[cfg(feature = "primitive-benchmarks")]
    pub(crate) generator_predicate_reuse: bool,
    #[cfg(feature = "primitive-benchmarks")]
    pub(crate) generator_plain_control_packets: bool,
    #[cfg(feature = "primitive-benchmarks")]
    pub(crate) condition_truthiness_reuse: bool,
    #[cfg(feature = "primitive-benchmarks")]
    pub(crate) integer_truthiness_reuse: bool,
    /// Canonical names of classes whose OWN typeidx is what `struct.new_default`
    /// actually used, so `ref.test` against them answers "was this object
    /// constructed by this class".
    ///
    /// ⛔ A SUPERSET OF [`Self::seam3_indexable`], AND NOT THE SAME QUESTION.
    /// Indexing needs allocation identity AND a storage model that may put the
    /// field in a GC slot; a TYPE TEST needs only the first. Deriving one from
    /// the other is how a js private brand lost its `ref.test` the moment js
    /// declared its fields to be own properties — a storage decision silently
    /// withdrawing an identity guarantee.
    ///
    /// Membership is still `published && no parent`, because under the
    /// base-constructor model a derived class's instance carries the BASE's
    /// rtt: `new Sub()` allocates as `Base`, so `ref.test Sub` is `false` for
    /// an object that plainly is one. M5 moves allocation to the most-derived
    /// constructor, at which point the parent clause goes from both sets.
    pub(crate) rtt_testable: std::collections::HashSet<String>,
    /// Every class NORMALIZED ONCE, during the declaration pass, keyed by
    /// canonical name. This is the single class model: computed before any
    /// body compiles, so a class's member set is knowable regardless of
    /// compilation order, and augmenting types (traits / mixins / promoted
    /// fields) can be resolved by name. See flexclassplan.md §3a, §4c.
    normalized_classes: NormalizedClasses,
    deferred_class_bodies: Option<deferred_class_bodies::Owner>,
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
    /// Source body receipts are valid only during this consumed Compiler's
    /// compile_with_imports call, while its borrowed immutable Module is alive.
    /// Cloned/normalized/generated bodies never inherit a source receipt.
    source_atomic_free_bodies: HashSet<(usize, usize)>,
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
    /// A bare return in a constructor returns the allocated instance.
    /// The chunk index keeps nested function bodies outside this context.
    pub(crate) constructor_this_ctx: Option<(usize, u16)>,
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
    /// instead of in `crate::primitives::namespaces`.
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
    user_namespace_lookup_cache: std::cell::RefCell<user_tree_cache::UserTreeCache>,
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
    compact_source_lines: bool,
    array_call_tail_identity: bool,
    interpolation_call_tail_identity: bool,
    fused_closure_analysis: bool,
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
    compile_certificate: Option<Arc<std::sync::Mutex<crate::compile_dependencies::Certificate>>>,
    local_module_export_overlays: HashSet<(String,String)>,
    module_exports: HashMap<String, HashMap<String, (String, String)>>,
    /// Constant-value snapshot of host module exports — `ExportEntry::Value`
    /// entries from `flatten_module_value_exports`. Populated at compile time
    /// and used during `collect_host_imports` to route value exports into
    /// `host_const_bindings` instead of `host_import_bindings`.
    module_value_exports: HashMap<String, HashMap<String, vybe_runtime::Value>>,
    /// Per-FRAME control-flow bookkeeping, indexed by chunk index — see
    /// [`FrameControlFlow`].
    frame_control_flow: FrameControlFlowStore,
    js_arguments_bindings: Vec<Option<JsArgumentsBinding>>,
}

/// §16.2.1.3 wildcard — `import * as alias from "module"`.
#[derive(Debug, Clone)]
pub struct HostWildcardImport {
    pub alias: String,
    pub module: String,
}

/// Control-flow bookkeeping that describes ONE frame, addressed by the chunk
/// it describes.
///
/// ⛔ Every field here names something FRAME-RELATIVE — a local slot, an index
/// into another of these vectors, or a structured-label depth — so it is
/// meaningless in any other chunk. It lived on the `Compiler` and was
/// hand-saved at each site that switches `self.current` to a nested chunk;
/// those sites disagreed about which fields to save, and the ones nobody saved
/// LEAKED: a lambda compiled inside a `Using` inherited the enclosing block's
/// `ResourceDispose { slot }`, so its implicit return emitted the enclosing
/// frame's `LOCAL_GET 77` into a chunk with nine locals and trapped. Keying
/// the state by frame is what makes that unrepresentable: switching
/// `self.current` switches the bookkeeping with it and there is nothing left
/// to remember to save.
///
/// This is not a language quirk and carries no directive: a `return` inside a
/// nested function runs that function's `finally` blocks and never the
/// enclosing function's, in every language that has both.
#[derive(Debug, Default)]
struct FrameControlFlow {
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
}

/// [`FrameControlFlow`] per frame, addressed by chunk index.
///
/// `empty` is what a frame with no entry yet reads as. An entry is created on
/// first WRITE — most chunks never open a `try`, a `using` or an async body —
/// and a read takes `&self`, so it cannot create one.
///
/// ⛔ INVARIANT: a chunk index never moves. `self.chunks` is only ever pushed
/// to (no `remove`/`swap`/`retain`/`insert`/`drain`), which is what makes the
/// index a stable name for a frame. A path that shifted them would silently
/// re-attach every frame's bookkeeping to a different chunk.
#[derive(Debug, Default)]
struct FrameControlFlowStore {
    frames: Vec<FrameControlFlow>,
    empty: FrameControlFlow,
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

/// A `try`-with-`finally` join point (see [`FrameControlFlow::finally_joins`]).
#[derive(Debug)]
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
            // ⛔ THIS LIST AND THE HOST'S ANCHOR LIST HAD DRIFTED APART.
            //
            // `ecma_globals::register` publishes `__ctor_<Name>` for all of the
            // names below; this predicate read only the ten above, so a bare
            // `Map` / `Set` / `DataView` resolved to something that was not the
            // canonical constructor — no `name`, no `prototype`. Invisible
            // while `instanceof` answered from a `__type` string; the moment
            // that stamp was deleted from the ECMA host it became
            // `new Map() instanceof Map === false`.
            //
            // Two lists, one authority, and only one of them was maintained.
            // The host is the authority: it is what actually installs the
            // anchors. Anything added there must be added here.
            | "Map"
            | "Set"
            | "ArrayBuffer"
            | "SharedArrayBuffer"
            | "DataView"
            | "Int8Array"
            | "Uint8Array"
            | "Uint8ClampedArray"
            | "Int16Array"
            | "Uint16Array"
            | "Int32Array"
            | "Uint32Array"
            | "Float32Array"
            | "Float64Array"
            | "BigInt64Array"
            | "BigUint64Array"
            | "Promise"
            | "WeakMap"
            | "WeakSet"
            | "WeakRef"
            | "FinalizationRegistry"
            // WHATWG constructors, anchored by `platforms/web`'s plugin
            // `finalize` — same contract as the ECMA ones above.
            | "URL"
            | "URLSearchParams"
            | "TextEncoder"
            | "TextDecoder"
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
fn collect_atomic_place_idents(stmts: &[Statement], out: &mut impl AtomicPlaceSink) {
    for stmt in stmts {
        collect_atomic_place_in_stmt(stmt, out);
    }
}

fn collect_atomic_place_in_stmt(stmt: &Statement, out: &mut impl AtomicPlaceSink) {
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
        StmtKind::FunctionDecl { body, .. } => {
            out.source_body(body);
            collect_atomic_place_idents(body, out)
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            out.begin_class();
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
            out.end_class();
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

fn collect_atomic_place_in_expr(expr: &Expression, out: &mut impl AtomicPlaceSink) {
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
                out.place_name(name);
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

// The module prescan and function prescan share one visitor. Only source
// procedure bodies keep their identity through emission; class bodies are
// normalized/copied, so their per-function scan must remain unconditional.
trait AtomicPlaceSink {
    fn begin_class(&mut self) {}
    fn end_class(&mut self) {}
    fn place_name(&mut self, name: &str);
    fn source_body(&mut self, _body: &[Statement]) {}
}
impl AtomicPlaceSink for HashSet<String> {
    fn place_name(&mut self, name: &str) {
        self.insert(name.to_owned());
    }
}
struct AtomicModuleScan<'a> {
    class_depth: usize,
    names: &'a mut HashSet<String>,
    bodies: &'a mut HashSet<(usize, usize)>,
}
impl AtomicPlaceSink for AtomicModuleScan<'_> {
    fn begin_class(&mut self) {
        self.class_depth += 1;
    }
    fn end_class(&mut self) {
        self.class_depth -= 1;
    }
    fn place_name(&mut self, name: &str) {
        self.names.insert(name.to_owned());
    }
    fn source_body(&mut self, body: &[Statement]) {
        if self.class_depth == 0 {
            self.bodies.insert((body.as_ptr() as usize, body.len()));
        }
    }
}

struct AtomicReceiptRetirement<'a>(&'a mut HashSet<(usize, usize)>);
impl AtomicPlaceSink for AtomicReceiptRetirement<'_> {
    fn place_name(&mut self, _: &str) {}
    fn source_body(&mut self, body: &[Statement]) {
        self.0.remove(&(body.as_ptr() as usize, body.len()));
    }
}

impl Compiler {
    fn retire_source_body_receipts(&mut self, stmt: &Statement) {
        if self.source_atomic_free_bodies.is_empty()
            || matches!(stmt.kind, StmtKind::ClassDecl { .. }
                | StmtKind::StructDecl { .. } | StmtKind::ModuleDecl { .. }) { return; }
        collect_atomic_place_in_stmt(stmt, &mut AtomicReceiptRetirement(&mut self.source_atomic_free_bodies));
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
                let mut local_names: HashSet<String> = params.iter().map(|p| p.name.clone()).collect();
                collect_declared_names(body, &mut local_names);
                collect_all_idents_in_stmts(body, &mut Filtered { locals: &local_names, out });
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
            let mut local_names: HashSet<String> = params.iter().map(|p| p.name.clone()).collect();
            collect_declared_names(body, &mut local_names);
            collect_all_idents_in_stmts(body, &mut Filtered { locals: &local_names, out });
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
            vybe_ast::InterpolPart::Expr(e) | vybe_ast::InterpolPart::Formatted(e, _) => {
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

pub(crate) fn expr_has_closure_with_this(expr: &Expression) -> bool {
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
        ExprKind::RefOf(place) => match &**place {
            PlaceExpr::Ident(_) => {}
            PlaceExpr::Member { object, .. } | PlaceExpr::Deref(object) => {
                collect_closure_captured_in_expr(object, out);
            }
            PlaceExpr::Index { object, index, .. } => {
                collect_closure_captured_in_expr(object, out);
                collect_closure_captured_in_expr(index, out);
            }
        },
        // `const K = class { m(){ return msg; } }` closes over the enclosing
        // frame just as the declaration form does — same members, same rule.
        ExprKind::ClassExpr { members, .. } => {
            collect_closure_captured_in_class_members(members, out);
        }
        ExprKind::Lambda { params, body, .. } => {
            let mut local_names: HashSet<String> = params.iter().map(|p| p.name.clone()).collect();
            if let LambdaBody::Block(stmts) = body { collect_declared_names(stmts, &mut local_names); }
            let mut filtered = Filtered { locals: &local_names, out };
            match body { LambdaBody::Block(stmts) => collect_all_idents_in_stmts(stmts, &mut filtered), LambdaBody::Expr(e) => collect_all_idents_in_expr(e, &mut filtered) }
        }
        ExprKind::FunctionExpr(stmt) => {
            if let StmtKind::FunctionDecl { params, body, .. } = &stmt.kind {
                let mut local_names: HashSet<String> = params.iter().map(|p| p.name.clone()).collect();
                collect_declared_names(body, &mut local_names);
                collect_all_idents_in_stmts(body, &mut Filtered { locals: &local_names, out });
            }
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::IsType { expr, .. }
        | ExprKind::RefLoad(expr) => collect_closure_captured_in_expr(expr, out),
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
                            let mut local_names: HashSet<String> = params.iter().map(|p| p.name.clone()).collect();
                            collect_declared_names(body, &mut local_names);
                            collect_all_idents_in_stmts(body, &mut Filtered { locals: &local_names, out });
                        }
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Sequence(exprs) | ExprKind::Tuple(exprs) => {
            for e in exprs {
                collect_closure_captured_in_expr(e, out);
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let vybe_ast::InterpolPart::Expr(e)
                | vybe_ast::InterpolPart::Formatted(e, _) = part
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

// The identifier visitor writes to a sink. Capture discovery filters locals
// directly; nested class-member discovery can still use a temporary borrowed
// set. Both paths use the same traversal, preserving syntax coverage.
fn collect_all_idents_in_stmts<'ast>(stmts: &'ast [Statement], out: &mut impl IdentSink<'ast>) {
    for stmt in stmts {
        collect_all_idents_in_stmt(stmt, out);
    }
}

fn collect_all_idents_in_stmt<'ast>(stmt: &'ast Statement, out: &mut impl IdentSink<'ast>) {
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

fn collect_all_idents_in_expr<'ast>(expr: &'ast Expression, out: &mut impl IdentSink<'ast>) {
    match &expr.kind {
        ExprKind::RefOf(place) => match &**place {
            PlaceExpr::Ident(name) => { out.insert(name.as_str()); }
            PlaceExpr::Member { object, .. } | PlaceExpr::Deref(object) => {
                collect_all_idents_in_expr(object, out);
            }
            PlaceExpr::Index { object, index, .. } => {
                collect_all_idents_in_expr(object, out);
                collect_all_idents_in_expr(index, out);
            }
        },
        ExprKind::Ident(name) => {
            out.insert(name.as_str());
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::IsType { expr, .. }
        | ExprKind::RefLoad(expr) => collect_all_idents_in_expr(expr, out),
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
                            let mut all_idents: HashSet<&str> = HashSet::new();
                            collect_all_idents_in_stmts(body, &mut all_idents);
                            let mut local_names: HashSet<String> =
                                params.iter().map(|p| p.name.clone()).collect();
                            collect_declared_names(body, &mut local_names);
                            for name in all_idents {
                                if !local_names.contains(name) {
                                    out.insert(name);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Sequence(exprs) | ExprKind::Tuple(exprs) => {
            for e in exprs {
                collect_all_idents_in_expr(e, out);
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let vybe_ast::InterpolPart::Expr(e)
                | vybe_ast::InterpolPart::Formatted(e, _) = part
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


struct Filtered<'a> { locals: &'a HashSet<String>, out: &'a mut HashSet<String> }
impl<'ast> IdentSink<'ast> for Filtered<'_> { fn insert(&mut self, name: &'ast str) -> bool { if self.locals.contains(name) || self.out.contains(name) { false } else { self.out.insert(name.to_owned()) } } }

trait IdentSink<'ast> { fn insert(&mut self, name: &'ast str) -> bool; }
impl<'ast> IdentSink<'ast> for HashSet<&'ast str> { fn insert(&mut self, name: &'ast str) -> bool { HashSet::insert(self,name) } }

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
    /// Advisory final-bytecode proofs for an explicitly proof-aware backend.
    /// Existing writers/VM consumers ignore these and preserve boxed storage.
    pub word_local_proofs: Vec<vybe_ast::word_local_transport::WordLocalChunkProof>,
    pub host_imports: HostImportMetadata,
    pub component_source: bool,
    pub component_binary: Option<std::sync::Arc<vybe_ast::canon::ComponentBinary>>,
    /// What the module DECLARED about presenting a UI
    /// ([`vybe_ast::Directives::app_shell`]), carried out so the embedder can
    /// read it without re-walking the tree. `None` states nothing — the
    /// document answers instead.
    pub app_shell: Option<vybe_ast::AppShell>,
}

/// One completed source producer within a prepared unit. These ranges report
/// progress only: later finalization still relocates imports/globals/types.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EmissionProgress {
    pub source_index: usize,
    pub function_declaration: bool,
    pub appended_chunks: std::ops::Range<usize>,
    pub entry_bytes: std::ops::Range<usize>,
}

mod prepared_emission;
pub use prepared_emission::{PreparedEmission, PreparedUnit, PreparedUnitProducer};

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
        Self::strip_global_namespace_prefix_view(name).to_owned()
    }

    /// Read-only name resolution does not need an allocated lowercase copy
    /// just to recognize the two ASCII global-namespace markers.
    pub(crate) fn strip_global_namespace_prefix_view(name: &str) -> &str {
        let trimmed = name.trim();
        if trimmed.get(..8).is_some_and(|prefix| prefix.eq_ignore_ascii_case("global::")) {
            return trimmed[8..].trim();
        }
        if trimmed.get(..7).is_some_and(|prefix| prefix.eq_ignore_ascii_case("global.")) {
            return trimmed[7..].trim();
        }
        trimmed
    }

    pub fn with_profile(profile: LanguageProfile) -> Self {
        Self::with_shared_profile(std::sync::Arc::new(profile))
    }

    pub fn with_shared_profile(profile: std::sync::Arc<LanguageProfile>) -> Self {
        Self {
            argument_emitters: argument_emission::Selection::default(),
            helper_parser: None,
            helper_templates: helper_templates::Owner::default(),
            last_lambda_body_chunk: None,
            chunks: vec![Chunk::new("<script>")],
            scopes: vec![Scope::new(scope::profile_variable_fold(&profile))],
            current: 0,
            loops: Vec::new(),
            loop_states: Vec::new(),
            label_depth: 0,
            function_label_base: 0,
            line: 1,
            defined_globals: body_symbols::BodySymbols::new(),
            module_variable_names: HashSet::new(),
            gui_auto_name_counter: 0,
            const_globals: HashSet::new(),
            global_const_values: std::collections::HashMap::new(),
            in_strict: false,
            directives: vec![vybe_ast::Directives::default()],
            canon_section: Vec::new(),
            core_export_callees: Vec::new(),
            canon_functypes: Vec::new(),
            canon_valtypes: Vec::new(),
            component_funcs: Vec::new(),
            component_func_links: Vec::new(),
            component_instance_links: None,
            in_typeof_operand: false,
            has_discarded_builtin_emits: profile.builtins.values().any(|def| def.discarded_emit.is_some()),
            want_i32_condition: false,
            gave_i32_condition: false,
            gave_i32_word: false,
            program_lexical_names: HashSet::new(),
            defined_functions: body_symbols::FunctionSymbols::new(),
            immutable_function_bindings: immutable_calls::Bindings::default(),
            run_module_end_hooks: true,
            function_param_modes: HashMap::default(),
            function_param_types: HashMap::default(),
            function_min_arity: HashMap::default(),
            function_signatures: HashMap::default(),
            rest_fixed_arities: BTreeSet::new(),
            function_return_types: HashMap::default(),
            fortran_interface_overloads: HashMap::new(),
            fortran_operator_overloads: HashMap::new(),
            constructor_signatures: HashMap::new(),
            defined_classes: body_symbols::BodySymbols::new(),
            abstract_classes: HashSet::new(),
            defined_class_methods: HashSet::new(),
            classes_with_late_static_binding: HashSet::new(),
            classes_with_cooperative_super: HashSet::new(),
            classes_with_indexer: HashSet::new(),
            classes_with_index_setter: HashSet::new(),
            program_protocol_slots: HashSet::new(),
            program_has_getattr: false,
            program_has_setattr: false,
            program_has_callmissing: false,
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
            statement_trace: None,
            current_result_slot: None,
            current_ref_out_params: None,
            pending_classes: HashMap::new(),
            seam3_indexable: std::collections::HashSet::new(),
            receiver_preserving_classes: std::collections::HashSet::new(),
            #[cfg(feature = "primitive-benchmarks")]
            inherited_gc_fields: true,
            #[cfg(feature = "primitive-benchmarks")]
            compact_parent_forwarding: true,
            #[cfg(feature = "primitive-benchmarks")]
            static_class_member_writes: true,
            #[cfg(feature = "primitive-benchmarks")]
            console_string_producers: true,
            #[cfg(feature = "primitive-benchmarks")]
            console_fixed_numeric: true,
            #[cfg(feature = "primitive-benchmarks")]
            compound_numeric_proof: true,
            #[cfg(feature = "primitive-benchmarks")]
            comparison_bool_reuse: true,
            #[cfg(feature = "primitive-benchmarks")]
            generator_predicate_reuse: true,
            #[cfg(feature = "primitive-benchmarks")]
            generator_plain_control_packets: true,
            #[cfg(feature = "primitive-benchmarks")]
            condition_truthiness_reuse: true,
            #[cfg(feature = "primitive-benchmarks")]
            integer_truthiness_reuse: true,
            rtt_testable: std::collections::HashSet::new(),
            normalized_classes: NormalizedClasses::default(),
            deferred_class_bodies: None,
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
            source_atomic_free_bodies: HashSet::new(),
            module_atomic_word_globals: HashSet::new(),
            current_closure_captured_locals: HashSet::new(),

            multi_return_functions: HashMap::new(),
            generator_functions: HashSet::new(),
            method_fn_kinds: HashMap::new(),
            js_derived_ctor_ctx: None,
            constructor_this_ctx: None,
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
            user_namespace_lookup_cache: Default::default(),
            current_module_imports: Vec::new(),
            active_namespaces: None,
            compile_certificate: None,
            local_module_export_overlays: HashSet::new(),
            module_exports: HashMap::new(),
            module_value_exports: HashMap::new(),
            frame_control_flow: FrameControlFlowStore::default(),
            uses_proxy: false,
            compact_source_lines: false,
            array_call_tail_identity: false,
            interpolation_call_tail_identity: false,
            fused_closure_analysis: false,
            bigint_enabled: false,
            js_arguments_bindings: Vec::new(),
        }
    }

    /// Apply explicit compiler lookup options before declaration/lowering.
    pub fn with_optimizations(mut self, options: vybe_runtime::registry::CompilerOptimizations) -> Self {
        self.compact_source_lines = options.compact_source_lines;
        self.array_call_tail_identity = options.array_call_tail_identity;
        self.interpolation_call_tail_identity = options.interpolation_call_tail_identity;
        self.fused_closure_analysis = options.fused_closure_analysis;
        self.defined_globals.configure_folded(options.symbol_membership_index);
        self.defined_classes.configure_folded(options.symbol_membership_index);
        self.defined_functions.configure_folded(options.symbol_membership_index);
        self.user_namespace_lookup_cache
            .borrow_mut()
            .configure(options.user_namespace_cache);
        self
    }

    /// Configure a fresh unit owner from an immutable request selection.
    pub fn with_selected_helpers(mut self, selected: &helper_templates::Selected) -> Self {
        self.helper_parser = Some(selected.selection.parser);
        self.helper_templates = selected.owner();
        self
    }
    pub(crate) fn with_selected_argument_emitters(mut self, selection: argument_emission::Selection) -> Self {
        self.argument_emitters = selection;
        self
    }
    /// Explicit exact-parser opt-in; absent for legacy PEG and ordinary callers.
    pub fn with_argument_emitters(mut self, parser: vybe_runtime::registry::ParseFn,
        definitions: &'static [argument_emission::Definition]) -> Self {
        self.argument_emitters = argument_emission::Selection::new(parser, definitions);
        self
    }
    pub fn with_helper_templates(mut self, parser: vybe_runtime::registry::ParseFn, definitions: &'static [helper_templates::Definition], pool: std::sync::Arc<helper_templates::Pool>) -> Self {
        self.helper_parser = Some(parser);
        self.helper_templates = helper_templates::Owner::new(definitions, pool);
        self
    }
    fn try_owned_helper(&mut self, name: &str, argc: u8, line: u32, request: vybe_runtime::registry::emitter_results::EmitRequest) -> Option<helper_templates::EmissionReceipt> {
        let Some(parser) = self.helper_parser else { return None };
        // Runtime-helper lowering uses a different reference expansion and
        // frame contract. Keep that path inline rather than sharing a body
        // captured for ordinary source compilation.
        if polyfills::is_compiling_runtime_helper() { return None; }
        let Some((namespace, _)) = name.split_once('.') else { return None };
        let Some(producer) = crate::languages::emit_dispatch_for(namespace) else { return None };
        let chunk = &self.chunks[self.current];
        let context = helper_templates::Context { profile: &self.profile, directives: self.directives.last().expect("module directive frame"), receiver: chunk.module_receiver_abi, own: chunk.module_instance_fields_are_own_properties, prototype: chunk.module_members_on_prototype };
        self.helper_templates.emit_with_result(name, parser, producer, argc, request, &context, &mut self.chunks, self.current, line)
    }
    pub fn with_run_module_end_hooks(mut self, enabled: bool) -> Self {
        self.run_module_end_hooks = enabled;
        self
    }

    /// The control-flow bookkeeping of the frame being compiled.
    fn frame_cf(&self) -> &FrameControlFlow {
        self.frame_control_flow
            .frames
            .get(self.current)
            .unwrap_or(&self.frame_control_flow.empty)
    }

    /// The same, for writing — the frame's entry is created here on demand.
    fn frame_cf_mut(&mut self) -> &mut FrameControlFlow {
        let frames = &mut self.frame_control_flow.frames;
        if frames.len() <= self.current {
            frames.resize_with(self.current + 1, FrameControlFlow::default);
        }
        &mut frames[self.current]
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
    pub(crate) fn with_compile_certificate(mut self, certificate: Arc<std::sync::Mutex<crate::compile_dependencies::Certificate>>) -> Self {
        self.compile_certificate=Some(certificate);self
    }
    fn compiler_diagnostic(&self,message:String) {
        if self.compile_certificate.as_ref().is_some_and(|c|c.lock().unwrap().capture_diagnostic(&message)) {return;}
        eprintln!("{message}");
    }
    fn recorded_module_export(&self,module:&str,name:&str)->Option<&(String,String)> {
        let target=self.module_exports.get(module).and_then(|m|m.get(name));
        if let Some(c)=&self.compile_certificate {
            if !self.local_module_export_overlays.contains(&(module.into(),name.into())) { c.lock().unwrap().record_export(module,name,target); }
        }
        target
    }
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

    /// The middle stage between `--dump-ast` and `--dump`: what the declaration
    /// pass actually BUILT from the AST. A receiver that "lost its type" is
    /// almost always a name present in the AST but absent from — or in the
    /// wrong one of — these tables, which neither other dump can show.
    ///
    /// Sorted, because `HashMap` order is arbitrary and a diagnostic that
    /// reorders between runs cannot be diffed.
    pub fn dump_pending_classes(&self) {
        println!("== pending classes ==");
        let mut names: Vec<&String> = self.pending_classes.keys().collect();
        names.sort();
        for name in names {
            let pc = &self.pending_classes[name];
            println!("{name}");
            println!("  parent:  {}", pc.parent.as_deref().unwrap_or("-"));
            let mut fields: Vec<String> = pc
                .instance_field_types
                .iter()
                .map(|(k, v)| format!("{k} -> {}", v.hint))
                .collect();
            fields.sort();
            println!("  fields:  {}", join_or_none(&fields));
            // ⛔ THE LINE ABOVE IS NOT THE EMITTED TYPE. It maps over
            // `instance_field_types` — field → TYPE HINT — so a language
            // without type annotations (js, python) prints `(none)` even when
            // the fields exist. That has been misread as "the emitted struct
            // type has zero fields" by three separate readings, including the
            // recorded "M6-pre BLOCKER MEASURED ON THE WRONG Vec".
            //
            // `TypeEntry.fields` is the structure the wasm writer actually
            // consumes and the one whose LENGTH decides whether an engine
            // accepts the module. Printing it beside the hints is what makes
            // the difference between the two visible instead of inferred.
            let emitted: Option<&Vec<String>> = self.chunks[0]
                .types
                .iter()
                .find(|t| t.name == *name)
                .map(|t| &t.fields);
            match emitted {
                Some(f) if !f.is_empty() => {
                    println!(
                        "  emitted: {} (TypeEntry.fields, {} slot(s))",
                        f.join(", "),
                        f.len()
                    )
                }
                Some(_) => println!("  emitted: (none) — TypeEntry.fields is EMPTY"),
                None => println!("  emitted: (no TypeEntry registered)"),
            }
            let mut statics: Vec<String> = pc
                .static_field_types
                .iter()
                .map(|(k, v)| format!("{k} -> {v}"))
                .collect();
            statics.sort();
            println!("  statics: {}", join_or_none(&statics));
            let mut members = pc.instance_member_names.clone();
            members.sort();
            println!("  members: {}", join_or_none(&members));
            let mut nested = pc.nested_types.clone();
            nested.sort();
            println!("  nested:  {}", join_or_none(&nested));
            // The two facts that decide whether a member READ of this class
            // becomes `struct.get <typeidx> <fieldidx>` or stays a string key.
            // A class can be fully published, with `TypeEntry.fields` above
            // showing real slots, and still be read by name because one of
            // these is missing — which neither the fields line nor `--dump`
            // can show.
            let canon = self.canon(name);
            println!(
                "  normal:  {}",
                if self.normalized_classes.contains_key(&canon) {
                    "in normalized_classes"
                } else {
                    "ABSENT — the property check fails closed, so reads stay keyed"
                }
            );
            println!(
                "  seam3:   {}",
                if self.seam3_indexable.contains(&canon) {
                    "licensed — reads may index"
                } else {
                    "NO LICENCE — reads stay keyed"
                }
            );
        }
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
    pub fn compile_with_imports(self, module: &Module) -> Result<CompileResult, String> {
        self.compile_with_source_body(module, std::borrow::Cow::Borrowed(&module.body), None)
    }

    /// Registered metadata-only normalization retains borrowed source bodies.
    /// Source-class donation remains exclusive to the consuming AST path.
    pub(crate) fn compile_with_imports_prepared(self, module: &Module,
        parser: vybe_runtime::registry::ParseFn) -> Result<CompileResult, String> {
        self.begin_emission_prepared(module, parser)?.finish()
    }

    pub(crate) fn begin_emission_prepared<'a>(self, module: &'a Module,
        parser: vybe_runtime::registry::ParseFn) -> Result<PreparedEmission<'a>, String> {
        let matches_profile = self.profile.uses_normalize_class
            && crate::languages::find_by_name(&self.profile.name).is_some_and(|language|
                std::ptr::fn_addr_eq(language.parse, parser));
        let callback = matches_profile
            .then(|| crate::frontend_preparation::snapshot(parser).capabilities.deferred_class_bodies)
            .flatten()
            // Header-only selection; Owner::new performs the complete
            // expression-scope proof once before omitting any installed body.
            .filter(|_| deferred_class_bodies::has_installed_library_owned_candidate(
                &module.body, self.profile.case_sensitive));
        self.prepare_source_emission(module, std::borrow::Cow::Borrowed(&module.body), callback)
    }

    /// Same borrowed prepared entry used by CaaS, exposed only for parity gates.
    #[cfg(feature = "service-verification")]
    pub fn compile_with_imports_prepared_for_validation(self, module: &Module,
        parser: vybe_runtime::registry::ParseFn) -> Result<CompileResult, String> {
        self.compile_with_imports_prepared(module, parser)
    }

    /// Consumes a caller-owned AST; unregistered parsers retain legacy normalization.
    pub(crate) fn compile_with_imports_owned(self, module: Module,
        parser: vybe_runtime::registry::ParseFn) -> Result<CompileResult, String> {
        self.begin_emission_owned_prepared(module, parser)?.finish()
    }

    pub(crate) fn begin_emission_owned_prepared(self, mut module: Module,
        parser: vybe_runtime::registry::ParseFn) -> Result<PreparedEmission<'static>, String> {
        let body = std::mem::take(&mut module.body);
        let matches_profile = self.profile.uses_normalize_class
            && crate::languages::find_by_name(&self.profile.name).is_some_and(|language|
                std::ptr::fn_addr_eq(language.parse, parser));
        let callback = matches_profile.then(||crate::frontend_preparation::snapshot(parser).capabilities.deferred_class_bodies).flatten();
        self.prepare_source_emission(&module, std::borrow::Cow::Owned(body), callback)
    }

    /// Validation only: production callers already carry this admission proof.
    #[cfg(feature = "service-verification")]
    pub fn compile_with_imports_owned_for_validation(self, module: Module,
        parser: vybe_runtime::registry::ParseFn) -> Result<CompileResult, String> {
        if self.profile.uses_normalize_class
            && crate::languages::find_by_name(&self.profile.name).is_some_and(|language| std::ptr::fn_addr_eq(language.parse, parser))
            && crate::frontend_preparation::snapshot(parser).capabilities.deferred_class_bodies.is_some()
            && ((deferred_class_bodies::has_candidate_class_headers(&module.body)
                && crate::dynamic::independent_lowering(&module))
                || deferred_class_bodies::has_installed_library_owned_candidate(&module.body,self.profile.case_sensitive)) {
            self.compile_with_imports_owned(module, parser)
        } else {
            self.compile_with_imports(&module)
        }
    }

    /// Prepare one validated source unit once, then emit its complete common-AST
    /// declarations/statements incrementally. Pending chunks are not relocated.
    pub fn begin_emission<'a>(self, module: &'a Module) -> Result<PreparedEmission<'a>, String> {
        self.prepare_source_emission(module, std::borrow::Cow::Borrowed(&module.body), None)
    }

    /// Transfer a validated parser/source-unit body without cloning its AST.
    /// This has the same unregistered normalization policy as `compile_with_imports`.
    pub fn begin_emission_owned(self, mut module: Module) -> Result<PreparedEmission<'static>, String> {
        let body = std::mem::take(&mut module.body);
        self.prepare_source_emission(&module, std::borrow::Cow::Owned(body), None)
    }

    fn compile_with_source_body(self, module: &Module,
        source_body: std::borrow::Cow<'_, [Statement]>,
        deferred_normalizer: Option<deferred_class_bodies::NormalizeFn>) -> Result<CompileResult, String> {
        self.prepare_source_emission(module, source_body, deferred_normalizer)?.finish()
    }

    fn prepare_source_emission<'a>(mut self, module: &Module,
        source_body: std::borrow::Cow<'a, [Statement]>,
        deferred_normalizer: Option<deferred_class_bodies::NormalizeFn>) -> Result<PreparedEmission<'a>, String> {
        let _source_lines = vybe_runtime::source_lines::LineEmissionScope::enter(self.compact_source_lines);
        if self.compact_source_lines {
            // The entry Chunk exists before options are selected. Later
            // function/helper chunks inherit this thread-local emission scope.
            for chunk in &mut self.chunks { chunk.lines.enable_compact(); }
        }
        let _debug_compile = vybe_runtime::debugger::DebugPhase::current("compiler total");
        let timing = compiler_timing_path();
        self.statement_trace = timing.as_ref().and_then(|path| {
            std::env::var("VYBE_COMPILER_DEBUG_FUNCTION").ok().map(|function| {
                std::sync::Arc::new(StatementTraceConfig { path: path.clone(), function })
            })
        });
        let total_started = Instant::now();
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
        // Lower the receiver DIRECTIVE to the receiver ABI, once, onto the
        // module chunk. The emitters that need this answer are free functions
        // holding `&mut [Chunk]` or the imports chunk, never a `Compiler` — so
        // asking each of them to carry a flag would be the "33 sites had to
        // write it, 23 forgot" shape one more time.
        //
        // ⛔ THIS IS PER-BUNDLE, NOT PER-UNIT — unlike `variable_fold` above,
        // which is genuinely per-module. `chunks[0]` is one slot for the whole
        // bundle, so in a multi-language bundle the LAST unit compiled wins.
        // Correct for a single-language program. Languages differ in this
        // directive, so a bundle that mixes a unit which declares
        // `UniversalParameter` with one that does not takes whichever compiled
        // last; such a bundle needs this moved onto each unit's own chunks
        // before it is trustworthy.
        if !self.chunks.is_empty() {
            self.chunks[0].module_receiver_abi = if self.universal_receiver() {
                vybe_runtime::chunk::ReceiverAbi::Parameter
            } else {
                vybe_runtime::chunk::ReceiverAbi::Ambient
            };
        }
        // The canon section is module-level DATA, not policy and not code: it
        // has no execution position, so it is captured here rather than
        // emitted, and published to the chunks once compilation is done.
        self.canon_section = common::canon::lower_section(&module.canon.defs)?;
        self.core_export_callees = common::canon::collect_core_export_callees(&module.canon.defs);
        self.core_export_callees.extend(common::canon::attach_resource_metadata(
            &module.canon.defs,
            &module.canon.types,
            &mut self.canon_section,
        )?);
        let (fts, vts) = common::canon::lower_types(&module.canon.types)?;
        self.canon_functypes = fts;
        self.canon_valtypes = vts;
        self.component_funcs = module.canon.funcs.clone();
        self.component_instance_links = common::canon::lower_component_instance_links(
            module.canon.instance_links.as_deref(),
        )?;
        self.component_func_links = module
            .canon
            .func_links
            .iter()
            .map(|link| vybe_runtime::canon_def::ComponentFuncLink {
                import: link.import,
                supplied: link.supplied,
                expected_type: link.expected_type,
                supplied_type: link.supplied_type,
            })
            .collect();
        self.current_module_imports = module.imports.clone();
        // Whether a global holds a pointer cell is a WHOLE-MODULE property, but
        // compilation is one forward pass: `function r(){ global $g; echo $g; }`
        // is emitted before `$g = 1; w($g);` promotes `$g`, so the read would
        // correctly see "not a cell" and emit no autoderef — and then read the
        // cell object raw at runtime.
        //
        // Establish it up front, the same way `collect_addr_taken_idents`
        // already establishes address-taken LOCALS before a body is compiled.
        // Resolve nested address-taking through its lexical scopes. A later
        // function can promote a global, but a shadowing local cannot.
        let prescan_started = Instant::now();
        let debug_prescans = vybe_runtime::debugger::DebugPhase::current("compiler prescans");
        // Preserve existing policy admission; only source-derived facts are
        // collected by the typed-unit transaction below.
        let binds_ecma_proxy = self.profile.esm_defaults.iter().any(|entry| {
            matches!(entry, vybe_runtime::profile::EsmDefault::Namespace { alias, module }
                if alias == "Proxy" && module == "ecma:proxy")
        });
        let bigint_enabled = self.bigint_widens_mixes()
            || self.profile.builtins.values().any(|def| {
                matches!(&def.emit, vybe_runtime::profile::BuiltinEmit::HostCall(module, _)
                    if module == "ecma:bigint")
            });
        let features = source_feature_preparation::SourceFeaturePreparation::new(
            bigint_enabled, binds_ecma_proxy, self.uses_proxy,
        );
        let mut units = unit_preparation::TypedUnitPreparation::new(
            source_body,
            module.directives.variable_fold(),
            std::mem::take(&mut self.module_atomic_word_globals),
            std::mem::take(&mut self.source_atomic_free_bodies),
            features,
        );
        while units.collect_next().is_some() {}
        let unit_preparation::PreparedTypedUnits {
            body: source_body, references, atomic_names, source_bodies,
            features: source_features, functions: prepared_function_headers,
        } = units.seal()?;
        self.module_addr_taken_globals = references.into_iter()
            .map(|name| self.pointer_binding_key(&name))
            .collect();
        // Same forward-pass argument, for ATOMIC places: a module-level name
        // any statement (including one inside a nested function or lambda)
        // targets with an atomic op becomes a shared-memory word at its
        // declaration. It also enters `module_addr_taken_globals` so readers
        // compiled BEFORE the promotion emit the deref, which passes through
        // untouched until the wrap exists.
        self.module_atomic_word_globals = atomic_names;
        self.source_atomic_free_bodies = source_bodies;
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

        self.bigint_enabled = source_features.bigint_enabled;
        self.uses_proxy = source_features.uses_proxy;
        if let Some(path) = timing.as_deref() {
            write_compiler_timing(path, "prescans", prescan_started.elapsed());
        }
        drop(debug_prescans);

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
        let link_started = Instant::now();
        let debug_link = vybe_runtime::debugger::DebugPhase::current("compiler resolve");
        self.link_with_body(module, source_body.as_ref());
        drop(debug_link);
        if let Some(path) = timing.as_deref() {
            write_compiler_timing(path, "link", link_started.elapsed());
        }

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
        let has_partial = source_body.iter().any(
            |s| matches!(&s.kind, StmtKind::ClassDecl { modifiers, .. } if modifiers.is_partial),
        );
        let merge_started = Instant::now();
        let mut merged_body: std::borrow::Cow<'_, [Statement]> = if has_partial {
            // Partial merging copies procedure bodies; their old address receipts
            // must not survive release of the original owned source buffer.
            self.source_atomic_free_bodies.clear();
            std::borrow::Cow::Owned(merge_partial_classes(source_body.as_ref(), self.case_sensitive))
        } else { source_body };
        // Prepared borrowed compilation selects installed-library-only headers;
        // Owner proves expression scopes before publishing any metadata. Public
        // legacy borrowed compilation supplies no callback and stays unchanged.
        if !has_partial {
            if let Some(callback) = deferred_normalizer {
                self.deferred_class_bodies = deferred_class_bodies::Owner::new(&self, callback, &merged_body);
            }
        }

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
        let builtin_signatures = (!self.profile.builtin_signatures.is_empty()).then(|| {
            Statement::new(StmtKind::InterfaceDecl {
                // Reserved spelling — a user type of this name is not expressible
                // in any surface language here, so registering it as a type name
                // cannot shadow a real class (php's own `PDOStatement` would).
                name: "__vybe_builtin_signatures".to_string(),
                parents: Vec::new(),
                members: self.profile.builtin_signatures.clone(),
                decorators: Vec::new(),
            })
        });
        #[cfg(test)]
        let (mut merged_body, builtin_signatures) = if COPY_MODULE_BODY_BASELINE.with(|copy| copy.get()) {
            let mut body = merged_body.into_owned();
            if let Some(declaration) = builtin_signatures { body.push(declaration); }
            (std::borrow::Cow::Owned(body), None::<Statement>)
        } else { (merged_body, builtin_signatures) };
        // Keep the appended declaration without copying the source AST merely
        // to obtain a contiguous slice. Each phase sees source declarations
        // first and the builtin signature interface last, as before.
        let bodies = [merged_body.as_ref(), builtin_signatures.as_slice()];
        if let Some(path) = timing.as_deref() {
            write_compiler_timing(path, "merge_body", merge_started.elapsed());
        }

        let declarations_started = Instant::now();
        let _debug_declarations = vybe_runtime::debugger::DebugPhase::current("compiler declarations");
        for body in bodies { self.predeclare_type_names(body, None); }
        for body in bodies { self.collect_module_variable_names(body); }
        for body in bodies { self.predeclare_module_variable_type_hints(body); }
        self.collect_reflection_metadata(merged_body.iter().chain(builtin_signatures.iter()));

        // Multi-value pre-scan: any function whose every explicit `Return`
        // is a same-arity tuple literal is a candidate for the WASM
        // multi-value ABI. We only opt in when the language profile
        // requests it — other languages keep tuple-as-heap-object semantics.
        if self.profile.multi_value_tuple_returns {
            for body in bodies { self.collect_multi_return_functions(body); }
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
        // Partial merging and class normalization above do not rewrite
        // top-level function headers. Reuse their sealed source-unit facts.
        self.register_function_declarations(prepared_function_headers);
        for body in bodies { self.predeclare_interface_signatures_in_body(body); }
        drop(_debug_declarations);
        if let Some(path) = timing.as_deref() {
            write_compiler_timing(path, "declaration_passes", declarations_started.elapsed());
        }

        // Pre-collect every rest-parameter arity in the program so call-site
        // rest packing is emitted even when the callee (e.g. a `const f =
        // (...xs) => ...` arrow) is declared after a hoisted function body that
        // calls it. Without this, those calls silently drop arguments.
        {
            let mut rest_arities = Vec::new();
            for body in bodies { crate::ast::collect_rest_param_arities(body, &mut rest_arities); }
            for arity in rest_arities {
                self.rest_fixed_arities.insert(arity);
            }
        }

        // All original-AST analyses (including reflection and rest arities)
        // have completed in their original order. Move Vec allocations intact.
        if let Some(owner) = self.deferred_class_bodies.take() {
            match &mut merged_body {
                std::borrow::Cow::Owned(body) => owner.finish(&mut self, body)?,
                std::borrow::Cow::Borrowed(_) => owner.finish_metadata_only()?,
            }
        }

        // Publish field layouts before top-level function bodies can use them.
        // Roots are published first, followed by eligible local parent chains.
        // Sort each pass by canonical name so type indices are reproducible.
        // Re-publication during class compilation verifies the same layout.
        let publish_storage_started = Instant::now();
        {
            // Keep the lookup registry installed: private/static member naming
            // consults it while publication writes to the chunk type table.
            let declarations = self.normalized_classes.clone();
            let mut parentless_classes: Vec<_> = declarations
                .iter()
                .filter(|(_, class)| class.parent.is_none())
                .map(|(key, class)| (key, std::borrow::Cow::Borrowed(class)))
                .collect();
            #[cfg(test)]
            if COPY_MODULE_BODY_BASELINE.with(|copy| copy.get()) {
                for (_, class) in &mut parentless_classes { class.to_mut(); }
            }
            parentless_classes.sort_by(|(a, _), (b, _)| a.cmp(b));
            for (_, class) in &parentless_classes {
                self.publish_class_storage(class);
            }
            // Publish each local single-parent chain in parent-first order.
            // Only a parent with a stable indexed layout can license a child;
            // own-property, external and unresolved parents retain their path.
            let mut remaining: Vec<_> = declarations.iter()
                .filter(|(_, class)| class.parent.is_some() && class.bases.len() <= 1)
                .collect();
            remaining.sort_by(|(a, _), (b, _)| a.cmp(b));
            loop {
                let mut advanced = false;
                remaining.retain(|(_, class)| {
                    let parent = self.canon(class.parent.as_deref().unwrap());
                    if self.seam3_indexable.contains(&parent) || self.receiver_preserving_classes.contains(&parent) {
                        self.publish_class_storage(class);
                        advanced = true;
                        false
                    } else { true }
                });
                if !advanced { break; }
            }
        }
        if let Some(path) = timing.as_deref() {
            write_compiler_timing(path, "publish_parentless_class_storage", publish_storage_started.elapsed(),);
        }

        Ok(PreparedEmission::new(self, merged_body, builtin_signatures,
            module.canon.component_source, module.canon.binary.clone(), timing, total_started))
    }

    fn finish_emission(mut self, component_source: bool,
        component_binary: Option<std::sync::Arc<vybe_ast::canon::ComponentBinary>>,
        timing: Option<String>, total_started: Instant) -> Result<CompileResult, String> {
        let _debug_compile = vybe_runtime::debugger::DebugPhase::current("compiler total");
        let _source_lines = vybe_runtime::source_lines::LineEmissionScope::enter(self.compact_source_lines);
        // Auto-call entry point if defined
        let ep_started = Instant::now();
        if let Some(ref ep) = self.profile.entry_point.clone() {
            let has_ep = self.defined_globals.contains(ep)
                || (!self.case_sensitive && self.defined_globals.contains_ascii_folded(ep));
            if has_ep {
                self.emit_var_get(ep);
                let abi = crate::primitives::class_context::module_receiver_abi(&self.chunks);
                let line = self.line;
                let recv =
                    crate::primitives::callable::emit_callback_receiver(self.chunk(), abi, line);
                self.emit_direct_callable_invoke(recv);
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
                    self.class_get(
                        class_slots::ObjSource::Stack,
                        &class_slots::ClassSlot::internal(&ep_canon),
                    );
                    let abi = crate::primitives::class_context::module_receiver_abi(&self.chunks);
                    let line = self.line;
                    let recv = crate::primitives::callable::emit_callback_receiver(
                        self.chunk(),
                        abi,
                        line,
                    );
                    self.emit_direct_callable_invoke(recv);
                    self.emit(Op::DROP);
                }
            }
        }
        if let Some(path) = timing.as_deref() {
            write_compiler_timing(path, "entrypoint_emit", ep_started.elapsed());
        }

        // Lifecycle hooks belong to the source module's declared policy.
        // Shutdown callbacks precede the final flush so their output remains
        // buffered. Other modules must not acquire these globals or imports.
        let module_end_started = Instant::now();
        if self.run_module_end_hooks {
            if let Some(vybe_ast::ShutdownPolicy::Callbacks { registry, order }) =
                self.directives().shutdown_callbacks.clone()
            {
                let saved = self.current;
                self.current = 0;
                self.emit_run_shutdown_callbacks(&registry, order);
                self.current = saved;
            }
        }

        let line = self.line;
        if self.run_module_end_hooks && self.directives().flush_output_buffers == Some(true) {
            common::io::emit_ob_flush_all(&mut self.chunks, 0, line);
        }

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
        if let Some(path) = timing.as_deref() {
            write_compiler_timing(path, "module_end_emit", module_end_started.elapsed());
        }

        self.emit_null();
        self.emit(Op::RETURN);
        // Take the max of the scope's highest slot and whatever raw local
        // slots compiler_common helpers (e.g. `invoke::emit_invoke_method`)
        // reserved directly on the chunk — those bypass `Scope` but still
        // need the VM to reserve slots at call-frame entry.
        let ns = self.scope().next_slot;
        self.chunks[0].finalize_local_count(ns);
        let local_names = &self.scopes.last().unwrap().defined_names;
        crate::primitives::bundle::finalize_source_local_names(&mut self.chunks[0], local_names);
        // Skip helper linking when compiling runtime helper source.
        // Re-running finalization here would call back into helper
        // compilation and recurse through `Compiler::compile`.
        // and recurse forever. Cheap thread-local guard since polyfill
        // compilation is single-threaded at vybex build time.
        let finalize_started = Instant::now();
        let helper_finalize_started = Instant::now();
        let _debug_finalize = vybe_runtime::debugger::DebugPhase::current("compiler finalize");
        let _debug_helpers = vybe_runtime::debugger::DebugPhase::current("finalize helpers");
        if !crate::primitives::polyfills::is_compiling_runtime_helper() {
            // The exclusion list is profile data — a language that supplies
            // its own implementation of a shared helper names it there, rather
            // than the shared crate carrying a per-language table.
            let mut excluded: Vec<String> = self.profile.excluded_runtime_helpers.clone();
            if let Some(certificate) = &self.compile_certificate {
                excluded.extend(certificate.lock().unwrap().runtime_helper_exports(&self.chunks)
                    .into_iter().map(str::to_owned));
            }
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
        if let Some(path) = timing.as_deref() {
            write_compiler_timing(
                path,
                "finalize.runtime_helpers",
                helper_finalize_started.elapsed(),
            );
        }
        drop(_debug_helpers);
        // Finish structural class layouts before discovering and normalizing
        // global operands. Descriptor singleton declarations must use the final
        // inherited field/method layout and participate in the same global
        // index space as the code that reads them. Finalizing descriptors after
        // globals forces the writer to synthesize a separate representation.
        // Appending descriptor rows never changes existing class type indices.
        let class_finalize_started = Instant::now();
        self.finalize_type_prefixes();
        self.append_descriptor_type_rows();
        if let Some(path) = timing.as_deref() {
            write_compiler_timing(
                path,
                "finalize.class_metadata",
                class_finalize_started.elapsed(),
            );
        }
        // Relocate helper-owned canonical rows before import unification moves
        // imports out of the chunks that own their local type metadata.
        common::canon::resolve_emitter_imports(&mut self.chunks,
            &mut self.canon_section, &mut self.canon_functypes, &mut self.canon_valtypes)?;
        let normalize_imports_started = Instant::now();
        let _debug_imports = vybe_runtime::debugger::DebugPhase::current("finalize imports");
        let global_operands = Self::normalize_import_table_and_collect_globals(&mut self.chunks);
        drop(_debug_imports);
        if let Some(path) = timing.as_deref() {
            write_compiler_timing(
                path,
                "finalize.normalize_import_table",
                normalize_imports_started.elapsed(),
            );
        }
        let global_table_started = Instant::now();
        let _debug_globals = vybe_runtime::debugger::DebugPhase::current("finalize globals");
        match global_operands.as_ref() {
            Some(operands) => common::globals::finalize_global_table_with_operands(&mut self.chunks, operands),
            // No import rewrite scanned this unit. Keep the original global
            // discovery pass instead of first collecting a separate record.
            None => common::globals::finalize_global_table(&mut self.chunks),
        }
        drop(_debug_globals);
        if let Some(path) = timing.as_deref() {
            write_compiler_timing(path, "finalize.global_table", global_table_started.elapsed());
        }
        // Publish this UNIT's receiver ABI to every chunk it produced, on the
        // same principle as the global index space and canon section below: a
        // chunk that was compiled under an ABI must be able to say which one.
        //
        // ⛔ STAMPING ONLY `chunks[0]` WAS A REAL BUG, not a latent one.
        // `chunks[0]` is one slot for the whole BUNDLE, so a wast unit bundled
        // with a js entry unit inherited js's answer: measured by Fathom as
        // `m.add(2,3) → 2` on a wast export — `0 + 2`, the receiver sitting in
        // slot 0 and both arguments shifted one place — while js was flipped.
        // A wast export must not move because a js directive changed.
        let unit_abi = if self.universal_receiver() {
            vybe_runtime::chunk::ReceiverAbi::Parameter
        } else {
            vybe_runtime::chunk::ReceiverAbi::Ambient
        };
        let unit_members_on_prototype = self.members_on_prototype();
        let unit_fields_are_own_properties = self.instance_fields_are_own_properties();
        let stamp_chunks_started = Instant::now();
        for chunk in &mut self.chunks {
            chunk.module_receiver_abi = unit_abi;
            chunk.module_members_on_prototype = unit_members_on_prototype;
            chunk.module_instance_fields_are_own_properties = unit_fields_are_own_properties;
        }
        if let Some(path) = timing.as_deref() {
            write_compiler_timing(path, "finalize.stamp_chunks", stamp_chunks_started.elapsed(),);
        }
        // The canon section, published to every chunk on the same principle as
        // the global index space above: a chunk carrying a canonidx must be
        // able to say what that index means.
        let canon_started = Instant::now();
        let mut canon = std::mem::take(&mut self.canon_section);
        let decls = std::mem::take(&mut self.core_export_callees);
        common::canon::resolve_core_export_callees(&mut self.chunks, &decls, &mut canon)?;
        common::canon::install_canon_section(&mut self.chunks, canon);
        let fts = std::mem::take(&mut self.canon_functypes);
        let vts = std::mem::take(&mut self.canon_valtypes);
        let cfuncs = std::mem::take(&mut self.component_funcs);
        common::canon::install_type_space(&mut self.chunks, fts, vts, cfuncs);
        common::canon::install_component_func_links(
            &mut self.chunks,
            std::mem::take(&mut self.component_func_links),
        );
        common::canon::install_component_instance_links(
            &mut self.chunks,
            self.component_instance_links.take(),
        )?;
        if let Some(path) = timing.as_deref() {
            write_compiler_timing(path, "finalize.canon", canon_started.elapsed());
        }
        let host_imports_started = Instant::now();
        let host_imports = self.collected_host_imports();
        if let Some(path) = timing.as_deref() {
            write_compiler_timing(
                path,
                "finalize.collected_host_imports",
                host_imports_started.elapsed(),
            );
        }
        // Frame 0 is the module's own declaration (installed above from
        // `module.directives`); an in-source `Directive` with `Module` scope
        // writes through to it, so reading the base frame answers for the whole
        // unit however it was stated.
        let app_shell = self.directives.first().and_then(|d| d.app_shell);
        drop(_debug_finalize);
        if dump_classes_enabled() {
            self.dump_pending_classes();
        }
        if self.compact_source_lines {
            let started = Instant::now();
            for chunk in &mut self.chunks { chunk.lines.compact(); }
            if let Some(path) = timing.as_deref() {
                write_compiler_timing(path, "finalize.source_lines", started.elapsed());
            }
        }
        if let Some(path) = timing.as_deref() {
            write_compiler_timing(path, "finalize", finalize_started.elapsed());
            write_compiler_timing(path, "compile_with_imports_total", total_started.elapsed());
        }
        Ok(CompileResult {
            chunks: self.chunks,
            // Eligibility does not certify physical storage. Populate only
            // after final slot/access provenance is implemented and verified.
            word_local_proofs: Vec::new(),
            host_imports,
            component_source,
            component_binary,
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

impl Compiler {
    #[cfg(feature = "primitive-benchmarks")]
    pub fn benchmark_condition_truthiness_reuse(&mut self, enabled: bool) {
        self.condition_truthiness_reuse = enabled;
    }

    #[cfg(feature = "primitive-benchmarks")]
    pub fn benchmark_negated_condition_chunks(mut self, repeats: usize) -> Vec<Chunk> {
        let mut root = Chunk::new("condition-owner");
        functions::emit_ref_func(&mut root, 1, 0, 0);
        root.emit_op(Op::RETURN, 0);
        self.chunks = vec![root, Chunk::new("condition")];
        self.current = 1;
        let input = self.define_local("input");
        self.chunk().arity = 1;
        self.chunk().param_count = 1;
        for n in 0..repeats {
            self.emit_u16(Op::LOCAL_GET, input);
            self.emit_negated_condition();
            let line = self.line;
            ops::emit_i32_to_bool(self.chunk(), line);
            if n + 1 < repeats { self.emit(Op::DROP); }
        }
        self.emit(Op::RETURN);
        bundle::finalize_with_runtime_helpers(&mut self.chunks);
        globals::declare_free_globals(&mut self.chunks);
        globals::normalize_global_table(&mut self.chunks);
        self.chunks
    }

    #[cfg(feature = "primitive-benchmarks")]
    pub fn benchmark_integer_truthiness_reuse(&mut self, enabled: bool) {
        self.integer_truthiness_reuse = enabled;
    }

    /// Negation of a proven i32 needs only eqz; no dynamic coercion.
    pub(crate) fn emit_i32_negated_condition(&mut self) {
        #[cfg(feature = "primitive-benchmarks")]
        if !self.integer_truthiness_reuse {
            self.emit_negated_condition();
            return;
        }
        self.emit(Op::I32_EQZ);
    }

    /// Integer XOR is zero exactly when the two integer bit patterns match.
    /// Comparisons retain the same i32 operand conversions while producing 0/1.
    pub(crate) fn emit_i32_xor_condition(&mut self, negate: bool) {
        #[cfg(feature = "primitive-benchmarks")]
        if !self.integer_truthiness_reuse {
            self.emit(Op::I32_XOR);
            if negate {
                self.emit_negated_condition();
            } else {
                let line = self.line;
                ops::emit_dyn_to_bool(self.chunk(), line);
            }
            return;
        }
        self.emit(if negate { Op::I32_EQ } else { Op::I32_NE });
    }

    /// Negate a dynamic condition once: emit_dyn_not already performs ToBoolean.
    pub(crate) fn emit_negated_condition(&mut self) {
        let line = self.line;
        #[cfg(feature = "primitive-benchmarks")]
        if !self.condition_truthiness_reuse {
            ops::emit_dyn_to_bool(self.chunk(), line);
        }
        ops::emit_dyn_not(self.chunk(), line);
    }

    /// The arguments of a PARENT-CONSTRUCTOR call, when `expr` is one:
    /// `SuperCall { method: None, args }`, `SuperCall { method: Some(ctor) }`
    /// where `ctor` is the profile's constructor name / `new` / `__init__`
    /// (case per `case_sensitive`), or `Call { callee: Super, args }`.
    /// A `super.method(..)` call is NOT one and answers `None`. This is the
    /// single predicate: the constructor prologue and the `SuperCall`
    /// emitter agree on what a parent-constructor call is by construction.
    pub(crate) fn parent_ctor_call_args<'a>(&self, expr: &'a Expression) -> Option<&'a [Argument]> {
        match &expr.kind {
            ExprKind::SuperCall { method, args } => {
                let is_ctor = match method {
                    None => true,
                    Some(m) => {
                        let ctor_name = &self.profile.constructor_name;
                        if self.case_sensitive {
                            m == ctor_name || m == "new" || m == "__init__"
                        } else {
                            m.eq_ignore_ascii_case(ctor_name)
                                || m.eq_ignore_ascii_case("new")
                                || m.eq_ignore_ascii_case("__init__")
                        }
                    }
                };
                is_ctor.then_some(args.as_slice())
            }
            ExprKind::Call { callee, args, .. } if matches!(callee.kind, ExprKind::Super) => {
                Some(args.as_slice())
            }
            _ => None,
        }
    }

    /// `stmt` is an expression statement that is a parent-constructor call.
    pub(crate) fn is_parent_ctor_call_stmt(&self, stmt: &Statement) -> bool {
        match &stmt.kind {
            StmtKind::Expr(expr) => self.parent_ctor_call_args(expr).is_some(),
            _ => false,
        }
    }

    /// A constructor body contains a parent-constructor call statement.
    pub(crate) fn body_has_parent_ctor_call(&self, body: &[Statement]) -> bool {
        body.iter().any(|stmt| self.is_parent_ctor_call_stmt(stmt))
    }
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
    let mut preparation = declaration_preparation::PartialClassPreparation::new(case_sensitive, body.len());
    preparation.consume(body);
    preparation.finish()
}

#[cfg(test)]
fn merge_partial_classes_whole_module_oracle(body: &[Statement], case_sensitive: bool) -> Vec<Statement> {
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

#[cfg(test)]
mod closure_capture_shape_tests {
    use super::*;

    #[test]
    fn borrowed_object_method_body_preserves_receiver_and_outer_binding() {
        std::thread::Builder::new().stack_size(32 * 1024 * 1024).spawn(|| {
            vybe_language_js::register();
            let source = "let x = 7; const object = { value: 5, read() { return this.value + x; } }; if (object.read() !== 12) { throw new Error('method capture'); } x = 9; if (object.read() !== 14) { throw new Error('live capture'); }";
            let module = vybe_language_js::parse(source).unwrap();
            let profile = crate::profile::parse_profile(vybe_language_js::profile_source()).unwrap();
            let chunks = Compiler::with_profile(profile).compile(&module).unwrap();
            let mut vm = vybe_runtime::VM::new();
            platforms::register_platforms_all(&mut vm);
            vm.run(chunks).unwrap();
        }).unwrap().join().unwrap();
    }

    #[test]
    fn successive_and_nested_closures_preserve_capture_storage() {
        std::thread::Builder::new().stack_size(32 * 1024 * 1024).spawn(|| {
            vybe_language_php::register();
            for (source, expected) in [
                ("<?php $x = 7; $a = function() use ($x) { return $x; }; $x = 9; $b = function() use ($x) { return $x; }; return $a() * 10 + $b();", 79.0),
                ("<?php $x = 7; $a = function() use (&$x) { $x += 1; return $x; }; $b = function() use (&$x) { $x += 2; return $x; }; $a(); return $b();", 10.0),
                ("<?php $x = 7; $outer = function() use ($x) { return function() use ($x) { return $x; }; }; $x = 9; $inner = $outer(); return $inner();", 7.0),
            ] {
                let module = vybe_language_php::parse(source).unwrap();
                let profile = crate::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
                let chunks = Compiler::with_profile(profile).compile(&module).unwrap();
                let mut vm = vybe_runtime::VM::new();
                platforms::register_platforms_all(&mut vm);
                let result = vm.run(chunks).unwrap();
                assert_eq!(result.as_f64(), expected, "capture storage changed for {source}");
            }
        }).unwrap().join().unwrap();
    }

    #[test]
    fn addressed_places_preserve_root_and_index_captures() {
        let body = Expression::new(ExprKind::Tuple(vec![
            Expression::new(ExprKind::RefOf(Box::new(PlaceExpr::Ident("destination".into())))),
            Expression::new(ExprKind::RefOf(Box::new(PlaceExpr::Member {
                object: Box::new(Expression::ident("record")),
                field: "value".into(), null_safe: false,
            }))),
            Expression::new(ExprKind::RefOf(Box::new(PlaceExpr::Index {
                object: Box::new(Expression::ident("array")),
                index: Box::new(Expression::ident("index")), null_safe: false,
            }))),
            Expression::new(ExprKind::RefOf(Box::new(PlaceExpr::Deref(Box::new(Expression::ident("pointer")))))),
        ]));
        let lambda = Expression::new(ExprKind::Lambda {
            params: vec![], body: LambdaBody::Expr(Box::new(body)),
            is_async: false, captures: vec![],
        });
        let mut captured = HashSet::new();
        collect_closure_captured_in_expr(&lambda, &mut captured);
        assert_eq!(captured, ["destination", "record", "array", "index", "pointer"]
            .into_iter().map(str::to_string).collect());
    }

    #[test]
    fn typed_reference_and_tuple_children_preserve_captured_identifiers() {
        let member = Expression::new(ExprKind::Member {
            object: Box::new(Expression::new(ExprKind::RefLoad(Box::new(
                Expression::new(ExprKind::Cast {
                    expr: Box::new(Expression::new(ExprKind::Ident("pointer".into()))),
                    type_name: "record*".into(),
                }),
            )))),
            field: "value".into(),
            null_safe: false,
        });
        let body = Expression::new(ExprKind::Tuple(vec![member,
            Expression::new(ExprKind::IsType {
                expr: Box::new(Expression::new(ExprKind::Ident("other".into()))),
                type_name: "record".into(),
            }),
            Expression::new(ExprKind::RefLoad(Box::new(
                Expression::new(ExprKind::Ident("local".into())),
            ))),
        ]));
        let lambda = Expression::new(ExprKind::Lambda {
            params: vec![crate::ast::Param {
                name: "local".into(), type_hint: None, default: None,
                pass_by: vybe_ast::PassBy::Value, is_rest: false,
                is_kwargs: false, is_optional: false, is_nullable: false,
            }], body: LambdaBody::Expr(Box::new(body)),
            is_async: false, captures: vec![],
        });
        let expected = HashSet::from(["pointer".to_string(), "other".to_string()]);
        let mut captured = HashSet::new();
        collect_closure_captured_in_expr(&lambda, &mut captured);
        assert_eq!(captured, expected);
        let wrapped = Expression::new(ExprKind::Tuple(vec![Expression::new(ExprKind::Cast {
            expr: Box::new(Expression::new(ExprKind::RefLoad(Box::new(lambda)))),
            type_name: "callable".into(),
        })]));
        captured.clear();
        collect_closure_captured_in_expr(&wrapped, &mut captured);
        assert_eq!(captured, expected);
    }
}

#[cfg(test)]
mod borrowed_body_tests {
    use super::*;

    #[test]
    #[ignore = "local WordPress compiler method diagnostic, not a timing benchmark"]
    fn wordpress_compiler_method_profile() {
        std::thread::Builder::new().stack_size(32 * 1024 * 1024).spawn(|| {
            vybe_language_php::register();
            let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../examples/webroot/wordpress/wp-includes");
            for name in ["class-wp-hook.php", "class-wp-query.php", "formatting.php"] {
                let source = std::fs::read_to_string(root.join(name)).unwrap();
                let module = vybe_language_php::parse(&source).unwrap();
                let profile = crate::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
                let report = Arc::new(std::sync::Mutex::new(
                    vybe_runtime::debugger::DebugReport::default()));
                {
                    let _scope = vybe_runtime::debugger::DebugReportScope::enter(Some(report.clone()));
                    std::hint::black_box(compile_mode(&module, profile, false));
                }
                let mut report = report.lock().unwrap();
                assert!(report.active.is_empty());
                let mut methods = std::mem::take(&mut report.completed).into_iter()
                    .filter(|(label, _)| label.starts_with("compiler method "))
                    .collect::<Vec<_>>();
                methods.sort_unstable_by(|a, b| b.1.cmp(&a.1));
                for (label, elapsed) in methods.into_iter().take(8) {
                    eprintln!("compiler_method_diagnostic file={name} wall_ms={:.3} label={label}",
                        elapsed.as_secs_f64() * 1000.0);
                }
            }
        }).unwrap().join().unwrap();
    }

    fn compile_mode(module: &Module, profile: vybe_runtime::profile::LanguageProfile, copy: bool) -> Vec<Chunk> {
        struct Reset(bool);
        impl Drop for Reset {
            fn drop(&mut self) { COPY_MODULE_BODY_BASELINE.with(|flag| flag.set(self.0)); }
        }
        let _reset = Reset(COPY_MODULE_BODY_BASELINE.with(|flag| flag.replace(copy)));
        Compiler::with_profile(profile).compile(module).unwrap()
    }

    #[test]
    fn class_registry_snapshot_preserves_lookup_and_detaches_only_on_write() {
        std::thread::Builder::new().stack_size(32 * 1024 * 1024).spawn(|| {
            vybe_language_php::register();
            let module = vybe_language_php::parse(
                "<?php class Box { public function read() { return 7; } }"
            ).unwrap();
            let profile = crate::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
            let mut compiler = Compiler::with_profile(profile);
            compiler.predeclare_type_names(&module.body, None);
            let key = compiler.canon("Box");
            let declarations = compiler.normalized_classes.clone();
            assert!(std::sync::Arc::ptr_eq(&declarations.0, &compiler.normalized_classes.0));
            let previous_name = declarations.get(&key).unwrap().name.clone();
            compiler.normalized_classes.get_mut(&key).unwrap().name = "Changed".into();
            assert_eq!(declarations.get(&key).unwrap().name, previous_name);
            assert_eq!(compiler.normalized_classes.get(&key).unwrap().name, "Changed");
            assert!(!std::sync::Arc::ptr_eq(&declarations.0, &compiler.normalized_classes.0));
        }).unwrap().join().unwrap();
    }

    #[cfg(target_os = "macos")]
    #[test]
    #[ignore = "paired compilation CPU measurement on local WordPress modules"]
    fn paired_wordpress_body_storage_cpu() {
        fn cpu_ns() -> u128 {
            unsafe extern "C" { fn clock_gettime(id: i32, time: *mut [i64; 2]) -> i32; }
            let mut time = [0i64; 2];
            assert_eq!(unsafe { clock_gettime(16, &mut time) }, 0);
            time[0] as u128 * 1_000_000_000 + time[1] as u128
        }
        std::thread::Builder::new().stack_size(32 * 1024 * 1024).spawn(|| {
            vybe_language_php::register();
            let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../examples/webroot/wordpress/wp-includes");
            for name in ["class-wp-hook.php", "class-wp-query.php", "formatting.php"] {
                let source = std::fs::read_to_string(root.join(name)).unwrap();
                let module = vybe_language_php::parse(&source).unwrap();
                let profile = crate::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
                let mut samples = [Vec::new(), Vec::new()];
                for round in 0..7 {
                    for copy in if round % 2 == 0 { [true, false] } else { [false, true] } {
                        let start = cpu_ns();
                        let chunks = compile_mode(&module, profile.clone(), copy);
                        std::hint::black_box(&chunks);
                        drop(chunks);
                        let elapsed = cpu_ns() - start;
                        if round != 0 { samples[copy as usize].push(elapsed); }
                    }
                }
                for values in &mut samples { values.sort_unstable(); }
                eprintln!("module_body file={name} copied_cpu_ms={:.3} borrowed_cpu_ms={:.3} ratio={:.3}",
                    samples[1][3] as f64 / 1e6, samples[0][3] as f64 / 1e6,
                    samples[0][3] as f64 / samples[1][3] as f64);
            }
        }).unwrap().join().unwrap();
    }

    #[test]
    fn borrowed_segments_match_previous_copy_for_php_reference_outputs() {
        std::thread::Builder::new().stack_size(32 * 1024 * 1024).spawn(|| {
            vybe_language_php::register();
            for source in [
                "<?php parse_str('', $result); return is_array($result);",
                "<?php class Box { private $value = 7; public function read() { return $this->value; } } $box = new Box(); return $box->read();",
                "<?php function parse_wrapper($input, &$output) { parse_str($input, $output); } $result = null; parse_wrapper('', $result); return is_array($result);",
                "<?php function parse_wrapper($input, &$output) { parse_str($input, $output); } function options($input, $defaults) { parse_wrapper($input, $parsed); return array_merge($defaults, $parsed); } return options('', ['optimize' => false]);",
            ] {
                let module = vybe_language_php::parse(source).unwrap();
                let original = format!("{module:?}");
                let profile = crate::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
                let baseline = compile_mode(&module, profile.clone(), true);
                let borrowed = compile_mode(&module, profile, false);
                assert_eq!(baseline.len(), borrowed.len());
                for (a, b) in baseline.iter().zip(&borrowed) {
                    assert_eq!(a.code, b.code, "emitted instruction bytes changed for {source}");
                    assert_eq!(format!("{:?}", a.constants), format!("{:?}", b.constants));
                    assert_eq!(format!("{:?}", a.imports), format!("{:?}", b.imports));
                }
                let run = |chunks| {
                    let mut vm = vybe_runtime::VM::new();
                    platforms::register_platforms_all(&mut vm);
                    vm.run(chunks).map(|value| format!("{value:?}")).map_err(|error| error.to_string())
                };
                assert_eq!(run(baseline), run(borrowed), "runtime changed for {source}");
                assert_eq!(format!("{module:?}"), original, "compilation changed the input AST");
            }
        }).unwrap().join().unwrap();
    }
}

#[cfg(test)]
mod staged_diagnostic_tests {
    use super::*;
    #[test]
    fn speculative_compiler_buffers_diagnostics_for_actual_include() {
        let certificate=Arc::new(std::sync::Mutex::new(crate::compile_dependencies::Certificate::speculative()));
        let compiler=Compiler::with_profile(crate::profile::parse_profile("[info]\nname='diagnostic-test'\n[compiler]\n").unwrap()).with_compile_certificate(certificate.clone());
        compiler.compiler_diagnostic("unreached diagnostic".into());
        assert_eq!(certificate.lock().unwrap().take_diagnostics(),["unreached diagnostic"]);
        assert!(certificate.lock().unwrap().take_diagnostics().is_empty());
    }
    #[test]
    fn actual_namespace_call_records_selected_and_missing_export_dependencies() {
        use vybe_runtime::{ExportEntry,ModuleRecord};
        fn compile(graph:&HashMap<String,ModuleRecord>)->Arc<std::sync::Mutex<crate::compile_dependencies::Certificate>> {
            let certificate=Arc::new(std::sync::Mutex::new(crate::compile_dependencies::Certificate::speculative()));
            let mut compiler=Compiler::with_profile(crate::profile::parse_profile("[info]\nname='certificate-test'\n[compiler]\n").unwrap())
                .with_module_exports(crate::bundle::flatten_module_exports(graph)).with_compile_certificate(certificate.clone());
            compiler.host_namespace_aliases.insert("pkg".into(),"adapter".into());
            let member=Expression::new(ExprKind::Member{object:Box::new(Expression::ident("pkg")),field:"f".into(),null_safe:false});
            let call=Expression::new(ExprKind::Call{callee:Box::new(member),args:vec![],optional:false});
            let module=Module{name:"test".into(),language:Lang::Unknown,body:vec![Statement::new(StmtKind::Expr(call))],imports:vec![],directives:Default::default(),canon:Default::default()};
            compiler.compile_with_imports(&module).unwrap();certificate
        }
        let mut graph=HashMap::new();
        graph.insert("adapter".into(),ModuleRecord::new_synthetic("adapter"));
        graph.insert("host".into(),ModuleRecord::new_synthetic("host"));
        let missing=compile(&graph);assert!(missing.lock().unwrap().validates(&graph));
        graph.get_mut("host").unwrap().exports.insert("terminal".into(),ExportEntry::Function{idx:1});
        graph.get_mut("adapter").unwrap().exports.insert("f".into(),ExportEntry::Indirect{from:"host".into(),name:"terminal".into()});
        assert!(!missing.lock().unwrap().validates(&graph));
        let selected=compile(&graph);assert!(selected.lock().unwrap().validates(&graph));
        graph.get_mut("adapter").unwrap().exports.insert("f".into(),ExportEntry::Function{idx:2});
        assert!(!selected.lock().unwrap().validates(&graph));
    }

}

pub mod memory_span_set;

// Generic verification code only; no language crate dependency and no normal
// build API or runtime state. The caller supplies an actual registered producer.
#[cfg(any(test,feature="service-verification"))]
mod helper_direct_frame_policy_verification {
    use super::*;
    use vybe_runtime::registry::{ParseFn,EmitDispatchFn};

    #[derive(Debug)]
    pub struct Report {
        pub direct_calls:usize,
        pub helper_chunks:usize,
        pub scope_next_slot:u16,
        pub local_count:u16,
        pub scratch_high_water:u16,
        pub escaping_scope_next_slot:u16,
        pub rooted_result:String,
    }
    fn key(_: &[Chunk],_:u32)->Vec<u8>{vec![1]}
    fn build(chunks:&mut Vec<Chunk>,line:u32)->usize {
        let mut body=Chunk::new("__frame_policy_probe");
        body.arity=2;body.param_count=2;body.local_count=2;
        body.takes_receiver=false;body.direct_entry_has_captures=Some(false);
        body.emit_op_u16(Op::LOCAL_GET,0,line);body.emit_op(Op::RETURN,line);
        let index=chunks.len();chunks.push(body);index
    }
    fn no_factory()->Chunk{panic!("contextual factory unused")}
    fn no_caller(_: &mut Chunk,_:usize,_:u32){panic!("contextual caller unused")}
    fn escaping_invoke(chunks:&mut Vec<Chunk>,current:usize,_:usize,_:u32){chunks[current].local_count+=3;}
    fn compiler(direct:bool,parser:ParseFn,producer:EmitDispatchFn,namespace:&'static str,operation:&'static str)->Compiler {
        let definitions=Box::leak(vec![helper_templates::Definition {
            contract:if direct {helper_templates::Contract::ContextualDirectCallAppendOnlyUnit{key,build}}
                else {helper_templates::Contract::ContextualAppendOnlyUnit{key,build,invoke:escaping_invoke}},
            operation,namespace,version:1,parser,producer,arity:2,
            helper_name:"__frame_policy_probe",factory:no_factory,caller:no_caller,
        }].into_boxed_slice());
        Compiler::with_profile(vybe_runtime::profile::parse_profile(
            "[info]\nname='helper-frame-verification'\n[compiler]").unwrap())
            .with_helper_templates(parser,definitions,std::sync::Arc::new(helper_templates::Pool::default()))
    }
    pub fn verify(parser:ParseFn,producer:EmitDispatchFn,namespace:&'static str,operation:&'static str)->Result<Report,String> {
        if operation.split_once('.').map(|(prefix,_)|prefix)!=Some(namespace){return Err("operation namespace mismatch".into());}
        if !crate::languages::emit_dispatch_for(namespace).is_some_and(|actual|std::ptr::fn_addr_eq(actual,producer)) {
            return Err("verification producer is not registered for namespace".into());
        }
        // Actual Compiler entry, builtin dispatch, Scope allocator and checked
        // direct helper path; not an extracted or mirrored scope model.
        let mut c=compiler(true,parser,producer,namespace,operation);
        let root=c.define_local("root");assert_eq!(root,0);
        c.chunks[0].local_count=11;c.chunks[0].scratch_high_water=17;
        c.emit_const(vybe_runtime::Value::String(std::sync::Arc::from("rooted")));
        c.emit_u16(Op::LOCAL_SET,root);
        for _ in 0..20 {
            // A live local and a nested call result survive new strings
            // allocated in both argument positions before each direct call.
            c.emit_u16(Op::LOCAL_GET,root);
            c.emit_const(vybe_runtime::Value::String(std::sync::Arc::from("in")));
            c.emit_const(vybe_runtime::Value::String(std::sync::Arc::from("ner")));
            c.emit(Op::STRING_CONCAT);c.emit_common(operation,2,13);
            c.emit_const(vybe_runtime::Value::String(std::sync::Arc::from("out")));
            c.emit_const(vybe_runtime::Value::String(std::sync::Arc::from("er")));
            c.emit(Op::STRING_CONCAT);c.emit_common(operation,2,29);c.emit(Op::DROP);
            assert_eq!(c.scopes[0].next_slot,1);
            assert_eq!(c.chunks[0].local_count,11);
            assert_eq!(c.chunks[0].scratch_high_water,17);
        }
        let direct_calls=c.chunks[0].direct_defined_calls.len();assert_eq!(direct_calls,40);
        let helper_chunks=c.chunks.len()-1;assert_eq!(helper_chunks,1);
        let scope_next_slot=c.scopes[0].next_slot;
        let local_count=c.chunks[0].local_count;
        let scratch_high_water=c.chunks[0].scratch_high_water;
        c.emit_u16(Op::LOCAL_GET,root);c.emit(Op::RETURN);
        let mut vm=vybe_runtime::VM::new();
        let value=vm.run(c.chunks).map_err(|e|e.to_string())?;
        let vybe_runtime::Value::String(rooted)=value else {return Err("rooted result is not a string".into());};
        if rooted.as_ref()!="rooted" {return Err("rooted value changed across nested helper arguments".into());}
        let rooted_result=rooted.to_string();
        // Old arbitrary adapters may leave live/escaping caller locals: they
        // must still reserve Scope slots and retain their scratch high-water.
        let mut old=compiler(false,parser,producer,namespace,operation);
        old.scopes[0].next_slot=1;old.chunks[0].local_count=11;
        old.emit_common(operation,2,13);
        assert_eq!(old.chunks[0].local_count,14);assert_eq!(old.scopes[0].next_slot,14);
        assert!(old.chunks[0].scratch_high_water>=14);
        Ok(Report {direct_calls,helper_chunks,scope_next_slot,local_count,scratch_high_water,
            escaping_scope_next_slot:old.scopes[0].next_slot,rooted_result})
    }
}
#[cfg(feature="service-verification")]
pub use helper_direct_frame_policy_verification::Report as HelperDirectFrameScopeReport;
/// Feature-only native acceptance entry. Caller registers its real frontend
/// and supplies its exact parse/dispatch identities; no normal API is added.
#[cfg(feature="service-verification")]
pub fn verify_helper_direct_frame_scope(parser:vybe_runtime::registry::ParseFn,
    producer:vybe_runtime::registry::EmitDispatchFn,namespace:&'static str,operation:&'static str,
)->Result<HelperDirectFrameScopeReport,String>{
    helper_direct_frame_policy_verification::verify(parser,producer,namespace,operation)
}

#[cfg(any(test,feature="service-verification"))]
mod helper_typed_direct_verification {
    use super::*;
    use vybe_runtime::registry::{ParseFn,EmitDispatchFn,emitter_results::ResultRepresentation};
    #[derive(Debug)]pub struct Case {
        pub result:ResultRepresentation,pub mode:&'static str,pub calls:usize,pub helpers:usize,
        pub scope:u16,pub locals:u16,pub high_water:u16,pub boolean_casts:usize,pub drops:usize,pub returned:String,
    }
    fn key(_: &[Chunk],_:u32)->Vec<u8>{Vec::new()}
    fn boxed(chunks:&mut Vec<Chunk>,line:u32)->usize{
        let mut c=Chunk::new("__typed_frame_probe_body");c.arity=0;c.param_count=0;c.local_count=0;c.takes_receiver=false;c.direct_entry_has_captures=Some(false);c.emit_bool_const(true,line);c.emit_op(Op::RETURN,line);let i=chunks.len();chunks.push(c);i
    }
    fn unused_factory()->Chunk{panic!("typed contextual helper has no standalone factory")}
    fn unused_caller(_: &mut Chunk,_:usize,_:u32){panic!("typed direct helper has no arbitrary caller")}
    fn compiler(parser:ParseFn,producer:EmitDispatchFn,namespace:&'static str,operation:&'static str,result:ResultRepresentation)->Compiler{
        let build=boxed;
        let defs=Box::leak(vec![helper_templates::Definition{contract:helper_templates::Contract::ContextualTypedDirectCallAppendOnlyUnit{key,build,result},operation,namespace,version:1,parser,producer,arity:0,helper_name:"__typed_frame_probe_body",factory:unused_factory,caller:unused_caller}].into_boxed_slice());
        let profile=format!("[info]\nname='typed-helper-proof'\n[compiler]\n[builtins]\n__typed_frame_probe={{emit='common:{operation}',discarded_emit='common:{operation}',min_args=0,max_args=0}}\n");
        Compiler::with_profile(vybe_runtime::profile::parse_profile(&profile).unwrap()).with_helper_templates(parser,defs,std::sync::Arc::new(helper_templates::Pool::default()))
    }
    pub fn verify(parser:ParseFn,producer:EmitDispatchFn,namespace:&'static str,operation:&'static str,check_code:fn(&[Chunk])->Result<(),String>)->Result<Vec<Case>,String>{
        if operation.split_once('.').map(|(prefix,_)|prefix)!=Some(namespace){return Err("operation namespace mismatch".into());}
        if !crate::languages::emit_dispatch_for(namespace).is_some_and(|actual|std::ptr::fn_addr_eq(actual,producer)){return Err("typed verification producer is not registered".into());}
        let mut reports=Vec::new();
        for result in [ResultRepresentation::BoxedBoolean]{for mode in ["value","condition","discarded-profile"]{
            let mut c=compiler(parser,producer,namespace,operation,result);
            let root=c.define_local("root");assert_eq!(root,0);c.chunks[0].local_count=11;c.chunks[0].scratch_high_water=17;
            c.emit_const(vybe_runtime::Value::String(std::sync::Arc::from("rooted")));c.emit_u16(Op::LOCAL_SET,root);
            let call=vybe_ast::Expression::new(vybe_ast::ExprKind::Call{callee:Box::new(vybe_ast::Expression::ident("__typed_frame_probe")),args:Vec::new(),optional:false});
            match mode {
                "value"=>c.compile_expr(&call)?,
                "condition"=>{
                    c.want_i32_condition=true;c.compile_expr(&call)?;assert!(c.gave_i32_condition,"typed Boolean must reach the ordinary Compiler condition report");
                    c.chunks[0].emit_if(17);c.emit_const(vybe_runtime::Value::String(std::sync::Arc::from("selected")));c.emit_u16(Op::LOCAL_SET,root);c.chunks[0].emit_end(17);c.emit_u16(Op::LOCAL_GET,root);
                },
                "discarded-profile"=>{assert!(c.has_discarded_builtin_emits);c.compile_expr_stmt(&call)?;c.emit_u16(Op::LOCAL_GET,root);},
                _=>unreachable!(),
            }
            assert_eq!(c.scopes[0].next_slot,1);assert_eq!(c.chunks[0].local_count,11);assert_eq!(c.chunks[0].scratch_high_water,17);
            let calls=c.chunks[0].direct_defined_calls.len();let helpers=c.chunks.len()-1;assert_eq!(calls,1);assert_eq!(helpers,1);
            let boolean_casts=c.chunks[0].imports.iter().filter(|i|i.module=="wasm:js-boolean"&&i.name=="cast").count();assert_eq!(boolean_casts,usize::from(mode=="condition"&&result==ResultRepresentation::BoxedBoolean));
            let mut drops=0;let mut ip=0;while ip<c.chunks[0].code.len(){let op=Op::from_code_at(&c.chunks[0].code,ip);drops+=usize::from(op==Op::DROP);ip+=4+op.operand_format().size_in(&c.chunks[0].code,ip+4);}assert_eq!(drops,usize::from(mode=="discarded-profile"));
            c.emit(Op::RETURN);check_code(&c.chunks)?;let mut vm=vybe_runtime::VM::new();vybe_runtime::js_builtins::register(&mut vm);vybe_platform_ecma::register(&mut vm);
            let missing:Vec<_>=c.chunks.iter().flat_map(|chunk|chunk.imports.iter()).filter_map(|import|vm.resolve_import_target(&import.module,&import.name).err().map(|e|format!("{}:{}: {e}",import.module,import.name))).collect();
            assert!(missing.is_empty(),"all missing actual provider bindings before VM execution: {missing:?}");let value=vm.run(c.chunks).map_err(|e|e.to_string())?;
            match mode {
                "value"=>match result{ResultRepresentation::BoxedBoolean=>assert!(matches!(&value,vybe_runtime::Value::Bool(true))),ResultRepresentation::RawI32Boolean=>assert!(matches!(&value,vybe_runtime::Value::I32(1))),_=>unreachable!()},
                "condition"=>assert!(matches!(&value,vybe_runtime::Value::String(s)if s.as_ref()=="selected")),
                "discarded-profile"=>assert!(matches!(&value,vybe_runtime::Value::String(s)if s.as_ref()=="rooted")),
                _=>unreachable!(),
            }
            reports.push(Case{result,mode,calls,helpers,scope:1,locals:11,high_water:17,boolean_casts,drops,returned:format!("{value:?}")});
        }}
        Ok(reports)
    }
}
#[cfg(feature="service-verification")]
pub use helper_typed_direct_verification::Case as HelperTypedDirectScopeCase;
/// Feature-only actual Compiler value/condition/profile-discard result gate.
#[cfg(feature="service-verification")]
pub fn verify_helper_typed_direct_scope(parser:vybe_runtime::registry::ParseFn,producer:vybe_runtime::registry::EmitDispatchFn,namespace:&'static str,operation:&'static str,check_code:fn(&[Chunk])->Result<(),String>)->Result<Vec<HelperTypedDirectScopeCase>,String>{helper_typed_direct_verification::verify(parser,producer,namespace,operation,check_code)}
