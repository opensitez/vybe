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

macro_rules! inst {
    ($self:expr, $($path:ident)::+ $(, $arg:expr)*) => {{
        crate::emitter::instructions::$($path)::+(&mut $self.chunks[$self.current], $self.line $(, $arg)*)
    }};
}

macro_rules! fn_call {
    ($self:expr, $module:literal, $name:literal, $argc:expr) => {{
        crate::emitter::instructions::host::CapabilityContext::get()
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
mod bindings;
mod builtins;
mod calls;
mod class_context;
pub mod class_normalize; // cross-language class normalisation (was crate::common::classes)
mod classes;
mod control_flow;
mod dotnet_calls;
mod emit_helpers;
mod events;
mod expressions;
mod lambdas;
mod link;
mod metadata;
mod operators;
mod overloads;
mod php_lang;
mod resolver;
mod scope;
mod statements;
mod type_inference;

use crate::ast::*;
use crate::compiler::scope::Scope;
use crate::emitter as common;
#[allow(unused_imports)]
use crate::emitter::instructions as inst;
use crate::emitter::loops::LoopState;
use crate::profile::*;
use std::collections::{BTreeSet, HashMap, HashSet};
use std::sync::Arc;
use vybe_bytecode::chunk::Import as BytecodeImport;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

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
    /// Type hints for instance fields, keyed by canonical field name.
    /// Used when implicit-self resolution turns a bare field name into
    /// `this.<field>` so member access keeps the original receiver type.
    instance_field_types: HashMap<String, String>,
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
}

impl CallSignature {
    fn from_params(params: &[Param]) -> Self {
        Self {
            param_names: params
                .iter()
                .map(|param| param.name.trim_start_matches('$').to_string())
                .collect(),
            param_defaults: params.iter().map(|param| param.default.clone()).collect(),
            min_arity: params
                .iter()
                .take_while(|param| param.default.is_none() && !param.is_rest)
                .count(),
            has_rest: params.last().is_some_and(|param| param.is_rest),
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
}

#[derive(Debug, Clone, Default)]
pub(crate) struct ReflectionMethodMetadata {
    pub decorators: Vec<Expression>,
    pub params: Vec<ReflectionParamMetadata>,
    pub is_static: bool,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct ReflectionMemberMetadata {
    pub decorators: Vec<Expression>,
    pub is_static: bool,
    pub can_write: bool,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct ReflectionConstructorMetadata {
    pub param_types: Vec<String>,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct ReflectionTypeMetadata {
    pub parents: Vec<String>,
    pub decorators: Vec<Expression>,
    pub interfaces: Vec<String>,
    pub nested_types: Vec<String>,
    pub constructors: Vec<ReflectionConstructorMetadata>,
    pub is_value_type: bool,
    pub methods: HashMap<String, ReflectionMethodMetadata>,
    pub properties: HashMap<String, ReflectionMemberMetadata>,
    pub fields: HashMap<String, ReflectionMemberMetadata>,
}

#[derive(Debug, Clone)]
pub(crate) enum ReflectionBinding {
    Type(String),
    Constructor {
        type_name: String,
        #[allow(dead_code)]
        param_types: Vec<String>,
    },
    Method {
        type_name: String,
        method_name: String,
    },
    Property {
        type_name: String,
        property_name: String,
    },
    Field {
        type_name: String,
        field_name: String,
    },
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
    loops: Vec<LoopCtx>,
    loop_states: Vec<LoopState>,
    label_depth: u32,
    function_label_base: u32,
    pub(crate) line: u32,
    pub(crate) defined_globals: HashSet<String>,
    const_globals: HashSet<String>,
    in_strict: bool,
    /// True while compiling the operand of a `typeof`. `typeof undeclaredName`
    /// must evaluate to `"undefined"`, never throw — so the unresolvable-binding
    /// ReferenceError in `emit_var_get` is suppressed in this context.
    in_typeof_operand: bool,
    /// Every name the program lexically declares (`let`/`const`/`var`/params/
    /// etc.) as a local, across all scopes — populated only for languages with
    /// `unresolved_reference_throws`. A name in this set that is unresolvable in
    /// the current scope is provably an out-of-scope user binding (never an
    /// untracked host global), so reading it is a ReferenceError even in sloppy
    /// mode (§9.1.1.4.6 applies in both strict and sloppy). See `emit_var_get`.
    program_lexical_names: HashSet<String>,

    shared_global_slots: HashMap<String, u16>,
    shared_global_names: Vec<String>,
    pub(crate) defined_functions: HashSet<String>,
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
    /// Classes declaring an index operator (`operator []` / `__getitem__`).
    /// Indexing one of these is a method call, not a key lookup — resolved
    /// from the receiver's static type so arrays, dicts and strings keep the
    /// plain index path with no runtime probe.
    pub(crate) classes_with_indexer: HashSet<String>,
    /// Classes declaring an index *setter* (`operator []=` / `__setitem__`).
    /// Kept apart from `classes_with_indexer` — a class may define either
    /// half on its own.
    pub(crate) classes_with_index_setter: HashSet<String>,
    global_type_hints: HashMap<String, String>,
    /// Map from member name → containing namespace name.
    /// Used for bare-name resolution within modules/namespaces/enums.
    /// E.g. `Main` inside `Module Program` resolves to `Program.Main`.
    /// `Green` inside `enum TColor` resolves to `TColor.Green`.
    /// Models the WASM Component Model's namespace-scoped imports.
    enum_members: HashMap<String, String>,
    /// Reverse enum lookup: enum type -> underlying integer -> member name.
    enum_value_names: HashMap<String, HashMap<i64, String>>,
    enum_flags: HashSet<String>,
    pub(crate) reflection_types: HashMap<String, ReflectionTypeMetadata>,
    pub(crate) attribute_usage: HashMap<String, AttributeUsageMetadata>,
    pub(crate) reflection_bindings: HashMap<String, ReflectionBinding>,
    case_sensitive: bool,
    pub(crate) profile: LanguageProfile,
    current_func_name: Option<String>,
    current_result_slot: Option<u16>,
    current_ref_out_params: Option<Vec<u16>>,
    pending_classes: HashMap<String, PendingClass>,
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
    php_function_globals: Vec<HashSet<String>>,
    array_bindings: HashMap<String, ArrayBindingMetadata>,
    /// Label for the next loop to be pushed (set by StmtKind::Labeled).
    pending_label: Option<String>,
    with_targets: Vec<u16>,
    capture_by_value_vars: Vec<String>,
    capture_locals: HashMap<u8, u16>,
    closure_env_names: Vec<String>,
    /// Shared env array for the current outer function: holds locals that
    /// are captured by inner closures. All reads/writes of these locals
    /// go through array.get/array.set so mutations are visible to closures.
    shared_env_slot: Option<u16>,
    shared_env_names: Vec<String>,
    pointer_cell_bindings: HashMap<usize, HashSet<String>>,

    /// Names of locals/params in the function currently being compiled whose
    /// address is taken somewhere in the body (`&v`). Populated by a pre-scan
    /// at function entry. Such bindings are promoted to a pointer cell *once*
    /// at their declaration (and params at entry) rather than lazily at the
    /// first `&v` use — taking the address inside a loop would otherwise
    /// re-wrap the cell every iteration and orphan prior mutations.
    current_addr_taken_locals: HashSet<String>,
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
    host_const_bindings: HashMap<String, vybe_bytecode::Value>,
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
    /// Snapshot of the current module's source imports.
    ///
    /// Used for narrow source-shape decisions that depend on the ambient
    /// framework surface, such as WinForms form inference for bare VB/C#
    /// classes inside a module that explicitly imports System.Windows.Forms.
    current_module_imports: Vec<Import>,
    /// JS-only: set when the module references `new Proxy(...)`. Member /
    /// Index reads + writes route through `emitter::js::proxy_adapter`
    /// for runtime trap dispatch. Off → direct `STRUCT_GET` / `ARRAY_GET`
    /// (zero overhead for non-Proxy code paths).
    pub(crate) uses_proxy: bool,
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
    module_value_exports: HashMap<String, HashMap<String, vybe_bytecode::Value>>,
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
        StmtKind::Assign { targets, value } => {
            for t in targets {
                collect_addr_taken_in_expr(t, out);
            }
            collect_addr_taken_in_expr(value, out);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            collect_addr_taken_in_expr(target, out);
            collect_addr_taken_in_expr(value, out);
        }
        StmtKind::Block(stmts) => collect_addr_taken_idents(stmts, out),
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
            ClassMember::Field { init: Some(init), .. } => {
                collect_closure_captured_in_expr(init, out)
            }
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
        StmtKind::Assign { targets, value } => {
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
        ExprKind::This => true,
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
        StmtKind::Assign { targets, value } => {
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
        ExprKind::New { class, args } => {
            if let ExprKind::Ident(name) = &class.kind {
                if name == "Proxy" {
                    return true;
                }
            }
            args.iter().any(|a| expr_uses_proxy(&a.value))
        }
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
        StmtKind::Assign { targets, value } => {
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
            chunks: vec![Chunk::new("<script>")],
            scopes: vec![Scope::new()],
            current: 0,
            loops: Vec::new(),
            loop_states: Vec::new(),
            label_depth: 0,
            function_label_base: 0,
            line: 1,
            defined_globals: HashSet::new(),
            const_globals: HashSet::new(),
            in_strict: false,
            in_typeof_operand: false,
            program_lexical_names: HashSet::new(),
            shared_global_slots: HashMap::new(),
            shared_global_names: Vec::new(),
            defined_functions: HashSet::new(),
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
            classes_with_indexer: HashSet::new(),
            classes_with_index_setter: HashSet::new(),
            global_type_hints: HashMap::new(),
            enum_members: HashMap::new(),
            enum_value_names: HashMap::new(),
            enum_flags: HashSet::new(),
            reflection_types: HashMap::new(),
            attribute_usage: HashMap::new(),
            reflection_bindings: HashMap::new(),
            case_sensitive: profile.case_sensitive,
            profile,
            current_func_name: None,
            current_result_slot: None,
            current_ref_out_params: None,
            pending_classes: HashMap::new(),
            current_class: None,
            current_namespace: None,
            current_class_implicit_self: false,
            current_member_is_static: false,
            static_local_bindings: Vec::new(),
            php_function_globals: Vec::new(),
            array_bindings: HashMap::new(),
            pending_label: None,
            with_targets: Vec::new(),
            capture_by_value_vars: Vec::new(),
            capture_locals: HashMap::new(),
            closure_env_names: Vec::new(),
            shared_env_slot: None,
            shared_env_names: Vec::new(),
            pointer_cell_bindings: HashMap::new(),
            current_addr_taken_locals: HashSet::new(),
            current_closure_captured_locals: HashSet::new(),

            multi_return_functions: HashMap::new(),
            generator_functions: HashSet::new(),
            method_fn_kinds: HashMap::new(),
            js_derived_ctor_ctx: None,
            generator_param_counts: HashMap::new(),
            host_import_bindings: HashMap::new(),
            host_const_bindings: HashMap::new(),
            host_namespace_aliases: HashMap::new(),
            host_package_roots: HashMap::new(),
            tree_mounts: HashMap::new(),
            ambient_tree_roots: Vec::new(),
            source_type_aliases: HashMap::new(),
            current_module_imports: Vec::new(),
            module_exports: HashMap::new(),
            module_value_exports: HashMap::new(),
            active_finally_blocks: Vec::new(),
            finally_joins: Vec::new(),
            fired_finally_indices: Vec::new(),
            catch_depth: 0,
            active_async_try_depth: 0,
            uses_proxy: false,
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

    fn reserve_shared_global_name(&mut self, name: &str) {
        if self.shared_global_slots.contains_key(name) {
            return;
        }
        let slot = self.shared_global_names.len() as u16;
        let owned = name.to_string();
        self.shared_global_slots.insert(owned.clone(), slot);
        self.shared_global_names.push(owned);
    }

    #[allow(dead_code)]
    fn reserve_shared_global_binding_pattern(&mut self, pattern: &BindingPattern) {
        match pattern {
            BindingPattern::Ident(name) => self.reserve_shared_global_name(&self.canon(name)),
            BindingPattern::Object(props) => {
                for prop in props {
                    if let Some(value) = prop.value.as_ref() {
                        self.reserve_shared_global_binding_pattern(value);
                    } else {
                        self.reserve_shared_global_name(&self.canon(&prop.key));
                    }
                }
            }
            BindingPattern::Array(items) => {
                for item in items {
                    match item {
                        ArrayPatternElem::Pattern(pattern, _) => {
                            self.reserve_shared_global_binding_pattern(pattern);
                        }
                        ArrayPatternElem::Rest(name) => {
                            self.reserve_shared_global_name(&self.canon(name));
                        }
                        ArrayPatternElem::Hole => {}
                    }
                }
            }
        }
    }

    #[allow(dead_code)]
    fn reserve_shared_global_names_in_body(&mut self, body: &[Statement]) {
        for stmt in body {
            match &stmt.kind {
                StmtKind::Block(stmts) => {
                    self.reserve_shared_global_names_in_body(stmts);
                }
                StmtKind::NamespaceDecl { body, .. } => {
                    self.reserve_shared_global_names_in_body(body);
                }
                StmtKind::FunctionDecl { name, .. }
                | StmtKind::ClassDecl { name, .. }
                | StmtKind::StructDecl { name, .. }
                | StmtKind::EnumDecl { name, .. }
                | StmtKind::ModuleDecl { name, .. } => {
                    self.reserve_shared_global_name(&self.canon(name));
                }
                StmtKind::VarDecl { declarations, .. } => {
                    for decl in declarations {
                        self.reserve_shared_global_binding_pattern(&decl.pattern);
                    }
                }
                _ => {}
            }
        }
    }

    fn reserve_runtime_global_names(&mut self) {
        for name in [
            "__vb_file_path_by_handle",
            "__vb_file_eof_by_handle",
            "__vb_record_rows_by_handle",
            "__vb_record_next_index_by_handle",
            "__vb_record_current_index_by_handle",
        ] {
            self.reserve_shared_global_name(name);
        }
    }

    fn seed_shared_global_constants(&self, chunk: &mut Chunk) {
        for name in &self.shared_global_names {
            chunk.add_constant(Value::String(Arc::from(name.as_str())));
        }
    }

    fn shared_global_slot(&self, name: &str) -> u16 {
        *self
            .shared_global_slots
            .get(name)
            .unwrap_or_else(|| panic!("missing shared global slot for {name}"))
    }

    fn global_name_const_idx(&mut self, name: &str) -> u16 {
        self.shared_global_slots
            .get(name)
            .copied()
            .unwrap_or_else(|| self.str_const(name))
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
        value_exports: HashMap<String, HashMap<String, vybe_bytecode::Value>>,
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
        self.case_sensitive = self.profile.case_sensitive;
        self.current_module_imports = module.imports.clone();

        // Pre-scan: detect `new Proxy(...)` anywhere in the module so the
        // Member / Index emit sites can route through the proxy dispatcher
        // even when the access appears before the construction in source
        // order. JS profile only — `Proxy` is a JS construct.
        if self.profile.name == "js" {
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

        // Register the .NET BCL class wrappers (Object → … → Form, Button, …)
        // before walking the user body, so user code that writes
        // `Inherits Form` finds a real `Form` class with a real ctor chain.
        // Gated on `profile.namespaces.use_dotnet` so non-.NET languages
        // don't get the names installed in their global scope.
        if self.profile.namespaces.use_dotnet {
            self.register_dotnet_classes()?;
        }

        if crate::registry::plib::module_uses_plib_gcl(module) {
            self.register_plib_gcl_classes()?;
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
        let merged_body = if has_partial {
            merge_partial_classes(&module.body, self.case_sensitive)
        } else {
            module.body.clone()
        };

        self.collect_reflection_metadata(&merged_body);

        // Multi-value pre-scan: any function whose every explicit `Return`
        // is a same-arity tuple literal is a candidate for the WASM
        // multi-value ABI. We only opt in when the language profile
        // requests it — other languages keep tuple-as-heap-object semantics.
        if self.profile.multi_value_tuple_returns {
            self.collect_multi_return_functions(&merged_body);
        }

        self.predeclare_type_names(&merged_body, None);
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

        self.reserve_runtime_global_names();
        let shared_global_names = self.shared_global_names.clone();
        for name in shared_global_names {
            self.chunks[0].add_constant(Value::String(Arc::from(name)));
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
                self.emit_u8(Op::CALL_REF, 0);
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
                    self.emit_u16(Op::STRUCT_GET, key);
                    self.emit_u8(Op::CALL_REF, 0);
                    self.emit(Op::DROP);
                }
            }
        }

        self.emit(Op::NULL);
        self.emit(Op::RETURN);
        // Take the max of the scope's highest slot and whatever raw local
        // slots compiler_common helpers (e.g. `invoke::emit_invoke_method`)
        // reserved directly on the chunk — those bypass `Scope` but still
        // need the VM to reserve slots at call-frame entry.
        let ns = self.scope().next_slot;
        self.chunks[0].finalize_local_count(ns);
        // Skip helper linking when compiling runtime helper source.
        // Re-running finalization here would call back into helper
        // compilation and recurse through `Compiler::compile`.
        // and recurse forever. Cheap thread-local guard since polyfill
        // compilation is single-threaded at vybex build time.
        if !crate::emitter::runtime_helpers::is_compiling_runtime_helper() {
            if self.profile.name == "c" {
                common::bundle::finalize_with_runtime_helpers_excluding(
                    &mut self.chunks,
                    &["__stdlib_sprintf"],
                );
            } else {
                common::bundle::finalize_with_runtime_helpers(&mut self.chunks);
            }
        }
        Self::normalize_import_table(&mut self.chunks);
        let host_imports = self.collected_host_imports();
        Ok(CompileResult {
            chunks: self.chunks,
            host_imports,
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
    if saw_any {
        arity
    } else {
        None
    }
}

fn is_hoisted_deconstruction_block(stmts: &[Statement]) -> bool {
    let call_stmt = match stmts {
        [Statement {
            kind: StmtKind::Expr(expr),
            ..
        }] => expr,
        [Statement {
            kind: StmtKind::VarDecl { .. },
            ..
        }, Statement {
            kind: StmtKind::Expr(expr),
            ..
        }] => expr,
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
