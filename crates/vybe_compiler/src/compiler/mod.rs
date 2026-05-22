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

mod classes;
mod expressions;
mod calls;
mod events;
#[path = "../php/compiler.rs"]
mod php;

use std::sync::Arc;
use std::collections::{BTreeSet, HashMap, HashSet};
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use crate::profile::*;
use crate::emitter as common;
use crate::ast::*;
use crate::scope::Scope;

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
}

// ════════════════════════════════════════════════════════════════════════════
// Pending class bookkeeping
// ════════════════════════════════════════════════════════════════════════════
struct PendingClass {
    parent: Option<String>,
    enclosing_class: Option<String>,
    fields: Vec<String>,
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
}

#[derive(Debug, Clone)]
pub(crate) struct CallSignature {
    param_names: Vec<String>,
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
            min_arity: params.iter().take_while(|param| param.default.is_none() && !param.is_rest).count(),
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
    Constructor { type_name: String, param_types: Vec<String> },
    Method { type_name: String, method_name: String },
    Property { type_name: String, property_name: String },
    Field { type_name: String, field_name: String },
    Parameter { type_name: String, method_name: String, index: usize },
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

// ════════════════════════════════════════════════════════════════════════════
// Compiler
// ════════════════════════════════════════════════════════════════════════════

pub struct Compiler {
    chunks: Vec<Chunk>,
    scopes: Vec<Scope>,
    current: usize,
    loops: Vec<LoopCtx>,
    loop_states: Vec<common::loops::LoopState>,
    /// Current label stack depth — incremented on every BLOCK/LOOP, decremented on END.
    /// Used to compute BR_LABEL depth for break/continue.
    label_depth: u32,
    /// Label depth at the entry of the current function body. RETURN
    /// must drain back to this — the VM's label_stack is global, and
    /// leaving function-local BLOCKs on it pollutes the caller's
    /// br_label depths (caller's `br 0` would land on a stale callee
    /// BLOCK target). Saved/restored across nested function decls.
    function_label_base: u32,
    line: u32,
    defined_globals: HashSet<String>,
    shared_global_slots: HashMap<String, u16>,
    shared_global_names: Vec<String>,
    defined_functions: HashSet<String>,
    function_param_modes: HashMap<String, Vec<PassBy>>,
    function_param_types: HashMap<String, Vec<Option<String>>>,
    function_min_arity: HashMap<String, usize>,
    function_signatures: HashMap<String, Vec<CallSignature>>,
    rest_fixed_arities: BTreeSet<u8>,
    function_return_types: HashMap<String, String>,
    fortran_interface_overloads: HashMap<String, Vec<FortranInterfaceOverload>>,
    fortran_operator_overloads: HashMap<String, Vec<FortranInterfaceOverload>>,
    constructor_signatures: HashMap<String, Vec<CallSignature>>,
    defined_classes: HashSet<String>,
    /// Names of methods defined on any user class — used to avoid value method
    /// hijacking (e.g. user class `Calc.Add()` shouldn't match array `add`).
    defined_class_methods: HashSet<String>,
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
    array_bindings: HashMap<String, ArrayBindingMetadata>,
    /// Label for the next loop to be pushed (set by StmtKind::Labeled).
    pending_label: Option<String>,
    with_targets: Vec<u16>,
    capture_by_value_vars: Vec<String>,
    pointer_cell_bindings: HashMap<usize, HashSet<String>>,


    /// Functions whose every explicit `Return` carries an `ExprKind::Tuple`
    /// of the same arity. Populated by a pre-pass before any function is
    /// compiled so both callee (set `chunk.result_arity`, push N values
    /// without packing) and caller (destructure directly off the stack)
    /// can agree on the multi-value ABI at emit time.
    multi_return_functions: HashMap<String, u8>,
    /// Functions compiled with `chunk.is_generator = true` — tracked
    /// by canonical name so `for v in gen()` call-site emission knows
    /// to use the `RESUME`-loop iterator protocol rather than the
    /// array-index protocol.
    generator_functions: HashSet<String>,
    /// ESM host-module import bindings: canon(local) → (module, func).
    /// Populated from user `import { X } from "wasi:foo"` statements.
    /// A direct call to `X` compiles to `CALL_IMPORT`; read-as-value
    /// (`const f = X`) reads the global that `host_imports::install`
    /// places under the same key.
    host_import_bindings: HashMap<String, (String, String)>,
    /// ESM wildcard namespace aliases: canon(alias) → module specifier.
    /// `import * as cli from "wasi:cli"` records `cli` → `"wasi:cli"`.
    /// Module Namespace Object access `cli.field` is resolved at compile
    /// time to `CALL_IMPORT (wasi:cli, field)` with no receiver pushed.
    /// Bare-value access of `cli` uses a runtime namespace object built
    /// by `host_imports::install` (reflection path).
    host_namespace_aliases: HashMap<String, String>,
    /// Component-Model package roots: canon(prefix) → module_root.
    /// Populated by the Linker from profile `PackageRoot` defaults
    /// (e.g. `{"vybe": "vybe:", "wasi": "wasi:", "wasm": "wasm:"}`).
    /// Phase 3 will wire `calls.rs`'s qualified-chain path to consume
    /// this map instead of `profile.namespaces.host_packages`.
    host_package_roots: HashMap<String, String>,
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
    /// Active finally blocks for the current control-flow path.
    ///
    /// Used to make early returns execute structured `finally` bodies
    /// even though the VM's TRY_START handler currently ignores the
    /// reserved finally offset operand.
    active_finally_blocks: Vec<FinallyAction>,
    /// Nesting depth of the catch body currently being compiled.
    catch_depth: usize,
    /// JS async-function wrapper try depth currently active for the
    /// function body being compiled. Explicit returns inside that body
    /// must emit matching TRY_END opcodes before RETURN so the VM does
    /// not retain stale handlers from the callee frame.
    active_async_try_depth: usize,
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

/// §16.2.1 named — `import { name as local } from "module"`.
#[derive(Debug, Clone)]
pub struct HostImportNamed {
    pub local: String,
    pub module: String,
    pub func: String,
}

/// AST scan: returns true if the statement (or anything nested within
/// it) constructs a `Proxy` (i.e. contains `new Proxy(...)`). Used to
/// gate the Member / Index proxy dispatcher emit so non-Proxy code
/// keeps the zero-overhead direct-opcode path.
fn stmt_uses_proxy(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Expr(e) => expr_uses_proxy(e),
        StmtKind::Return(opt) => opt.as_ref().map_or(false, expr_uses_proxy),
        StmtKind::Throw { expr, cause } => {
            expr.as_ref().map_or(false, expr_uses_proxy)
                || cause.as_ref().map_or(false, expr_uses_proxy)
        }
        StmtKind::VarDecl { declarations, .. } => declarations.iter().any(|d| {
            d.init.as_ref().map_or(false, expr_uses_proxy)
        }),
        StmtKind::Assign { value, .. } | StmtKind::CompoundAssign { value, .. } => {
            expr_uses_proxy(value)
        }
        StmtKind::Block(stmts) => stmts.iter().any(stmt_uses_proxy),
        StmtKind::If { cond, then_body, elifs, else_body } => {
            expr_uses_proxy(cond)
                || then_body.iter().any(stmt_uses_proxy)
                || elifs.iter().any(|(c, b)| expr_uses_proxy(c) || b.iter().any(stmt_uses_proxy))
                || else_body.as_ref().map_or(false, |b| b.iter().any(stmt_uses_proxy))
        }
        StmtKind::While { cond, body, .. } => {
            expr_uses_proxy(cond) || body.iter().any(stmt_uses_proxy)
        }
        StmtKind::DoWhile { cond, body, .. } => {
            expr_uses_proxy(cond) || body.iter().any(stmt_uses_proxy)
        }
        StmtKind::For { init, cond, update, body, .. } => {
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
                getter.as_ref().map_or(false, |b| b.iter().any(stmt_uses_proxy))
                    || setter.as_ref().map_or(false, |s| s.body.iter().any(stmt_uses_proxy))
            }
            _ => false,
        }),
        StmtKind::Try { body, catches, finally, .. } => {
            body.iter().any(stmt_uses_proxy)
                || catches.iter().any(|c| c.body.iter().any(stmt_uses_proxy))
                || finally.as_ref().map_or(false, |b| b.iter().any(stmt_uses_proxy))
        }
        _ => false,
    }
}

fn expr_uses_proxy(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::New { class, args } => {
            if let ExprKind::Ident(name) = &class.kind {
                if name == "Proxy" { return true; }
            }
            args.iter().any(|a| expr_uses_proxy(&a.value))
        }
        ExprKind::Call { callee, args, .. } => {
            expr_uses_proxy(callee) || args.iter().any(|a| expr_uses_proxy(&a.value))
        }
        ExprKind::Binary { left, right, .. } => {
            expr_uses_proxy(left) || expr_uses_proxy(right)
        }
        ExprKind::Unary { expr, .. } => expr_uses_proxy(expr),
        ExprKind::Member { object, .. } => expr_uses_proxy(object),
        ExprKind::Index { object, index, .. } => {
            expr_uses_proxy(object) || expr_uses_proxy(index)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            expr_uses_proxy(cond) || expr_uses_proxy(then) || expr_uses_proxy(else_)
        }
        ExprKind::Array(elems) => elems.iter().any(|e| expr_uses_proxy(&e.value)),
        ExprKind::Object(props) => props.iter().any(|p| match p {
            ObjectProperty::KeyValue { value, .. } => expr_uses_proxy(value),
            ObjectProperty::Computed { key, value } => expr_uses_proxy(key) || expr_uses_proxy(value),
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

impl Compiler {
    fn strip_global_namespace_prefix(name: &str) -> String {
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

    fn stdlib_global_name(stdlib_name: &str) -> String {
        match stdlib_name {
            "regex_replace_pat_first" => "__ecma_regexp_replace_pat_first".to_string(),
            "regex_split_pat_first" => "__ecma_regexp_split_pat_first".to_string(),
            "regex_match_all_pat_first" => "__ecma_regexp_match_all_pat_first".to_string(),
            _ => format!("__vybe_{}", stdlib_name),
        }
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
            defined_class_methods: HashSet::new(),
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
            array_bindings: HashMap::new(),
            pending_label: None,
            with_targets: Vec::new(),
            capture_by_value_vars: Vec::new(),
            pointer_cell_bindings: HashMap::new(),


            multi_return_functions: HashMap::new(),
            generator_functions: HashSet::new(),
            host_import_bindings: HashMap::new(),
            host_namespace_aliases: HashMap::new(),
            host_package_roots: HashMap::new(),
            source_type_aliases: HashMap::new(),
            current_module_imports: Vec::new(),
            module_exports: HashMap::new(),
            active_finally_blocks: Vec::new(),
            catch_depth: 0,
            active_async_try_depth: 0,
            uses_proxy: false,
        }
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
                | StmtKind::ModuleDecl { name, .. }
                 => {
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

    fn predeclare_interface_signatures_in_body(&mut self, body: &[Statement]) {
        for stmt in body {
            self.predeclare_interface_signatures_in_stmt(stmt);
        }
    }

    fn predeclare_interface_signatures_in_stmt(&mut self, stmt: &Statement) {
        match &stmt.kind {
            StmtKind::Block(body)
            | StmtKind::NamespaceDecl { body, .. }
            | StmtKind::FunctionDecl { body, .. } => {
                self.predeclare_interface_signatures_in_body(body);
            }
            StmtKind::InterfaceDecl { name, members, .. } => {
                self.register_interface_method_signatures(name, members);
            }
            StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::ModuleDecl { members, .. } => {
                self.predeclare_interface_signatures_in_members(members);
            }
            _ => {}
        }
    }

    fn predeclare_interface_signatures_in_members(&mut self, members: &[ClassMember]) {
        for member in members {
            match member {
                ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                    self.predeclare_interface_signatures_in_stmt(stmt);
                }
                _ => {}
            }
        }
    }

    fn register_interface_method_signatures(&mut self, interface_name: &str, members: &[InterfaceMember]) {
        let interface_canonical = self.canon(interface_name);
        let operator_symbol = self.fortran_interface_operator_symbol(interface_name);

        for member in members {
            let InterfaceMember::Method {
                name,
                params,
                return_type,
                signature_source,
                ..
            } = member else {
                continue;
            };

            let target_name = signature_source.as_ref().unwrap_or(name);
            let target_canonical = self.canon(target_name);
            let canonical_names = if self.profile.name == "fortran" && !interface_name.is_empty() {
                vec![target_canonical.clone(), interface_canonical.clone()]
            } else {
                vec![self.canon(name)]
            };

            if let Some(source_name) = signature_source.as_ref() {
                let source_canonical = self.canon(source_name);
                for canonical in &canonical_names {
                    if let Some(source_modes) = self.function_param_modes.get(&source_canonical).cloned() {
                        self.function_param_modes
                            .entry(canonical.clone())
                            .or_insert(source_modes);
                    }
                    if let Some(source_types) = self.function_param_types.get(&source_canonical).cloned() {
                        self.function_param_types
                            .entry(canonical.clone())
                            .or_insert(source_types);
                    }
                    if let Some(min_arity) = self.function_min_arity.get(&source_canonical).copied() {
                        self.function_min_arity.entry(canonical.clone()).or_insert(min_arity);
                    }
                    if let Some(signatures) = self.function_signatures.get(&source_canonical).cloned() {
                        self.function_signatures
                            .entry(canonical.clone())
                            .or_insert(signatures);
                    }
                    if let Some(source_return_type) = self.function_return_types.get(&source_canonical).cloned() {
                        self.function_return_types
                            .entry(canonical.clone())
                            .or_insert(source_return_type);
                    }
                }
            } else {
                let param_modes: Vec<PassBy> = params.iter().map(|param| param.pass_by).collect();
                let param_types: Vec<Option<String>> = params.iter().map(|param| param.type_hint.clone()).collect();
                let min_arity = params
                    .iter()
                    .take_while(|param| param.default.is_none() && !param.is_rest)
                    .count();
                let signature = CallSignature::from_params(params);

                for canonical in &canonical_names {
                    self.function_param_modes
                        .entry(canonical.clone())
                        .or_insert_with(|| param_modes.clone());
                    self.function_param_types
                        .entry(canonical.clone())
                        .or_insert_with(|| param_types.clone());
                    self.function_min_arity
                        .entry(canonical.clone())
                        .or_insert(min_arity);

                    let signatures = self.function_signatures.entry(canonical.clone()).or_default();
                    if !signatures.iter().any(|existing| {
                        existing.param_names == signature.param_names
                            && existing.min_arity == signature.min_arity
                            && existing.has_rest == signature.has_rest
                    }) {
                        signatures.push(signature.clone());
                    }

                    if let Some(return_type) = return_type.as_ref() {
                        self.function_return_types
                            .entry(canonical.clone())
                            .or_insert_with(|| return_type.clone());
                    }
                }
            }

            if self.profile.name == "fortran" && !interface_name.is_empty() {
                let overload = FortranInterfaceOverload {
                    target_name: target_canonical,
                    min_arity: params
                        .iter()
                        .take_while(|param| param.default.is_none() && !param.is_rest)
                        .count(),
                    param_types: params.iter().map(|param| param.type_hint.clone()).collect(),
                };

                if let Some(symbol) = operator_symbol.as_ref() {
                    let overloads = self.fortran_operator_overloads.entry(symbol.clone()).or_default();
                    if !overloads.iter().any(|existing| existing.target_name == overload.target_name) {
                        overloads.push(overload);
                    }
                } else {
                    let overloads = self
                        .fortran_interface_overloads
                        .entry(interface_canonical.clone())
                        .or_default();
                    if !overloads.iter().any(|existing| existing.target_name == overload.target_name) {
                        overloads.push(overload);
                    }
                }
            }
        }
    }

    fn fortran_interface_operator_symbol(&self, name: &str) -> Option<String> {
        let trimmed = name.trim();
        let lower = trimmed.to_ascii_lowercase();
        if !lower.starts_with("operator(") || !trimmed.ends_with(')') {
            return None;
        }
        let start = trimmed.find('(')? + 1;
        let end = trimmed.rfind(')')?;
        Some(trimmed[start..end].trim().to_string())
    }

    fn normalize_fortran_dispatch_type(&self, type_hint: &str) -> String {
        let resolved = self.resolve_source_type_alias(type_hint);
        let normalized = Self::normalize_type_hint(&resolved);
        let trimmed = normalized.trim();

        if let Some(inner) = trimmed
            .strip_prefix("type(")
            .and_then(|rest| rest.strip_suffix(')'))
            .or_else(|| trimmed.strip_prefix("class(").and_then(|rest| rest.strip_suffix(')')))
        {
            return self.canon(inner.trim());
        }

        if trimmed == "int" || trimmed.starts_with("integer") {
            return "integer".to_string();
        }
        if matches!(trimmed, "real" | "float" | "double" | "double precision")
            || trimmed.starts_with("real(")
        {
            return "real".to_string();
        }
        if trimmed == "bool" || trimmed.starts_with("logical") {
            return "logical".to_string();
        }

        self.canon(trimmed)
    }

    fn fortran_overload_target_param_types(&self, overload: &FortranInterfaceOverload) -> Vec<Option<String>> {
        self.function_param_types
            .get(&overload.target_name)
            .cloned()
            .filter(|param_types| !param_types.is_empty())
            .unwrap_or_else(|| overload.param_types.clone())
    }

    fn resolve_fortran_overload_target_with_fallback(
        &self,
        overloads: &[FortranInterfaceOverload],
        arg_exprs: &[Expression],
        allow_unknown_fallback: bool,
    ) -> Option<String> {
        let arg_types: Vec<Option<String>> = arg_exprs
            .iter()
            .map(|expr| self.infer_expr_type_hint(expr).map(|hint| self.normalize_fortran_dispatch_type(&hint)))
            .collect();
        let has_known_arg_types = arg_types.iter().any(Option::is_some);

        let mut best: Option<(&FortranInterfaceOverload, usize)> = None;
        let mut ambiguous = false;

        for overload in overloads {
            if arg_exprs.len() < overload.min_arity {
                continue;
            }

            let param_types = self.fortran_overload_target_param_types(overload);
            if !param_types.is_empty() && param_types.len() != arg_exprs.len() {
                continue;
            }

            let mut score = 0usize;
            let mut compatible = true;
            for (arg_type, param_type) in arg_types.iter().zip(param_types.iter()) {
                let Some(param_type) = param_type.as_ref() else {
                    continue;
                };
                let param_key = self.normalize_fortran_dispatch_type(param_type);
                let Some(arg_key) = arg_type.as_ref() else {
                    continue;
                };
                if arg_key == &param_key {
                    score += 2;
                    continue;
                }
                compatible = false;
                break;
            }

            if !compatible {
                continue;
            }

            match best {
                None => {
                    best = Some((overload, score));
                    ambiguous = false;
                }
                Some((_, best_score)) if score > best_score => {
                    best = Some((overload, score));
                    ambiguous = false;
                }
                Some((_, best_score)) if score == best_score => {
                    ambiguous = true;
                }
                _ => {}
            }
        }

        if let Some((overload, _)) = best {
            if !ambiguous || overloads.len() == 1 {
                return Some(overload.target_name.clone());
            }
        }

        (allow_unknown_fallback && !has_known_arg_types && overloads.len() == 1)
            .then(|| overloads[0].target_name.clone())
    }

    fn resolve_fortran_overload_target(
        &self,
        overloads: &[FortranInterfaceOverload],
        arg_exprs: &[Expression],
    ) -> Option<String> {
        self.resolve_fortran_overload_target_with_fallback(overloads, arg_exprs, true)
    }

    pub(super) fn resolve_fortran_interface_target(
        &self,
        name: &str,
        arg_exprs: &[Expression],
    ) -> Option<String> {
        (self.profile.name == "fortran")
            .then(|| self.canon(name))
            .and_then(|canonical| self.fortran_interface_overloads.get(&canonical))
            .and_then(|overloads| self.resolve_fortran_overload_target(overloads, arg_exprs))
    }

    pub(super) fn resolve_fortran_operator_target(
        &self,
        op: &BinOp,
        arg_exprs: &[Expression],
    ) -> Option<String> {
        let symbol = match op {
            BinOp::Add => "+",
            BinOp::Sub => "-",
            BinOp::Mul => "*",
            BinOp::Div => "/",
            BinOp::Pow => "**",
            _ => return None,
        };

        self.fortran_operator_overloads
            .get(symbol)
            .and_then(|overloads| self.resolve_fortran_overload_target_with_fallback(overloads, arg_exprs, false))
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
        *self.shared_global_slots.get(name).unwrap_or_else(|| {
            panic!("missing shared global slot for {name}")
        })
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

        // Pre-pass: merge `Partial Class` declarations sharing the same name.
        // Walker-driven: only runs when at least one ClassDecl in the module
        // is flagged `modifiers.is_partial = true` (VB/C# walkers set this
        // on `Partial Class`; other languages leave it false and skip the
        // merge entirely). After merging, the body has exactly one ClassDecl
        // per class name with all fields/methods pooled together.
        let has_partial = module.body.iter().any(|s| {
            matches!(&s.kind, StmtKind::ClassDecl { modifiers, .. } if modifiers.is_partial)
        });
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
        self.predeclare_interface_signatures_in_body(&merged_body);
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

        for stmt in &merged_body {
            if matches!(&stmt.kind, StmtKind::FunctionDecl { .. }) {
                continue;
            }
            self.compile_stmt(stmt)?;
        }

        // Auto-call entry point if defined
        if let Some(ref ep) = self.profile.entry_point.clone() {
            let has_ep = self.defined_globals.contains(ep)
                || (!self.case_sensitive && self.defined_globals.iter().any(|g| g.eq_ignore_ascii_case(ep)));
            if has_ep {
                self.emit_var_get(ep);
                self.emit_u8(Op::CALL_REF, 0);
                self.emit(Op::DROP);
            }
        }

        self.emit(Op::NULL);
        self.emit(Op::HALT);
        // Take the max of the scope's highest slot and whatever raw local
        // slots compiler_common helpers (e.g. `invoke::emit_invoke_method`)
        // reserved directly on the chunk — those bypass `Scope` but still
        // need the VM to reserve slots at call-frame entry.
        let locals = self.scope().next_slot.max(self.chunks[0].local_count);
        self.chunks[0].local_count = locals;
        // Skip stdlib bundling when compiling polyfill source. Polyfills
        // ARE stdlib chunks (extracted via `emitter::stdlib::build_polyfill`)
        // — re-running `finalize_with_stdlib` here would call back into
        // `build_stdlib` → `build_polyfill` → `Compiler::compile` → here
        // and recurse forever. Cheap thread-local guard since polyfill
        // compilation is single-threaded at vybex build time.
        if !crate::emitter::stdlib::is_compiling_polyfill() {
            common::bundle::finalize_with_stdlib(&mut self.chunks);
        }
        let host_imports = self.collected_host_imports();
        Ok(CompileResult {
            chunks: self.chunks,
            host_imports,
        })
    }

    /// Drain the compiler's host-import metadata into the shape the VM
    /// setup expects.
    fn collected_host_imports(&self) -> HostImportMetadata {
        let mut named: Vec<HostImportNamed> = self.host_import_bindings.iter()
            .map(|(local, (module, func))| HostImportNamed {
                local: local.clone(),
                module: module.clone(),
                func: func.clone(),
            })
            .collect();
        named.sort_by(|a, b| a.local.cmp(&b.local));
        let mut wildcard: Vec<HostWildcardImport> = self.host_namespace_aliases.iter()
            .map(|(alias, module)| HostWildcardImport {
                alias: alias.clone(),
                module: module.clone(),
            })
            .collect();
        wildcard.sort_by(|a, b| a.alias.cmp(&b.alias));
        HostImportMetadata { named, wildcard }
    }

    fn predeclare_type_names(&mut self, body: &[Statement], namespace: Option<&str>) {
        for stmt in body {
            match &stmt.kind {
                StmtKind::NamespaceDecl { name, body } => {
                    let member = self.canon(name).replace('\\', ".");
                    let qualified = namespace
                        .map(|prefix| format!("{prefix}.{member}"))
                        .unwrap_or(member);
                    self.predeclare_type_names(body, Some(&qualified));
                }
                StmtKind::ClassDecl { name, .. } | StmtKind::StructDecl { name, .. } => {
                    let member = self.canon(name);
                    self.defined_globals.insert(member.clone());
                    self.defined_classes.insert(member.clone());
                    if let StmtKind::StructDecl { members, .. } = &stmt.kind {
                        self.predeclare_struct_surface(&member, members);
                    }
                    if let Some(prefix) = namespace {
                        let qualified = format!("{prefix}.{member}");
                        self.defined_globals.insert(qualified.clone());
                        self.defined_classes.insert(qualified);
                    }
                }
                StmtKind::ModuleDecl { name, members, .. } => {
                    let member = self.canon(name);
                    self.defined_globals.insert(member.clone());
                    self.defined_classes.insert(member.clone());
                    self.register_module_static_container(&member, members);
                    if let Some(prefix) = namespace {
                        let qualified = format!("{prefix}.{member}");
                        self.defined_globals.insert(qualified.clone());
                        self.defined_classes.insert(qualified);
                    }
                }
                _ => {}
            }
        }
    }

    fn predeclare_struct_surface(&mut self, name: &str, members: &[ClassMember]) {
        let mut fields = Vec::new();
        let mut instance_member_names = Vec::new();
        let mut instance_pointer_method_names = Vec::new();
        let mut instance_field_types = HashMap::new();
        let mut static_fields = Vec::new();
        let mut static_field_types = HashMap::new();
        let mut static_method_names = Vec::new();

        for member in members {
            match member {
                ClassMember::Field { name, type_hint, modifiers, .. } => {
                    let field_name = self.canon(name);
                    if modifiers.is_shared {
                        static_fields.push(field_name.clone());
                        if let Some(type_hint) = type_hint.as_ref() {
                            static_field_types
                                .insert(field_name, Self::normalize_type_hint(type_hint));
                        }
                    } else {
                        fields.push(field_name.clone());
                        if let Some(type_hint) = type_hint.as_ref() {
                            instance_field_types
                                .insert(field_name, Self::normalize_type_hint(type_hint));
                        }
                    }
                }
                ClassMember::Method(stmt) => {
                    if let StmtKind::FunctionDecl { name: method_name, modifiers, params, .. } = &stmt.kind {
                        let canonical = self.canon(method_name);
                        if modifiers.is_shared {
                            static_method_names.push(canonical);
                        } else {
                            if params
                                .first()
                                .and_then(|param| param.type_hint.as_deref())
                                .is_some_and(|type_hint| type_hint.trim_start().starts_with('*'))
                            {
                                instance_pointer_method_names.push(canonical.clone());
                            }
                            instance_member_names.push(canonical);
                        }
                    }
                }
                _ => {}
            }
        }

        self.defined_globals.insert(format!("{}$arity0", name));
        self.pending_classes.entry(name.to_string()).or_insert(PendingClass {
            parent: None,
            enclosing_class: self.current_class.clone(),
            fields,
            is_value_type: true,
            instance_member_names,
            instance_pointer_method_names,
            instance_field_types,
            static_fields,
            static_field_types,
            static_method_names,
            instance_method_overloads: HashMap::new(),
            static_method_overloads: HashMap::new(),
            nested_types: Vec::new(),
            statics: Vec::new(),
        });
    }

    fn register_module_static_container(&mut self, module_name: &str, members: &[ClassMember]) {
        let mut module_static_fields: Vec<String> = Vec::new();
        let mut module_static_field_types: HashMap<String, String> = HashMap::new();
        let mut module_static_methods: Vec<String> = Vec::new();
        let mut module_nested_types: Vec<String> = Vec::new();

        for member in members {
            match member {
                ClassMember::Field { name, type_hint, .. } => {
                    let field_name = self.canon(name);
                    module_static_fields.push(field_name.clone());
                    if let Some(type_hint) = type_hint.as_ref() {
                        module_static_field_types
                            .insert(field_name, Self::normalize_type_hint(type_hint));
                    }
                }
                ClassMember::Const { name, type_hint, .. } => {
                    let const_name = self.canon(name);
                    module_static_fields.push(const_name.clone());
                    if let Some(type_hint) = type_hint.as_ref() {
                        module_static_field_types
                            .insert(const_name, Self::normalize_type_hint(type_hint));
                    }
                }
                ClassMember::Method(stmt) => {
                    if let StmtKind::FunctionDecl { name: method_name, params, return_type, .. } = &stmt.kind {
                        let method_canon = self.canon(method_name);
                        module_static_methods.push(method_canon.clone());
                        self.function_param_modes
                            .entry(method_canon.clone())
                            .or_insert_with(|| params.iter().map(|param| param.pass_by).collect());
                        self.function_min_arity
                            .entry(method_canon.clone())
                            .or_insert_with(|| {
                                params
                                    .iter()
                                    .take_while(|param| param.default.is_none() && !param.is_rest)
                                    .count()
                            });
                        if let Some(return_type) = return_type.clone() {
                            self.function_return_types
                                .entry(method_canon)
                                .or_insert(return_type);
                        }
                    }
                }
                ClassMember::NestedType(stmt) => {
                    if let Some(type_name) = match &stmt.kind {
                        StmtKind::ClassDecl { name, .. }
                        | StmtKind::StructDecl { name, .. }
                        | StmtKind::EnumDecl { name, .. }
                        | StmtKind::InterfaceDecl { name, .. }
                        | StmtKind::ModuleDecl { name, .. } => Some(self.canon(name)),
                        _ => None,
                    } {
                        module_nested_types.push(type_name);
                    }
                    if let StmtKind::InterfaceDecl { name, members, .. } = &stmt.kind {
                        self.register_interface_method_signatures(name, members);
                    }
                }
                _ => {}
            }
        }

        self.pending_classes.insert(module_name.to_string(), PendingClass {
            parent: None,
            enclosing_class: self.current_class.clone(),
            fields: Vec::new(),
            is_value_type: false,
            instance_member_names: Vec::new(),
            instance_pointer_method_names: Vec::new(),
            instance_field_types: HashMap::new(),
            static_fields: module_static_fields,
            static_field_types: module_static_field_types,
            static_method_names: module_static_methods,
            instance_method_overloads: HashMap::new(),
            static_method_overloads: HashMap::new(),
            nested_types: module_nested_types,
            statics: Vec::new(),
        });
    }

    /// The Linker phase — ECMA-262 §16.2.1.5 Link adapted for Vybe.
    ///
    /// Populates the three resolver maps (`host_import_bindings`,
    /// `host_namespace_aliases`, `host_package_roots`) from two
    /// sources, in order:
    ///
    ///   1. **Profile defaults** (`profile.esm_defaults`) — the
    ///      language's ambient pre-declared imports. For JS,
    ///      `console → wasi:cli` and `Math → ecma:math`. For VB,
    ///      `System` as a `PackageRoot`.
    ///   2. **User imports** (`module.imports`) — `import { X } from
    ///      "wasi:foo"` etc. Walked last so they shadow profile
    ///      defaults on key collision (ECMA-262 §16.2 lexical bindings
    ///      override module-scope defaults).
    ///
    /// `HashMap::insert` on a duplicate key replaces the value, so
    /// walking profile-then-user gives spec-correct shadowing for
    /// free.
    ///
    /// Runs before any bytecode is emitted.
    fn link(&mut self, module: &crate::ast::Module) {
        // Phase A.1: ambient profile defaults.
        let defaults = self.profile.esm_defaults.clone();
        for d in &defaults {
            match d {
                crate::profile::EsmDefault::Named { local, module: m, name } => {
                    let key = self.canon(local);
                    self.host_import_bindings.insert(key, (m.clone(), name.clone()));
                }
                crate::profile::EsmDefault::Namespace { alias, module: m } => {
                    let key = self.canon(alias);
                    self.host_namespace_aliases.insert(key, m.clone());
                }
                crate::profile::EsmDefault::PackageRoot { prefix, module_root } => {
                    // Component Model package names are lowercase by
                    // spec; store + look up in lowercase regardless of
                    // the language's case sensitivity.
                    let key = prefix.to_ascii_lowercase();
                    self.host_package_roots.insert(key, module_root.clone());
                }
            }
        }

        // Phase A.2: user imports — shadow profile defaults on key
        // collision. Resolves host-specifier paths (wasi:* / wasm:* /
        // vybe:*) directly, and Adapter-module paths (node:*, etc.)
        // by walking the re-export chain in `module_exports` to the
        // ultimate target. Relative paths still resolve at bundle
        // load time.
        let bare_aliases = self.profile.bare_module_aliases.clone();
        let normalize_bare = |path: &str| -> String {
            // Profile-driven: JS routes `'fs'` → `'node:fs'` via the
            // [bare_module_aliases] table; Python's profile leaves it
            // empty so `import os` keeps Python's stdlib semantics.
            bare_aliases.get(path).cloned().unwrap_or_else(|| path.to_string())
        };
        for imp in &module.imports {
            match &imp.kind {
                crate::ast::ImportKind::Simple { path, alias: Some(alias) } => {
                    self.source_type_aliases.insert(self.canon(alias), path.clone());
                }
                crate::ast::ImportKind::Named { path, names, .. } => {
                    let path = normalize_bare(path);
                    if is_host_specifier(&path) {
                        for n in names {
                            let raw_local = n.alias.as_ref().unwrap_or(&n.name).clone();
                            let key = self.canon(&raw_local);
                            self.host_import_bindings.insert(key, (path.clone(), n.name.clone()));
                        }
                    } else if let Some(adapter_exports) = self.module_exports.get(&path).cloned() {
                        // Adapter module: each name is a pre-resolved
                        // `(final_module, final_name)` pair courtesy
                        // of the Indirect chain walker in the Bundle.
                        for n in names {
                            let raw_local = n.alias.as_ref().unwrap_or(&n.name).clone();
                            let key = self.canon(&raw_local);
                            if let Some(target) = adapter_exports.get(&n.name).cloned() {
                                self.host_import_bindings.insert(key, target);
                            }
                            // Unresolved export — leave it; Phase 8
                            // will surface a link error here.
                        }
                    }
                    // Relative / file-system imports — bundle-level
                    // resolver handles them by inlining sources.
                }
                crate::ast::ImportKind::Wildcard { path, alias } => {
                    let path = normalize_bare(path);
                    if !is_host_specifier(&path) { continue; }
                    if let Some(ns) = alias {
                        let key = self.canon(ns);
                        self.host_namespace_aliases.insert(key, path);
                    }
                }
                // Default + Simple: no meaning for host modules; skip.
                crate::ast::ImportKind::Default { .. }
                | crate::ast::ImportKind::Simple { .. } => {}
            }
        }
    }

    pub(super) fn resolve_source_type_alias(&self, name: &str) -> String {
        let normalized = Self::strip_global_namespace_prefix(name);
        let trimmed = normalized.trim().replace('\\', ".");
        let (head, tail) = trimmed
            .split_once('.')
            .map(|(head, tail)| (head.trim(), Some(tail.trim())))
            .unwrap_or((trimmed.as_str(), None));
        let (alias_head, suffix) = head
            .strip_suffix("()")
            .map(|bare| (bare.trim_end(), "()"))
            .unwrap_or((head, ""));
        let key = self.canon(alias_head);
        let Some(target) = self.source_type_aliases.get(&key) else {
            return trimmed;
        };
        match tail {
            Some(tail) if !tail.is_empty() => format!("{}{}.{}", target, suffix, tail),
            _ => format!("{}{}", target, suffix),
        }
    }

    fn parse_pascal_array_bound_token(token: &str) -> Option<(i64, bool)> {
        let trimmed = token.trim();
        if let Ok(value) = trimmed.parse::<i64>() {
            return Some((value, false));
        }

        if trimmed.starts_with('\'') && trimmed.ends_with('\'') && trimmed.len() >= 3 {
            let inner = &trimmed[1..trimmed.len() - 1];
            let unescaped = inner.replace("''", "'");
            let mut chars = unescaped.chars();
            if let (Some(ch), None) = (chars.next(), chars.next()) {
                return Some((ch as i64, true));
            }
        }

        None
    }

    fn pascal_array_type_hint_metadata(&self, type_hint: &str) -> Option<PascalArrayBoundsMetadata> {
        let trimmed = type_hint.trim().trim_end_matches('?').trim();
        let lowered = trimmed.to_ascii_lowercase();
        if !lowered.starts_with("array") {
            return None;
        }

        let Some(bracket_start) = trimmed.find('[') else {
            return Some(PascalArrayBoundsMetadata {
                is_fixed: false,
                dimensions: Vec::new(),
            });
        };
        let bracket_end = trimmed[bracket_start + 1..].find(']')? + bracket_start + 1;
        let mut dimensions = Vec::new();
        for dim in trimmed[bracket_start + 1..bracket_end].split(',').map(str::trim).filter(|dim| !dim.is_empty()) {
            let (lower, upper) = dim.split_once("..")?;
            let (lower, lower_is_char) = Self::parse_pascal_array_bound_token(lower)?;
            let (upper, upper_is_char) = Self::parse_pascal_array_bound_token(upper)?;
            if lower_is_char != upper_is_char {
                return None;
            }
            let length = if upper >= lower {
                (upper - lower + 1) as usize
            } else {
                0
            };
            dimensions.push(PascalArrayDimensionMetadata {
                first_index: lower,
                length,
                uses_char_ordinal: lower_is_char,
            });
        }
        Some(PascalArrayBoundsMetadata {
            is_fixed: !dimensions.is_empty(),
            dimensions,
        })
    }

    fn pascal_ordinal_index_expr(index: Expression) -> Expression {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("Ord")),
            args: vec![Argument::positional(index)],
            optional: false,
        })
    }

    fn pascal_indexed_type_hint(type_hint: &str) -> Option<String> {
        let trimmed = type_hint.trim().trim_end_matches('?').trim();
        if !trimmed.to_ascii_lowercase().starts_with("array") {
            return None;
        }

        let bracket_start = trimmed.find('[')?;
        let bracket_end = trimmed[bracket_start + 1..].find(']')? + bracket_start + 1;
        let dims = trimmed[bracket_start + 1..bracket_end].trim();
        let after_bracket = trimmed[bracket_end + 1..].trim();
        let of_pos = after_bracket.to_ascii_lowercase().find("of")?;
        let element_type = after_bracket[of_pos + 2..].trim();
        if element_type.is_empty() {
            return None;
        }

        if let Some((_, remaining_dims)) = dims.split_once(',') {
            Some(format!("array[{}] of {}", remaining_dims.trim(), element_type))
        } else {
            Some(element_type.to_string())
        }
    }

    fn collect_reflection_metadata(&mut self, body: &[Statement]) {
        self.reflection_types.clear();
        self.attribute_usage.clear();
        self.reflection_bindings.clear();
        for stmt in body {
            self.collect_reflection_stmt(stmt, None);
        }
    }

    fn collect_reflection_stmt(&mut self, stmt: &Statement, parent_runtime_name: Option<&str>) {
        match &stmt.kind {
            StmtKind::ClassDecl {
                name,
                parents,
                interfaces,
                members,
                decorators,
                ..
            } => {
                let runtime_name = self.reflection_runtime_type_name(name, parent_runtime_name);
                self.record_reflection_type(
                    &runtime_name,
                    parents,
                    interfaces,
                    decorators,
                    members,
                    false,
                );
            }
            StmtKind::StructDecl {
                name, interfaces, members, decorators, ..
            } => {
                let runtime_name = self.reflection_runtime_type_name(name, parent_runtime_name);
                self.record_reflection_type(
                    &runtime_name,
                    &[],
                    interfaces,
                    decorators,
                    members,
                    true,
                );
            }
            StmtKind::InterfaceDecl {
                name,
                parents,
                decorators,
                ..
            } => {
                let runtime_name = self.reflection_runtime_type_name(name, parent_runtime_name);
                let metadata = ReflectionTypeMetadata {
                    parents: parents
                    .iter()
                    .map(|parent| self.reflection_runtime_type_name(parent, None))
                    .collect(),
                    decorators: decorators.clone(),
                    interfaces: parents
                        .iter()
                        .map(|parent| self.reflection_runtime_type_name(parent, None))
                        .collect(),
                    ..ReflectionTypeMetadata::default()
                };
                self.reflection_types.insert(runtime_name.clone(), metadata);
                let usage = self.extract_attribute_usage(decorators);
                self.attribute_usage.insert(runtime_name, usage);
            }
            StmtKind::EnumDecl {
                name,
                interfaces,
                decorators,
                body_members,
                ..
            } => {
                let runtime_name = self.reflection_runtime_type_name(name, parent_runtime_name);
                self.record_reflection_type(
                    &runtime_name,
                    &[],
                    interfaces,
                    decorators,
                    body_members,
                    true,
                );
            }
            StmtKind::NamespaceDecl { name, body } => {
                let namespace_runtime = self.reflection_runtime_type_name(name, parent_runtime_name);
                for nested in body {
                    self.collect_reflection_stmt(nested, Some(namespace_runtime.as_str()));
                }
            }
            StmtKind::Block(body) => {
                for nested in body {
                    self.collect_reflection_stmt(nested, parent_runtime_name);
                }
            }
            _ => {}
        }
    }

    fn record_reflection_type(
        &mut self,
        runtime_name: &str,
        parents: &[String],
        interfaces: &[String],
        decorators: &[Expression],
        members: &[ClassMember],
        is_value_type: bool,
    ) {
        let mut metadata = ReflectionTypeMetadata {
            parents: parents
            .iter()
            .map(|parent| self.reflection_runtime_type_name(parent, None))
            .collect(),
            interfaces: interfaces
                .iter()
                .map(|parent| self.reflection_runtime_type_name(parent, None))
                .collect(),
            decorators: decorators.to_vec(),
            is_value_type,
            ..ReflectionTypeMetadata::default()
        };
        let mut nested_types: Vec<&Statement> = Vec::new();

        for member in members {
            match member {
                ClassMember::Method(stmt) => {
                    if let StmtKind::FunctionDecl {
                        name,
                        params,
                        modifiers,
                        ..
                    } = &stmt.kind
                    {
                        let mut method_decorators = Vec::new();
                        let mut param_decorators: HashMap<usize, Vec<Expression>> = HashMap::new();
                        for decorator in &modifiers.decorators {
                            if let Some((index, attr)) = self.unpack_param_decorator_carrier(decorator) {
                                param_decorators.entry(index).or_default().push(attr);
                            } else {
                                method_decorators.push(decorator.clone());
                            }
                        }
                        metadata.methods.insert(
                            name.clone(),
                            ReflectionMethodMetadata {
                                decorators: method_decorators,
                                is_static: modifiers.is_static,
                                params: params
                                    .iter()
                                    .enumerate()
                                    .map(|(index, param)| ReflectionParamMetadata {
                                        name: param.name.clone(),
                                        decorators: param_decorators.remove(&index).unwrap_or_default(),
                                    })
                                    .collect(),
                            },
                        );
                    }
                }
                ClassMember::Property {
                    name,
                    setter,
                    modifiers,
                    ..
                } => {
                    metadata.properties.insert(
                        name.clone(),
                        ReflectionMemberMetadata {
                            decorators: modifiers.decorators.clone(),
                            is_static: modifiers.is_static,
                            can_write: setter.is_some(),
                        },
                    );
                }
                ClassMember::Field {
                    name,
                    modifiers,
                    ..
                } => {
                    metadata.fields.insert(
                        name.clone(),
                        ReflectionMemberMetadata {
                            decorators: modifiers.decorators.clone(),
                            is_static: modifiers.is_static,
                            can_write: true,
                        },
                    );
                }
                ClassMember::Constructor { params, .. } => {
                    metadata.constructors.push(ReflectionConstructorMetadata {
                        param_types: params
                            .iter()
                            .map(|param| self.reflection_runtime_type_name(
                                param.type_hint.as_deref().unwrap_or("Object"),
                                None,
                            ))
                            .collect(),
                    });
                }
                ClassMember::NestedType(stmt) => {
                    let nested_runtime = match &stmt.kind {
                        StmtKind::ClassDecl { name, .. }
                        | StmtKind::StructDecl { name, .. }
                        | StmtKind::InterfaceDecl { name, .. }
                        | StmtKind::EnumDecl { name, .. } => {
                            Some(self.reflection_runtime_type_name(name, Some(runtime_name)))
                        }
                        _ => None,
                    };
                    if let Some(nested_runtime) = nested_runtime {
                        metadata.nested_types.push(nested_runtime);
                    }
                    nested_types.push(stmt);
                }
                _ => {}
            }
        }

        self.reflection_types.insert(runtime_name.to_string(), metadata);
        let usage = self.extract_attribute_usage(decorators);
        self.attribute_usage.insert(runtime_name.to_string(), usage);
        for stmt in nested_types {
            self.collect_reflection_stmt(stmt, Some(runtime_name));
        }
    }

    fn extract_attribute_usage(&self, decorators: &[Expression]) -> AttributeUsageMetadata {
        let mut usage = AttributeUsageMetadata::default();

        for decorator in decorators {
            let ExprKind::New { args, .. } = &decorator.kind else {
                continue;
            };
            let Some(attr_type) = self.reflection_attribute_type_name(decorator) else {
                continue;
            };
            if !attr_type.eq_ignore_ascii_case("System.AttributeUsageAttribute") {
                continue;
            }

            for arg in args {
                match arg.name.as_deref() {
                    Some("AllowMultiple") => {
                        usage.allow_multiple = matches!(arg.value.kind, ExprKind::Lit(Literal::Bool(true)));
                    }
                    Some("Inherited") => {
                        usage.inherited = matches!(arg.value.kind, ExprKind::Lit(Literal::Bool(true)));
                    }
                    _ => {}
                }
            }
        }

        usage
    }

    pub(crate) fn reflection_runtime_type_name(&self, type_name: &str, parent_runtime_name: Option<&str>) -> String {
        let global_stripped = Self::strip_global_namespace_prefix(type_name);
        let trimmed = global_stripped.trim().trim_end_matches('?').trim();
        let mut without_generics = String::with_capacity(trimmed.len());
        let mut depth = 0usize;
        for ch in trimmed.chars() {
            match ch {
                '<' => depth += 1,
                '>' => depth = depth.saturating_sub(1),
                _ if depth == 0 => without_generics.push(ch),
                _ => {}
            }
        }
        let normalized = match without_generics.trim() {
            "int" | "Int32" => "Int32",
            "uint" | "UInt32" => "UInt32",
            "long" | "Int64" => "Int64",
            "ulong" | "UInt64" => "UInt64",
            "short" | "Int16" => "Int16",
            "ushort" | "UInt16" => "UInt16",
            "byte" | "Byte" => "Byte",
            "sbyte" | "SByte" => "SByte",
            "float" | "Single" => "Single",
            "double" | "Double" => "Double",
            "decimal" | "Decimal" => "Decimal",
            "bool" | "Boolean" => "Boolean",
            "char" | "Char" => "Char",
            "string" | "String" => "String",
            "object" | "Object" => "Object",
            other => other,
        };
        let normalized = normalized.strip_prefix("System.System.").unwrap_or(normalized);
        if let Some(parent) = parent_runtime_name {
            let leaf = normalized.rsplit('.').next().unwrap_or(normalized).trim();
            return format!("{parent}.{leaf}");
        }
        if normalized.starts_with("System.") {
            normalized.to_string()
        } else {
            format!("System.{}", normalized)
        }
    }

    pub(crate) fn reflection_attribute_type_name(&self, expr: &Expression) -> Option<String> {
        let class = match &expr.kind {
            ExprKind::New { class, .. } => class.as_ref(),
            _ => return None,
        };

        let raw_name = match &class.kind {
            ExprKind::Ident(name) => name.clone(),
            ExprKind::Member { .. } => self.flatten_member_chain(class).join("."),
            _ => return None,
        };

        if !raw_name.contains('.') {
            let mut matches: Vec<String> = self
                .reflection_types
                .keys()
                .filter(|known| known.rsplit('.').next().is_some_and(|leaf| leaf == raw_name))
                .cloned()
                .collect();
            matches.sort();
            matches.dedup();
            if matches.len() == 1 {
                return matches.into_iter().next();
            }
        }

        Some(self.reflection_runtime_type_name(&raw_name, None))
    }

    fn unpack_param_decorator_carrier(&self, expr: &Expression) -> Option<(usize, Expression)> {
        let ExprKind::New { class, args } = &expr.kind else {
            return None;
        };
        let ExprKind::Ident(name) = &class.kind else {
            return None;
        };
        if name != "__vybe_param_attribute" || args.len() != 2 {
            return None;
        }
        let ExprKind::Lit(Literal::Int(index)) = args[0].value.kind else {
            return None;
        };
        Some((index.max(0) as usize, args[1].value.clone()))
    }

    pub(crate) fn reflection_type_lookup_name(&self, type_name: &str) -> String {
        self.reflection_runtime_type_name(type_name, None)
    }

    pub(crate) fn reflection_type_metadata(&self, type_name: &str) -> Option<&ReflectionTypeMetadata> {
        let lookup = self.reflection_type_lookup_name(type_name);
        self.reflection_types.get(&lookup)
    }

    pub(crate) fn reflection_type_short_name(&self, type_name: &str) -> String {
        let trimmed = type_name.trim().trim_end_matches('?').trim();
        let without_namespace = trimmed.rsplit('.').next().unwrap_or(trimmed).trim();
        let without_generics = without_namespace.split('<').next().unwrap_or(without_namespace).trim();
        without_generics.to_string()
    }

    pub(crate) fn reflection_type_full_name(&self, type_name: &str) -> String {
        self.reflection_runtime_type_name(type_name, None)
    }

    pub(crate) fn reflection_is_enum_type(&self, type_name: &str) -> bool {
        let lookup = self.reflection_type_lookup_name(type_name);
        let short_name = self.canon(&self.reflection_type_short_name(type_name));
        self.enum_value_names.contains_key(&lookup)
            || self.enum_value_names.contains_key(&short_name)
            || self.enum_value_names.keys().any(|known| known.eq_ignore_ascii_case(&lookup) || known.eq_ignore_ascii_case(&short_name))
    }

    pub(crate) fn reflection_is_value_type(&self, type_name: &str) -> bool {
        let lookup = self.reflection_type_lookup_name(type_name);
        if self.reflection_types.get(&lookup).is_some_and(|meta| meta.is_value_type) {
            return true;
        }
        matches!(
            lookup.as_str(),
            "System.Boolean"
                | "System.Byte"
                | "System.SByte"
                | "System.Int16"
                | "System.UInt16"
                | "System.Int32"
                | "System.UInt32"
                | "System.Int64"
                | "System.UInt64"
                | "System.Single"
                | "System.Double"
                | "System.Decimal"
                | "System.Char"
                | "System.DateTime"
                | "System.TimeSpan"
                | "System.Guid"
        )
    }

    pub(crate) fn reflection_base_type_name(&self, type_name: &str) -> Option<String> {
        self.reflection_type_metadata(type_name)
            .and_then(|meta| meta.parents.first().cloned())
    }

    pub(crate) fn reflection_nested_type_name(&self, type_name: &str, nested_name: &str) -> Option<String> {
        let parent = self.reflection_type_lookup_name(type_name);
        let desired = nested_name.trim();
        self.reflection_types.keys().find(|candidate| {
            candidate
                .strip_prefix(&(parent.clone() + "."))
                .is_some_and(|leaf| leaf.eq_ignore_ascii_case(desired))
        }).cloned()
    }

    pub(crate) fn reflection_generic_argument_types(&self, type_name: &str) -> Vec<String> {
        let trimmed = type_name.trim();
        let Some(start) = trimmed.find('<') else {
            return Vec::new();
        };
        let Some(end) = trimmed.rfind('>') else {
            return Vec::new();
        };
        let inner = &trimmed[start + 1..end];
        let mut parts = Vec::new();
        let mut current = String::new();
        let mut depth = 0usize;
        for ch in inner.chars() {
            match ch {
                '<' => {
                    depth += 1;
                    current.push(ch);
                }
                '>' => {
                    depth = depth.saturating_sub(1);
                    current.push(ch);
                }
                ',' if depth == 0 => {
                    let part = current.trim();
                    if !part.is_empty() {
                        parts.push(self.reflection_type_full_name(part));
                    }
                    current.clear();
                }
                _ => current.push(ch),
            }
        }
        let part = current.trim();
        if !part.is_empty() {
            parts.push(self.reflection_type_full_name(part));
        }
        parts
    }

    pub(crate) fn reflection_open_generic_type_name(&self, type_name: &str) -> String {
        self.reflection_type_lookup_name(type_name)
    }

    pub(crate) fn reflection_interfaces(&self, type_name: &str) -> Vec<String> {
        self.reflection_type_metadata(type_name)
            .map(|meta| meta.interfaces.clone())
            .unwrap_or_default()
    }

    pub(crate) fn reflection_is_assignable_from(&self, target_type: &str, candidate_type: &str) -> bool {
        let target = self.reflection_type_lookup_name(target_type);
        let mut pending = vec![self.reflection_type_lookup_name(candidate_type)];
        let mut visited = HashSet::new();

        while let Some(current) = pending.pop() {
            if !visited.insert(current.clone()) {
                continue;
            }
            if current.eq_ignore_ascii_case(&target) {
                return true;
            }
            if let Some(meta) = self.reflection_types.get(&current) {
                pending.extend(meta.parents.iter().cloned());
                pending.extend(meta.interfaces.iter().cloned());
            }
        }

        false
    }

    fn module_imports_namespace(&self, namespace: &str) -> bool {
        self.current_module_imports.iter().any(|import| match &import.kind {
            ImportKind::Simple { path, .. }
            | ImportKind::Named { path, .. }
            | ImportKind::Wildcard { path, .. }
            | ImportKind::Default { path, .. } => path.eq_ignore_ascii_case(namespace),
        })
    }

    fn should_infer_winforms_form(&self, name: &str, parents: &[String]) -> bool {
        if !parents.is_empty()
            || !self.profile.namespaces.use_dotnet
            || !matches!(self.profile.name.as_str(), "vb" | "csharp")
            || !self.module_imports_namespace("System.Windows.Forms")
        {
            return false;
        }

        // Real VB/C# WinForms code commonly omits the explicit base type in
        // the user-authored partial while the surrounding project/designer
        // model still treats the class as a form. Keep the inference narrow:
        // only classes that follow the standard *Form / FormN naming shape
        // opt into the existing Form adapter wrapper.
        name.to_ascii_lowercase().contains("form")
    }

    pub(crate) fn reflection_constructor_for_types(&self, type_name: &str, param_types: &[String]) -> Option<ReflectionBinding> {
        let lookup = self.reflection_type_lookup_name(type_name);
        let normalized_params: Vec<String> = param_types
            .iter()
            .map(|param| self.reflection_type_lookup_name(param))
            .collect();
        let meta = self.reflection_types.get(&lookup)?;
        let ctor = meta.constructors.iter().find(|ctor| {
            ctor.param_types.len() == normalized_params.len()
                && ctor
                    .param_types
                    .iter()
                    .zip(normalized_params.iter())
                    .all(|(left, right)| left.eq_ignore_ascii_case(right))
        })?;
        Some(ReflectionBinding::Constructor {
            type_name: lookup,
            param_types: ctor.param_types.clone(),
        })
    }

    // ════════════════════════════════════════════════════════════════════════
    // Multi-value tuple returns (opt-in via `multi_value_tuple_returns`)
    // ════════════════════════════════════════════════════════════════════════

    /// `true` if `iter` is an `ExprKind::Call` to an ident that names
    /// a function previously compiled with `is_generator = true`. Used
    /// by `for v in gen():` to pick the stack-switching iterator path.
    fn is_direct_generator_call(&self, iter: &Expression) -> bool {
        if let ExprKind::Call { callee, .. } = &iter.kind {
            if let ExprKind::Ident(n) = &callee.kind {
                return self.generator_functions.contains(&self.canon(n));
            }
        }
        false
    }

    fn emit_with_target_get(&mut self, name: &str) -> bool {
        let Some(slot) = self.with_targets.last().copied() else {
            return false;
        };
        self.emit_u16(Op::LOCAL_GET, slot);
        let idx = self.str_const(&self.canon(name));
        self.emit_u16(Op::STRUCT_GET, idx);
        true
    }

    fn emit_with_target_set(&mut self, name: &str) -> bool {
        let Some(slot) = self.with_targets.last().copied() else {
            return false;
        };
        let value_slot = self.define_local("__with_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        let idx = self.str_const(&self.canon(name));
        self.emit_u16(Op::STRUCT_SET, idx);
        true
    }


    /// Emit a `for v in gen():` loop that drives the generator via
    ///   block $exit
    ///     loop $loop
    ///       local.get $cont
    ///       gen.next            ;; pushes (value, has_more)
    ///       br_if 0             ;; break out when has_more == 0
    ///       local.set $v        ;; assign yielded value
    ///       <body>
    ///       br $loop
    ///     end
    ///   end
    fn compile_generator_for_in(
        &mut self,
        var: &str,
        key: Option<&str>,
        iter: &Expression,
        body: &[Statement],
        else_body: Option<&[Statement]>,
    ) -> Result<(), String> {
        use crate::ast::ExprKind;
        // Compile and stash the continuation.
        let (callee, args) = match &iter.kind {
            ExprKind::Call { callee, args, .. } => (callee, args),
            _ => unreachable!("compile_generator_for_in expects Call"),
        };
        self.compile_call(callee, args)?;
        let cont_slot = self.define_local("__gen_cont");
        self.emit_u16(Op::LOCAL_SET, cont_slot); self.emit(Op::DROP);

        self.compile_generator_for_in_cont(var, key, cont_slot, body, else_body)
    }

    fn compile_generator_for_in_cont(
        &mut self,
        var: &str,
        key: Option<&str>,
        cont_slot: u16,
        body: &[Statement],
        else_body: Option<&[Statement]>,
    ) -> Result<(), String> {

        let key_index_slot = self.maybe_define_php_generator_key_index_slot(key);

        let line = self.line;
        let block_patch = self.chunk().emit_block(line);
        let (loop_patch, _) = self.chunk().emit_loop_s(line);
        self.label_depth += 2;

        // Advance the generator. GEN_NEXT pops cont and pushes (value, has_more).
        self.emit_u16(Op::LOCAL_GET, cont_slot);
        self.emit(Op::GEN_NEXT);
        let has_more_slot = self.define_local("__gen_has_more");
        self.emit_u16(Op::LOCAL_SET, has_more_slot); self.emit(Op::DROP);
        let value_slot = self.define_local("__gen_value");
        self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);

        if self.is_php_profile() {
            self.emit_php_generator_foreach_state(cont_slot, has_more_slot, value_slot);
        } else {
            self.emit_u16(Op::LOCAL_GET, has_more_slot);
            self.emit(Op::DYN_TO_BOOL);
            self.emit(Op::DYN_NOT);
            // br_if_label 1 → jump to $exit when has_more was 0.
            self.emit_u8(Op::BR_IF_LABEL, 1);
        }

        if let Some(key_name) = key {
            let key_slot = self.define_local(key_name);
            if self.is_php_profile() {
                self.emit_php_generator_key_binding(key_slot, value_slot, key_index_slot);
            } else {
                self.emit(Op::NULL);
                self.emit_u16(Op::LOCAL_SET, key_slot); self.emit(Op::DROP);
            }
        }

        let var_slot = self.define_local(var);
        if self.is_php_profile() {
            self.emit_php_generator_value_binding(var_slot, value_slot);
        } else {
            self.emit_u16(Op::LOCAL_GET, value_slot);
            self.emit_u16(Op::LOCAL_SET, var_slot); self.emit(Op::DROP);
        }

        if let Some(key_index_slot) = key_index_slot {
            self.emit_u16(Op::LOCAL_GET, key_index_slot);
            self.emit_const(Value::F64(1.0));
            self.emit(Op::DYN_ADD);
            self.emit_u16(Op::LOCAL_SET, key_index_slot); self.emit(Op::DROP);
        }

        // Compile loop body inside a `$body` block so `continue` can
        // target it without rerunning the advance.
        let body_block = self.chunk().emit_block(line);
        self.label_depth += 1;
        let break_depth = self.label_depth - 2; // $exit
        let continue_depth = self.label_depth - 0; // $body
        self.loops.push(LoopCtx {
            label: self.pending_label.take(),
            break_label_depth: break_depth,
            continue_label_depth: continue_depth,
            did_break_slot: None,
        });
        for s in body { self.compile_stmt(s)?; }
        self.loops.pop();
        self.chunk().emit_end(line);
        self.chunk().patch_block(body_block);
        self.label_depth -= 1;


        // Continue the loop.
        self.emit_u8(Op::BR_LABEL, 0);
        self.chunk().emit_end(line);
        self.chunk().patch_loop(loop_patch);
        self.chunk().emit_end(line);
        self.chunk().patch_block(block_patch);
        self.label_depth -= 2;

        if let Some(else_stmts) = else_body {
            for s in else_stmts { self.compile_stmt(s)?; }
        }
        Ok(())
    }

    /// Arity of the currently-compiling function if the pre-scan tagged
    /// it multi-return, else `None`. Driven off `current_func_name` so
    /// it automatically tracks function boundaries without a parallel
    /// stack.
    fn current_multi_return_arity(&self) -> Option<u8> {
        let name = self.current_func_name.as_deref()?;
        self.multi_return_functions.get(name).copied()
    }

    pub(super) fn multi_return_arity_for_callee(&self, callee: &Expression) -> Option<u8> {
        match &callee.kind {
            ExprKind::Ident(name) => self.multi_return_functions.get(&self.canon(name)).copied(),
            ExprKind::Member { object, field, .. } => {
                if let ExprKind::Ident(object_name) = &object.kind {
                    let qualified = self.canon(&format!("{}.{}", object_name, field));
                    if let Some(&arity) = self.multi_return_functions.get(&qualified) {
                        return Some(arity);
                    }
                }

                self.multi_return_functions.get(&self.canon(field)).copied()
            }
            _ => None,
        }
    }

    /// Emit the CALL for a multi-value receive context *without* the
    /// trailing repack that `compile_expr` would normally add. The
    /// destructure path consumes the N raw stack values directly.
    pub(super) fn compile_call_raw(&mut self, value: &Expression) -> Result<(), String> {
        if let ExprKind::Call { callee, args, .. } = &value.kind {
            self.compile_call(callee, args)
        } else {
            self.compile_expr(value)
        }
    }

    /// Pack the top-N stack values — produced by a multi-value CALL —
    /// into a single array/tuple so downstream uses see the expected
    /// single-value semantics. The last pushed value becomes element
    /// `n-1`; order matches what a destructure would assign.
    pub(super) fn pack_multi_value_result(&mut self, n: u8) {
        let line = self.line;
        // Reserve N consecutive slots via the existing scope helper —
        // `emit_pack_n` stashes each stack value into a slot, then
        // rebuilds the array from those slots in declaration order.
        let mut first = 0u16;
        for i in 0..n {
            let s = self.define_local("__mv_pack");
            if i == 0 { first = s; }
        }
        common::collections::emit_pack_n(&mut self.chunks, self.current, n as u16, first, line);
    }

    /// Return `Some((N, [ident...]))` when `targets`/`value` match the
    /// "multi-value receive" shape:
    ///   * exactly one target, a tuple-destructure of N plain identifiers
    ///   * value is a direct `Ident(name)` call to a function the pre-scan
    ///     tagged multi-return with matching arity N
    /// For any other shape we return `None` and fall through to the
    /// existing heap-tuple destructuring path.
    fn detect_multi_value_receive(&self, targets: &[Expression], value: &Expression)
        -> Option<(u8, Vec<String>)>
    {
        if targets.len() != 1 { return None; }
        let idents = match &targets[0].kind {
            ExprKind::Destructure(DestructurePattern::Array(pats)) => {
                let mut names = Vec::with_capacity(pats.len());
                for p in pats {
                    match p {
                        ArrayPatternElem::Pattern(BindingPattern::Ident(n), _) => {
                            names.push(n.clone());
                        }
                        _ => return None,
                    }
                }
                names
            }
            _ => return None,
        };
        let multi_n = match &value.kind {
            ExprKind::Call { callee, args, .. } => {
                let _ = args;
                self.multi_return_arity_for_callee(callee)?
            }
            _ => return None,
        };
        if multi_n as usize != idents.len() { return None; }
        Some((multi_n, idents))
    }

    /// Walk top-level function declarations and record every function
    /// whose explicit `Return` statements all carry a tuple literal of
    /// the same arity. Those functions opt into the WASM multi-value
    /// ABI: callee sets `chunk.result_arity = N` and pushes the tuple
    /// elements unpacked; caller destructures directly off the stack.
    fn collect_multi_return_functions(&mut self, stmts: &[Statement]) {
        for stmt in stmts {
            if let StmtKind::FunctionDecl { name, body, .. } = &stmt.kind {
                if let Some(arity) = uniform_tuple_return_arity(body) {
                    let cname = self.canon(name);
                    self.multi_return_functions.insert(cname, arity);
                }
            }
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Helpers
    // ════════════════════════════════════════════════════════════════════════

    fn scope(&self) -> &Scope { self.scopes.last().unwrap() }
    fn scope_mut(&mut self) -> &mut Scope { self.scopes.last_mut().unwrap() }

    fn pointer_binding_key(&self, name: &str) -> String {
        if self.case_sensitive {
            name.to_string()
        } else {
            self.canon(name)
        }
    }

    fn binding_uses_pointer_cell(&self, name: &str) -> bool {
        let key = self.pointer_binding_key(name);
        self.pointer_cell_bindings
            .get(&self.current)
            .is_some_and(|bindings| bindings.contains(&key))
    }

    fn mark_pointer_cell_binding(&mut self, name: &str) {
        let key = self.pointer_binding_key(name);
        self.pointer_cell_bindings.entry(self.current).or_default().insert(key);
    }

    fn resolve_named_local_slot(&self, name: &str) -> Option<u16> {
        self.scope().resolve(name).or_else(|| {
            if self.case_sensitive {
                None
            } else {
                self.scope().resolve_ci(name)
            }
        })
    }

    fn promote_local_binding_to_pointer_cell(&mut self, name: &str) -> Option<u16> {
        let slot = self.resolve_named_local_slot(name)?;
        if !self.binding_uses_pointer_cell(name) {
            common::references::emit_cell_new_from_local(&mut self.chunks, self.current, slot, self.line);
            self.emit_u16(Op::LOCAL_SET, slot);
            self.emit(Op::DROP);
            self.mark_pointer_cell_binding(name);
        }
        Some(slot)
    }

    fn promote_global_binding_to_pointer_cell(&mut self, name: &str) -> bool {
        let canon_name = self.canon(name);
        if !self.defined_globals.contains(&canon_name) {
            return false;
        }

        if !self.binding_uses_pointer_cell(name) {
            let value_slot = self.define_local("__ref_global_value");
            let idx = self.str_const(&canon_name);
            self.emit_u16(Op::GLOBAL_GET, idx);
            self.emit_u16(Op::LOCAL_SET, value_slot);
            self.emit(Op::DROP);
            common::references::emit_cell_new(&mut self.chunks, self.current, value_slot, self.line);
            self.emit_u16(Op::GLOBAL_SET, idx);
            self.emit(Op::DROP);
            self.mark_pointer_cell_binding(name);
        }

        true
    }

    pub(super) fn emit_wrap_top_of_stack_in_pointer_cell(&mut self) {
        let value_slot = self.define_local("__ref_cell_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit(Op::DROP);
        common::references::emit_cell_new(&mut self.chunks, self.current, value_slot, self.line);
    }

    pub(super) fn emit_autoderef_pointer_cell(&mut self) {
        let obj_slot = self.define_local("__ref_autoderef_obj");
        self.emit_u16(Op::LOCAL_SET, obj_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit(Op::REF_IS_OBJECT);
        let not_object = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        let kind_key = self.str_const("__ref_kind");
        self.emit_u16(Op::STRUCT_GET, kind_key);
        self.emit_const(Value::String(Arc::from("cell")));
        self.emit(Op::DYN_EQ);
        let not_cell = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        common::references::emit_cell_load(&mut self.chunks, self.current, self.line);
        let done = self.emit_jump(Op::BR);

        self.patch_jump(not_cell);
        self.patch_jump(not_object);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.patch_jump(done);
    }

    pub(super) fn compile_address_of_expr(&mut self, expr: &Expression) -> Result<(), String> {
        match &expr.kind {
            ExprKind::Ident(name) => {
                let canon_name = self.canon(name);
                if self.defined_functions.contains(&canon_name) {
                    self.emit_var_get(name);
                    return Ok(());
                }
                if let Some(slot) = self.promote_local_binding_to_pointer_cell(name) {
                    self.emit_u16(Op::LOCAL_GET, slot);
                    return Ok(());
                }
                if self.promote_global_binding_to_pointer_cell(name) {
                    let idx = self.str_const(&canon_name);
                    self.emit_u16(Op::GLOBAL_GET, idx);
                    return Ok(());
                }
            }
            ExprKind::Unary { op: UnaryOp::Deref, expr } => {
                self.compile_expr(expr)?;
                return Ok(());
            }
            _ => {}
        }

        self.compile_expr(expr)?;
        self.emit_wrap_top_of_stack_in_pointer_cell();
        Ok(())
    }

    pub(super) fn compile_deref_expr(&mut self, expr: &Expression) -> Result<(), String> {
        self.compile_expr(expr)?;
        self.emit_autoderef_pointer_cell();
        Ok(())
    }

    /// Define a local in the current scope AND sync the current chunk's
    /// `local_count` to the new high-water mark.
    ///
    /// Why this exists: helpers in `emitter/` (`emit_invoke_method`,
    /// `emit_get_range`, `emit_array_pair`, `emit_stdlib_call_*`) allocate
    /// scratch slots starting at `chunk.local_count`. If `chunk.local_count`
    /// isn't kept in sync with `scope.next_slot` during compilation, those
    /// scratch slots overlap named locals (params, rest-collection slots,
    /// user `let` bindings) and silently corrupt them.
    ///
    /// This is the historical root cause of the variadic-param-corruption
    /// bug — see `tests/js/test_variadic_bug.rs`. Maintaining the
    /// invariant `chunk.local_count >= scope.next_slot` at all times
    /// makes every helper using `chunk.local_count` for scratch correct
    /// by construction.
    pub(crate) fn define_local(&mut self, name: &str) -> u16 {
        let slot = self.scopes.last_mut().unwrap().define(name);
        let high = self.scopes.last().unwrap().next_slot;
        let cur = self.current;
        if high > self.chunks[cur].local_count {
            self.chunks[cur].local_count = high;
        }
        slot
    }

    /// Stack: [coll, idx] → [coll, idx_norm]. For languages where
    /// negative array indices wrap from the end (Python `arr[-1]`,
    /// Ruby, PHP). Maps return length 0 from `ARRAY_LENGTH` so this
    /// is a no-op on dict-style collections (negative integer keys
    /// stay negative). Strings return char count → `s[-1]` works.
    pub(crate) fn emit_negative_index_wrap(&mut self) {
        let line = self.line;
        let arr_slot = self.define_local("__neg_idx_arr");
        let idx_slot = self.define_local("__neg_idx_i");
        // Stash [coll, idx] into locals (LOCAL_SET peeks; DROP pops).
        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);
        // if idx < 0: idx = arr.length + idx
        self.emit_u16(Op::LOCAL_GET, idx_slot);
        self.emit_const(Value::I32(0));
        self.emit(Op::DYN_LT);
        self.emit(Op::DYN_TO_BOOL);
        self.emit(Op::DYN_NOT);
        let block_p = self.chunk().emit_block(line);
        self.label_depth += 1;
        self.chunk().emit_br_if(0, line); // skip wrap if !(idx < 0)
        self.emit_u16(Op::LOCAL_GET, arr_slot);
        common::collections::emit_array_length(&mut self.chunks[self.current], line);
        self.emit_u16(Op::LOCAL_GET, idx_slot);
        self.emit(Op::DYN_ADD);
        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
        self.chunk().emit_end(line);
        self.chunk().patch_block(block_p);
        self.label_depth -= 1;
        // Re-push [arr, idx_norm] for the caller's emit_get.
        self.emit_u16(Op::LOCAL_GET, arr_slot);
        self.emit_u16(Op::LOCAL_GET, idx_slot);
    }

    /// Emit RETURN, draining any function-local BLOCK/LOOP labels first.
    /// Without this, an early `return` inside an `if` (or any nested
    /// block) leaves stale labels on the VM's global label_stack —
    /// the caller's later `br_label N` then targets the callee's
    /// orphaned BLOCK and jumps into garbage bytecode.
    pub(crate) fn emit_return(&mut self) {
        let line = self.line;
        let drain = self.label_depth.saturating_sub(self.function_label_base);
        for _ in 0..drain {
            self.chunk().emit_end(line);
        }
        self.emit(Op::RETURN);
    }

    fn emit_active_finally_blocks(&mut self) -> Result<(), String> {
        if self.active_finally_blocks.is_empty() {
            return Ok(());
        }

        let original = self.active_finally_blocks.clone();
        for idx in (0..original.len()).rev() {
            self.active_finally_blocks = original[..idx].to_vec();
            self.emit_finally_action(&original[idx])?;
        }
        self.active_finally_blocks = original;
        Ok(())
    }

    fn emit_finally_action(&mut self, action: &FinallyAction) -> Result<(), String> {
        match action {
            FinallyAction::Statements(stmts) => {
                for stmt in stmts {
                    self.compile_stmt(stmt)?;
                }
            }
            FinallyAction::ResourceDispose { slot, method, line } => {
                self.label_depth += 1;
                common::errors::emit_resource_dispose(self.chunk(), *slot, method, *line);
                self.label_depth -= 1;
            }
        }
        Ok(())
    }

    fn emit_return_through_finally(&mut self, result_count: usize) -> Result<(), String> {
        let slots: Vec<u16> = (0..result_count)
            .map(|idx| self.define_local(&format!("__return_val_{}", idx)))
            .collect();
        for idx in (0..result_count).rev() {
            self.emit_u16(Op::LOCAL_SET, slots[idx]);
            self.emit(Op::DROP);
        }

        if !self.active_finally_blocks.is_empty() {
            self.emit_active_finally_blocks()?;
        }

        for slot in &slots {
            self.emit_u16(Op::LOCAL_GET, *slot);
        }

        let ref_out_slots = self.current_ref_out_params.clone().unwrap_or_default();
        if !ref_out_slots.is_empty() && self.current_multi_return_arity().is_none() {
            for slot in &ref_out_slots {
                self.emit_u16(Op::LOCAL_GET, *slot);
            }
            let pack_count = result_count + ref_out_slots.len();
            let mut first = 0u16;
            for index in 0..pack_count {
                let slot = self.define_local(&format!("__return_pack_{}", index));
                if index == 0 {
                    first = slot;
                }
            }
            common::collections::emit_pack_n(
                &mut self.chunks,
                self.current,
                pack_count as u16,
                first,
                self.line,
            );
        }
        if self.current_chunk_is_js_async() {
            for _ in 0..self.active_async_try_depth {
                self.emit(Op::TRY_END);
            }
        }
        self.emit_return();
        Ok(())
    }

    fn current_chunk_is_js_async(&self) -> bool {
        self.is_js_profile() && self.chunks[self.current].is_async
    }

    /// Same as `define_local` but with a type hint — sugar around
    /// `Scope::define_typed`. Keeps the sync invariant.
    pub(crate) fn define_local_typed(&mut self, name: &str, type_hint: Option<String>) -> u16 {
        let slot = self.scopes.last_mut().unwrap().define_typed(name, type_hint);
        let high = self.scopes.last().unwrap().next_slot;
        let cur = self.current;
        if high > self.chunks[cur].local_count {
            self.chunks[cur].local_count = high;
        }
        slot
    }
    fn chunk(&mut self) -> &mut Chunk { &mut self.chunks[self.current] }

    fn reserve_local_slot(&mut self, slot: u16) {
        self.chunks[self.current].local_count = self.chunks[self.current]
            .local_count
            .max(slot + 1);
    }

    fn emit(&mut self, op: Op) { let l = self.line; self.chunks[self.current].emit_op(op, l); }
    fn emit_u16(&mut self, op: Op, v: u16) { let l = self.line; self.chunks[self.current].emit_op_u16(op, v, l); }
    fn emit_u8(&mut self, op: Op, v: u8) { let l = self.line; self.chunks[self.current].emit_op_u8(op, v, l); }
    fn emit_const(&mut self, val: Value) { let idx = self.chunks[self.current].add_constant(val); self.emit_u16(Op::CONST, idx); }
    fn emit_jump(&mut self, op: Op) -> usize { let l = self.line; self.chunks[self.current].emit_jump(op, l) }
    fn patch_jump(&mut self, o: usize) { self.chunks[self.current].patch_jump(o); }
    fn emit_loop(&mut self, t: usize) { let l = self.line; self.chunks[self.current].emit_loop(t, l); }

    /// Compute BR_LABEL depth for `break`.
    fn break_depth(&self, label: Option<&str>) -> Option<u8> {
        let ctx = if let Some(lbl) = label {
            self.loops.iter().rev().find(|c| c.label.as_deref() == Some(lbl))?
        } else {
            self.loops.last()?
        };
        Some((self.label_depth - ctx.break_label_depth) as u8)
    }

    /// Compute BR_LABEL depth for `continue`.
    fn continue_depth(&self, label: Option<&str>) -> Option<u8> {
        let ctx = if let Some(lbl) = label {
            self.loops.iter().rev().find(|c| c.label.as_deref() == Some(lbl))?
        } else {
            self.loops.last()?
        };
        Some((self.label_depth - ctx.continue_label_depth) as u8)
    }

    #[allow(dead_code)]
    fn current_offset(&self) -> usize { self.chunks[self.current].current_offset() }
    fn str_const(&mut self, s: &str) -> u16 { self.chunks[self.current].add_constant(Value::String(Arc::from(s))) }
    fn shared_str_const(&mut self, s: &str) -> u16 { self.chunks[0].add_constant(Value::String(Arc::from(s))) }

    fn import(&mut self, module: &str, name: &str) -> u16 { self.chunks[0].add_import(module, name) }
    fn emit_host_call(&mut self, idx: u16, argc: u8) {
        let l = self.line;
        self.chunks[self.current].emit_op_u16(Op::CALL_IMPORT, idx, l);
        self.chunks[self.current].emit(argc, l);
    }

    /// Resolve a qualified identifier to a Component Model host call
    /// `(module, function)` pair when its first segment matches the
    /// profile's `host_packages` list, else `None`.
    ///
    /// Walker conventions: PHP passes backslash-separated names
    /// (`Vybe\Http\Request\method`), other languages should normalize
    /// their separator to `\` before this point (TODO for Python / C# /
    /// etc.). This keeps the resolver language-agnostic.
    ///
    /// Mapping:
    /// - `[Vybe, Http, Request, method]` → `("vybe:http/request", "method")`
    /// - `[Vybe, Math, cos]`             → `("ecma:math", "cos")`
    /// - `[Wasi, Cli, log]`              → `("wasi:cli", "log")`
    ///
    /// First join is `:` (package → interface), further joins use `/`,
    /// last segment is the function name. Everything is lowercased.
    fn resolve_component_model_call(&self, name: &str) -> Option<(String, String)> {
        if !name.contains('\\') { return None; }
        let parts: Vec<&str> = name.split('\\').collect();
        if parts.len() < 2 { return None; }

        // Consult the Linker's `host_package_roots` map instead of
        // `profile.namespaces.host_packages`. Populated at link time
        // from `EsmDefault::PackageRoot` entries (which the profile
        // loader auto-translates from the legacy list). Component
        // Model package names are lowercase by spec — match
        // case-insensitively regardless of the language's case rules.
        let first_key = parts[0].to_ascii_lowercase();
        if !self.host_package_roots.contains_key(&first_key) { return None; }

        let lower: Vec<String> = parts.iter().map(|s| s.to_ascii_lowercase()).collect();
        let (func, path) = lower.split_last()?;
        if path.is_empty() { return None; }

        let module = if path.len() == 1 {
            path[0].clone()
        } else {
            let mut m = path[0].clone();
            m.push(':');
            m.push_str(&path[1]);
            for p in &path[2..] {
                m.push('/');
                m.push_str(p);
            }
            m
        };
        Some((module, func.clone()))
    }

    // ── Crate-private accessors used by `dotnet_register` ──────────────
    //
    // The .NET class registration logic lives in a sibling file
    // (`dotnet_register.rs`) but operates on Compiler internals. These
    // helpers expose just the bits that registration needs without
    // making the underlying fields `pub`.
    pub(crate) fn chunks_mut(&mut self) -> &mut Vec<Chunk> { &mut self.chunks }
    pub(crate) fn current_line(&self) -> u32 { self.line }
    pub(crate) fn note_defined_global(&mut self, name: &str) {
        self.defined_globals.insert(name.to_string());
    }
    pub(crate) fn note_defined_class(&mut self, name: &str) {
        self.defined_classes.insert(name.to_string());
    }
    pub(crate) fn note_pending_class(&mut self, name: &str, parent: Option<String>) {
        self.pending_classes.insert(name.to_string(), PendingClass {
            parent,
            enclosing_class: self.current_class.clone(),
            fields: Vec::new(),
            is_value_type: false,
            instance_member_names: Vec::new(),
            instance_pointer_method_names: Vec::new(),
            instance_field_types: HashMap::new(),
            static_fields: Vec::new(),
            static_field_types: HashMap::new(),
            static_method_names: Vec::new(),
            instance_method_overloads: HashMap::new(),
            static_method_overloads: HashMap::new(),
            nested_types: Vec::new(),
            statics: Vec::new(),
        });
    }

    /// Push the canonical event-registry key for a control expression.
    /// Used by AddHandler / RemoveHandler so the GUI host indexes handlers by
    /// the source-stable identifier (field name, class name for `Me`, etc.)
    /// rather than the runtime `.Name` property — renaming a control after
    /// the handler is wired must NOT break dispatch.
    ///
    pub(crate) fn canon(&self, name: &str) -> String {
        if self.case_sensitive { name.to_string() } else { name.to_lowercase() }
    }

    fn normalize_type_hint(type_hint: &str) -> String {
        type_hint.trim().to_lowercase()
    }

    pub(super) fn emit_default_value_for_type_hint(&mut self, type_hint: Option<&str>) {
        match type_hint.map(Self::normalize_type_hint).as_deref() {
            Some("integer") | Some("int") | Some("int32") | Some("longint") | Some("real")
            | Some("double") | Some("float") | Some("single") | Some("decimal")
            | Some("long") | Some("int64") | Some("short") | Some("int16")
            | Some("uint") | Some("uint32") | Some("ulong") | Some("uint64")
            | Some("ushort") | Some("uint16") | Some("byte") | Some("sbyte")
            | Some("char") => self.emit(Op::F64_CONST_0),
            Some("boolean") | Some("bool") => self.emit(Op::FALSE),
            Some(type_hint) if Self::is_string_type_hint(type_hint) => self.emit_const(Value::String(Arc::from(""))),
            _ => self.emit(Op::NULL),
        }
    }

    fn is_string_type_hint(type_hint: &str) -> bool {
        let normalized = Self::normalize_type_hint(type_hint);
        normalized == "string"
            || normalized == "system.string"
            || normalized.ends_with(".string")
            || normalized == "character"
            || normalized.starts_with("character(")
            || normalized.starts_with("character*")
    }

    fn is_numeric_type_hint(type_hint: &str) -> bool {
        matches!(
            Self::normalize_type_hint(type_hint).as_str(),
            "integer" | "int" | "int32" | "longint" | "real"
                | "double" | "float" | "single" | "decimal"
                | "long" | "int64" | "short" | "int16"
                | "uint" | "uint32" | "ulong" | "uint64"
                | "ushort" | "uint16" | "byte" | "sbyte"
        )
    }

    fn fortran_out_param_ctor_name(type_hint: &str) -> Option<String> {
        let normalized = Self::normalize_type_hint(type_hint);
        if normalized.ends_with("()")
            || Self::is_numeric_type_hint(&normalized)
            || Self::is_string_type_hint(&normalized)
            || matches!(normalized.as_str(), "boolean" | "bool")
        {
            return None;
        }

        if let Some(inner) = normalized
            .strip_prefix("type(")
            .and_then(|inner| inner.strip_suffix(')'))
        {
            return Some(inner.trim().to_string());
        }

        if let Some(inner) = normalized
            .strip_prefix("class(")
            .and_then(|inner| inner.strip_suffix(')'))
        {
            return Some(inner.trim().to_string());
        }

        Some(normalized)
    }

    fn maybe_initialize_fortran_out_param(&mut self, param: &Param) {
        if self.profile.name != "fortran" || param.pass_by != PassBy::Out {
            return;
        }

        let Some(type_hint) = param.type_hint.as_deref() else {
            return;
        };
        let Some(ctor_name) = Self::fortran_out_param_ctor_name(type_hint) else {
            return;
        };
        let Some(slot) = self.scope().resolve(&param.name) else {
            return;
        };
        if !(self.defined_classes.contains(&ctor_name)
            || self.defined_globals.contains(&ctor_name)
            || self.profile.lookup_known_type(&ctor_name).is_some())
        {
            return;
        }

        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit(Op::REF_IS_NULL);
        let skip_init = self.emit_jump(Op::BR_IF_FALSE);

        if let Some((module, func)) = self
            .profile
            .lookup_known_type(&ctor_name)
            .map(|(module, func)| (module.to_string(), func.to_string()))
        {
            let idx = self.import(&module, &func);
            self.emit_host_call(idx, 0);
        } else {
            let idx = self.global_name_const_idx(&ctor_name);
            self.emit_u16(Op::GLOBAL_GET, idx);
            self.emit_u8(Op::CALL_REF, 0);
        }
        self.emit_u16(Op::LOCAL_SET, slot);
        self.emit(Op::DROP);

        self.patch_jump(skip_init);
    }

    fn can_instantiate_fortran_ctor_name(&self, ctor_name: &str) -> bool {
        self.defined_classes.contains(ctor_name)
            || self.defined_globals.contains(ctor_name)
            || self.profile.lookup_known_type(ctor_name).is_some()
    }

    fn emit_fortran_ctor_call(&mut self, ctor_name: &str) {
        if let Some((module, func)) = self
            .profile
            .lookup_known_type(ctor_name)
            .map(|(module, func)| (module.to_string(), func.to_string()))
        {
            let idx = self.import(&module, &func);
            self.emit_host_call(idx, 0);
        } else {
            let idx = self.global_name_const_idx(ctor_name);
            self.emit_u16(Op::GLOBAL_GET, idx);
            self.emit_u8(Op::CALL_REF, 0);
        }
    }

    fn fortran_allocate_ctor_name(&self, target: &Expression) -> Option<String> {
        let type_hint = self.infer_expr_type_hint(target)?;
        let normalized = Self::normalize_type_hint(&type_hint);
        let element_hint = normalized.strip_suffix("()").unwrap_or(normalized.as_str()).trim();
        let ctor_name = Self::fortran_out_param_ctor_name(element_hint)?;
        self.can_instantiate_fortran_ctor_name(&ctor_name)
            .then_some(ctor_name)
    }

    fn emit_fortran_allocated_array(&mut self, dim_slots: &[u16], ctor_name: Option<&str>) {
        let line = self.line;
        self.emit_u16(Op::LOCAL_GET, dim_slots[0]);
        common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
        let array_slot = self.define_local("__fortran_alloc_array");
        self.emit_u16(Op::LOCAL_SET, array_slot);
        self.emit(Op::DROP);

        if dim_slots.len() == 1 && ctor_name.is_none() {
            self.emit_u16(Op::LOCAL_GET, array_slot);
            return;
        }

        let idx_slot = self.define_local("__fortran_alloc_idx");
        self.emit_const(Value::F64(0.0));
        self.emit_u16(Op::LOCAL_SET, idx_slot);
        self.emit(Op::DROP);

        let loop_top = self.chunks[self.current].current_offset();
        self.emit_u16(Op::LOCAL_GET, idx_slot);
        self.emit_u16(Op::LOCAL_GET, dim_slots[0]);
        self.emit(Op::DYN_LT);
        let loop_exit = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, array_slot);
        self.emit_u16(Op::LOCAL_GET, idx_slot);
        if dim_slots.len() > 1 {
            self.emit_fortran_allocated_array(&dim_slots[1..], ctor_name);
        } else if let Some(ctor_name) = ctor_name {
            self.emit_fortran_ctor_call(ctor_name);
        } else {
            self.emit(Op::NULL);
        }
        common::collections::emit_set(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, idx_slot);
        self.emit_const(Value::F64(1.0));
        self.emit(Op::F64_ADD);
        self.emit_u16(Op::LOCAL_SET, idx_slot);
        self.emit(Op::DROP);
        self.emit_loop(loop_top);
        self.patch_jump(loop_exit);

        self.emit_u16(Op::LOCAL_GET, array_slot);
    }

    fn expr_prefers_numeric_add(&self, expr: &Expression) -> bool {
        self.infer_expr_type_hint(expr)
            .as_deref()
            .is_some_and(Self::is_numeric_type_hint)
    }

    fn compile_expr_with_numeric_add_hint(
        &mut self,
        expr: &Expression,
        prefer_numeric_add: bool,
    ) -> Result<(), String> {
        if prefer_numeric_add {
            if let ExprKind::Binary {
                op: BinOp::Add,
                left,
                right,
            } = &expr.kind
            {
                self.compile_expr_with_numeric_add_hint(left, true)?;
                self.compile_expr_with_numeric_add_hint(right, true)?;
                self.emit(Op::F64_ADD);
                return Ok(());
            }
        }

        self.compile_expr(expr)
    }

    fn emit_assignment_type_coercion_for_target(&mut self, target: &Expression) {
        let ExprKind::Ident(name) = &target.kind else {
            return;
        };
        self.emit_assignment_type_coercion_for_ident(name);
    }

    fn emit_assignment_type_coercion_for_ident(&mut self, name: &str) {
        if self.lookup_array_binding(name).is_some() {
            return;
        }
        let Some(type_hint) = self.lookup_var_type_hint(name).map(str::to_string) else {
            return;
        };
        let normalized = Self::normalize_type_hint(&type_hint);
        self.emit(Op::DUP);
        self.emit(Op::REF_IS_NULL);
        let skip = self.emit_jump(Op::BR_IF_TRUE);
        let line = self.line;
        match normalized.as_str() {
            "integer" | "int" | "int32" | "longint" | "long" | "int64"
            | "short" | "int16" | "uint" | "uint32" | "ulong" | "uint64"
            | "ushort" | "uint16" | "byte" | "sbyte" => {
                let number_idx = self.import("ecma:number", "Number");
                self.emit_host_call(number_idx, 1);
                common::convert::emit_to_int(self.chunk(), line);
            }
            "real" | "double" | "float" | "single" | "decimal" => {
                let number_idx = self.import("ecma:number", "Number");
                self.emit_host_call(number_idx, 1);
            }
            _ => {}
        }
        self.patch_jump(skip);
    }

    fn emit_file_key_compare(&mut self, relation: FileKeyRelation) {
        match relation {
            FileKeyRelation::Equal => self.emit(Op::DYN_EQ),
            FileKeyRelation::Greater => self.emit(Op::DYN_GT),
            FileKeyRelation::GreaterOrEqual => self.emit(Op::DYN_GE),
            FileKeyRelation::Less => self.emit(Op::DYN_LT),
            FileKeyRelation::LessOrEqual => self.emit(Op::DYN_LE),
        }
    }

    fn emit_global_map_get_into_local(&mut self, map_name: &str, key_slot: u16, value_slot: u16) {
        let map_key = self.shared_global_slot(map_name);
        self.emit_ensure_global_map(map_name);
        self.emit_u16(Op::GLOBAL_GET, map_key);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit(Op::ARRAY_GET);
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit(Op::DROP);
    }

    fn emit_global_map_set_from_local(&mut self, map_name: &str, key_slot: u16, value_slot: u16) {
        let map_key = self.shared_global_slot(map_name);
        self.emit_ensure_global_map(map_name);
        self.emit_u16(Op::GLOBAL_GET, map_key);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit(Op::ARRAY_SET);
        self.emit(Op::DROP);
    }

    fn emit_global_map_set_const(&mut self, map_name: &str, key_slot: u16, value: Value) {
        let map_key = self.shared_global_slot(map_name);
        self.emit_ensure_global_map(map_name);
        self.emit_u16(Op::GLOBAL_GET, map_key);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit_const(value);
        self.emit(Op::ARRAY_SET);
        self.emit(Op::DROP);
    }

    fn emit_global_map_set_null(&mut self, map_name: &str, key_slot: u16) {
        let map_key = self.shared_global_slot(map_name);
        self.emit_ensure_global_map(map_name);
        self.emit_u16(Op::GLOBAL_GET, map_key);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit(Op::NULL);
        self.emit(Op::ARRAY_SET);
        self.emit(Op::DROP);
    }

    fn emit_record_rows_cache(&mut self, file_slot: u16, rows_slot: u16, len_slot: u16) {
        let line = self.line;
        let path_map_key = self.shared_global_slot("__vb_file_path_by_handle");

        self.emit_global_map_get_into_local("__vb_record_rows_by_handle", file_slot, rows_slot);
        self.emit_u16(Op::LOCAL_GET, rows_slot);
        self.emit(Op::REF_IS_NULL);
        let cached_rows = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_ensure_global_map("__vb_file_path_by_handle");
        self.emit_u16(Op::GLOBAL_GET, path_map_key);
        self.emit_u16(Op::LOCAL_GET, file_slot);
        self.emit(Op::ARRAY_GET);
        let read_file_idx = self.import("wasi:filesystem", "readFile");
        self.emit_host_call(read_file_idx, 1);
        self.emit_const(Value::String(Arc::from("\n")));
        self.emit(Op::STR_SPLIT);
        self.emit_u16(Op::LOCAL_SET, rows_slot);
        self.emit(Op::DROP);

        let skip_trim = self.chunk().emit_block(line);
        self.emit_u16(Op::LOCAL_GET, rows_slot);
        common::collections::emit_array_length(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, len_slot);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, len_slot);
        self.emit(Op::I32_CONST_0);
        self.emit(Op::DYN_GT);
        self.emit(Op::DYN_NOT);
        self.chunk().emit_br_if(0, line);
        self.emit_u16(Op::LOCAL_GET, rows_slot);
        self.emit_u16(Op::LOCAL_GET, len_slot);
        self.emit(Op::I32_CONST_1);
        self.emit(Op::I32_SUB);
        common::collections::emit_get(&mut self.chunks, self.current, line);
        self.emit_const(Value::String(Arc::from("")));
        self.emit(Op::DYN_EQ);
        self.emit(Op::DYN_NOT);
        self.chunk().emit_br_if(0, line);
        self.emit_u16(Op::LOCAL_GET, rows_slot);
        common::collections::emit_pop(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);
        self.chunk().emit_end(line);
        self.chunk().patch_block(skip_trim);

        self.emit_global_map_set_from_local("__vb_record_rows_by_handle", file_slot, rows_slot);
        self.patch_jump(cached_rows);

        self.emit_u16(Op::LOCAL_GET, rows_slot);
        common::collections::emit_array_length(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, len_slot);
        self.emit(Op::DROP);
    }

    fn emit_record_assign_nulls(&mut self, variables: &[String]) {
        for variable in variables {
            self.emit(Op::NULL);
            self.emit_var_set(variable);
        }
    }

    fn emit_record_assign_values_from_local(&mut self, values_slot: u16, variables: &[String]) {
        for (index, variable) in variables.iter().enumerate() {
            self.emit_u16(Op::LOCAL_GET, values_slot);
            self.emit_const(Value::F64(index as f64));
            self.emit(Op::ARRAY_GET);
            self.emit_assignment_type_coercion_for_ident(variable);
            self.emit_var_set(variable);
        }
    }

    fn emit_record_rewrite_field_format(&mut self, field_format: Option<&RecordFieldFormat>) {
        let Some(field_format) = field_format else {
            return;
        };

        let number_idx = self.import("ecma:number", "Number");
        let to_fixed_idx = self.import("ecma:number", "toFixed");
        self.emit_host_call(number_idx, 1);
        self.emit_const(Value::F64(field_format.decimal_places as f64));
        self.emit_host_call(to_fixed_idx, 2);
    }

    fn vb_fixed_string_len(type_hint: &str) -> Option<i32> {
        let normalized = Self::normalize_type_hint(type_hint);
        let (base, len) = normalized.split_once('*')?;
        let base = base.trim();
        if base != "string" && base != "system.string" && !base.ends_with(".string") {
            return None;
        }
        len.trim().parse::<i32>().ok().filter(|len| *len >= 0)
    }

    fn emit_vb_fixed_string_adjust_from_stack(&mut self, target_len: i32, align_right: bool) {
        let line = self.line;
        let value_slot = self.define_local("__vb_fixed_string_value");
        let to_string = self.import("ecma:string", "String");
        let pad_idx = self.import("ecma:string", if align_right { "padStart" } else { "padEnd" });

        self.emit_host_call(to_string, 1);
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, value_slot);
    common::strings::emit_length(self.chunk(), line);
        self.emit_const(Value::I32(target_len));
        self.emit(Op::DYN_GT);
        let no_truncate = self.emit_jump(Op::BR_IF_FALSE);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit(Op::I32_CONST_0);
        self.emit_const(Value::I32(target_len));
    common::strings::emit_substring(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit(Op::DROP);
        self.patch_jump(no_truncate);

        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_const(Value::I32(target_len));
        self.emit_const(Value::String(Arc::from(" ")));
        self.emit_host_call(pad_idx, 3);
    }

    fn compile_vb_fixed_string_stmt(
        &mut self,
        target: &Expression,
        value: &Expression,
        align_right: bool,
    ) -> Result<(), String> {
        let ExprKind::Ident(name) = &target.kind else {
            self.compile_expr(value)?;
            self.emit(Op::DROP);
            return Ok(());
        };
        let Some(type_hint) = self.lookup_var_type_hint(name) else {
            self.compile_expr(value)?;
            self.compile_assign_target(target)?;
            return Ok(());
        };
        let Some(target_len) = Self::vb_fixed_string_len(type_hint) else {
            self.compile_expr(value)?;
            self.compile_assign_target(target)?;
            return Ok(());
        };

        self.compile_expr(value)?;
        self.emit_vb_fixed_string_adjust_from_stack(target_len, align_right);
        self.compile_assign_target(target)
    }

    fn compile_vb_mid_stmt(
        &mut self,
        target: &Expression,
        start: &Expression,
        count: &Expression,
        value: &Expression,
    ) -> Result<(), String> {
        let line = self.line;
        let target_slot = self.define_local("__vb_mid_target");
        let start_slot = self.define_local("__vb_mid_start");
        let count_slot = self.define_local("__vb_mid_count");
        let value_slot = self.define_local("__vb_mid_value");
        let prefix_slot = self.define_local("__vb_mid_prefix");
        let replace_slot = self.define_local("__vb_mid_replace");
        let to_string = self.import("ecma:string", "String");

        self.compile_expr(target)?;
        self.emit_u16(Op::LOCAL_SET, target_slot);
        self.emit(Op::DROP);
        self.compile_expr(start)?;
    common::convert::emit_to_int(self.chunk(), line);
        self.emit_const(Value::I32(1));
        self.emit(Op::I32_SUB);
        self.emit_u16(Op::LOCAL_SET, start_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, start_slot);
        self.emit(Op::I32_CONST_0);
        self.emit(Op::DYN_GE);
        let start_ok = self.emit_jump(Op::BR_IF_TRUE);
        self.emit(Op::I32_CONST_0);
        self.emit_u16(Op::LOCAL_SET, start_slot);
        self.emit(Op::DROP);
        self.patch_jump(start_ok);

        self.compile_expr(value)?;
        self.emit_host_call(to_string, 1);
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit(Op::DROP);

        self.compile_expr(count)?;
    common::convert::emit_to_int(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, count_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, count_slot);
        self.emit_const(Value::I32(0));
        self.emit(Op::DYN_GE);
        let use_explicit_count = self.emit_jump(Op::BR_IF_TRUE);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        common::strings::emit_length(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, replace_slot);
        self.emit(Op::DROP);
        let count_done = self.emit_jump(Op::BR);
        self.patch_jump(use_explicit_count);
        self.emit_u16(Op::LOCAL_GET, count_slot);
        self.emit_u16(Op::LOCAL_SET, replace_slot);
        self.emit(Op::DROP);
        self.patch_jump(count_done);

        self.emit_u16(Op::LOCAL_GET, target_slot);
        self.emit(Op::I32_CONST_0);
        self.emit_u16(Op::LOCAL_GET, start_slot);
        common::strings::emit_substring(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, prefix_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, prefix_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit(Op::DYN_ADD);
        self.emit_u16(Op::LOCAL_GET, target_slot);
        self.emit_u16(Op::LOCAL_GET, start_slot);
        self.emit_u16(Op::LOCAL_GET, replace_slot);
        self.emit(Op::I32_ADD);
        self.emit_u16(Op::LOCAL_GET, target_slot);
        common::strings::emit_length(self.chunk(), line);
        common::strings::emit_substring(self.chunk(), line);
        self.emit(Op::DYN_ADD);

        if let ExprKind::Ident(name) = &target.kind {
            if let Some(type_hint) = self.lookup_var_type_hint(name) {
                if let Some(target_len) = Self::vb_fixed_string_len(type_hint) {
                    self.emit_vb_fixed_string_adjust_from_stack(target_len, false);
                }
            }
        }

        self.compile_assign_target(target)
    }

    fn compile_vb_err_raise_stmt(&mut self, args: &[Argument]) -> Result<(), String> {
        if let Some(description) = args.get(2).or_else(|| args.get(1)).or_else(|| args.first()) {
            self.compile_expr(&description.value)?;
        } else {
            self.emit_const(Value::String(Arc::from("")));
        }

        self.emit_js_exception_ctor_from_message_value("Exception")?;

        if let Some(number) = args.first() {
            self.emit(Op::DUP);
            self.compile_expr(&number.value)?;
            let key = self.str_const("number");
            self.emit_u16(Op::STRUCT_SET, key);
            self.emit(Op::DROP);
        }

        if let Some(source) = args.get(1) {
            self.emit(Op::DUP);
            self.compile_expr(&source.value)?;
            let key = self.str_const("source");
            self.emit_u16(Op::STRUCT_SET, key);
            self.emit(Op::DROP);
        }

        let line = self.line;
        common::errors::emit_throw(self.chunk(), line);
        Ok(())
    }

    pub(super) fn is_collection_like_type_hint(type_hint: &str) -> bool {
        let normalized = Self::normalize_type_hint(type_hint);
        let bare = normalized
            .split('<')
            .next()
            .unwrap_or(normalized.as_str())
            .trim_end_matches('?');
        let terminal = bare.rsplit('.').next().unwrap_or(bare);
        Self::is_string_type_hint(type_hint)
            || matches!(
                terminal,
                "list"
                    | "arraylist"
                    | "dictionary"
                    | "queue"
                    | "stack"
                    | "hashset"
                    | "set"
                    | "collection"
                    | "icollection"
                    | "readonlycollection"
                    | "enumerable"
                    | "ienumerable"
                    | "readonlylist"
                    | "ilist"
                    | "array"
            )
            || bare.ends_with("[]")
            || normalized.ends_with("()")
    }

    fn is_dictionary_type_hint(type_hint: &str) -> bool {
        let normalized = Self::normalize_type_hint(type_hint);
        normalized.contains("dictionary") || normalized.ends_with("hashtable")
    }

    fn is_sorted_dictionary_type_hint(type_hint: &str) -> bool {
        Self::normalize_type_hint(type_hint).contains("sorteddictionary")
    }

    fn is_sorted_set_type_hint(type_hint: &str) -> bool {
        Self::normalize_type_hint(type_hint).contains("sortedset")
    }

    pub(super) fn is_pascal_set_type_hint(type_hint: &str) -> bool {
        Self::normalize_type_hint(type_hint).starts_with("set of ")
    }

    fn is_case_insensitive_string_key_type_hint(type_hint: &str) -> bool {
        Self::normalize_type_hint(type_hint).contains("#ordinalignorecase")
    }

    fn expr_uses_case_insensitive_string_keys(&self, expr: &Expression) -> bool {
        match &expr.kind {
            ExprKind::Ident(name) => self
                .lookup_var_type_hint(name)
                .is_some_and(Self::is_case_insensitive_string_key_type_hint),
            _ => self
                .infer_expr_type_hint(expr)
                .as_deref()
                .is_some_and(Self::is_case_insensitive_string_key_type_hint),
        }
    }

    fn compile_collection_key(&mut self, owner: &Expression, key: &Expression) -> Result<(), String> {
        self.compile_array_index_operand_for_owner(owner, key)?;
        if self.expr_uses_case_insensitive_string_keys(owner) {
            let line = self.line;
            common::strings::emit_to_lower(self.chunk(), line);
        }
        Ok(())
    }

    fn emit_php_array_assoc_key_tracking(&mut self, obj_slot: u16, key_slot: u16, line: u32) {
        let missing_slot = self.emit_php_array_assoc_key_is_missing(obj_slot, key_slot, line);
        self.emit_php_array_assoc_key_tracking_when_missing(obj_slot, key_slot, missing_slot, line);
    }

    fn emit_php_array_assoc_key_is_missing(&mut self, obj_slot: u16, key_slot: u16, line: u32) -> u16 {
        let missing_tmp = self.define_local("__php_assoc_key_missing");

        self.emit_const(Value::Bool(false));
        self.emit_u16(Op::LOCAL_SET, missing_tmp);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit(Op::REF_IS_STRING);
        let not_string = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        common::collections::emit_get(&mut self.chunks, self.current, line);
        self.emit(Op::DUP);
        self.emit(Op::REF_IS_NULL);
        let value_is_null = self.emit_jump(Op::BR_IF_TRUE);
        self.emit(Op::REF_IS_UNDEFINED);
        let after_missing_check = self.emit_jump(Op::BR);
        self.patch_jump(value_is_null);
        self.emit(Op::DROP);
        self.emit_const(Value::Bool(true));
        self.patch_jump(after_missing_check);
        self.emit_u16(Op::LOCAL_SET, missing_tmp);
        self.emit(Op::DROP);

        self.patch_jump(not_string);
        missing_tmp
    }

    fn emit_php_array_assoc_key_tracking_when_missing(
        &mut self,
        obj_slot: u16,
        key_slot: u16,
        missing_slot: u16,
        line: u32,
    ) {
        let tracker_tmp = self.define_local("__php_assoc_keys_csv");
        let keys_tmp = self.define_local("__php_assoc_keys");
        let tracker_key = self.str_const("vybe$assoc_keys_csv");
        let keys_key = self.str_const("__keys");

        self.emit_u16(Op::LOCAL_GET, missing_slot);
        self.emit(Op::DYN_TO_BOOL);
        self.emit(Op::DYN_NOT);
        let done = self.emit_jump(Op::BR_IF_TRUE);

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u16(Op::STRUCT_GET, keys_key);
        self.emit_u16(Op::LOCAL_SET, keys_tmp);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, keys_tmp);
        self.emit(Op::DUP);
        self.emit(Op::REF_IS_NULL);
        let no_keys = self.emit_jump(Op::BR_IF_TRUE);
        self.emit(Op::REF_IS_UNDEFINED);
        let no_keys_undef = self.emit_jump(Op::BR_IF_TRUE);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, keys_tmp);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        common::collections::emit_push(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);
        let keys_done = self.emit_jump(Op::BR);

        self.patch_jump(no_keys);
        self.emit(Op::DROP);
        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
        self.emit_u16(Op::LOCAL_SET, keys_tmp);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u16(Op::LOCAL_GET, keys_tmp);
        self.emit_u16(Op::STRUCT_SET, keys_key);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, keys_tmp);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        common::collections::emit_push(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);
        let keys_init_done = self.emit_jump(Op::BR);

        self.patch_jump(no_keys_undef);
        self.emit(Op::DROP);
        common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
        self.emit_u16(Op::LOCAL_SET, keys_tmp);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u16(Op::LOCAL_GET, keys_tmp);
        self.emit_u16(Op::STRUCT_SET, keys_key);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, keys_tmp);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        common::collections::emit_push(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);

        self.patch_jump(keys_done);
        self.patch_jump(keys_init_done);

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u16(Op::STRUCT_GET, tracker_key);
        self.emit_u16(Op::LOCAL_SET, tracker_tmp);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, tracker_tmp);
        self.emit(Op::DUP);
        self.emit(Op::REF_IS_NULL);
        let tracker_is_null = self.emit_jump(Op::BR_IF_TRUE);
        self.emit(Op::REF_IS_UNDEFINED);
        let tracker_presence_done = self.emit_jump(Op::BR);
        self.patch_jump(tracker_is_null);
        self.emit(Op::DROP);
        self.emit_const(Value::Bool(true));
        self.patch_jump(tracker_presence_done);
        let no_tracker = self.emit_jump(Op::BR_IF_TRUE);

        self.emit_u16(Op::LOCAL_GET, tracker_tmp);
        self.emit_const(Value::String(Arc::from("\x1F")));
        common::strings::emit_str_concat(self.chunk(), line);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        common::strings::emit_str_concat(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, tracker_tmp);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u16(Op::LOCAL_GET, tracker_tmp);
    self.emit_u16(Op::STRUCT_SET, tracker_key);
    self.emit(Op::DROP);
        let after_track = self.emit_jump(Op::BR);

        self.patch_jump(no_tracker);
        self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit_u16(Op::LOCAL_SET, tracker_tmp);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit_u16(Op::LOCAL_GET, tracker_tmp);
    self.emit_u16(Op::STRUCT_SET, tracker_key);
    self.emit(Op::DROP);

        self.patch_jump(after_track);
        self.patch_jump(done);
    }

    fn emit_php_promote_empty_array_for_string_key(&mut self, obj_slot: u16, key_slot: u16, line: u32) {
        let is_array_idx = self.import("ecma:array", "isArray");
        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.chunk().emit_op_u16(Op::CALL_IMPORT, is_array_idx, line);
        self.chunk().emit(1, line);
        let done = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit(Op::REF_IS_STRING);
        let not_string = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, obj_slot);
        self.emit(Op::ARRAY_LENGTH);
        self.emit(Op::I32_CONST_0);
        self.emit(Op::DYN_NE);
        let not_empty = self.emit_jump(Op::BR_IF_TRUE);
        let map_new_idx = self.import("ecma:map", "new");
        self.chunk().emit_op_u16(Op::CALL_IMPORT, map_new_idx, line);
        self.chunk().emit(0, line);
        self.emit_u16(Op::LOCAL_SET, obj_slot);
        self.emit(Op::DROP);

        self.patch_jump(not_empty);
        self.patch_jump(not_string);
        self.patch_jump(done);
    }

    pub(super) fn is_callable_type_hint(type_hint: &str) -> bool {
        let normalized = Self::normalize_type_hint(type_hint);
        if normalized.ends_with("()") {
            return false;
        }
        let lower = normalized.to_ascii_lowercase();
        if lower.starts_with("procedure(") || lower == "procedure" {
            return true;
        }
        let leaf = lower.rsplit('.').next().unwrap_or(lower.as_str());
        let bare = leaf
            .split('<')
            .next()
            .unwrap_or(leaf)
            .split('(')
            .next()
            .unwrap_or(leaf)
            .trim();
        matches!(bare, "func" | "action" | "eventhandler" | "predicate" | "comparison" | "converter")
            || lower.contains(" delegate")
    }

    fn lookup_var_type_hint(&self, name: &str) -> Option<&str> {
        if let Some(binding) = self.static_local_binding(name) {
            if let Some(type_hint) = binding.type_hint.as_deref() {
                return Some(type_hint);
            }
        }
        if let Some(type_hint) = self.scope().resolve_type(name) {
            return Some(type_hint);
        }
        if !self.case_sensitive {
            if let Some(type_hint) = self.scope().resolve_type_ci(name) {
                return Some(type_hint);
            }
        }
        for scope in self.scopes.iter().rev().skip(1) {
            if let Some(type_hint) = scope.resolve_type(name) {
                return Some(type_hint);
            }
            if !self.case_sensitive {
                if let Some(type_hint) = scope.resolve_type_ci(name) {
                    return Some(type_hint);
                }
            }
        }
        if let Some(type_hint) = self.lookup_implicit_self_field_type_hint(name) {
            return Some(type_hint);
        }
        let cname = self.canon(name);
        self.global_type_hints.get(&cname).map(|s| s.as_str())
    }

    pub(super) fn has_accessible_local_binding(&self, name: &str) -> bool {
        if self.static_local_binding(name).is_some() {
            return true;
        }
        self.scopes.iter().rev().any(|scope| {
            scope.resolve(name).is_some()
                || (!self.case_sensitive && scope.resolve_ci(name).is_some())
        })
    }

    fn static_local_binding(&self, name: &str) -> Option<&StaticLocalBinding> {
        let canon_name = self.canon(name);
        self.static_local_bindings
            .iter()
            .rev()
            .find_map(|bindings| bindings.get(&canon_name))
    }

    fn has_static_local_binding(&self, name: &str) -> bool {
        self.static_local_binding(name).is_some()
    }

    fn array_binding_key(&self, name: &str) -> String {
        let canon_name = self.canon(name);
        if self.scopes.len() > 1 {
            let class_name = self.current_class.as_deref().unwrap_or("<module>");
            let func_name = self.current_func_name.as_deref().unwrap_or("<top>");
            format!("{}::{}::{}", self.canon(class_name), self.canon(func_name), canon_name)
        } else {
            canon_name
        }
    }

    fn record_array_binding(&mut self, name: &str, metadata: ArrayBindingMetadata) {
        let key = self.array_binding_key(name);
        self.array_bindings.insert(key, metadata);
    }

    fn lookup_array_binding(&self, name: &str) -> Option<&ArrayBindingMetadata> {
        let key = self.array_binding_key(name);
        self.array_bindings.get(&key).or_else(|| self.array_bindings.get(&self.canon(name)))
    }

    fn pascal_array_index_bounds_for_owner(&self, owner: &Expression) -> Option<PascalArrayBoundsMetadata> {
        if self.profile.name != "pascal" {
            return None;
        }

        if let ExprKind::Ident(name) = &owner.kind {
            if let Some(bounds) = self.lookup_array_binding(name).and_then(|binding| binding.pascal_bounds.clone()) {
                return Some(bounds);
            }
        }

        self.infer_expr_type_hint(owner)
            .as_deref()
            .and_then(|type_hint| self.pascal_array_type_hint_metadata(type_hint))
    }

    fn profile_array_index_semantics(&self) -> Option<ArrayIndexSemantics> {
        match self.profile.name.as_str() {
            _ => None,
        }
    }

    fn normalized_array_index_operand_for_owner(&self, owner: &Expression, index: Expression) -> Expression {
        if let Some(bounds) = self.pascal_array_index_bounds_for_owner(owner) {
            if let Some(dimension) = bounds.dimensions.first() {
                let normalized_index = if dimension.uses_char_ordinal {
                    Self::pascal_ordinal_index_expr(index)
                } else {
                    index
                };
                return normalize_array_index_operand(
                    normalized_index,
                    ArrayIndexSemantics {
                        first_index: dimension.first_index,
                    },
                );
            }
        }

        if let Some(semantics) = self.profile_array_index_semantics() {
            return normalize_array_index_operand(index, semantics);
        }

        index
    }

    fn compile_array_index_operand(&mut self, index: &Expression) -> Result<(), String> {
        if let Some(semantics) = self.profile_array_index_semantics() {
            let normalized = normalize_array_index_operand(index.clone(), semantics);
            self.compile_expr(&normalized)
        } else {
            self.compile_expr(index)
        }
    }

    fn compile_array_index_operand_for_owner(&mut self, owner: &Expression, index: &Expression) -> Result<(), String> {
        let normalized = self.normalized_array_index_operand_for_owner(owner, index.clone());
        self.compile_expr(&normalized)
    }

    fn compile_setlength(&mut self, target: &Expression, len_expr: &Expression) -> Result<(), String> {
        let line = self.line;
        let old_slot = self.define_local("__setlength_old");
        let new_len_slot = self.define_local("__setlength_new_len");
        let new_slot = self.define_local("__setlength_new");
        let copy_len_slot = self.define_local("__setlength_copy_len");
        let idx_slot = self.define_local("__setlength_idx");

        self.compile_expr(target)?;
        self.emit_u16(Op::LOCAL_SET, old_slot);
        self.emit(Op::DROP);

        self.compile_expr(len_expr)?;
        self.emit_u16(Op::LOCAL_SET, new_len_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, new_len_slot);
        common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
        self.emit_u16(Op::LOCAL_SET, new_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, old_slot);
        self.emit(Op::REF_IS_ARRAY);
        let skip_copy = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, old_slot);
        common::collections::emit_len(&mut self.chunks, self.current, line);
        self.emit_u16(Op::LOCAL_GET, new_len_slot);
        common::math::emit_min(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, copy_len_slot);
        self.emit(Op::DROP);

        self.emit(Op::I32_CONST_0);
        self.emit_u16(Op::LOCAL_SET, idx_slot);
        self.emit(Op::DROP);

        let copy_block = self.chunk().emit_block(line);
        let (copy_loop, _) = self.chunk().emit_loop_s(line);
        self.emit_u16(Op::LOCAL_GET, idx_slot);
        self.emit_u16(Op::LOCAL_GET, copy_len_slot);
        self.emit(Op::DYN_LT);
        self.emit(Op::DYN_NOT);
        self.chunk().emit_br_if(1, line);

        self.emit_u16(Op::LOCAL_GET, new_slot);
        self.emit_u16(Op::LOCAL_GET, idx_slot);
        self.emit_u16(Op::LOCAL_GET, old_slot);
        self.emit_u16(Op::LOCAL_GET, idx_slot);
        common::collections::emit_get(&mut self.chunks, self.current, line);
        common::collections::emit_set(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, idx_slot);
        self.emit_const(Value::F64(1.0));
        self.emit(Op::DYN_ADD);
        self.emit_u16(Op::LOCAL_SET, idx_slot);
        self.emit(Op::DROP);
        self.chunk().emit_br(0, line);
        self.chunk().emit_end(line);
        self.chunk().patch_loop(copy_loop);
        self.chunk().emit_end(line);
        self.chunk().patch_block(copy_block);

        self.patch_jump(skip_copy);

        self.emit_u16(Op::LOCAL_GET, new_slot);
        if let ExprKind::Ident(name) = &target.kind {
            self.emit_var_set(name);
        } else {
            self.compile_assign_target(target)?;
        }

        self.emit(Op::NULL);
        Ok(())
    }

    fn ensure_static_local_binding(
        &mut self,
        name: &str,
        type_hint: Option<String>,
    ) -> Result<StaticLocalBinding, String> {
        let canon_name = self.canon(name);
        let normalized_type_hint = type_hint
            .as_deref()
            .map(Self::normalize_type_hint);

        if let Some(existing) = self
            .static_local_bindings
            .last_mut()
            .and_then(|bindings| bindings.get_mut(&canon_name))
        {
            if existing.type_hint.is_none() {
                existing.type_hint = normalized_type_hint;
            }
            return Ok(existing.clone());
        }

        let func_name = self
            .current_func_name
            .as_deref()
            .map(|name| self.canon(name))
            .unwrap_or_else(|| "anon".to_string());
        let Some(bindings) = self.static_local_bindings.last_mut() else {
            return Err(format!("static local `{name}` declared outside a function"));
        };
        let global_name = format!("__staticlocal_{}_{}_{}", self.current, func_name, canon_name);
        let binding = StaticLocalBinding {
            init_flag_name: format!("{}__init", global_name),
            global_name,
            type_hint: normalized_type_hint,
        };
        bindings.insert(canon_name, binding.clone());
        Ok(binding)
    }

    fn emit_vb_fixed_array_initializer(&mut self, bounds: &[Expression]) -> Result<(), String> {
        let line = self.line;
        if bounds.is_empty() {
            self.emit(Op::NULL);
            return Ok(());
        }

        self.compile_expr(&bounds[0])?;
        self.emit_const(Value::F64(1.0));
        self.emit(Op::DYN_ADD);

        if bounds.len() == 1 {
            common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
            return Ok(());
        }

        let len_slot = self.define_local("__vb_md_len");
        self.emit_u16(Op::LOCAL_SET, len_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, len_slot);
        common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
        let array_slot = self.define_local("__vb_md_array");
        self.emit_u16(Op::LOCAL_SET, array_slot);
        self.emit(Op::DROP);

        let index_slot = self.define_local("__vb_md_index");
        self.emit(Op::F64_CONST_0);
        self.emit_u16(Op::LOCAL_SET, index_slot);
        self.emit(Op::DROP);

        let fill_block = self.chunk().emit_block(line);
        let (fill_loop, _) = self.chunk().emit_loop_s(line);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_u16(Op::LOCAL_GET, len_slot);
        self.emit(Op::DYN_LT);
        self.emit(Op::DYN_NOT);
        self.chunk().emit_br_if(1, line);

        self.emit_u16(Op::LOCAL_GET, array_slot);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_vb_fixed_array_initializer(&bounds[1..])?;
        common::collections::emit_set(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_const(Value::F64(1.0));
        self.emit(Op::DYN_ADD);
        self.emit_u16(Op::LOCAL_SET, index_slot);
        self.emit(Op::DROP);
        self.chunk().emit_br(0, line);
        self.chunk().emit_end(line);
        self.chunk().patch_loop(fill_loop);
        self.chunk().emit_end(line);
        self.chunk().patch_block(fill_block);

        self.emit_u16(Op::LOCAL_GET, array_slot);
        Ok(())
    }

    fn emit_pascal_fixed_array_initializer(&mut self, dimensions: &[PascalArrayDimensionMetadata]) -> Result<(), String> {
        let line = self.line;
        if dimensions.is_empty() {
            self.emit(Op::NULL);
            return Ok(());
        }

        self.emit_const(Value::F64(dimensions[0].length as f64));

        if dimensions.len() == 1 {
            common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
            return Ok(());
        }

        let len_slot = self.define_local("__pascal_md_len");
        self.emit_u16(Op::LOCAL_SET, len_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, len_slot);
        common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
        let array_slot = self.define_local("__pascal_md_array");
        self.emit_u16(Op::LOCAL_SET, array_slot);
        self.emit(Op::DROP);

        let index_slot = self.define_local("__pascal_md_index");
        self.emit(Op::F64_CONST_0);
        self.emit_u16(Op::LOCAL_SET, index_slot);
        self.emit(Op::DROP);

        let fill_block = self.chunk().emit_block(line);
        let (fill_loop, _) = self.chunk().emit_loop_s(line);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_u16(Op::LOCAL_GET, len_slot);
        self.emit(Op::DYN_LT);
        self.emit(Op::DYN_NOT);
        self.chunk().emit_br_if(1, line);

        self.emit_u16(Op::LOCAL_GET, array_slot);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_pascal_fixed_array_initializer(&dimensions[1..])?;
        common::collections::emit_set(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_const(Value::F64(1.0));
        self.emit(Op::DYN_ADD);
        self.emit_u16(Op::LOCAL_SET, index_slot);
        self.emit(Op::DROP);
        self.chunk().emit_br(0, line);
        self.chunk().emit_end(line);
        self.chunk().patch_loop(fill_loop);
        self.chunk().emit_end(line);
        self.chunk().patch_block(fill_block);

        self.emit_u16(Op::LOCAL_GET, array_slot);
        Ok(())
    }

    fn emit_fortran_fixed_array_initializer(&mut self, bounds: &[Expression]) -> Result<(), String> {
        let line = self.line;
        if bounds.is_empty() {
            self.emit(Op::NULL);
            return Ok(());
        }

        self.compile_expr(&bounds[0])?;

        if bounds.len() == 1 {
            common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
            return Ok(());
        }

        let len_slot = self.define_local("__fortran_md_len");
        self.emit_u16(Op::LOCAL_SET, len_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, len_slot);
        common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
        let array_slot = self.define_local("__fortran_md_array");
        self.emit_u16(Op::LOCAL_SET, array_slot);
        self.emit(Op::DROP);

        let index_slot = self.define_local("__fortran_md_index");
        self.emit(Op::F64_CONST_0);
        self.emit_u16(Op::LOCAL_SET, index_slot);
        self.emit(Op::DROP);

        let fill_block = self.chunk().emit_block(line);
        let (fill_loop, _) = self.chunk().emit_loop_s(line);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_u16(Op::LOCAL_GET, len_slot);
        self.emit(Op::DYN_LT);
        self.emit(Op::DYN_NOT);
        self.chunk().emit_br_if(1, line);

        self.emit_u16(Op::LOCAL_GET, array_slot);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_fortran_fixed_array_initializer(&bounds[1..])?;
        common::collections::emit_set(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_const(Value::F64(1.0));
        self.emit(Op::DYN_ADD);
        self.emit_u16(Op::LOCAL_SET, index_slot);
        self.emit(Op::DROP);
        self.chunk().emit_br(0, line);
        self.chunk().emit_end(line);
        self.chunk().patch_loop(fill_loop);
        self.chunk().emit_end(line);
        self.chunk().patch_block(fill_block);

        self.emit_u16(Op::LOCAL_GET, array_slot);
        Ok(())
    }

    fn is_fortran_fixed_array_synth_init(expr: &Expression) -> bool {
        matches!(
            &expr.kind,
            ExprKind::Call { callee, args, optional: false }
                if matches!(&callee.kind, ExprKind::Ident(name) if name == "Array")
                    && args.len() == 2
                    && matches!(args[1].value.kind, ExprKind::Lit(Literal::Int(0)))
        )
    }

    fn emit_var_decl_initializer_value(&mut self, decl: &VarDeclarator, resolved_type_hint: Option<&str>) -> Result<(), String> {
        if let Some(ref init_expr) = decl.init {
            if self.profile.name == "fortran"
                && decl.array_bounds.as_ref().is_some_and(|bounds| !bounds.is_empty())
                && Self::is_fortran_fixed_array_synth_init(init_expr)
            {
                self.emit_fortran_fixed_array_initializer(decl.array_bounds.as_ref().expect("checked non-empty bounds"))?;
            } else {
                self.compile_expr_with_value_copy(init_expr)?;
                self.maybe_promote_pascal_array_literal_to_set(decl.type_hint.as_deref(), init_expr);
            }
        } else if let Some(ref bounds) = decl.array_bounds {
            if self.profile.name == "fortran" {
                self.emit_fortran_fixed_array_initializer(bounds)?;
            } else if self.profile.name == "vb" && bounds.len() > 1 {
                self.emit_vb_fixed_array_initializer(bounds)?;
            } else if let Some(size_expr) = bounds.first() {
                let line = self.line;
                self.compile_expr(size_expr)?;
                self.emit_const(Value::F64(1.0));
                self.emit(Op::DYN_ADD);
                common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
            } else {
                self.emit(Op::NULL);
            }
        } else {
            let resolved_type_hint = resolved_type_hint
                .map(str::to_string)
                .or_else(|| decl.type_hint.as_deref().map(|type_hint| self.resolve_source_type_alias(type_hint)));
            let effective_type_hint = resolved_type_hint.as_deref().or(decl.type_hint.as_deref());

            if let Some(metadata) = effective_type_hint
                .and_then(|type_hint| self.pascal_array_type_hint_metadata(type_hint))
                .filter(|metadata| metadata.is_fixed)
            {
                if metadata.dimensions.len() > 1 {
                    self.emit_pascal_fixed_array_initializer(&metadata.dimensions)?;
                } else if let Some(dimension) = metadata.dimensions.first() {
                    let line = self.line;
                    self.emit_const(Value::F64(dimension.length as f64));
                    common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
                } else {
                    self.emit(Op::NULL);
                }
            } else if effective_type_hint.and_then(Self::vb_fixed_string_len).is_some() {
                self.emit_const(Value::String(Arc::from("")));
            } else if let Some(type_name) = decl
                .type_hint
                .as_deref()
                .and_then(|type_hint| self.user_value_type_name_from_hint(type_hint))
            {
                let ctor_global = {
                    let overload = format!("{}$arity0", type_name);
                    if self.defined_globals.contains(&overload) {
                        overload
                    } else {
                        type_name.clone()
                    }
                };
                let idx = self.str_const(&ctor_global);
                self.emit_u16(Op::GLOBAL_GET, idx);
                self.emit_u8(Op::CALL_REF, 0);
                return Ok(());
            } else {
                match effective_type_hint.map(|s| s.to_lowercase()).as_deref() {
                    Some("integer") | Some("int") | Some("longint") | Some("real") | Some("double") | Some("float") => {
                        self.emit(Op::F64_CONST_0);
                    }
                    Some("boolean") | Some("bool") => self.emit(Op::FALSE),
                    Some(type_hint) if Self::is_string_type_hint(type_hint) => self.emit_const(Value::String(Arc::from(""))),
                    _ => self.emit(Op::NULL),
                }
            }
        }
        if let Some(target_len) = decl.type_hint.as_deref().and_then(Self::vb_fixed_string_len) {
            self.emit_vb_fixed_string_adjust_from_stack(target_len, false);
        }
        Ok(())
    }

    fn lookup_implicit_self_field_type_hint(&self, name: &str) -> Option<&str> {
        if !self.current_class_implicit_self {
            return None;
        }

        let canon_name = self.canon(name);
        let mut current = self.current_class.as_deref();
        while let Some(class_name) = current {
            let pending = self.pending_classes.get(class_name)?;
            if let Some(type_hint) = pending.instance_field_types.get(&canon_name) {
                return Some(type_hint.as_str());
            }
            current = pending.parent.as_deref();
        }
        None
    }

    pub(super) fn prefers_type_qualified_member_lookup(&self, type_name: &str, member_name: &str) -> bool {
        if self.enum_member_ordinal(type_name, member_name).is_some() {
            return true;
        }

        let type_canon = self.canon(type_name);
        let Some(pending) = self.pending_classes.get(&type_canon).or_else(|| {
            self.pending_classes
                .iter()
                .find(|(name, _)| name.eq_ignore_ascii_case(type_name) || name.eq_ignore_ascii_case(&type_canon))
                .map(|(_, pending)| pending)
        }) else {
            return false;
        };

        let member_canon = self.canon(member_name);
        pending.static_fields.iter().any(|name| name == &member_canon)
            || pending.static_method_names.iter().any(|name| self.canon(name) == member_canon)
            || pending.static_method_overloads.contains_key(&member_canon)
            || pending.nested_types.iter().any(|name| self.canon(name) == member_canon)
    }

    fn expr_terminal_type_name(expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(name) => Some(name.rsplit('.').next().unwrap_or(name).to_string()),
            ExprKind::Member { field, .. } => Some(field.clone()),
            _ => None,
        }
    }

    fn infer_dotnet_factory_return_type(&self, callee: &Expression) -> Option<String> {
        if !self.profile.namespaces.use_dotnet {
            return None;
        }
        let ExprKind::Member { object, field, .. } = &callee.kind else {
            return None;
        };
        let class_name = Self::expr_terminal_type_name(object)?;
        if class_name.eq_ignore_ascii_case("TimeSpan")
            && matches!(field.as_str(), "FromDays" | "FromHours" | "FromMinutes" | "FromSeconds" | "FromMilliseconds" | "Zero")
        {
            return Some("TimeSpan".into());
        }
        if class_name.eq_ignore_ascii_case("DateTime")
            && matches!(field.as_str(), "Now" | "UtcNow" | "Today" | "Parse")
        {
            return Some("DateTime".into());
        }
        if class_name.eq_ignore_ascii_case("Convert")
            && field.eq_ignore_ascii_case("ToDateTime")
        {
            return Some("DateTime".into());
        }
        if class_name.eq_ignore_ascii_case("Guid")
            && matches!(field.as_str(), "Empty" | "NewGuid" | "Parse")
        {
            return Some("Guid".into());
        }
        if class_name.eq_ignore_ascii_case("Version")
            && matches!(field.as_str(), "Parse")
        {
            return Some("Version".into());
        }
        None
    }

    fn infer_function_return_type(&self, callee: &Expression) -> Option<String> {
        match &callee.kind {
            ExprKind::Ident(name) => {
                if self.profile.name == "vb" {
                    if name.eq_ignore_ascii_case("command") || name.eq_ignore_ascii_case("environ") {
                        return Some("String".into());
                    }
                    if name.eq_ignore_ascii_case("timer") {
                        return Some("Double".into());
                    }
                }
                self.function_return_types.get(&self.canon(name)).cloned()
            }
            ExprKind::Member { object, field, .. } => {
                if let Some(receiver_type) = self.infer_expr_type_hint(object) {
                    let receiver_trimmed = receiver_type.trim().trim_end_matches('?').trim();
                    let receiver_base = receiver_trimmed
                        .split('<')
                        .next()
                        .unwrap_or(receiver_trimmed)
                        .trim();
                    let receiver_key = self
                        .resolve_pending_class_name_for_type_hint(&receiver_type)
                        .unwrap_or_else(|| self.canon(receiver_base));
                    let qualified = self.canon(&format!("{}.{}", receiver_key, field));
                    if let Some(return_type) = self.function_return_types.get(&qualified) {
                        return Some(return_type.clone());
                    }
                }
                if let ExprKind::Ident(object_name) = &object.kind {
                    let qualified = self.canon(&format!("{}.{}", object_name, field));
                    if let Some(return_type) = self.function_return_types.get(&qualified) {
                        return Some(return_type.clone());
                    }
                }
                self.function_return_types.get(&self.canon(field)).cloned()
            }
            _ => None,
        }
    }

    fn infer_array_element_type_hint<'a>(&self, values: impl IntoIterator<Item = &'a Expression>) -> String {
        let mut element_type: Option<String> = None;
        for value in values {
            let inferred = self.infer_expr_type_hint(value).unwrap_or_else(|| "object".into());
            match &element_type {
                None => element_type = Some(inferred),
                Some(existing)
                    if Self::normalize_type_hint(existing) == Self::normalize_type_hint(&inferred) => {}
                Some(_) => {
                    element_type = Some("object".into());
                    break;
                }
            }
        }
        element_type.unwrap_or_else(|| "object".into())
    }

    fn member_access_path(expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(name) => Some(name.clone()),
            ExprKind::Member { object, field, .. } => {
                let prefix = Self::member_access_path(object)?;
                Some(format!("{prefix}.{field}"))
            }
            _ => None,
        }
    }

    fn infer_vb_runtime_member_type_hint(&self, expr: &Expression) -> Option<String> {
        let path = Self::member_access_path(expr)?;
        match self.canon(&path).as_str() {
            "environment.currentdirectory"
            | "environment.newline"
            | "environment.machinename"
            | "environment.username"
            | "environment.osversion"
            | "system.environment.currentdirectory"
            | "system.environment.newline"
            | "system.environment.machinename"
            | "system.environment.username"
            | "system.environment.osversion"
            | "app.path"
            | "app.title" => Some("string".into()),
            "environment.processorcount"
            | "environment.tickcount"
            | "system.environment.processorcount"
            | "system.environment.tickcount"
            | "screen.width"
            | "screen.height" => Some("integer".into()),
            _ => None,
        }
    }

    fn infer_expr_type_hint(&self, expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(name) => self.lookup_var_type_hint(name).map(str::to_string),
            ExprKind::Lit(Literal::Int(_)) => Some("int".into()),
            ExprKind::Lit(Literal::Float(_)) => Some("double".into()),
            ExprKind::Lit(Literal::Str(_)) => Some("string".into()),
            ExprKind::Lit(Literal::Bool(_)) => Some("bool".into()),
            ExprKind::Lit(Literal::Char(_)) => Some("char".into()),
            ExprKind::Cast { type_name, .. } => Some(type_name.clone()),
            ExprKind::Unary { op: UnaryOp::Neg | UnaryOp::Pos, expr } => self.infer_expr_type_hint(expr),
            ExprKind::RefOf(place) => {
                let pointee_type = match place.as_ref() {
                    PlaceExpr::Ident(name) => self.lookup_var_type_hint(name).map(str::to_string),
                    PlaceExpr::Member { object, field, null_safe } => self.infer_expr_type_hint(&Expression::new(ExprKind::Member {
                        object: object.clone(),
                        field: field.clone(),
                        null_safe: *null_safe,
                    })),
                    PlaceExpr::Index { object, index, null_safe } => self.infer_expr_type_hint(&Expression::new(ExprKind::Index {
                        object: object.clone(),
                        index: index.clone(),
                        null_safe: *null_safe,
                    })),
                    PlaceExpr::Deref(expr) => self.infer_expr_type_hint(expr).map(|type_hint| {
                        type_hint
                            .trim()
                            .trim_end_matches('?')
                            .trim()
                            .trim_start_matches('*')
                            .trim_start_matches('^')
                            .trim()
                            .to_string()
                    }),
                }?;
                Some(format!("*{}", pointee_type.trim()))
            }
            ExprKind::Unary { op: UnaryOp::AddrOf, expr } => self
                .infer_expr_type_hint(expr)
                .map(|type_hint| format!("*{}", type_hint.trim().trim_end_matches('?').trim())),
            ExprKind::Unary { op: UnaryOp::Deref, expr } | ExprKind::RefLoad(expr) => self
                .infer_expr_type_hint(expr)
                .map(|type_hint| {
                    type_hint
                        .trim()
                        .trim_end_matches('?')
                        .trim()
                        .trim_start_matches('*')
                        .trim_start_matches('^')
                        .trim()
                        .to_string()
                }),
            ExprKind::New { class, .. } => Self::expr_terminal_type_name(class),
            ExprKind::Array(elements) => Some(format!(
                "{}()",
                self.infer_array_element_type_hint(elements.iter().map(|item| &item.value))
            )),
            ExprKind::Call { callee, args, .. } => {
                if matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("Array")) {
                    return Some(format!(
                        "{}()",
                        self.infer_array_element_type_hint(args.iter().map(|arg| &arg.value))
                    ));
                }
                if self.profile.parens_for_index
                    && args.len() == 1
                    && self
                        .infer_expr_type_hint(callee)
                        .as_deref()
                        .map(Self::normalize_type_hint)
                        .is_some_and(|type_hint| type_hint.ends_with("()") && !Self::is_callable_type_hint(&type_hint))
                {
                    return self
                        .infer_expr_type_hint(callee)
                        .and_then(|type_hint| type_hint.trim().trim_end_matches('?').trim().strip_suffix("()").map(str::to_string));
                }
                self.infer_function_return_type(callee)
                    .or_else(|| self.infer_dotnet_factory_return_type(callee))
            }
            ExprKind::Index { object, .. } => self
                .infer_expr_type_hint(object)
                .and_then(|type_hint| {
                    let trimmed = type_hint.trim().trim_end_matches('?').trim();
                    trimmed
                        .strip_suffix("()")
                        .map(str::to_string)
                        .or_else(|| Self::pascal_indexed_type_hint(trimmed))
                }),
            ExprKind::Member { object, field, .. } => {
                if let Some(type_hint) = self.infer_vb_runtime_member_type_hint(expr) {
                    return Some(type_hint);
                }
                if let Some(receiver_type) = self.infer_expr_type_hint(object) {
                    if let Some(class_name) = self.resolve_pending_class_name_for_type_hint(&receiver_type) {
                        if let Some(type_hint) = self
                            .pending_classes
                            .get(class_name.as_str())
                            .and_then(|pending| pending.instance_field_types.get(&self.canon(field)))
                        {
                            return Some(type_hint.clone());
                        }
                    }
                }
                let enum_type = Self::expr_terminal_type_name(object)?;
                self.enum_value_names
                    .contains_key(&self.canon(&enum_type))
                    .then_some(enum_type)
            }
            ExprKind::Binary { op, left, right } if matches!(op, BinOp::BitOr | BinOp::BitAnd | BinOp::BitXor) => {
                let left_type = self.infer_expr_type_hint(left)?;
                let right_type = self.infer_expr_type_hint(right)?;
                if left_type.eq_ignore_ascii_case(&right_type)
                    && self.enum_value_names.contains_key(&self.canon(&left_type))
                {
                    Some(left_type)
                } else {
                    None
                }
            }
            _ => None,
        }
    }

    fn user_value_type_name_from_hint(&self, type_hint: &str) -> Option<String> {
        let resolved = self.resolve_source_type_alias(type_hint);
        let trimmed = resolved.trim().trim_end_matches('?').trim();
        if trimmed.starts_with('*')
            || trimmed.starts_with('^')
            || trimmed.starts_with("[]")
            || trimmed.starts_with("map[")
            || trimmed.starts_with("chan ")
            || trimmed.starts_with("func(")
        {
            return None;
        }

        if let Some(class_name) = self.resolve_pending_class_name_for_type_hint(type_hint) {
            if self
                .pending_classes
                .get(&class_name)
                .is_some_and(|pending| pending.is_value_type)
            {
                return Some(class_name);
            }
        }

        for candidate in [
            Some(trimmed),
            trimmed.rsplit('.').next().filter(|segment| *segment != trimmed),
        ]
        .into_iter()
        .flatten()
        {
            let canonical = self.canon(candidate);
            if self.pending_classes.get(&canonical).is_some_and(|pending| pending.is_value_type) {
                return Some(canonical);
            }
            if let Some((name, _)) = self
                .pending_classes
                .iter()
                .find(|(name, pending)| pending.is_value_type && name.eq_ignore_ascii_case(candidate))
            {
                return Some(name.clone());
            }
        }
        None
    }

    fn expr_user_value_type_name(&self, expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Ident(name) => self
                .lookup_var_type_hint(name)
                .and_then(|type_hint| self.user_value_type_name_from_hint(type_hint)),
            _ => self
                .infer_expr_type_hint(expr)
                .and_then(|type_hint| self.user_value_type_name_from_hint(&type_hint)),
        }
    }

    fn expr_is_array_like(&self, expr: &Expression) -> bool {
        if self
            .infer_expr_type_hint(expr)
            .as_deref()
            .map(Self::normalize_type_hint)
            .is_some_and(|type_hint| type_hint.ends_with("()") && !Self::is_callable_type_hint(&type_hint))
        {
            return true;
        }

        match &expr.kind {
            ExprKind::Array(_) => true,
            ExprKind::Ident(name) => self.lookup_array_binding(name).is_some(),
            ExprKind::Index { object, index, .. } => {
                matches!(index.kind, ExprKind::Slice { .. }) && self.expr_is_array_like(object)
            }
            ExprKind::Call { callee, .. } => {
                matches!(&callee.kind, ExprKind::Ident(name)
                    if matches!(self.canon(name).as_str(), "array" | "str_split" | "str_getcsv"))
            }
            ExprKind::Binary { op, left, right }
                if matches!(op, BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Div | BinOp::Pow) =>
            {
                self.expr_is_array_like(left) || self.expr_is_array_like(right)
            }
            _ => false,
        }
    }

    fn vb_generic_type_display_name(&self, type_hint: &str) -> Option<String> {
        let trimmed = type_hint.trim().trim_end_matches('?').trim();
        let short_name = trimmed.rsplit('.').next().unwrap_or(trimmed).trim();

        let angle_arity = self.reflection_generic_argument_types(trimmed).len();
        if angle_arity > 0 {
            let base = short_name.split('<').next().unwrap_or(short_name).trim();
            return Some(format!("{base}`{angle_arity}"));
        }

        let lowered = trimmed.to_lowercase();
        let marker = "(of ";
        let start = lowered.find(marker)?;
        let base = trimmed[..start].trim().rsplit('.').next().unwrap_or(trimmed[..start].trim()).trim();
        let inner = trimmed.get(start + marker.len()..trimmed.len().saturating_sub(1))?;
        let mut depth = 0usize;
        let mut arity = 1usize;
        for ch in inner.chars() {
            match ch {
                '(' | '<' => depth += 1,
                ')' | '>' => depth = depth.saturating_sub(1),
                ',' if depth == 0 => arity += 1,
                _ => {}
            }
        }
        Some(format!("{base}`{arity}"))
    }

    fn vb_reflection_display_type_name(&self, type_name: &str) -> Option<String> {
        let trimmed = type_name.trim().trim_end_matches('?').trim();
        let short_target = trimmed.rsplit('.').next().unwrap_or(trimmed).trim();
        self.reflection_types
            .keys()
            .find(|candidate| {
                candidate.eq_ignore_ascii_case(trimmed)
                    || candidate
                        .rsplit('.')
                        .next()
                        .is_some_and(|leaf| leaf.eq_ignore_ascii_case(short_target))
            })
            .map(|candidate| self.reflection_type_short_name(candidate))
    }

    fn vb_typename_from_type_hint(&self, type_hint: &str) -> Option<String> {
        let resolved = self.resolve_source_type_alias(type_hint);
        let trimmed = resolved.trim().trim_end_matches('?').trim();

        if let Some(element_type) = trimmed.strip_suffix("()") {
            return self
                .vb_typename_from_type_hint(element_type.trim())
                .map(|name| format!("{name}()"));
        }

        let normalized = Self::normalize_type_hint(trimmed);
        let primitive = match normalized.as_str() {
            "integer" | "int" | "int32" | "longint" | "system.int32" => Some("Integer"),
            "long" | "int64" | "system.int64" => Some("Long"),
            "short" | "int16" | "system.int16" => Some("Short"),
            "ushort" | "uint16" | "system.uint16" => Some("UShort"),
            "uint" | "uint32" | "system.uint32" => Some("UInteger"),
            "ulong" | "uint64" | "system.uint64" => Some("ULong"),
            "byte" | "system.byte" => Some("Byte"),
            "sbyte" | "system.sbyte" => Some("SByte"),
            "single" | "float" | "system.single" => Some("Single"),
            "double" | "real" | "system.double" => Some("Double"),
            "decimal" | "system.decimal" => Some("Decimal"),
            "boolean" | "bool" | "system.boolean" => Some("Boolean"),
            "char" | "system.char" => Some("Char"),
            "string" | "system.string" => Some("String"),
            "datetime" | "date" | "system.datetime" => Some("Date"),
            "object" | "system.object" => Some("Object"),
            _ => None,
        };
        if let Some(name) = primitive {
            return Some(name.into());
        }

        if let Some(name) = self.vb_generic_type_display_name(trimmed) {
            return Some(name);
        }

        if let Some(name) = self.vb_reflection_display_type_name(trimmed) {
            return Some(name);
        }

        let short_target = trimmed.rsplit('.').next().unwrap_or(trimmed).trim();
        if let Some((display_name, _)) = self.pending_classes.iter().find(|(candidate, _)| {
            candidate.eq_ignore_ascii_case(trimmed)
                || candidate
                    .rsplit('.')
                    .next()
                    .is_some_and(|leaf| leaf.eq_ignore_ascii_case(short_target))
        }) {
            return Some(display_name.rsplit('.').next().unwrap_or(display_name).to_string());
        }

        if self.reflection_type_metadata(trimmed).is_some() || self.reflection_is_enum_type(trimmed) {
            return Some(self.reflection_type_short_name(trimmed));
        }
        None
    }

    fn vb_typename_from_expr(&self, expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::Lit(Literal::Int(_)) => Some("Integer".into()),
            ExprKind::Lit(Literal::Float(_)) => Some("Double".into()),
            ExprKind::Lit(Literal::Str(_)) => Some("String".into()),
            ExprKind::Lit(Literal::Bool(_)) => Some("Boolean".into()),
            ExprKind::Lit(Literal::Char(_)) => Some("Char".into()),
            ExprKind::Lit(Literal::Null | Literal::Undefined) => Some("Nothing".into()),
            _ => self
                .infer_expr_type_hint(expr)
                .and_then(|type_hint| self.vb_typename_from_type_hint(&type_hint)),
        }
    }

    fn vb_is_reference_type_hint(&self, type_hint: &str) -> bool {
        let resolved = self.resolve_source_type_alias(type_hint);
        let trimmed = resolved.trim().trim_end_matches('?').trim();
        if trimmed.ends_with("()") {
            return true;
        }
        if self.reflection_is_enum_type(trimmed) || self.reflection_is_value_type(trimmed) {
            return false;
        }
        match self.vb_typename_from_type_hint(trimmed).as_deref() {
            Some("Integer" | "Long" | "Short" | "UShort" | "UInteger" | "ULong" | "Byte" | "SByte"
                | "Single" | "Double" | "Decimal" | "Boolean" | "Char" | "Date") => false,
            Some("String" | "Object") => true,
            Some(name) if name.ends_with("()") => true,
            Some(_) => true,
            None => false,
        }
    }

    fn vb_is_object_type_hint(&self, type_hint: &str) -> bool {
        let resolved = self.resolve_source_type_alias(type_hint);
        let trimmed = resolved.trim().trim_end_matches('?').trim();
        if trimmed.ends_with("()") {
            return true;
        }
        if self.reflection_is_enum_type(trimmed) || self.reflection_is_value_type(trimmed) {
            return false;
        }
        match self.vb_typename_from_type_hint(trimmed).as_deref() {
            Some("Object") => true,
            Some("String" | "Integer" | "Long" | "Short" | "UShort" | "UInteger" | "ULong" | "Byte"
                | "SByte" | "Single" | "Double" | "Decimal" | "Boolean" | "Char" | "Date") => false,
            Some(name) if name.ends_with("()") => true,
            Some(_) => true,
            None => false,
        }
    }

    fn vb_is_reference_expr(&self, expr: &Expression) -> Option<bool> {
        match &expr.kind {
            ExprKind::Lit(Literal::Int(_))
            | ExprKind::Lit(Literal::Float(_))
            | ExprKind::Lit(Literal::Bool(_))
            | ExprKind::Lit(Literal::Char(_))
            | ExprKind::Lit(Literal::Null | Literal::Undefined) => Some(false),
            ExprKind::Lit(Literal::Str(_)) => Some(true),
            _ => self
                .infer_expr_type_hint(expr)
                .map(|type_hint| self.vb_is_reference_type_hint(&type_hint)),
        }
    }

    fn vb_is_object_expr(&self, expr: &Expression) -> Option<bool> {
        match &expr.kind {
            ExprKind::Lit(Literal::Int(_))
            | ExprKind::Lit(Literal::Float(_))
            | ExprKind::Lit(Literal::Bool(_))
            | ExprKind::Lit(Literal::Char(_))
            | ExprKind::Lit(Literal::Str(_))
            | ExprKind::Lit(Literal::Null | Literal::Undefined) => Some(false),
            _ => self
                .infer_expr_type_hint(expr)
                .map(|type_hint| self.vb_is_object_type_hint(&type_hint)),
        }
    }

    pub(super) fn compile_expr_with_value_copy(&mut self, expr: &Expression) -> Result<(), String> {
        self.compile_expr(expr)?;
        let should_clone = matches!(
            expr.kind,
            ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Index { .. }
        );
        if should_clone {
            if let Some(type_name) = self.expr_user_value_type_name(expr) {
                self.emit_user_value_type_clone_from_stack(&type_name);
            }
        }
        Ok(())
    }

    fn emit_array_clone_from_stack(&mut self) {
        let source_slot = self.define_local("__array_clone_src");
        let len_slot = self.define_local("__array_clone_len");

        self.emit_u16(Op::LOCAL_SET, source_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, source_slot);
        common::collections::emit_len(&mut self.chunks, self.current, self.line);
        self.emit_u16(Op::LOCAL_SET, len_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, source_slot);
        self.emit_const(Value::F64(0.0));
        self.emit_u16(Op::LOCAL_GET, len_slot);
        common::collections::emit_slice(&mut self.chunks, self.current, self.line);
    }

    fn emit_user_value_type_clone_from_stack(&mut self, type_name: &str) {
        let Some((fields, instance_member_names)) = self.pending_classes.get(type_name).map(|pending| {
            (pending.fields.clone(), pending.instance_member_names.clone())
        }) else {
            return;
        };

        let source_slot = self.define_local("__value_type_src");
        self.emit_u16(Op::LOCAL_SET, source_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, source_slot);
        self.emit(Op::REF_IS_NULL);
        let source_is_null = self.emit_jump(Op::BR_IF_TRUE);

        self.emit_u16(Op::LOCAL_GET, source_slot);
        self.emit(Op::REF_IS_UNDEFINED);
        let source_is_undefined = self.emit_jump(Op::BR_IF_TRUE);

        let clone_slot = self.define_local("__value_type_clone");
        let line = self.line;
        common::classes::emit_new_typed_object(self.chunk(), clone_slot, type_name, line);

        for member_name in fields.iter().chain(instance_member_names.iter()) {
            let member_key = self.str_const(member_name);
            self.emit_u16(Op::LOCAL_GET, clone_slot);
            self.emit_u16(Op::LOCAL_GET, source_slot);
            self.emit_u16(Op::STRUCT_GET, member_key);
            self.emit_u16(Op::STRUCT_SET, member_key);
            self.emit(Op::DROP);
        }

        self.emit_u16(Op::LOCAL_GET, clone_slot);
        let done = self.emit_jump(Op::BR);

        self.patch_jump(source_is_null);
        self.emit_u16(Op::LOCAL_GET, source_slot);
        let null_done = self.emit_jump(Op::BR);

        self.patch_jump(source_is_undefined);
        self.emit_u16(Op::LOCAL_GET, source_slot);

        self.patch_jump(done);
        self.patch_jump(null_done);
    }

    fn expr_is_known_string_receiver(&self, expr: &Expression) -> bool {
        match &expr.kind {
            ExprKind::Lit(Literal::Str(_)) | ExprKind::Interpolation(_) => true,
            ExprKind::Ident(name) => self
                .lookup_var_type_hint(name)
                .is_some_and(Self::is_string_type_hint),
            _ => false,
        }
    }

    fn maybe_promote_pascal_array_literal_to_set(
        &mut self,
        type_hint: Option<&str>,
        value: &Expression,
    ) {
        if self.profile.name != "pascal" {
            return;
        }
        if !type_hint.is_some_and(Self::is_pascal_set_type_hint) {
            return;
        }
        if !matches!(value.kind, ExprKind::Array(_)) {
            return;
        }
        let idx = self.import("ecma:set", "fromIterable");
        self.emit_host_call(idx, 1);
    }

    pub(super) fn expr_is_pascal_set(&self, expr: &Expression) -> bool {
        if self.profile.name != "pascal" {
            return false;
        }

        match &expr.kind {
            ExprKind::Set(_) => true,
            ExprKind::Ident(name) => self
                .lookup_var_type_hint(name)
                .is_some_and(Self::is_pascal_set_type_hint),
            ExprKind::Binary { op, left, right }
                if matches!(op, BinOp::Add | BinOp::Mul | BinOp::Sub) =>
            {
                self.expr_is_pascal_set(left) && self.expr_is_pascal_set(right)
            }
            _ => false,
        }
    }

    fn emit_var_get(&mut self, name: &str) {
        // Local
        if let Some(slot) = self.scope().resolve(name) {
            self.emit_u16(Op::LOCAL_GET, slot);
            if self.binding_uses_pointer_cell(name) {
                common::references::emit_cell_load(&mut self.chunks, self.current, self.line);
            }
            return;
        }
        if !self.case_sensitive {
            if let Some(slot) = self.scope().resolve_ci(name) {
                self.emit_u16(Op::LOCAL_GET, slot);
                if self.binding_uses_pointer_cell(name) {
                    common::references::emit_cell_load(&mut self.chunks, self.current, self.line);
                }
                return;
            }
        }
        // Upvalue (closure capture)
        if self.scopes.len() > 1 {
            if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, name) {
                self.emit_u8(Op::UPVALUE_GET, uv);
                return;
            }
        }
        if let Some(binding) = self.static_local_binding(name) {
            let global_name = binding.global_name.clone();
            let idx = self.global_name_const_idx(&global_name);
            self.emit_u16(Op::GLOBAL_GET, idx);
            return;
        }
        // Implicit self field — when inside a class method and the name is a
        // field of the current class, read from `me.<name>`. This is what
        // languages like VB do for unqualified field access. Without this,
        // dotted-name resolution that returns InstanceMember { local: "field" }
        // would fall through to global_get and read null.
        if self.current_class_implicit_self && self.is_class_field(name) {
            if self.emit_self_ref() {
                let cname = self.canon(name);
                let idx = self.str_const(&cname);
                self.emit_u16(Op::STRUCT_GET, idx);
                return;
            }
        }
        // Static field of the current class — `Count++` inside `Counter`
        // ctor reads `Counter.Count` (struct_get on the class global).
        // Without this, the bare name falls through to global_get and
        // returns null because the static field lives on the class
        // struct, not the module's global namespace.
        if let Some(class_name) = self.is_class_static_field(name) {
            let class_idx = self.global_name_const_idx(&class_name);
            self.emit_u16(Op::GLOBAL_GET, class_idx);
            let field_idx = self.str_const(&self.canon(name));
            self.emit_u16(Op::STRUCT_GET, field_idx);
            return;
        }
        // Bare static method in class scope — `Double(x)` inside
        // `class Converter` resolves to `Converter.Double`.
        if let Some(class_name) = self.is_class_static_method(name) {
            let class_idx = self.global_name_const_idx(&class_name);
            self.emit_u16(Op::GLOBAL_GET, class_idx);
            let method_idx = self.str_const(&self.canon(name));
            self.emit_u16(Op::STRUCT_GET, method_idx);
            return;
        }
        let cname = self.canon(name);
        let shadows_named_global = self.defined_globals.contains(&cname)
            || self.defined_functions.contains(&cname)
            || self.defined_classes.contains(&cname);
        if !shadows_named_global && self.emit_with_target_get(name) {
            return;
        }
        // Known type used as a value (e.g. `e instanceof RangeError`) — emit
        // the type name as a string so vybe:object:instanceOf can look it up
        // via its String fallback. Without this, `RangeError` would become
        // `global_get` of a nonexistent global → null.
        // Only do this when the name isn't shadowed by an actual global
        // (e.g. `Dim list As New List(Of String)` shadows the `list` type name).
        if self.profile.known_types.contains_key(name)
            && !self.defined_globals.contains(name)
            && !self.defined_globals.contains(&cname)
        {
            self.emit_const(Value::String(Arc::from(name)));
            return;
        }
        // Global — canonicalize name for case-insensitive languages
        let idx = self.global_name_const_idx(&cname);
        self.emit_u16(Op::GLOBAL_GET, idx);
        if self.binding_uses_pointer_cell(name) {
            common::references::emit_cell_load(&mut self.chunks, self.current, self.line);
        }
    }


    fn emit_ensure_global_map(&mut self, name: &str) {
        let key = self.shared_global_slot(name);
        self.emit_u16(Op::GLOBAL_GET, key);
        self.emit(Op::DUP);
        self.emit(Op::REF_IS_NULL);
        let has_map = self.emit_jump(Op::BR_IF_FALSE);

        self.emit(Op::DROP);
        common::collections::emit_map_new(&mut self.chunks, self.current, self.line);
        self.emit(Op::DUP);
        self.emit_u16(Op::GLOBAL_SET, key);
        self.emit(Op::DROP);

        self.patch_jump(has_map);
    }
    fn emit_var_set(&mut self, name: &str) {
        // Local
        if let Some(slot) = self.scope().resolve(name) {
            if self.binding_uses_pointer_cell(name) {
                let value_slot = self.define_local("__ref_cell_set_value");
                self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, slot);
                common::references::emit_cell_store(&mut self.chunks, self.current, value_slot, self.line);
                self.emit(Op::DROP);
            } else {
                self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
            }
            return;
        }
        if !self.case_sensitive {
            if let Some(slot) = self.scope().resolve_ci(name) {
                if self.binding_uses_pointer_cell(name) {
                    let value_slot = self.define_local("__ref_cell_set_value");
                    self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, slot);
                    common::references::emit_cell_store(&mut self.chunks, self.current, value_slot, self.line);
                    self.emit(Op::DROP);
                } else {
                    self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                }
                return;
            }
        }
        // Upvalue (closure capture)
        if self.scopes.len() > 1 {
            if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, name) {
                self.emit_u8(Op::UPVALUE_SET, uv);
                self.emit(Op::DROP);
                return;
            }
        }
        if let Some(binding) = self.static_local_binding(name) {
            let global_name = binding.global_name.clone();
            let idx = self.global_name_const_idx(&global_name);
            self.emit_u16(Op::GLOBAL_SET, idx);
            self.emit(Op::DROP);
            return;
        }
        if self.current_class_implicit_self && self.is_class_field(name) {
            let value_slot = self.define_local("__implicit_self_value");
            self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);
            if self.emit_self_ref() {
                self.emit_u16(Op::LOCAL_GET, value_slot);
                let cname = self.canon(name);
                let idx = self.str_const(&cname);
                self.emit_u16(Op::STRUCT_SET, idx);
                self.emit(Op::DROP);
                return;
            }
            self.emit_u16(Op::LOCAL_GET, value_slot);
        }
        // Static field of the current class — write through to
        // `<ClassName>.<name>` instead of falling to global_set.
        if let Some(class_name) = self.is_class_static_field(name) {
            // Stack: [value]. Need [class_obj, value] for STRUCT_SET.
            let value_slot = self.define_local("__static_set_value");
            self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);
            let class_idx = self.global_name_const_idx(&class_name);
            self.emit_u16(Op::GLOBAL_GET, class_idx);
            self.emit_u16(Op::LOCAL_GET, value_slot);
            let bare_name = self.canon(name);
            let field_idx = self.str_const(&bare_name);
            self.emit_u16(Op::STRUCT_SET, field_idx);
            self.emit(Op::DROP);
            if self.defined_globals.contains(&bare_name) {
                let global_idx = self.global_name_const_idx(&bare_name);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.emit_u16(Op::GLOBAL_SET, global_idx);
                self.emit(Op::DROP);
            }
            return;
        }
        let cname = self.canon(name);
        let shadows_named_global = self.defined_globals.contains(&cname)
            || self.defined_functions.contains(&cname)
            || self.defined_classes.contains(&cname);
        if !shadows_named_global && self.emit_with_target_set(name) {
            return;
        }
        // Global — canonicalize name for case-insensitive languages
        if self.scopes.len() == 1 {
            self.defined_globals.insert(cname.clone());
        }
        if self.binding_uses_pointer_cell(name) {
            let value_slot = self.define_local("__ref_global_set_value");
            self.emit_u16(Op::LOCAL_SET, value_slot);
            self.emit(Op::DROP);
            let idx = self.global_name_const_idx(&cname);
            self.emit_u16(Op::GLOBAL_GET, idx);
            common::references::emit_cell_store(&mut self.chunks, self.current, value_slot, self.line);
            self.emit(Op::DROP);
            return;
        }
        let idx = self.global_name_const_idx(&cname);
        self.emit_u16(Op::GLOBAL_SET, idx);
        self.emit(Op::DROP);
    }

    /// Walk up the scope chain to find a variable in a parent scope.
    /// Returns the upvalue index in the current scope if found.
    fn resolve_upvalue(&mut self, scope_idx: usize, name: &str) -> Option<u8> {
        if scope_idx == 0 { return None; }
        let parent = scope_idx - 1;
        // Check parent's locals
        let found_local = if self.case_sensitive {
            self.scopes[parent].resolve(name)
        } else {
            self.scopes[parent].resolve(name).or_else(|| self.scopes[parent].resolve_ci(name))
        };
        if let Some(slot) = found_local {
            self.scopes[parent].mark_captured(slot);
            return Some(self.scopes[scope_idx].add_upvalue(slot as u8, true));
        }
        // Recurse up
        if let Some(uv) = self.resolve_upvalue(parent, name) {
            return Some(self.scopes[scope_idx].add_upvalue(uv, false));
        }
        None
    }

    /// Returns the owning class name when `name` is a static field of
    /// the currently-compiling class (or one of its ancestors). Used by
    /// `emit_var_get` / `emit_var_set` to rewrite bare references to
    /// `ClassName.name` so static state lives on the class struct.
    fn is_class_static_field(&self, name: &str) -> Option<String> {
        if let Some(ref class_name) = self.current_class {
            let mut owner = Some(class_name.clone());
            while let Some(start) = owner {
                let mut current = Some(start.as_str());
                while let Some(cn) = current {
                    if let Some(pc) = self.pending_classes.get(cn) {
                        if pc.static_fields.iter().any(|f| {
                            if self.case_sensitive { f == name } else { f.eq_ignore_ascii_case(name) }
                        }) {
                            return Some(cn.to_string());
                        }
                        current = pc.parent.as_deref();
                    } else {
                        break;
                    }
                }
                owner = self.next_enclosing_class_name(&start);
            }
        }
        None
    }

    fn is_class_static_field_type_hint(&self, name: &str) -> Option<String> {
        if let Some(ref class_name) = self.current_class {
            let mut owner = Some(class_name.clone());
            while let Some(start) = owner {
                let mut current = Some(start.as_str());
                while let Some(cn) = current {
                    if let Some(pc) = self.pending_classes.get(cn) {
                        let canon = self.canon(name);
                        if let Some(type_hint) = pc.static_field_types.get(&canon) {
                            return Some(type_hint.clone());
                        }
                        current = pc.parent.as_deref();
                    } else {
                        break;
                    }
                }
                owner = self.next_enclosing_class_name(&start);
            }
        }
        None
    }

    fn is_class_nested_type(&self, name: &str) -> Option<String> {
        if let Some(ref class_name) = self.current_class {
            let mut owner = Some(class_name.clone());
            while let Some(start) = owner {
                let mut current = Some(start.as_str());
                while let Some(cn) = current {
                    if let Some(pc) = self.pending_classes.get(cn) {
                        if pc.nested_types.iter().any(|n| {
                            if self.case_sensitive { n == name } else { n.eq_ignore_ascii_case(name) }
                        }) {
                            return Some(cn.to_string());
                        }
                        current = pc.parent.as_deref();
                    } else {
                        break;
                    }
                }
                owner = self.next_enclosing_class_name(&start);
            }
        }
        None
    }

    pub(super) fn generic_static_member_key(&self, type_expr: &str, field: &str) -> Option<String> {
        let expr = type_expr.trim();
        if !expr.contains('<') || !expr.contains('>') {
            return None;
        }

        let base = expr.split('<').next().map(str::trim).unwrap_or(expr);
        let base_canon = self.canon(base);
        if !self.defined_classes.contains(&base_canon) {
            return None;
        }

        let field_canon = self.canon(field);
        let has_static = self.pending_classes.get(base)
            .or_else(|| self.pending_classes.get(base_canon.as_str()))
            .map(|pc| pc.static_fields.iter().any(|f| f == &field_canon))
            .unwrap_or(false);
        if !has_static {
            return None;
        }

        let compact_type: String = expr.chars().filter(|c| !c.is_whitespace()).collect();
        let type_canon = self.canon(&compact_type);
        Some(format!("__gstatic_{}_{}", type_canon, field_canon))
    }

    /// Returns the owning class when `name` is a static method of the
    /// currently compiling class (or one of its ancestors).
    fn is_class_static_method(&self, name: &str) -> Option<String> {
        if let Some(ref class_name) = self.current_class {
            let mut owner = Some(class_name.clone());
            while let Some(start) = owner {
                let mut current = Some(start.as_str());
                while let Some(cn) = current {
                    if let Some(pc) = self.pending_classes.get(cn) {
                        if pc.static_method_names.iter().any(|m| {
                            if self.case_sensitive { m == name } else { m.eq_ignore_ascii_case(name) }
                        }) {
                            return Some(cn.to_string());
                        }
                        current = pc.parent.as_deref();
                    } else {
                        break;
                    }
                }
                owner = self.next_enclosing_class_name(&start);
            }
        }
        None
    }

    fn next_enclosing_class_name(&self, class_name: &str) -> Option<String> {
        self.pending_classes
            .get(class_name)
            .and_then(|pc| pc.enclosing_class.clone())
            .or_else(|| class_name.rsplit_once('.').map(|(outer, _)| outer.to_string()))
    }

    /// Check if a name is a field of the current class (for implicit self resolution).
    fn is_class_field(&self, name: &str) -> bool {
        if !self.current_class_implicit_self { return false; }
        if let Some(ref class_name) = self.current_class {
            let mut current = Some(class_name.as_str());
            while let Some(cn) = current {
                if let Some(pc) = self.pending_classes.get(cn) {
                    if pc.fields.iter().any(|f| {
                        if self.case_sensitive { f == name } else { f.eq_ignore_ascii_case(name) }
                    }) {
                        return true;
                    }
                    current = pc.parent.as_deref();
                } else {
                    break;
                }
            }
        }
        false
    }

    fn emit_self_ref(&mut self) -> bool {
        let self_kw = self.profile.self_keyword.clone();
        if let Some(self_slot) = self.scope().resolve(&self_kw)
            .or_else(|| self.scope().resolve_ci(&self_kw))
        {
            self.emit_u16(Op::LOCAL_GET, self_slot);
            return true;
        }
        if self.scopes.len() > 1 {
            if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, &self_kw) {
                self.emit_u8(Op::UPVALUE_GET, uv);
                return true;
            }
        }
        false
    }

    fn is_js_profile(&self) -> bool {
        self.profile.name == "js"
    }

    fn is_php_profile(&self) -> bool {
        self.profile.name == "php"
    }

    fn is_python_profile(&self) -> bool {
        self.profile.name == "python"
    }

    fn emit_condition_truthiness_from_stack(&mut self) {
        if !self.is_python_profile() && !self.is_php_profile() {
            self.emit(Op::DYN_TO_BOOL);
            return;
        }

        if self.is_php_profile() {
            let value_slot = self.define_local("__php_truth_value");
            let keys_slot = self.define_local("__php_truth_keys");
            let tracker_slot = self.define_local("__php_truth_tracker");

            let typeof_idx = self.import("ecma:value", "typeof");
            let array_len_idx = self.import("ecma:array", "length");
            let has_own_idx = self.import("ecma:object", "hasOwn");
            let keys_key = self.str_const("__keys");
            let tracker_key = self.str_const("vybe$assoc_keys_csv");

            self.emit_u16(Op::LOCAL_SET, value_slot);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, value_slot);
            self.emit_host_call(typeof_idx, 1);
            self.emit_const(Value::String(Arc::from("object")));
            self.emit(Op::STR_EQUALS);
            let object_case = self.emit_jump(Op::BR_IF_TRUE);

            self.emit_u16(Op::LOCAL_GET, value_slot);
            self.emit(Op::DYN_TO_BOOL);
            let end = self.emit_jump(Op::BR);

            self.patch_jump(object_case);

            self.emit_u16(Op::LOCAL_GET, value_slot);
            self.emit_host_call(array_len_idx, 1);
            self.emit(Op::DYN_TO_BOOL);
            let non_empty_collection = self.emit_jump(Op::BR_IF_TRUE);

            self.emit_u16(Op::LOCAL_GET, value_slot);
            self.emit_u16(Op::STRUCT_GET, keys_key);
            self.emit_u16(Op::LOCAL_SET, keys_slot);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, keys_slot);
            self.emit(Op::DUP);
            self.emit(Op::REF_IS_NULL);
            let no_keys = self.emit_jump(Op::BR_IF_TRUE);
            self.emit(Op::REF_IS_UNDEFINED);
            let no_keys_undef = self.emit_jump(Op::BR_IF_TRUE);
            self.emit(Op::DROP);
            self.emit_u16(Op::LOCAL_GET, keys_slot);
            self.emit(Op::ARRAY_LENGTH);
            self.emit(Op::DYN_TO_BOOL);
            let non_empty_keys = self.emit_jump(Op::BR_IF_TRUE);
            let after_keys = self.emit_jump(Op::BR);

            self.patch_jump(no_keys);
            self.emit(Op::DROP);
            let after_keys_null = self.emit_jump(Op::BR);

            self.patch_jump(no_keys_undef);
            self.emit(Op::DROP);

            self.patch_jump(after_keys);
            self.patch_jump(after_keys_null);

            self.emit_u16(Op::LOCAL_GET, value_slot);
            self.emit_u16(Op::STRUCT_GET, tracker_key);
            self.emit_u16(Op::LOCAL_SET, tracker_slot);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, tracker_slot);
            self.emit(Op::DUP);
            self.emit(Op::REF_IS_NULL);
            let no_tracker = self.emit_jump(Op::BR_IF_TRUE);
            self.emit(Op::REF_IS_UNDEFINED);
            let no_tracker_undef = self.emit_jump(Op::BR_IF_TRUE);
            self.emit(Op::DROP);
            self.emit_u16(Op::LOCAL_GET, tracker_slot);
            self.emit(Op::DYN_TO_BOOL);
            let non_empty_tracker = self.emit_jump(Op::BR_IF_TRUE);
            let after_tracker = self.emit_jump(Op::BR);

            self.patch_jump(no_tracker);
            self.emit(Op::DROP);
            let after_tracker_null = self.emit_jump(Op::BR);

            self.patch_jump(no_tracker_undef);
            self.emit(Op::DROP);

            self.patch_jump(after_tracker);
            self.patch_jump(after_tracker_null);

            self.emit_u16(Op::LOCAL_GET, value_slot);
            self.emit_const(Value::String(Arc::from("__proto__")));
            self.emit_host_call(has_own_idx, 2);
            let has_proto = self.emit_jump(Op::BR_IF_TRUE);

            self.emit_const(Value::Bool(false));
            let object_end = self.emit_jump(Op::BR);

            self.patch_jump(non_empty_collection);
            self.emit_const(Value::Bool(true));
            let collection_end = self.emit_jump(Op::BR);

            self.patch_jump(non_empty_keys);
            self.emit_const(Value::Bool(true));
            let keys_end = self.emit_jump(Op::BR);

            self.patch_jump(non_empty_tracker);
            self.emit_const(Value::Bool(true));
            let tracker_end = self.emit_jump(Op::BR);

            self.patch_jump(has_proto);
            self.emit_const(Value::Bool(true));

            self.patch_jump(collection_end);
            self.patch_jump(keys_end);
            self.patch_jump(tracker_end);
            self.patch_jump(object_end);
            self.patch_jump(end);
            return;
        }

        let value_slot = self.define_local("__py_truth_value");
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.emit(Op::DROP);

        let typeof_idx = self.import("ecma:value", "typeof");
        let array_len_idx = self.import("ecma:array", "length");
        let has_own_idx = self.import("ecma:object", "hasOwn");

        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_host_call(typeof_idx, 1);
        self.emit_const(Value::String(Arc::from("object")));
        self.emit(Op::STR_EQUALS);
        let object_case = self.emit_jump(Op::BR_IF_TRUE);

        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit(Op::DYN_TO_BOOL);
        let end = self.emit_jump(Op::BR);

        self.patch_jump(object_case);

        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_host_call(array_len_idx, 1);
        self.emit(Op::DYN_TO_BOOL);
        let non_empty_collection = self.emit_jump(Op::BR_IF_TRUE);

        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_const(Value::String(Arc::from("__proto__")));
        self.emit_host_call(has_own_idx, 2);
        let has_proto = self.emit_jump(Op::BR_IF_TRUE);

        self.emit_const(Value::Bool(false));
        let object_end = self.emit_jump(Op::BR);

        self.patch_jump(non_empty_collection);
        self.emit_const(Value::Bool(true));
        let collection_end = self.emit_jump(Op::BR);

        self.patch_jump(has_proto);
        self.emit_const(Value::Bool(true));

        self.patch_jump(collection_end);
        self.patch_jump(object_end);
        self.patch_jump(end);
    }

    fn save_js_this(&mut self, local_name: &str) -> Option<u16> {
        if !self.is_js_profile() {
            return None;
        }
        let slot = self.scope().resolve(local_name)
            .unwrap_or_else(|| self.define_local(local_name));
        let idx = self.str_const("__js_this");
        self.emit_u16(Op::GLOBAL_GET, idx);
        self.emit_u16(Op::LOCAL_SET, slot);
        self.emit(Op::DROP);
        Some(slot)
    }

    fn set_js_this_from_stack(&mut self) {
        if !self.is_js_profile() {
            return;
        }
        let idx = self.str_const("__js_this");
        self.emit_u16(Op::GLOBAL_SET, idx);
        self.emit(Op::DROP);
    }

    fn restore_js_this(&mut self, slot: Option<u16>) {
        let Some(slot) = slot else { return; };
        let idx = self.str_const("__js_this");
        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit_u16(Op::GLOBAL_SET, idx);
        self.emit(Op::DROP);
    }

    fn flatten_member_chain(&self, expr: &Expression) -> Vec<String> {
        match &expr.kind {
            ExprKind::Ident(name) => Self::strip_global_namespace_prefix(name)
                .replace("::", ".")
                .split('.')
                .map(str::trim)
                .filter(|part| !part.is_empty())
                .map(ToString::to_string)
                .collect(),
            ExprKind::This => vec![self.profile.self_keyword.clone()],
            ExprKind::Super => vec![self.profile.base_keyword.clone().unwrap_or_else(|| "super".into())],
            ExprKind::Member { object, field, .. } => {
                let mut parts = self.flatten_member_chain(object);
                parts.push(field.clone());
                if parts.first().is_some_and(|part| part.eq_ignore_ascii_case("global")) {
                    parts.remove(0);
                }
                parts
            }
            _ => Vec::new(),
        }
    }

    /// Extract plain expressions from Argument slice.
    #[allow(dead_code)]
    fn arg_exprs(args: &[Argument]) -> Vec<&Expression> {
        args.iter().map(|a| &a.value).collect()
    }

    // ════════════════════════════════════════════════════════════════════════
    // Statement compilation
    // ════════════════════════════════════════════════════════════════════════

    fn compile_stmt(&mut self, stmt: &Statement) -> Result<(), String> {
        self.line = stmt.span.start_line;
        match &stmt.kind {
            // ── Expression statement ────────────────────────────────────
            StmtKind::Expr(expr) => {
                match &expr.kind {
                    ExprKind::Call { callee, args, .. }
                        if matches!(&callee.kind, ExprKind::Ident(name) if name == "__vb_lset_stmt") && args.len() == 2 => {
                        return self.compile_vb_fixed_string_stmt(&args[0].value, &args[1].value, false);
                    }
                    ExprKind::Call { callee, args, .. }
                        if matches!(&callee.kind, ExprKind::Ident(name) if name == "__vb_rset_stmt") && args.len() == 2 => {
                        return self.compile_vb_fixed_string_stmt(&args[0].value, &args[1].value, true);
                    }
                    ExprKind::Call { callee, args, .. }
                        if matches!(&callee.kind, ExprKind::Ident(name) if name == "__vb_mid_stmt") && args.len() == 4 => {
                        return self.compile_vb_mid_stmt(&args[0].value, &args[1].value, &args[2].value, &args[3].value);
                    }
                    ExprKind::Call { callee, args, .. }
                        if matches!(&callee.kind, ExprKind::Ident(name) if name == "__vb_err_raise") => {
                        return self.compile_vb_err_raise_stmt(args);
                    }
                    ExprKind::Ident(name)
                        if self.is_php_profile()
                            && (name.eq_ignore_ascii_case("exit")
                                || name.eq_ignore_ascii_case("die")) => {
                        self.emit(Op::NULL);
                        self.emit_return_through_finally(1)?;
                        return Ok(());
                    }
                    // Bare identifier that's a known function → call with 0 args
                    ExprKind::Ident(name) if self.defined_functions.contains(name.as_str()) => {
                        let saved_js_this = self.save_js_this("__js_stmt_prev_this");
                        if self.is_js_profile() {
                            let line = self.line;
                            common::expressions::emit_undefined(self.chunk(), line);
                            self.set_js_this_from_stack();
                        }
                        self.emit_var_get(name);
                        self.emit_u8(Op::CALL_REF, 0);
                        if saved_js_this.is_some() {
                            let result_slot = self.define_local("__js_stmt_result");
                            self.emit_u16(Op::LOCAL_SET, result_slot);
                            self.emit(Op::DROP);
                            self.restore_js_this(saved_js_this);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                        }
                        self.emit(Op::DROP);
                    }
                    // JS bare member statements evaluate the property access
                    // and discard the result; they are not implicit calls.
                    ExprKind::Member { object, field, .. } => {
                        if self.is_js_profile() {
                            self.compile_expr(expr)?;
                            self.emit(Op::DROP);
                            return Ok(());
                        }
                        self.compile_expr(object)?;
                        let field_name = self.canon(field);
                        let prop = self.str_const(&field_name);
                        self.emit(Op::DUP);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        let fn_tmp = self.define_local("__fn");
                        self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                        let obj_tmp = self.define_local("__obj");
                        self.reserve_local_slot(obj_tmp);
                        self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, fn_tmp);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.emit(Op::DROP);
                    }
                    _ => {
                        self.compile_expr(expr)?;
                        self.emit(Op::DROP);
                    }
                }
            }

            // ── Block ───────────────────────────────────────────────────
            StmtKind::Block(stmts) => {
                let all_decls = stmts.iter().all(|s| matches!(s.kind,
                    StmtKind::VarDecl { .. } | StmtKind::FunctionDecl { .. } |
                    StmtKind::ClassDecl { .. } | StmtKind::EnumDecl { .. }
                ));
                let hoisted_deconstruction = is_hoisted_deconstruction_block(stmts);
                if !all_decls && !hoisted_deconstruction { self.scope_mut().begin_scope(); }
                for s in stmts { self.compile_stmt(s)?; }
                if !all_decls && !hoisted_deconstruction { self.scope_mut().end_scope(); }
            }

            // ── Variable declarations ───────────────────────────────────
            StmtKind::VarDecl { declarations, kind } => {
                for decl in declarations {
                    self.compile_var_declarator(decl, kind)?;
                }
            }

            // ── Assignment ──────────────────────────────────────────────
            StmtKind::Assign { targets, value } => {
                if self.profile.name == "fortran" {
                    if let [target] = targets.as_slice() {
                        let is_whole_array_target = matches!(target.kind, ExprKind::Ident(_) | ExprKind::Member { .. })
                            && self.expr_is_array_like(target);
                        if is_whole_array_target {
                            let line = self.line;
                            let value_slot = self.define_local("__fortran_array_fill_value");
                            self.compile_expr(value)?;
                            self.emit_u16(Op::LOCAL_SET, value_slot);
                            self.emit(Op::DROP);

                            if !self.expr_is_array_like(value) {
                                self.emit_u16(Op::LOCAL_GET, value_slot);
                                self.emit(Op::REF_IS_ARRAY);
                                let value_is_runtime_array = self.emit_jump(Op::BR_IF_TRUE);

                                self.compile_expr(target)?;
                                let array_slot = self.define_local("__fortran_array_fill_target");
                                self.emit_u16(Op::LOCAL_SET, array_slot);
                                self.emit(Op::DROP);

                                self.emit_u16(Op::LOCAL_GET, array_slot);
                                self.emit_u16(Op::LOCAL_GET, value_slot);
                                self.emit_const(Value::I32(0));
                                self.emit_const(Value::I32(i32::MAX));
                                common::collections::emit_fill(&mut self.chunks, self.current, line);
                                self.compile_assign_target(target)?;
                                let done = self.emit_jump(Op::BR);

                                self.patch_jump(value_is_runtime_array);
                                self.emit_u16(Op::LOCAL_GET, value_slot);
                                self.emit_array_clone_from_stack();
                                self.compile_assign_target(target)?;
                                self.patch_jump(done);
                                return Ok(());
                            }

                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            self.emit_array_clone_from_stack();
                            self.compile_assign_target(target)?;
                            return Ok(());
                        }
                    }
                }
                if targets.len() == 1 {
                    if let ExprKind::Ident(name) = &targets[0].kind {
                        let binding_key = self.canon(name);
                        if let Some(binding) = self.resolve_reflection_binding_expr(value) {
                            self.reflection_bindings.insert(binding_key, binding);
                        } else {
                            self.reflection_bindings.remove(&binding_key);
                        }
                    }
                }
                if (self.profile.name == "csharp" || self.profile.name == "vb") && targets.len() == 1 {
                    if let ExprKind::Binary { op, left, right } = &value.kind {
                        if self.assign_target_matches_expr(&targets[0], left)
                            && self.is_csharp_delegate_handler_expr(right)
                        {
                            match op {
                                BinOp::Add => {
                                    self.compile_expr(left)?;
                                    self.compile_expr(right)?;
                                    common::delegates::emit_combine(&mut self.chunks, self.current, self.line);
                                    self.compile_assign_target(&targets[0])?;
                                    return Ok(());
                                }
                                BinOp::Sub => {
                                    self.compile_expr(left)?;
                                    self.compile_expr(right)?;
                                    common::delegates::emit_remove(&mut self.chunks, self.current, self.line);
                                    self.compile_assign_target(&targets[0])?;
                                    return Ok(());
                                }
                                _ => {}
                            }
                        }
                    }
                }
                // Multi-value receive: `a, b, c = callee(...)` where the
                // callee is a direct identifier call to a function the
                // pre-scan marked multi-return with matching arity. We
                // skip the heap-tuple alloc: compile the call, then let
                // each destructured element LOCAL_SET off the stack.
                if let Some((_arity, idents)) = self.detect_multi_value_receive(targets, value) {
                    // Compile the call inline so the Call-expression path
                    // in `expressions.rs` does NOT re-pack the results —
                    // we want the raw N values on the stack for direct
                    // destructuring.
                    self.compile_call_raw(value)?;
                    // Stack now holds [v0, v1, …, v(N-1)] with v(N-1) at
                    // TOS. Reverse assignment maps v_i to the i-th target.
                    // Inside a function, a fresh ident that doesn't already
                    // resolve should become a new local — C#'s
                    // `var (a, b) = f();` introduces new names, and this
                    // lets the walker emit a single Assign statement
                    // without juggling a Block + VarDecl pair.
                    let in_function = self.scopes.len() > 1;
                    for name in idents.iter().rev() {
                        if in_function
                            && self.scope().resolve(name).is_none()
                            && (self.case_sensitive || self.scope().resolve_ci(name).is_none())
                        {
                            self.define_local(name);
                        }
                        self.emit_var_set(name);
                    }
                } else {
                    let prefer_numeric_add = matches!(targets.as_slice(), [target] if self.expr_prefers_numeric_add(target));
                    self.compile_expr_with_numeric_add_hint(value, prefer_numeric_add)?;
                    if let [target] = targets.as_slice() {
                        self.emit_assignment_type_coercion_for_target(target);
                    }
                    if let [target] = targets.as_slice() {
                        if let ExprKind::Ident(name) = &target.kind {
                            if let Some(type_hint) = self.lookup_var_type_hint(name) {
                                if let Some(target_len) = Self::vb_fixed_string_len(type_hint) {
                                    self.emit_vb_fixed_string_adjust_from_stack(target_len, false);
                                }
                            }
                        }
                    }
                    if let [target] = targets.as_slice() {
                        if let ExprKind::Ident(name) = &target.kind {
                            let type_hint = self.lookup_var_type_hint(name).map(str::to_string);
                            self.maybe_promote_pascal_array_literal_to_set(type_hint.as_deref(), value);
                        }
                    }
                    for (i, target) in targets.iter().enumerate() {
                        if i < targets.len() - 1 { self.emit(Op::DUP); }
                        self.compile_assign_target(target)?;
                    }
                }
            }

            StmtKind::CompoundAssign { target, op, value } => {
                if (self.profile.name == "csharp" || self.profile.name == "vb")
                    && matches!(op, CompoundOp::Add | CompoundOp::Sub)
                    && self.is_csharp_delegate_handler_expr(value)
                {
                    match op {
                        CompoundOp::Add => {
                            self.compile_expr(target)?;
                            self.compile_expr(value)?;
                            common::delegates::emit_combine(&mut self.chunks, self.current, self.line);
                            self.compile_assign_target(target)?;
                            return Ok(());
                        }
                        CompoundOp::Sub => {
                            self.compile_expr(target)?;
                            self.compile_expr(value)?;
                            common::delegates::emit_remove(&mut self.chunks, self.current, self.line);
                            self.compile_assign_target(target)?;
                            return Ok(());
                        }
                        _ => {}
                    }
                }
                if matches!(op, CompoundOp::NullCoalesce) {
                    self.compile_expr(target)?;
                    let current_slot = self.define_local("__null_coalesce_current");
                    self.emit_u16(Op::LOCAL_SET, current_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, current_slot);
                    self.emit(Op::REF_IS_NULL);
                    let keep_current = self.emit_jump(Op::BR_IF_FALSE);
                    self.compile_expr(value)?;
                    let done = self.emit_jump(Op::BR);
                    self.patch_jump(keep_current);
                    self.emit_u16(Op::LOCAL_GET, current_slot);
                    self.patch_jump(done);
                    self.compile_assign_target(target)?;
                    return Ok(());
                }
                // Load current value
                self.compile_expr(target)?;
                let prefer_numeric_add = matches!(op, CompoundOp::Add) && self.expr_prefers_numeric_add(target);
                self.compile_expr_with_numeric_add_hint(value, prefer_numeric_add)?;
                if prefer_numeric_add {
                    self.emit(Op::F64_ADD);
                } else {
                    self.compile_compound_op(op);
                }
                self.compile_assign_target(target)?;
            }

            // ── If / Elif / Else (structured CF with label tracking) ──
            StmtKind::If { cond, then_body, elifs, else_body } => {
                let line = self.line;
                let outer = self.chunk().emit_block(line);
                self.label_depth += 1; // outer block

                // Then branch
                let then_block = self.chunk().emit_block(line);
                self.label_depth += 1;
                self.compile_expr(cond)?;
                self.emit_condition_truthiness_from_stack();
                self.emit(Op::DYN_NOT);
                let line = self.line;
                self.chunk().emit_br_if(0, line); // skip then if false
                self.scope_mut().begin_scope();
                for s in then_body { self.compile_stmt(s)?; }
                self.scope_mut().end_scope();
                if !elifs.is_empty() || else_body.is_some() {
                    let line = self.line;
                    self.chunk().emit_br(1, line); // to outer end
                }
                let line = self.line;
                self.chunk().emit_end(line);
                self.chunk().patch_block(then_block);
                self.label_depth -= 1;

                // Elif branches
                for (elif_cond, elif_body) in elifs {
                    let line = self.line;
                    let elif_block = self.chunk().emit_block(line);
                    self.label_depth += 1;
                    self.compile_expr(elif_cond)?;
                    self.emit_condition_truthiness_from_stack();
                    self.emit(Op::DYN_NOT);
                    let line = self.line;
                    self.chunk().emit_br_if(0, line);
                    self.scope_mut().begin_scope();
                    for s in elif_body { self.compile_stmt(s)?; }
                    self.scope_mut().end_scope();
                    let line = self.line;
                    self.chunk().emit_br(1, line); // to outer end
                    let line = self.line;
                    self.chunk().emit_end(line);
                    self.chunk().patch_block(elif_block);
                    self.label_depth -= 1;
                }

                if let Some(else_stmts) = else_body {
                    self.scope_mut().begin_scope();
                    for s in else_stmts { self.compile_stmt(s)?; }
                    self.scope_mut().end_scope();
                }

                let line = self.line;
                self.chunk().emit_end(line);
                self.chunk().patch_block(outer);
                self.label_depth -= 1;
            }

            // ── While (compiler_common::loops) ─────────────────────────
            StmtKind::While { cond, body, else_body } => {
                let line = self.line;
                let lp = common::loops::emit_loop_start(&mut self.chunks, self.current, line);
                // block + loop = 2 label stack entries
                let break_depth = self.label_depth + 1; // block is first (break target)
                let continue_depth = self.label_depth + 2; // loop is second (continue target)
                self.label_depth += 2;
                self.loop_states.push(lp);
                self.loops.push(LoopCtx { label: self.pending_label.take(), break_label_depth: break_depth, continue_label_depth: continue_depth, did_break_slot: None });
                self.compile_expr(cond)?;
                let line = self.line;
                common::loops::emit_loop_cond(&mut self.chunks, self.current, line);
                for s in body { self.compile_stmt(s)?; }
                self.loops.pop();
                let lp = self.loop_states.pop().unwrap();
                let line = self.line;
                common::loops::emit_loop_end(&mut self.chunks, self.current, lp, line);
                self.label_depth -= 2; // block + loop closed
                if let Some(else_stmts) = else_body {
                    for s in else_stmts { self.compile_stmt(s)?; }
                }
            }

            // ── For C-style (compiler_common::loops) ────────────────────
            StmtKind::For { init, cond, update, body } => {
                self.scope_mut().begin_scope();
                if let Some(init_stmt) = init { self.compile_stmt(init_stmt)?; }
                let loop_capture_name = if self.profile.name == "vb" {
                    init.as_ref().and_then(|stmt| match &stmt.kind {
                        StmtKind::VarDecl { declarations, .. } if declarations.len() == 1 => match &declarations[0].pattern {
                            BindingPattern::Ident(name) => Some(self.canon(name)),
                            _ => None,
                        },
                        _ => None,
                    })
                } else {
                    None
                };
                let line = self.line;
                // For C-style with update: use block { loop { cond, block $body { body }, update, br loop } }
                let block_patch = self.chunk().emit_block(line);
                self.label_depth += 1; // block
                let (loop_patch, _) = self.chunk().emit_loop_s(line);
                self.label_depth += 1; // loop
                let break_depth = self.label_depth - 1; // the block
                if let Some(c) = cond {
                    self.compile_expr(c)?;
                } else {
                    self.emit(Op::TRUE);
                }
                let line = self.line;
                common::loops::emit_loop_cond(&mut self.chunks, self.current, line);
                // Body block for continue-to-update
                let body_block = if update.is_some() {
                    let bp = self.chunk().emit_block(line);
                    self.label_depth += 1;
                    Some(bp)
                } else {
                    None
                };
                let continue_depth = self.label_depth; // innermost = continue target (body block or loop)
                let lp = common::loops::LoopState { block_patch, loop_patch, body_block_patch: body_block };
                self.loop_states.push(lp);
                self.loops.push(LoopCtx { label: self.pending_label.take(), break_label_depth: break_depth, continue_label_depth: continue_depth, did_break_slot: None });
                if let Some(loop_capture_name) = loop_capture_name.clone() {
                    self.capture_by_value_vars.push(loop_capture_name);
                }
                for s in body { self.compile_stmt(s)?; }
                if loop_capture_name.is_some() {
                    self.capture_by_value_vars.pop();
                }
                self.loops.pop();
                let lp = self.loop_states.pop().unwrap();
                // Close body block (continue lands here)
                if let Some(bp) = lp.body_block_patch {
                    self.chunk().emit_end(line);
                    self.chunk().patch_block(bp);
                    self.label_depth -= 1;
                }
                if let Some(u) = update { self.compile_expr(u)?; self.emit(Op::DROP); }
                let line = self.line;
                self.chunk().emit_br(0, line); // br loop
                self.chunk().emit_end(line); // end loop
                self.chunk().patch_loop(lp.loop_patch);
                self.label_depth -= 1;
                self.chunk().emit_end(line); // end block
                self.chunk().patch_block(lp.block_patch);
                self.label_depth -= 1;
                self.scope_mut().end_scope();
            }

            // ── ForIn / ForOf ───────────────────────────────────────────
            StmtKind::ForIn { var, key, iter, body, else_body, of, is_async, .. } => {
                // Specialisation: if `iter` is a direct call to a
                // function the pre-pass tagged as a true generator,
                // emit a `GEN_NEXT`-driven loop rather than the
                // array-index loop. This is the only path that makes
                // `for v in @generator_fn()` iterate lazily via the
                // WASM stack-switching coroutine machinery.
                if self.is_direct_generator_call(iter) {
                    self.compile_generator_for_in(var, key.as_deref(), iter, body, else_body.as_deref())?;
                } else {
                    let line = self.line;
                    self.compile_expr(iter)?;
                    let iter_slot = self.define_local("__forin_iter");
                    self.emit_u16(Op::LOCAL_SET, iter_slot); self.emit(Op::DROP);

                    let runtime_generator_done = if *of && key.is_none() {
                        // Large PHP foreach bodies can exceed the i16 reach of
                        // flat BR/BR_IF patching when we emit the runtime
                        // generator fast-path inline. Use structured label
                        // branches here so skipping the generator path does not
                        // depend on relative byte offsets.
                        let done_block = self.chunk().emit_block(line);
                        self.label_depth += 1;

                        let normal_path_gate = self.chunk().emit_block(line);
                        self.label_depth += 1;

                        self.emit_u16(Op::LOCAL_GET, iter_slot);
                        let is_gen_idx = self.import("ecma:value", "isGenerator");
                        self.emit_host_call(is_gen_idx, 1);
                        self.emit(Op::DYN_NOT);
                        self.emit_u8(Op::BR_IF_LABEL, 0);

                        self.compile_generator_for_in_cont(var, key.as_deref(), iter_slot, body, else_body.as_deref())?;
                        self.emit_u8(Op::BR_LABEL, 1);

                        self.chunk().emit_end(line);
                        self.chunk().patch_block(normal_path_gate);
                        self.label_depth -= 1;
                        Some(done_block)
                    } else {
                        None
                    };

                    // JS profile: route the iter through __vybe_iter_drain
                    // first. If the value has a user-defined `iterator()`
                    // (canonical name for `[Symbol.iterator]`), the
                    // polyfill calls it with `__js_this` correctly bound
                    // and returns the drained array. For built-ins
                    // (Array / Map / Set / String) it returns the input
                    // unchanged so iterForOf still produces the right
                    // shape. Only kicks in for `for ... of` (values
                    // path) — for-in over keys keeps standard semantics.
                    if self.is_js_profile() && *of && key.is_none() {
                        let drain_key = self.str_const("__vybe_iter_drain");
                        self.emit_u16(Op::GLOBAL_GET, drain_key);
                        self.emit_u16(Op::LOCAL_GET, iter_slot);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.emit_u16(Op::LOCAL_SET, iter_slot);
                        self.emit(Op::DROP);
                    }

                    self.emit_u16(Op::LOCAL_GET, iter_slot);

                    let iter_type_hint = match &iter.kind {
                        ExprKind::Ident(name) => self.lookup_var_type_hint(name).map(str::to_string),
                        _ => self.infer_expr_type_hint(iter),
                    };

                    let iterates_dictionary_entries = key.is_none() && *of && iter_type_hint
                        .as_deref()
                        .map(Self::is_dictionary_type_hint)
                        .unwrap_or(false);
                    let iterates_sorted_dictionary_entries = key.is_none() && *of && iter_type_hint
                        .as_deref()
                        .map(Self::is_sorted_dictionary_type_hint)
                        .unwrap_or(false);
                    let iterates_sorted_set_values = key.is_none() && *of && iter_type_hint
                        .as_deref()
                        .map(Self::is_sorted_set_type_hint)
                        .unwrap_or(false);

                    // Pick the polymorphic iteration primitive. All three
                    // dispatch on Array / Map / Ordinary uniformly so PHP
                    // assoc arrays, Python dicts, JS objects, Ruby hashes
                    // iterate correctly without per-language code.
                    //
                    //   for v in X       → values(X)        (Python for)
                    //   for k => v in X  → entries(X)       (PHP foreach, Ruby each_pair, JS for..of of Map/entries)
                    //   for k in X       → keys(X)          (JS for..in, Python dict iter-keys)
                    if key.is_some() || iterates_dictionary_entries {
                        common::collections::emit_iter_entries(&mut self.chunks, self.current, line);
                    } else if *of {
                        common::collections::emit_iter_values(&mut self.chunks, self.current, line);
                    } else {
                        common::collections::emit_iter_keys(&mut self.chunks, self.current, line);
                    }

                    if iterates_sorted_dictionary_entries {
                        self.emit_common("dotnet.sorted_dictionary_entries", 1, line);
                    } else if iterates_sorted_set_values {
                        common::collections::emit_sorted(&mut self.chunks, self.current, line);
                    }

                    let arr_slot = self.define_local("__forin_arr");
                    self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);
                    let idx_slot = self.define_local("__forin_idx");
                    // Allocate did_break slot BEFORE the for-in scaffolding
                    // so the assign-to-false initializer doesn't sit inside
                    // any of the for's blocks. Only when `else` is present
                    // — keeps the cost off the common case.
                    let did_break_slot = if else_body.is_some() {
                        let slot = self.define_local("__for_did_break");
                        self.emit(Op::FALSE);
                        self.emit_u16(Op::LOCAL_SET, slot);
                        self.emit(Op::DROP);
                        Some(slot)
                    } else { None };
                    let lp = common::loops::emit_for_in_start(
                        &mut self.chunks, self.current, arr_slot, idx_slot, line,
                    );
                    // for_in_start emits: block + loop + cond + block $body = 3 labels
                    let break_depth = self.label_depth + 1; // outer block
                    let continue_depth = self.label_depth + 3; // body block (innermost)
                    self.label_depth += 3;

                    if let Some(k_name) = key {
                        // Entries path: TOS is a [k, v] pair. Destructure
                        // into key_var and var, then run body.
                        //
                        // Stack at loop body entry: [pair]
                        //   DUP; index 0 → key_var
                        //   index 1 → value_var
                        let pair_slot = self.define_local("__forin_pair");
                        self.emit_u16(Op::LOCAL_SET, pair_slot); self.emit(Op::DROP);
                        // key = pair[0]
                        self.emit_u16(Op::LOCAL_GET, pair_slot);
                        self.emit_const(Value::I32(0));
                        common::collections::emit_get(&mut self.chunks, self.current, line);
                        let key_slot = self.define_local(k_name);
                        self.emit_u16(Op::LOCAL_SET, key_slot); self.emit(Op::DROP);

                        // var = pair[1]
                        self.emit_u16(Op::LOCAL_GET, pair_slot);
                        self.emit_const(Value::I32(1));
                        common::collections::emit_get(&mut self.chunks, self.current, line);
                        let var_slot = self.define_local(var);
                        self.emit_u16(Op::LOCAL_SET, var_slot); self.emit(Op::DROP);
                    } else if iterates_dictionary_entries {
                        let var_slot = self.define_local(var);
                        self.emit_u16(Op::LOCAL_SET, var_slot); self.emit(Op::DROP);
                    } else {
                        // Values path: TOS is the value, bind directly.
                        // `for await (let v of …)` per ECMA-262 §13.7.5
                        // performs `Await(value)` between iterator-step
                        // and binding. Emit the WASM JSPI suspend op so
                        // promise values unwrap before the body runs;
                        // non-promises pass through unchanged.
                        if *is_async {
                            self.emit(Op::PROMISE_SUSPEND);
                        }
                        let value_type_hint = iter_type_hint
                            .as_deref()
                            .and_then(|type_hint| type_hint.trim().trim_end_matches('?').trim().strip_suffix("()").map(str::to_string));
                        let var_slot = if let Some(type_hint) = value_type_hint {
                            self.define_local_typed(var, Some(type_hint))
                        } else {
                            self.define_local(var)
                        };
                        self.emit_u16(Op::LOCAL_SET, var_slot); self.emit(Op::DROP);
                    }

                    self.loop_states.push(lp);
                    self.loops.push(LoopCtx { label: self.pending_label.take(), break_label_depth: break_depth, continue_label_depth: continue_depth, did_break_slot });
                    for s in body { self.compile_stmt(s)?; }
                    self.loops.pop();
                    let lp = self.loop_states.pop().unwrap();
                    common::loops::emit_for_in_end(
                        &mut self.chunks, self.current, idx_slot, lp, line,
                    );
                    self.label_depth -= 3;
                    if let Some(else_stmts) = else_body {
                        // Python/Ruby for-else: skip else if any `break` fired.
                        // Wrap in `block { br_if 0 (if did_break); ...else... }`.
                        let dbs = did_break_slot.expect("did_break_slot allocated when else_body present");
                        let skip = self.chunk().emit_block(line);
                        self.label_depth += 1;
                        self.emit_u16(Op::LOCAL_GET, dbs);
                        self.emit(Op::DYN_TO_BOOL);
                        self.chunk().emit_br_if(0, line); // skip else if did_break
                        for s in else_stmts { self.compile_stmt(s)?; }
                        self.chunk().emit_end(line);
                        self.chunk().patch_block(skip);
                        self.label_depth -= 1;
                    }

                    if let Some(done_block) = runtime_generator_done {
                        self.chunk().emit_end(line);
                        self.chunk().patch_block(done_block);
                        self.label_depth -= 1;
                    }
                }
            }

            // ── DoWhile (compiler_common::loops) ────────────────────────
            StmtKind::DoWhile { body, cond, until } => {
                let line = self.line;
                let lp = common::loops::emit_do_loop_start(&mut self.chunks, self.current, line);
                let break_depth = self.label_depth + 1;
                let continue_depth = self.label_depth + 2;
                self.label_depth += 2;
                self.loop_states.push(lp);
                self.loops.push(LoopCtx { label: self.pending_label.take(), break_label_depth: break_depth, continue_label_depth: continue_depth, did_break_slot: None });
                for s in body { self.compile_stmt(s)?; }
                self.compile_expr(cond)?;
                self.loops.pop();
                let lp = self.loop_states.pop().unwrap();
                let line = self.line;
                common::loops::emit_do_loop_end(&mut self.chunks, self.current, lp, *until, line);
                self.label_depth -= 2;
            }

            // ── Switch / Select Case ────────────────────────────────────
            StmtKind::Switch { expr, cases, default } => {
                // Save switch expression to a local so checks can read it
                // without leaving it on the stack during body execution.
                self.compile_expr(expr)?;
                let sw_slot = self.define_local("__sw_expr");
                self.emit_u16(Op::LOCAL_SET, sw_slot); self.emit(Op::DROP);

                // Switch uses a BLOCK for break — push onto loop stack so break can find it
                let line = self.line;
                let switch_block = self.chunk().emit_block(line);
                self.label_depth += 1;
                let switch_lp = common::loops::LoopState { block_patch: switch_block, loop_patch: 0, body_block_patch: None };
                self.loop_states.push(switch_lp);
                self.loops.push(LoopCtx { label: self.pending_label.take(), break_label_depth: self.label_depth, continue_label_depth: self.label_depth, did_break_slot: None });

                // Merge legacy `default` field into the cases list.
                // New walkers emit default as a case with empty conditions
                // in source order. Old walkers may still use the separate
                // `default` field — append it at the end if present.
                let mut all_cases: Vec<&SwitchCase> = cases.iter().collect();
                let default_case_storage;
                if let Some(def) = default {
                    if !def.is_empty() && !cases.iter().any(|c| c.conditions.is_empty()) {
                        default_case_storage = SwitchCase { conditions: vec![], body: def.clone() };
                        all_cases.push(&default_case_storage);
                    }
                }

                // Phase 1: emit condition checks. Each matching condition
                // jumps to the corresponding body. A case with empty
                // conditions is the default — emit an unconditional jump.
                let mut body_jumps: Vec<Vec<usize>> = Vec::new();
                let _default_jump: Option<usize> = None;
                for case in all_cases.iter() {
                    if case.conditions.is_empty() {
                        // Default case — unconditional jump (deferred to after
                        // all condition checks so specific cases are tried first)
                        body_jumps.push(vec![]);
                        continue;
                    }
                    let mut match_patches = Vec::new();
                    for cond in &case.conditions {
                        match cond {
                            CaseCondition::Value(val) => {
                                self.emit_u16(Op::LOCAL_GET, sw_slot);
                                self.compile_expr(val)?;
                                self.emit(Op::DYN_EQ);
                                match_patches.push(self.emit_jump(Op::BR_IF_TRUE));
                            }
                            CaseCondition::Range { from, to } => {
                                self.emit_u16(Op::LOCAL_GET, sw_slot);
                                self.compile_expr(from)?;
                                self.emit(Op::DYN_GE);
                                let first = self.emit_jump(Op::BR_IF_FALSE);
                                self.emit_u16(Op::LOCAL_GET, sw_slot);
                                self.compile_expr(to)?;
                                self.emit(Op::DYN_LE);
                                match_patches.push(self.emit_jump(Op::BR_IF_TRUE));
                                self.patch_jump(first);
                            }
                            CaseCondition::Comparison { op, expr: cmp_expr } => {
                                self.emit_u16(Op::LOCAL_GET, sw_slot);
                                self.compile_expr(cmp_expr)?;
                                match op {
                                    ComparisonOp::Eq => self.emit(Op::DYN_EQ),
                                    ComparisonOp::NotEq => self.emit(Op::DYN_NE),
                                    ComparisonOp::Lt => self.emit(Op::DYN_LT),
                                    ComparisonOp::LtEq => self.emit(Op::DYN_LE),
                                    ComparisonOp::Gt => self.emit(Op::DYN_GT),
                                    ComparisonOp::GtEq => self.emit(Op::DYN_GE),
                                }
                                match_patches.push(self.emit_jump(Op::BR_IF_TRUE));
                            }
                        }
                    }
                    body_jumps.push(match_patches);
                }
                // No specific case matched → jump to default body (if any) or end
                let default_idx = all_cases.iter().position(|c| c.conditions.is_empty());
                let no_match = self.emit_jump(Op::BR);

                // Phase 2: emit all case bodies in order.
                let body_has_break = |body: &[Statement]| -> bool {
                    body.iter().any(|s| matches!(s.kind, StmtKind::Break(_)))
                };
                for (i, case) in all_cases.iter().enumerate() {
                    // Patch condition jumps to land at this body
                    for p in &body_jumps[i] { self.patch_jump(*p); }
                    // Default case: patch the no_match jump to land here
                    if case.conditions.is_empty() {
                        self.patch_jump(no_match);
                    }
                    for s in &case.body { self.compile_stmt(s)?; }
                    // Auto-break for non-fallthrough languages
                    if !self.profile.switch_fallthrough
                        && !body_has_break(&case.body)
                        && !case.body.is_empty()
                    {
                        // Break from switch → BR_LABEL to the switch block
                        if let Some(depth) = self.break_depth(None) {
                            let line = self.line;
                            self.chunk().emit_br(depth, line);
                        }
                    }
                }
                // If no default case, patch no_match to end
                if default_idx.is_none() {
                    self.patch_jump(no_match);
                }
                self.loops.pop();
                let switch_lp = self.loop_states.pop().unwrap();
                let line = self.line;
                self.chunk().emit_end(line);
                self.chunk().patch_block(switch_lp.block_patch);
                self.label_depth -= 1;
            }

            // ── Try / Catch / Finally ───────────────────────────────────
            StmtKind::Try { body, catches, else_body, finally } => {
                let line = self.line;
                let finally_exc_slot = if catches.is_empty() && finally.is_some() {
                    let slot = self.define_local("__try_finally_exc");
                    self.emit(Op::NULL);
                    self.emit_u16(Op::LOCAL_SET, slot);
                    self.emit(Op::DROP);
                    Some(slot)
                } else {
                    None
                };
                let catch_jump = common::errors::emit_try_start(&mut self.chunks[self.current], line);
                if let Some(fin) = finally.clone() {
                    self.active_finally_blocks.push(FinallyAction::Statements(fin));
                }
                for s in body { self.compile_stmt(s)?; }
                common::errors::emit_try_end(&mut self.chunks[self.current], line);
                // Python else: runs if no exception
                if let Some(else_stmts) = else_body {
                    for s in else_stmts { self.compile_stmt(s)?; }
                }
                let skip_to_finally = self.emit_jump(Op::BR);
                common::errors::patch_catch(&mut self.chunks[self.current], catch_jump);
                if catches.is_empty() {
                    if let Some(exc_slot) = finally_exc_slot {
                        self.emit_u16(Op::LOCAL_SET, exc_slot);
                        self.emit(Op::DROP);
                    } else {
                        self.emit(Op::DROP);
                    }
                } else {
                    // Multi-catch dispatch: each arm tests the exception's
                    // canonical __exception_type field. If it matches one of
                    // the arm's types, run the body; otherwise fall through
                    // to the next arm. The exception object is on TOS at
                    // every step. A catch-all arm (empty types or "Exception")
                    // catches everything. After all arms, any unmatched
                    // exception is re-thrown.
                    let mut end_patches: Vec<usize> = Vec::new();
                    for c in catches {
                        let types: Vec<&str> = c.types.iter()
                            .map(|t| common::errors::canonical_exception_name(t))
                            .collect();
                        let is_catch_all = types.is_empty()
                            || types.iter().any(|t| *t == "Exception");

                        let mut skip_arm_patches: Vec<usize> = Vec::new();
                        if !is_catch_all {
                            let mut to_body: Vec<usize> = Vec::new();
                            for ty in &types {
                                let mut expected_names = vec![(*ty).to_string()];
                                if !self.case_sensitive {
                                    let canon_ty = self.canon(ty);
                                    if canon_ty != *ty {
                                        expected_names.push(canon_ty);
                                    }
                                }

                                // Match if __exception_type === ty
                                for expected in &expected_names {
                                    self.emit(Op::DUP);
                                    let line = self.line;
                                    let key = self.str_const("__exception_type");
                                    self.chunks[self.current]
                                        .emit_op_u16(Op::STRUCT_GET, key, line);
                                    let v = self.str_const(expected);
                                    self.chunks[self.current]
                                        .emit_op_u16(Op::CONST, v, line);
                                    self.emit(Op::DYN_EQ);
                                    to_body.push(self.emit_jump(Op::BR_IF_TRUE));
                                }
                                // Or match if __type === ty (user class extends
                                // Exception — its ctor stamps __type via the
                                // class infrastructure but inherits
                                // __exception_type from the base ctor; checking
                                // both lets `catch (AppException)` find
                                // `throw new AppException(...)`).
                                for expected in &expected_names {
                                    self.emit(Op::DUP);
                                    let line = self.line;
                                    let key = self.str_const("__type");
                                    self.chunks[self.current]
                                        .emit_op_u16(Op::STRUCT_GET, key, line);
                                    let v = self.str_const(expected);
                                    self.chunks[self.current]
                                        .emit_op_u16(Op::CONST, v, line);
                                    self.emit(Op::DYN_EQ);
                                    to_body.push(self.emit_jump(Op::BR_IF_TRUE));
                                }
                                // Or match any name in the cross-language
                                // inheritance chain stamped by shared class
                                // emission. This lets `catch (BaseError)`
                                // match `throw new NotFoundError(...)`.
                                for expected in &expected_names {
                                    self.emit(Op::DUP);
                                    let line = self.line;
                                    let types_key = self.str_const("__types");
                                    self.chunks[self.current]
                                        .emit_op_u16(Op::STRUCT_GET, types_key, line);
                                    self.emit(Op::DUP);
                                    self.emit(Op::REF_IS_NULL);
                                    let has_types = self.emit_jump(Op::BR_IF_FALSE);
                                    self.emit(Op::DROP);
                                    self.emit(Op::FALSE);
                                    let done = self.emit_jump(Op::BR);
                                    self.patch_jump(has_types);
                                    let expected_const = Value::String(Arc::from(expected.as_str()));
                                    self.emit_const(expected_const);
                                    common::collections::emit_contains(&mut self.chunks, self.current, line);
                                    self.patch_jump(done);
                                    to_body.push(self.emit_jump(Op::BR_IF_TRUE));
                                }
                            }
                            skip_arm_patches.push(self.emit_jump(Op::BR));
                            for p in to_body { self.patch_jump(p); }
                        }

                        if let Some(ref var) = c.var_name {
                            self.scope_mut().begin_scope();
                            let slot = self.define_local(var);
                            self.emit(Op::DUP);
                            self.emit_u16(Op::LOCAL_SET, slot);
                            self.emit(Op::DROP);
                        } else {
                            self.scope_mut().begin_scope();
                        }

                        if let Some(cond) = &c.when_clause {
                            self.compile_expr(cond)?;
                            self.emit(Op::DYN_TO_BOOL);
                            skip_arm_patches.push(self.emit_jump(Op::BR_IF_FALSE));
                        }

                        self.emit(Op::DROP);
                        self.catch_depth += 1;
                        for s in &c.body { self.compile_stmt(s)?; }
                        self.catch_depth = self.catch_depth.saturating_sub(1);
                        self.scope_mut().end_scope();
                        end_patches.push(self.emit_jump(Op::BR));

                        for p in skip_arm_patches { self.patch_jump(p); }
                    }
                    // Fallthrough = no arm matched. Re-throw the exception.
                    let line = self.line;
                    common::errors::emit_throw(self.chunk(), line);
                    for p in end_patches { self.patch_jump(p); }
                }
                self.patch_jump(skip_to_finally);
                if finally.is_some() {
                    self.active_finally_blocks.pop();
                }
                if let Some(fin) = finally {
                    for s in fin { self.compile_stmt(s)?; }
                }
                if let Some(exc_slot) = finally_exc_slot {
                    if self.catch_depth > 0 {
                        return Ok(());
                    }
                    self.emit_u16(Op::LOCAL_GET, exc_slot);
                    self.emit(Op::REF_IS_NULL);
                    let done = self.emit_jump(Op::BR_IF_TRUE);
                    self.emit_u16(Op::LOCAL_GET, exc_slot);
                    let line = self.line;
                    common::errors::emit_throw(self.chunk(), line);
                    self.patch_jump(done);
                }
            }

            // ── Return ──────────────────────────────────────────────────
            StmtKind::Return(val) => {
                // Multi-value path: `return a, b, c` in a function the
                // pre-scan marked as multi-return. We push each element
                // separately (no heap tuple allocation) and let the VM's
                // `RETURN` pop N values per `chunk.result_arity`.
                let multi_n = self.current_multi_return_arity();
                if let (Some(n), Some(v)) = (multi_n, val) {
                    if let ExprKind::Tuple(elems) = &v.kind {
                        if elems.len() == n as usize {
                            for elem in elems { self.compile_expr(elem)?; }
                            self.emit_return_through_finally(n as usize)?;
                            return Ok(());
                        }
                    }
                }

                if let Some(v) = val {
                    self.compile_expr(v)?;
                } else if let Some(rs) = self.current_result_slot {
                    // ResultSlot return: return the result slot value
                    self.emit_u16(Op::LOCAL_GET, rs);
                } else if self.current_chunk_is_js_async() {
                    self.emit(Op::UNDEFINED);
                } else {
                    self.emit(Op::NULL);
                }

                if self.current_chunk_is_js_async() {
                    let resolve_idx = self.import("ecma:promise", "resolve");
                    self.emit_host_call(resolve_idx, 1);
                }
                self.emit_return_through_finally(1)?;
            }

            // ── Break ───────────────────────────────────────────────────
            StmtKind::Break(target) => {
                let line = self.line;
                match target {
                    // Exit Sub / Exit Function → RETURN (not a loop break)
                    BreakTarget::Kind(ExitKind::Sub) | BreakTarget::Kind(ExitKind::Function) => {
                        // Return with current result slot value, or null
                        if let Some(result_slot) = self.current_result_slot {
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                        } else {
                            self.emit(Op::NULL);
                        }
                        self.emit_return_through_finally(1)?;
                    }
                    BreakTarget::Implicit | BreakTarget::Kind(_) | BreakTarget::Level(_) => {
                        // If the targeted loop has a did_break slot (Python/
                        // Ruby for-else), record that break fired so the
                        // post-loop else clause is skipped.
                        if let Some(slot) = self.loops.last().and_then(|c| c.did_break_slot) {
                            self.emit(Op::TRUE);
                            self.emit_u16(Op::LOCAL_SET, slot);
                            self.emit(Op::DROP);
                        }
                        if let Some(depth) = self.break_depth(None) {
                            self.chunk().emit_br(depth, line);
                        }
                    }
                    BreakTarget::Label(label) => {
                        if let Some(slot) = self.loops.iter().rev()
                            .find(|c| c.label.as_deref() == Some(label))
                            .and_then(|c| c.did_break_slot)
                        {
                            self.emit(Op::TRUE);
                            self.emit_u16(Op::LOCAL_SET, slot);
                            self.emit(Op::DROP);
                        }
                        if let Some(depth) = self.break_depth(Some(label)) {
                            self.chunk().emit_br(depth, line);
                        }
                    }
                    BreakTarget::Value(expr) => {
                        self.compile_expr(expr)?;
                        self.emit_return_through_finally(1)?;
                    }
                }
            }

            // ── Continue ────────────────────────────────────────────────
            StmtKind::Continue(target) => {
                let line = self.line;
                match target {
                    ContinueTarget::Implicit | ContinueTarget::Kind(_) | ContinueTarget::Level(_) => {
                        if let Some(depth) = self.continue_depth(None) {
                            self.chunk().emit_br(depth, line);
                        }
                    }
                    ContinueTarget::Label(label) => {
                        if let Some(depth) = self.continue_depth(Some(&label)) {
                            self.chunk().emit_br(depth, line);
                        }
                    }
                }
            }

            // ── Throw ───────────────────────────────────────────────────
            StmtKind::Throw { expr, cause: _ } => {
                if let Some(v) = expr { self.compile_expr(v)?; } else { self.emit(Op::NULL); }
                let line = self.line;
                common::errors::emit_throw(&mut self.chunks[self.current], line);
            }

            // ── Function declaration ────────────────────────────────────
            StmtKind::FunctionDecl { name, params, return_type, body, modifiers: _, handles, is_async, is_generator, is_sub } => {
                self.compile_function_decl(name, params, return_type, body, *is_sub, *is_generator, handles, *is_async)?;
            }

            // ── Class declaration ───────────────────────────────────────
            StmtKind::ClassDecl { name, parents, interfaces, members, modifiers, .. } => {
                let cname = self.canon(name);
                self.defined_globals.insert(cname.clone());
                self.defined_classes.insert(cname.clone());
                let inferred_parents;
                let effective_parents: &[String] = if self.should_infer_winforms_form(name, parents) {
                    inferred_parents = vec!["Form".to_string()];
                    &inferred_parents
                } else {
                    parents
                };
                // Every language's profile has `uses_normalize_class = true`
                // after Phase 3. ClassDecl always goes through
                // walker → normalize_class → emit_class → compile_class.
                // If a new language is added that hasn't written its
                // normalizer yet, `emit_class_from_ast` returns an error
                // loudly rather than silently picking a legacy path.
                let span = stmt.span.clone();
                crate::common::classes::emit::emit_class_from_ast(
                    self,
                    span,
                    &cname,
                    effective_parents,
                    interfaces,
                    members,
                    modifiers,
                    self.profile.name == "fortran",
                )?;
            }

            // ── Interface declaration ───────────────────────────────────
            StmtKind::InterfaceDecl { .. } => {
                // No-op — interfaces are type-level only
            }

            // ── Enum declaration ────────────────────────────────────────
            // Compiles to a namespace object: Color = { Red: 0, Green: 1, Blue: 2 }
            // Bare member references (e.g. Pascal `c := Green`) are resolved at
            // compile time via the enum_members map.
            StmtKind::EnumDecl { name, members, is_flags, backing_type, interfaces, body_members, .. } => {
                let cname = self.canon(name);
                if *is_flags {
                    self.enum_flags.insert(cname.clone());
                } else {
                    self.enum_flags.remove(&cname);
                }

                match self.profile.name.as_str() {
                    "dart" => {
                        self.compile_dart_enum_decl(name, interfaces, body_members, members, stmt.span)?;
                        return Ok(());
                    }
                    "php" => {
                        self.compile_php_enum_decl(name, backing_type.as_ref(), interfaces, body_members, members, stmt.span)?;
                        return Ok(());
                    }
                    _ => {}
                }

                let mut next_val = 0i64;
                let mut value_names = HashMap::new();
                for m in members {
                    if let Some(ref v) = m.value {
                        if let ExprKind::Lit(Literal::Int(n)) = &v.kind {
                            next_val = *n;
                        }
                    }
                    next_val += 1;
                    let mname = self.canon(&m.name);
                    // Register member → enum type for bare-name resolution
                    self.enum_members.insert(mname, cname.clone());
                    value_names.insert(next_val - 1, m.name.clone());
                }
                self.enum_value_names.insert(cname.clone(), value_names);
                self.compile_shared_enum_decl(name, interfaces, body_members, members, stmt.span)?;
                self.defined_globals.insert(cname);
            }

            // ── Struct declaration (same as class) ──────────────────────
            // Structs compile through the same pipeline as classes: no
            // parent, no interfaces (struct `interfaces` list is ignored
            // by legacy compile_class anyway), same normalize → emit
            // path. Treated as a parent-less class by the walker's
            // normalize_class for the active language.
            StmtKind::StructDecl { name, members, .. } => {
                let cn = self.canon(name);
                self.defined_globals.insert(cn.clone());
                self.defined_classes.insert(cn.clone());
                let span = stmt.span.clone();
                crate::common::classes::emit::emit_class_from_ast(
                    self, span, &cn, &[], &[], members, &crate::ast::ClassModifiers::default(), true,
                )?;
            }

            // ── Module declaration (VB) ─────────────────────────────────
            // Models WASM Component Model: members are exports of the module.
            // - Members compile as globals (so call_ref works)
            // - Bare member names register in enum_members map → resolve to Module.Member
            // - A namespace struct is built so qualified `Module.Member` works too
            StmtKind::ModuleDecl { name, members, .. } => {
                let module_name = self.canon(name);
                self.defined_classes.insert(module_name.clone());
                self.register_module_static_container(&module_name, members);
                let mut member_names: Vec<String> = Vec::new();

                // First pass: compile all members as globals + collect names
                for m in members {
                    match m {
                        ClassMember::Method(stmt) => {
                            if let StmtKind::FunctionDecl { name: mname, .. } = &stmt.kind {
                                let mn = self.canon(mname);
                                let saved_class = self.current_class.clone();
                                let saved_implicit_self = self.current_class_implicit_self;
                                let saved_member_static = self.current_member_is_static;
                                self.current_class = Some(module_name.clone());
                                self.current_class_implicit_self = false;
                                self.current_member_is_static = true;
                                self.compile_stmt(stmt)?;
                                self.current_class = saved_class;
                                self.current_class_implicit_self = saved_implicit_self;
                                self.current_member_is_static = saved_member_static;
                                member_names.push(mn);
                            }
                        }
                        ClassMember::Field { name: fname, init, .. } => {
                            if let Some(init_expr) = init {
                                self.compile_expr(init_expr)?;
                            } else {
                                self.emit(Op::NULL);
                            }
                            let cname = self.canon(fname);
                            let idx = self.str_const(&cname);
                            self.emit_u16(Op::GLOBAL_SET, idx);
                            self.emit(Op::DROP);
                            self.defined_globals.insert(cname.clone());
                            member_names.push(cname);
                        }
                        ClassMember::Const { name: cname, value, .. } => {
                            // Compile value once, install as global
                            // `<Class>.<Const>` (legacy access path)
                            // AND stamp on the class object so PHP
                            // `Class::Const` static access (struct_get
                            // on class) resolves to the value.
                            self.compile_expr(value)?;
                            let val_slot = self.define_local("__class_const_val");
                            self.emit_u16(Op::LOCAL_SET, val_slot);
                            self.emit(Op::DROP);

                            let cn = self.canon(cname);
                            let idx = self.str_const(&cn);
                            self.emit_u16(Op::LOCAL_GET, val_slot);
                            self.emit_u16(Op::GLOBAL_SET, idx);
                            self.emit(Op::DROP);
                            self.defined_globals.insert(cn.clone());
                            member_names.push(cn.clone());

                            // Stamp on class object for static access.
                            // `name` here is the enclosing class name; on
                            // module-level Const blocks it's the module
                            // name, but the class object lookup will
                            // miss harmlessly in that case.
                            let class_canon = self.canon(name);
                            if self.defined_globals.contains(&class_canon) {
                                let cg_idx = self.str_const(&class_canon);
                                self.emit_u16(Op::GLOBAL_GET, cg_idx);
                                self.emit_u16(Op::LOCAL_GET, val_slot);
                                let field_idx = self.str_const(cname);
                                self.emit_u16(Op::STRUCT_SET, field_idx);
                                self.emit(Op::DROP);
                            }
                        }
                        ClassMember::NestedType(stmt) => {
                            // Nested types get their own globals; attach them to the
                            // module object so `Module.Type.Member` resolves through the
                            // same shared namespace path used by classes.
                            if let Some(cn) = match &stmt.kind {
                                StmtKind::ClassDecl { name: cname, .. }
                                | StmtKind::StructDecl { name: cname, .. }
                                | StmtKind::EnumDecl { name: cname, .. }
                                | StmtKind::InterfaceDecl { name: cname, .. }
                                | StmtKind::ModuleDecl { name: cname, .. } => Some(self.canon(cname)),
                                _ => None,
                            } {
                                member_names.push(cn);
                            }
                            self.compile_stmt(stmt)?;
                        }
                        ClassMember::Constructor { params, body, .. } => {
                            // Module-level constructor — compile as a function named after constructor_name
                            let ctor_stmt = Statement::new(StmtKind::FunctionDecl {
                                name: self.profile.constructor_name.clone(),
                                params: params.clone(),
                                return_type: None,
                                body: body.clone(),
                                modifiers: Modifiers::default(),
                                handles: Vec::new(),
                                is_async: false,
                                is_generator: false,
                                is_sub: true,
                            });
                            let saved_class = self.current_class.clone();
                            let saved_implicit_self = self.current_class_implicit_self;
                            let saved_member_static = self.current_member_is_static;
                            self.current_class = Some(module_name.clone());
                            self.current_class_implicit_self = false;
                            self.current_member_is_static = true;
                            self.compile_stmt(&ctor_stmt)?;
                            self.current_class = saved_class;
                            self.current_class_implicit_self = saved_implicit_self;
                            self.current_member_is_static = saved_member_static;
                            member_names.push(self.canon(&self.profile.constructor_name));
                        }
                        _ => {}
                    }
                }

                // Second pass: build namespace struct { member1: global, member2: global, ... }
                self.emit_u16(Op::STRUCT_NEW, 0);
                for mn in &member_names {
                    self.emit(Op::DUP);
                    let gidx = self.str_const(mn);
                    self.emit_u16(Op::GLOBAL_GET, gidx);
                    let key = self.str_const(mn);
                    self.emit_u16(Op::STRUCT_SET, key);
                    self.emit(Op::DROP);
                    // Register bare member → module name for qualified resolution
                    self.enum_members.insert(mn.clone(), module_name.clone());
                }
                let mod_idx = self.str_const(&module_name);
                self.emit_u16(Op::GLOBAL_SET, mod_idx);
                self.emit(Op::DROP);
                self.defined_globals.insert(module_name);
            }

            // ── Namespace declaration ───────────────────────────────────
            // C#/VB namespace: container of types. Compiles members as top-level globals
            // (matches .NET behavior — within the same compilation unit, bare type access
            // works without import). Also builds namespace struct for qualified access.
            StmtKind::NamespaceDecl { name, body } => {
                let local_ns_name = self.canon(name).replace('\\', ".");
                let ns_name = match self.current_namespace.as_deref() {
                    Some(prefix) if !prefix.is_empty() => format!("{prefix}.{local_ns_name}"),
                    _ => local_ns_name,
                };
                let mut member_names: Vec<(String, String, bool)> = Vec::new();
                let prev_namespace = self.current_namespace.clone();
                self.current_namespace = Some(ns_name.clone());
                for s in body {
                    // Track top-level type/function names declared in this namespace
                    match &s.kind {
                        StmtKind::ClassDecl { name: cn, .. }
                        | StmtKind::StructDecl { name: cn, .. }
                        | StmtKind::EnumDecl { name: cn, .. }
                        | StmtKind::InterfaceDecl { name: cn, .. }
                        | StmtKind::ModuleDecl { name: cn, .. } => {
                            let member_name = self.canon(cn);
                            member_names.push((
                                member_name.clone(),
                                format!("{ns_name}.{member_name}"),
                                true,
                            ));
                        }
                        StmtKind::FunctionDecl { name: cn, .. } => {
                            let member_name = self.canon(cn);
                            member_names.push((
                                member_name.clone(),
                                format!("{ns_name}.{member_name}"),
                                false,
                            ));
                        }
                        _ => {}
                    }
                    self.compile_stmt(s)?;
                }
                self.current_namespace = prev_namespace;

                for (member_name, qualified_name, is_type_like) in &member_names {
                    let source_idx = self.str_const(member_name);
                    let qualified_idx = self.str_const(qualified_name);
                    self.emit_u16(Op::GLOBAL_GET, source_idx);
                    self.emit_u16(Op::GLOBAL_SET, qualified_idx);
                    self.emit(Op::DROP);
                    self.defined_globals.insert(qualified_name.clone());
                    if *is_type_like {
                        self.defined_classes.insert(qualified_name.clone());
                    }
                }

                // Build namespace struct
                self.emit_u16(Op::STRUCT_NEW, 0);
                for (member_name, qualified_name, _) in &member_names {
                    self.emit(Op::DUP);
                    let gidx = self.str_const(qualified_name);
                    self.emit_u16(Op::GLOBAL_GET, gidx);
                    let key = self.str_const(member_name);
                    self.emit_u16(Op::STRUCT_SET, key);
                    self.emit(Op::DROP);
                }
                let ns_idx = self.str_const(&ns_name);
                self.emit_u16(Op::GLOBAL_SET, ns_idx);
                self.emit(Op::DROP);
                self.defined_globals.insert(ns_name.clone());

                let namespace_parts: Vec<&str> = ns_name.split('.').map(|part| part.trim()).filter(|part| !part.is_empty()).collect();
                if namespace_parts.len() > 1 {
                    for depth in 1..namespace_parts.len() {
                        let parent_name = self.canon(&namespace_parts[..depth].join("."));
                        let child_name = self.canon(&namespace_parts[..=depth].join("."));
                        let child_key = self.canon(namespace_parts[depth]);

                        if self.defined_globals.contains(&parent_name) {
                            let parent_idx = self.str_const(&parent_name);
                            self.emit_u16(Op::GLOBAL_GET, parent_idx);
                        } else {
                            self.emit_u16(Op::STRUCT_NEW, 0);
                        }
                        self.emit(Op::DUP);
                        let child_idx = self.str_const(&child_name);
                        self.emit_u16(Op::GLOBAL_GET, child_idx);
                        let key_idx = self.str_const(&child_key);
                        self.emit_u16(Op::STRUCT_SET, key_idx);
                        self.emit(Op::DROP);
                        let parent_idx = self.str_const(&parent_name);
                        self.emit_u16(Op::GLOBAL_SET, parent_idx);
                        self.emit(Op::DROP);
                        self.defined_globals.insert(parent_name);
                    }
                }
            }

            // ── Delegate declaration ────────────────────────────────────
            StmtKind::DelegateDecl { .. } => {
                // No-op — delegates are type-level
            }

            // ── With ────────────────────────────────────────────────────
            StmtKind::With { items, body, .. } => {
                self.scope_mut().begin_scope();
                if let Some(first) = items.first() {
                    self.compile_expr(&first.expr)?;
                    let slot = if let Some(ref var) = first.var {
                        self.define_local(var)
                    } else {
                        self.define_local("__with_target")
                    };
                    self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                    self.with_targets.push(slot);
                }
                for s in body { self.compile_stmt(s)?; }
                if !items.is_empty() {
                    self.with_targets.pop();
                }
                self.scope_mut().end_scope();
            }

            // ── Using ───────────────────────────────────────────────────
            // ECMA-334 §13.14: `using (var r = expr) { body; }` is
            // equivalent to:
            //
            //     var r = expr;
            //     try { body; } finally { r?.Dispose(); }
            //
            // Wrapping in real try/finally bytecode means an exception
            // escaping the body still triggers Dispose — matching the
            // C# semantic exercised by `using_disposes_on_exception`.
            // Cross-language: Python `with`, Java try-with-resources,
            // JS Stage 3 `using` share the same lowering.
            StmtKind::Using { var, resource, body } => {
                self.compile_expr(resource)?;
                let slot = self.define_local(var);
                self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);

                let line = self.line;
                let catch_jump = common::errors::emit_try_start(&mut self.chunks[self.current], line);
                self.active_finally_blocks.push(FinallyAction::ResourceDispose {
                    slot,
                    method: "Dispose".to_string(),
                    line,
                });
                for s in body { self.compile_stmt(s)?; }
                self.active_finally_blocks.pop();
                common::errors::emit_try_end(&mut self.chunks[self.current], line);
                let skip_to_finally = self.emit_jump(Op::BR);
                common::errors::patch_catch(&mut self.chunks[self.current], catch_jump);
                // Catch arm: dispose, then rethrow the exception
                // (which is on TOS after `patch_catch`).
                let exc_slot = self.define_local("__using_exc");
                self.emit_u16(Op::LOCAL_SET, exc_slot); self.emit(Op::DROP);
                self.label_depth += 1;
                common::errors::emit_resource_dispose(
                    self.chunk(), slot, "Dispose", line,
                );
                self.label_depth -= 1;
                self.emit_u16(Op::LOCAL_GET, exc_slot);
                common::errors::emit_throw(&mut self.chunks[self.current], line);
                // Normal-completion path: dispose, fall through.
                self.patch_jump(skip_to_finally);
                self.label_depth += 1;
                common::errors::emit_resource_dispose(
                    self.chunk(), slot, "Dispose", line,
                );
                self.label_depth -= 1;
            }

            // ── Lock ────────────────────────────────────────────────────
            StmtKind::Lock { body, .. } => {
                // No real locking in our VM — just compile body
                for s in body { self.compile_stmt(s)?; }
            }

            // ── ReDim ───────────────────────────────────────────────────
            // VB `ReDim arr(N)` allocates a fresh array of N+1 elements;
            // `ReDim Preserve arr(N)` allocates a new array AND copies the
            // old elements over (extending with defaults if growing). The
            // upper bound is inclusive (N → N+1 length).
            StmtKind::ReDim { array, bounds, preserve } => {
                if let Some(size_expr) = bounds.first() {
                    let line = self.line;
                    if *preserve {
                        // Allocate new array of N+1, then iterate the OLD
                        // array via compiler_common::loops::emit_for_in_start
                        // and copy each element into new[i] (bounded by
                        // new_len). This reuses the canonical for-in loop
                        // emit pattern that every other iteration site uses.
                        let old_slot = self.define_local("__redim_old");
                        let new_slot = self.define_local("__redim_new");
                        let new_len_slot = self.define_local("__redim_nlen");
                        let idx_slot = self.define_local("__redim_idx");
                        let old_len_slot = self.define_local("__redim_olen");
                        let fill_idx_slot = self.define_local("__redim_fill_idx");
                        let default_slot = self.define_local("__redim_default");

                        // old = arr
                        self.emit_var_get(array);
                        self.emit_u16(Op::LOCAL_SET, old_slot); self.emit(Op::DROP);
                        // new_len = N + 1
                        self.compile_expr(size_expr)?;
                        self.emit_const(Value::F64(1.0));
                        self.emit(Op::DYN_ADD);
                        self.emit_u16(Op::LOCAL_SET, new_len_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, old_slot);
                        common::collections::emit_len(&mut self.chunks, self.current, line);
                        self.emit_u16(Op::LOCAL_SET, old_len_slot); self.emit(Op::DROP);
                        // new = newWithLength(new_len) via common::collections
                        self.emit_u16(Op::LOCAL_GET, new_len_slot);
                        common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
                        self.emit_u16(Op::LOCAL_SET, new_slot); self.emit(Op::DROP);

                        // Iterate old array with the canonical for-in helper.
                        // The helper leaves [element] on the stack each pass
                        // and exposes the index in `idx_slot`.
                        let lp = common::loops::emit_for_in_start(
                            &mut self.chunks, self.current, old_slot, idx_slot, line);
                        // Stack: [element]. If idx >= new_len, drop and break
                        // (don't write past the new array). Otherwise
                        // new[idx] = element.
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_u16(Op::LOCAL_GET, new_len_slot);
                        self.emit(Op::DYN_LT);
                        let in_bounds = self.emit_jump(Op::BR_IF_TRUE);
                        // out of bounds: drop the element from for_in_start
                        self.emit(Op::DROP);
                        let after = self.emit_jump(Op::BR);
                        self.patch_jump(in_bounds);
                        // in bounds: new[idx] = element via common::collections::emit_set.
                        // Stack currently has [element].
                        let elem_slot = self.define_local("__redim_el");
                        self.emit_u16(Op::LOCAL_SET, elem_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, new_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        // emit_set preserves [val] — drop it.
                        self.emit(Op::DROP);
                        self.patch_jump(after);

                        common::loops::emit_for_in_end(
                            &mut self.chunks, self.current, idx_slot, lp, line);

                        // Fill any grown tail with the array's default value.
                        // Until arrays carry static element metadata, infer the
                        // default from the first existing element's runtime
                        // category: numbers -> 0, bools -> false, refs -> null.
                        self.emit(Op::NULL);
                        self.emit_u16(Op::LOCAL_SET, default_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, old_len_slot);
                        self.emit_const(Value::F64(0.0));
                        self.emit(Op::DYN_GT);
                        let no_default_seed = self.emit_jump(Op::BR_IF_FALSE);

                        self.emit_u16(Op::LOCAL_GET, old_slot);
                        self.emit(Op::I32_CONST_0);
                        common::collections::emit_get(&mut self.chunks, self.current, line);
                        let seed_slot = self.define_local("__redim_seed");
                        self.emit_u16(Op::LOCAL_SET, seed_slot); self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, seed_slot);
                        self.emit(Op::REF_IS_BOOL);
                        let bool_default = self.emit_jump(Op::BR_IF_TRUE);
                        self.emit_u16(Op::LOCAL_GET, seed_slot);
                        self.emit(Op::REF_IS_NUMBER);
                        let number_default = self.emit_jump(Op::BR_IF_TRUE);
                        self.emit(Op::NULL);
                        self.emit_u16(Op::LOCAL_SET, default_slot); self.emit(Op::DROP);
                        let default_done = self.emit_jump(Op::BR);

                        self.patch_jump(bool_default);
                        self.emit(Op::FALSE);
                        self.emit_u16(Op::LOCAL_SET, default_slot); self.emit(Op::DROP);
                        let bool_done = self.emit_jump(Op::BR);

                        self.patch_jump(number_default);
                        self.emit(Op::I32_CONST_0);
                        self.emit_u16(Op::LOCAL_SET, default_slot); self.emit(Op::DROP);

                        self.patch_jump(bool_done);
                        self.patch_jump(default_done);
                        self.patch_jump(no_default_seed);

                        self.emit_u16(Op::LOCAL_GET, old_len_slot);
                        self.emit_u16(Op::LOCAL_SET, fill_idx_slot); self.emit(Op::DROP);
                        let fill_block = self.chunk().emit_block(line);
                        let (fill_loop, _) = self.chunk().emit_loop_s(line);
                        self.emit_u16(Op::LOCAL_GET, fill_idx_slot);
                        self.emit_u16(Op::LOCAL_GET, new_len_slot);
                        self.emit(Op::DYN_LT);
                        self.emit(Op::DYN_NOT);
                        self.chunk().emit_br_if(1, line);

                        self.emit_u16(Op::LOCAL_GET, new_slot);
                        self.emit_u16(Op::LOCAL_GET, fill_idx_slot);
                        self.emit_u16(Op::LOCAL_GET, default_slot);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, fill_idx_slot);
                        self.emit_const(Value::F64(1.0));
                        self.emit(Op::DYN_ADD);
                        self.emit_u16(Op::LOCAL_SET, fill_idx_slot); self.emit(Op::DROP);
                        self.chunk().emit_br(0, line);
                        self.chunk().emit_end(line);
                        self.chunk().patch_loop(fill_loop);
                        self.chunk().emit_end(line);
                        self.chunk().patch_block(fill_block);

                        // arr = new
                        self.emit_u16(Op::LOCAL_GET, new_slot);
                        self.emit_var_set(array);
                    } else {
                        // ReDim arr(N) — non-preserving. N is the upper
                        // bound; length is N+1. Emit through
                        // `common::collections` (Phase D2).
                        let line = self.line;
                        self.compile_expr(size_expr)?;
                        self.emit_const(Value::F64(1.0));
                        self.emit(Op::DYN_ADD);
                        common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
                        self.emit_var_set(array);
                    }
                }
            }

            // ── Events ──────────────────────────────────────────────────
            // AddHandler/RemoveHandler are language-agnostic statements; the
            // canonical AST holds the control + handler as Expressions, so any
            // frontend (.NET, MAUI, Flutter, …) can produce the same node by
            // mapping its surface syntax (`Handles X.Y`, `obj.Y += h`, etc.).
            //
            // The handler is registered under the SOURCE-CODE-STABLE control
            // identifier (field name, class name for `Me`/`This`, or runtime

            StmtKind::Erase { array } => {
                let line = self.line;
                let Some(binding) = self.lookup_array_binding(array).cloned() else {
                    self.emit(Op::NULL);
                    self.emit_var_set(array);
                    self.emit(Op::DROP);
                    return Ok(());
                };

                if !binding.is_fixed {
                    self.emit(Op::NULL);
                    self.emit_var_set(array);
                    self.emit(Op::DROP);
                    return Ok(());
                }

                let old_slot = self.define_local("__erase_old");
                let len_slot = self.define_local("__erase_len");
                let new_slot = self.define_local("__erase_new");

                self.emit_var_get(array);
                self.emit_u16(Op::LOCAL_SET, old_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, old_slot);
                common::collections::emit_len(&mut self.chunks, self.current, line);
                self.emit_u16(Op::LOCAL_SET, len_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, len_slot);
                common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
                self.emit_u16(Op::LOCAL_SET, new_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, new_slot);
                self.emit_default_value_for_type_hint(
                    binding
                        .type_hint
                        .as_deref()
                        .map(|type_hint| type_hint.trim().trim_end_matches("()").trim()),
                );
                self.emit_const(Value::I32(0));
                self.emit_const(Value::I32(i32::MAX));
                common::collections::emit_fill(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, new_slot);
                self.emit_var_set(array);
                self.emit(Op::DROP);
            }
            // `__control_name` for general expressions). This decouples the
            // registry key from the runtime `.Name` property — renaming a
            // control via `btn.Name = "x"` doesn't break wired-up handlers.
            StmtKind::AddHandler { control, event, handler } => {
                self.compile_add_handler_stmt(control, event, handler)?;
            }

            StmtKind::RemoveHandler { control, event, handler } => {
                self.compile_remove_handler_stmt(control, event, handler)?;
            }

            StmtKind::RaiseEvent { event_name, args } => {
                self.compile_raise_event_stmt(event_name, args)?;
            }

            // ── VB legacy error handling ────────────────────────────────
            StmtKind::OnErrorResumeNext => { /* no-op in bytecode VM */ }
            StmtKind::OnErrorGoTo(_) => { /* no-op */ }
            StmtKind::GoTo(_) => { /* no-op — structured bytecode doesn't support arbitrary gotos */ }
            StmtKind::Label(_) => { /* no-op */ }

            // ── VB legacy file I/O ──────────────────────────────────────
            StmtKind::OpenFile { path, mode, file_number } => {
                let path_slot = self.define_local("__vb_open_path");
                let file_slot = self.define_local("__vb_open_file_number");
                let path_map_key = self.shared_global_slot("__vb_file_path_by_handle");
                let eof_map_key = self.shared_global_slot("__vb_file_eof_by_handle");
                let mode_text = match mode {
                    FileMode::Input => "Input",
                    FileMode::Output => "Output",
                    FileMode::Append => "Append",
                    FileMode::Binary => "Binary",
                    FileMode::Random => "Random",
                };

                self.compile_expr(path)?;
                self.emit_u16(Op::LOCAL_SET, path_slot);
                self.emit(Op::DROP);

                self.compile_expr(file_number)?;
                self.emit_u16(Op::LOCAL_SET, file_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, path_slot);
                self.emit_const(Value::String(Arc::from(mode_text)));
                self.emit_u16(Op::LOCAL_GET, file_slot);
                let idx = self.import("wasi:filesystem", "openFile");
                self.emit_host_call(idx, 3);
                self.emit(Op::DROP);

                self.emit_ensure_global_map("__vb_file_path_by_handle");
                self.emit_u16(Op::GLOBAL_GET, path_map_key);
                self.emit_u16(Op::LOCAL_GET, file_slot);
                self.emit_u16(Op::LOCAL_GET, path_slot);
                self.emit(Op::ARRAY_SET);
                self.emit(Op::DROP);

                self.emit_ensure_global_map("__vb_file_eof_by_handle");
                self.emit_u16(Op::GLOBAL_GET, eof_map_key);
                self.emit_u16(Op::LOCAL_GET, file_slot);
                self.emit_const(Value::Bool(false));
                self.emit(Op::ARRAY_SET);
                self.emit(Op::DROP);

                self.emit_global_map_set_null("__vb_record_rows_by_handle", file_slot);
                self.emit_global_map_set_null("__vb_record_next_index_by_handle", file_slot);
                self.emit_global_map_set_null("__vb_record_current_index_by_handle", file_slot);
            }
            StmtKind::CloseFile(file_num) => {
                let path_map_key = self.shared_global_slot("__vb_file_path_by_handle");
                let eof_map_key = self.shared_global_slot("__vb_file_eof_by_handle");
                if let Some(fnum) = file_num {
                    let file_slot = self.define_local("__vb_close_file_number");
                    self.compile_expr(fnum)?;
                    self.emit_u16(Op::LOCAL_SET, file_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, file_slot);
                    let idx = self.import("wasi:filesystem", "closeFile");
                    self.emit_host_call(idx, 1);
                    self.emit(Op::DROP);

                    self.emit_ensure_global_map("__vb_file_path_by_handle");
                    self.emit_u16(Op::GLOBAL_GET, path_map_key);
                    self.emit_u16(Op::LOCAL_GET, file_slot);
                    self.emit(Op::NULL);
                    self.emit(Op::ARRAY_SET);
                    self.emit(Op::DROP);

                    self.emit_ensure_global_map("__vb_file_eof_by_handle");
                    self.emit_u16(Op::GLOBAL_GET, eof_map_key);
                    self.emit_u16(Op::LOCAL_GET, file_slot);
                    self.emit_const(Value::Bool(false));
                    self.emit(Op::ARRAY_SET);
                    self.emit(Op::DROP);

                    self.emit_global_map_set_null("__vb_record_rows_by_handle", file_slot);
                    self.emit_global_map_set_null("__vb_record_next_index_by_handle", file_slot);
                    self.emit_global_map_set_null("__vb_record_current_index_by_handle", file_slot);
                } else {
                    self.emit(Op::NULL);
                    let idx = self.import("wasi:filesystem", "closeFile");
                    self.emit_host_call(idx, 1);
                    self.emit(Op::DROP);
                }
            }
            StmtKind::PrintFile { file_number, items } => {
                self.compile_expr(file_number)?;
                for item in items { self.compile_expr(item)?; }
                let idx = self.import("wasi:filesystem", "printFile");
                self.emit_host_call(idx, (items.len() + 1) as u8);
                self.emit(Op::DROP);
            }
            StmtKind::WriteFile { file_number, items } => {
                self.compile_expr(file_number)?;
                for item in items { self.compile_expr(item)?; }
                let idx = self.import("wasi:filesystem", "writeFile_handle");
                self.emit_host_call(idx, (items.len() + 1) as u8);
                self.emit(Op::DROP);
            }
            StmtKind::InputFile { file_number, variables } => {
                let file_slot = self.define_local("__vb_input_file_number");
                let values_slot = self.define_local("__vb_input_values");
                let rows_slot = self.define_local("__vb_input_rows");
                let len_slot = self.define_local("__vb_input_len");
                let idx_slot = self.define_local("__vb_input_idx");
                let eof_map_key = self.shared_global_slot("__vb_file_eof_by_handle");

                self.compile_expr(file_number)?;
                self.emit_u16(Op::LOCAL_SET, file_slot);
                self.emit(Op::DROP);

                self.emit_record_rows_cache(file_slot, rows_slot, len_slot);
                self.emit_global_map_get_into_local("__vb_record_next_index_by_handle", file_slot, idx_slot);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit(Op::REF_IS_NULL);
                let have_idx = self.emit_jump(Op::BR_IF_FALSE);
                self.emit(Op::I32_CONST_0);
                self.emit_u16(Op::LOCAL_SET, idx_slot);
                self.emit(Op::DROP);
                self.patch_jump(have_idx);

                self.emit_u16(Op::LOCAL_GET, file_slot);
                let idx = self.import("wasi:filesystem", "inputFile");
                self.emit_host_call(idx, 1);
                self.emit_u16(Op::LOCAL_SET, values_slot);
                self.emit(Op::DROP);

                for (index, variable) in variables.iter().enumerate() {
                    self.emit_u16(Op::LOCAL_GET, values_slot);
                    self.emit_const(Value::F64(index as f64));
                    self.emit(Op::ARRAY_GET);
                    self.emit_assignment_type_coercion_for_target(variable);
                    self.compile_assign_target(variable)?;
                }

                if variables.is_empty() {
                    self.emit_u16(Op::LOCAL_GET, values_slot);
                    self.emit(Op::DROP);
                }

                self.emit_ensure_global_map("__vb_file_eof_by_handle");
                self.emit_u16(Op::GLOBAL_GET, eof_map_key);
                self.emit_u16(Op::LOCAL_GET, file_slot);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit(Op::I32_CONST_1);
                self.emit(Op::I32_ADD);
                self.emit_u16(Op::LOCAL_SET, idx_slot);
                self.emit(Op::DROP);

                self.emit_global_map_set_from_local("__vb_record_next_index_by_handle", file_slot, idx_slot);

                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_u16(Op::LOCAL_GET, len_slot);
                self.emit(Op::DYN_GE);
                self.emit(Op::ARRAY_SET);
                self.emit(Op::DROP);
            }
            StmtKind::LineInput { file_number, variable } => {
                let file_slot = self.define_local("__vb_line_input_file_number");
                let rows_slot = self.define_local("__vb_line_input_rows");
                let len_slot = self.define_local("__vb_line_input_len");
                let idx_slot = self.define_local("__vb_line_input_idx");
                let eof_map_key = self.shared_global_slot("__vb_file_eof_by_handle");

                self.compile_expr(file_number)?;
                self.emit_u16(Op::LOCAL_SET, file_slot);
                self.emit(Op::DROP);

                self.emit_record_rows_cache(file_slot, rows_slot, len_slot);
                self.emit_global_map_get_into_local("__vb_record_next_index_by_handle", file_slot, idx_slot);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit(Op::REF_IS_NULL);
                let have_idx = self.emit_jump(Op::BR_IF_FALSE);
                self.emit(Op::I32_CONST_0);
                self.emit_u16(Op::LOCAL_SET, idx_slot);
                self.emit(Op::DROP);
                self.patch_jump(have_idx);

                self.emit_u16(Op::LOCAL_GET, file_slot);
                let idx = self.import("wasi:filesystem", "lineInput");
                self.emit_host_call(idx, 1);
                self.emit_var_set(variable);

                self.emit_ensure_global_map("__vb_file_eof_by_handle");
                self.emit_u16(Op::GLOBAL_GET, eof_map_key);
                self.emit_u16(Op::LOCAL_GET, file_slot);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit(Op::I32_CONST_1);
                self.emit(Op::I32_ADD);
                self.emit_u16(Op::LOCAL_SET, idx_slot);
                self.emit(Op::DROP);

                self.emit_global_map_set_from_local("__vb_record_next_index_by_handle", file_slot, idx_slot);

                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_u16(Op::LOCAL_GET, len_slot);
                self.emit(Op::DYN_GE);
                self.emit(Op::ARRAY_SET);
                self.emit(Op::DROP);
            }
            StmtKind::StartFile { file_number, key_index, key_value, relation } => {
                let line = self.line;
                let file_slot = self.define_local("__vb_start_file_number");
                let rows_slot = self.define_local("__vb_start_rows");
                let len_slot = self.define_local("__vb_start_len");
                let key_slot = self.define_local("__vb_start_key");
                let found_slot = self.define_local("__vb_start_found");
                let idx_slot = self.define_local("__vb_start_idx");
                let row_slot = self.define_local("__vb_start_row");
                let values_slot = self.define_local("__vb_start_values");

                self.compile_expr(file_number)?;
                self.emit_u16(Op::LOCAL_SET, file_slot);
                self.emit(Op::DROP);
                self.emit_record_rows_cache(file_slot, rows_slot, len_slot);

                self.compile_expr(key_value)?;
                self.emit_u16(Op::LOCAL_SET, key_slot);
                self.emit(Op::DROP);
                self.emit(Op::NULL);
                self.emit_u16(Op::LOCAL_SET, found_slot);
                self.emit(Op::DROP);

                let state = common::loops::emit_for_in_start(&mut self.chunks, self.current, rows_slot, idx_slot, line);
                self.emit_u16(Op::LOCAL_SET, row_slot);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, row_slot);
                self.emit_const(Value::String(Arc::from(",")));
                self.emit(Op::STR_SPLIT);
                self.emit_u16(Op::LOCAL_SET, values_slot);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, values_slot);
                self.emit_const(Value::F64(*key_index as f64));
                self.emit(Op::ARRAY_GET);
                self.emit_u16(Op::LOCAL_GET, key_slot);
                self.emit_file_key_compare(*relation);
                self.emit(Op::DYN_TO_BOOL);
                self.emit(Op::DYN_NOT);
                let no_match = self.emit_jump(Op::BR_IF_TRUE);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_u16(Op::LOCAL_SET, found_slot);
                self.emit(Op::DROP);
                self.chunks[self.current].emit_br(state.break_depth(0), line);
                self.patch_jump(no_match);
                common::loops::emit_for_in_end(&mut self.chunks, self.current, idx_slot, state, line);

                self.emit_global_map_set_null("__vb_record_current_index_by_handle", file_slot);
                self.emit_u16(Op::LOCAL_GET, found_slot);
                self.emit(Op::REF_IS_NULL);
                let found = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_global_map_set_from_local("__vb_record_next_index_by_handle", file_slot, len_slot);
                self.emit_global_map_set_const("__vb_file_eof_by_handle", file_slot, Value::Bool(true));
                let done = self.emit_jump(Op::BR);
                self.patch_jump(found);
                self.emit_global_map_set_from_local("__vb_record_next_index_by_handle", file_slot, found_slot);
                self.emit_global_map_set_const("__vb_file_eof_by_handle", file_slot, Value::Bool(false));
                self.patch_jump(done);
            }
            StmtKind::InputRecordFile { file_number, variables, key_index, key_value } => {
                let line = self.line;
                let file_slot = self.define_local("__vb_record_file_number");
                let rows_slot = self.define_local("__vb_record_rows");
                let len_slot = self.define_local("__vb_record_len");
                let idx_slot = self.define_local("__vb_record_idx");
                let row_slot = self.define_local("__vb_record_row");
                let values_slot = self.define_local("__vb_record_values");
                let found_slot = self.define_local("__vb_record_found");
                let key_slot = key_value.as_ref().map(|_| self.define_local("__vb_record_key"));

                self.compile_expr(file_number)?;
                self.emit_u16(Op::LOCAL_SET, file_slot);
                self.emit(Op::DROP);
                self.emit_record_rows_cache(file_slot, rows_slot, len_slot);

                if let Some(key_expr) = key_value {
                    let key_slot = key_slot.expect("key slot allocated when key_value exists");
                    let key_index = key_index.unwrap_or(0);

                    self.compile_expr(key_expr)?;
                    self.emit_u16(Op::LOCAL_SET, key_slot);
                    self.emit(Op::DROP);
                    self.emit(Op::NULL);
                    self.emit_u16(Op::LOCAL_SET, found_slot);
                    self.emit(Op::DROP);

                    let state = common::loops::emit_for_in_start(&mut self.chunks, self.current, rows_slot, idx_slot, line);
                    self.emit_u16(Op::LOCAL_SET, row_slot);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, row_slot);
                    self.emit_const(Value::String(Arc::from(",")));
                    self.emit(Op::STR_SPLIT);
                    self.emit_u16(Op::LOCAL_SET, values_slot);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, values_slot);
                    self.emit_const(Value::F64(key_index as f64));
                    self.emit(Op::ARRAY_GET);
                    self.emit_u16(Op::LOCAL_GET, key_slot);
                    self.emit_file_key_compare(FileKeyRelation::Equal);
                    self.emit(Op::DYN_TO_BOOL);
                    self.emit(Op::DYN_NOT);
                    let no_match = self.emit_jump(Op::BR_IF_TRUE);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit_u16(Op::LOCAL_SET, found_slot);
                    self.emit(Op::DROP);
                    self.chunks[self.current].emit_br(state.break_depth(0), line);
                    self.patch_jump(no_match);
                    common::loops::emit_for_in_end(&mut self.chunks, self.current, idx_slot, state, line);

                    self.emit_u16(Op::LOCAL_GET, found_slot);
                    self.emit(Op::REF_IS_NULL);
                    let found = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit_record_assign_nulls(variables);
                    self.emit_global_map_set_null("__vb_record_current_index_by_handle", file_slot);
                    self.emit_global_map_set_const("__vb_file_eof_by_handle", file_slot, Value::Bool(true));
                    let done = self.emit_jump(Op::BR);
                    self.patch_jump(found);
                    self.emit_u16(Op::LOCAL_GET, rows_slot);
                    self.emit_u16(Op::LOCAL_GET, found_slot);
                    self.emit(Op::ARRAY_GET);
                    self.emit_u16(Op::LOCAL_SET, row_slot);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, row_slot);
                    self.emit_const(Value::String(Arc::from(",")));
                    self.emit(Op::STR_SPLIT);
                    self.emit_u16(Op::LOCAL_SET, values_slot);
                    self.emit(Op::DROP);
                    self.emit_record_assign_values_from_local(values_slot, variables);
                    self.emit_global_map_set_from_local("__vb_record_current_index_by_handle", file_slot, found_slot);
                    self.emit_u16(Op::LOCAL_GET, found_slot);
                    self.emit(Op::I32_CONST_1);
                    self.emit(Op::I32_ADD);
                    self.emit_u16(Op::LOCAL_SET, idx_slot);
                    self.emit(Op::DROP);
                    self.emit_global_map_set_from_local("__vb_record_next_index_by_handle", file_slot, idx_slot);
                    self.emit_global_map_set_const("__vb_file_eof_by_handle", file_slot, Value::Bool(false));
                    self.patch_jump(done);
                } else {
                    self.emit_global_map_get_into_local("__vb_record_next_index_by_handle", file_slot, idx_slot);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit(Op::REF_IS_NULL);
                    let have_idx = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit(Op::I32_CONST_0);
                    self.emit_u16(Op::LOCAL_SET, idx_slot);
                    self.emit(Op::DROP);
                    self.patch_jump(have_idx);

                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit_u16(Op::LOCAL_GET, len_slot);
                    self.emit(Op::DYN_LT);
                    let has_row = self.emit_jump(Op::BR_IF_TRUE);
                    self.emit_record_assign_nulls(variables);
                    self.emit_global_map_set_null("__vb_record_current_index_by_handle", file_slot);
                    self.emit_global_map_set_const("__vb_file_eof_by_handle", file_slot, Value::Bool(true));
                    let done = self.emit_jump(Op::BR);
                    self.patch_jump(has_row);
                    self.emit_u16(Op::LOCAL_GET, rows_slot);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit(Op::ARRAY_GET);
                    self.emit_u16(Op::LOCAL_SET, row_slot);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, row_slot);
                    self.emit_const(Value::String(Arc::from(",")));
                    self.emit(Op::STR_SPLIT);
                    self.emit_u16(Op::LOCAL_SET, values_slot);
                    self.emit(Op::DROP);
                    self.emit_record_assign_values_from_local(values_slot, variables);
                    self.emit_global_map_set_from_local("__vb_record_current_index_by_handle", file_slot, idx_slot);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit(Op::I32_CONST_1);
                    self.emit(Op::I32_ADD);
                    self.emit_u16(Op::LOCAL_SET, idx_slot);
                    self.emit(Op::DROP);
                    self.emit_global_map_set_from_local("__vb_record_next_index_by_handle", file_slot, idx_slot);
                    self.emit_global_map_set_const("__vb_file_eof_by_handle", file_slot, Value::Bool(false));
                    self.patch_jump(done);
                }
            }
            StmtKind::RewriteRecordFile { file_number, items, field_formats } => {
                let line = self.line;
                let file_slot = self.define_local("__vb_rewrite_file_number");
                let rows_slot = self.define_local("__vb_rewrite_rows");
                let len_slot = self.define_local("__vb_rewrite_len");
                let current_slot = self.define_local("__vb_rewrite_current");
                let line_slot = self.define_local("__vb_rewrite_line");
                let items_slot = self.define_local("__vb_rewrite_items");
                let path_slot = self.define_local("__vb_rewrite_path");
                let path_map_key = self.shared_global_slot("__vb_file_path_by_handle");

                self.compile_expr(file_number)?;
                self.emit_u16(Op::LOCAL_SET, file_slot);
                self.emit(Op::DROP);
                self.emit_record_rows_cache(file_slot, rows_slot, len_slot);
                self.emit_global_map_get_into_local("__vb_record_current_index_by_handle", file_slot, current_slot);
                self.emit_u16(Op::LOCAL_GET, current_slot);
                self.emit(Op::REF_IS_NULL);
                let has_current = self.emit_jump(Op::BR_IF_FALSE);
                let done = self.emit_jump(Op::BR);
                self.patch_jump(has_current);

                for (index, item) in items.iter().enumerate() {
                    self.compile_expr(item)?;
                    self.emit_record_rewrite_field_format(field_formats.get(index).and_then(|format| format.as_ref()));
                }
                common::collections::emit_array_new(&mut self.chunks, self.current, items.len() as u16, line);
                self.emit_u16(Op::LOCAL_SET, items_slot);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, items_slot);
                self.emit_const(Value::String(Arc::from(",")));
                common::collections::emit_join(&mut self.chunks, self.current, line);
                self.emit_u16(Op::LOCAL_SET, line_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, rows_slot);
                self.emit_u16(Op::LOCAL_GET, current_slot);
                self.emit_u16(Op::LOCAL_GET, line_slot);
                common::collections::emit_set(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);

                self.emit_ensure_global_map("__vb_file_path_by_handle");
                self.emit_u16(Op::GLOBAL_GET, path_map_key);
                self.emit_u16(Op::LOCAL_GET, file_slot);
                self.emit(Op::ARRAY_GET);
                self.emit_u16(Op::LOCAL_SET, path_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, path_slot);
                self.emit_u16(Op::LOCAL_GET, rows_slot);
                self.emit_const(Value::String(Arc::from("\n")));
                common::collections::emit_join(&mut self.chunks, self.current, line);
                let write_file_idx = self.import("wasi:filesystem", "writeFile");
                self.emit_host_call(write_file_idx, 2);
                self.emit(Op::DROP);
                self.patch_jump(done);
            }

            // ── Export ──────────────────────────────────────────────────
            StmtKind::Export { declaration, default, .. } => {
                if let Some(decl) = declaration {
                    self.compile_stmt(decl)?;
                }
                if let Some(expr) = default {
                    self.compile_expr(expr)?;
                    let idx = self.str_const("default");
                    self.emit_u16(Op::GLOBAL_SET, idx);
                    self.emit(Op::DROP);
                }
            }

            // ── Labeled statement ───────────────────────────────────────
            StmtKind::Labeled { label, body } => {
                // Store label so the next loop push picks it up.
                self.pending_label = Some(label.clone());
                self.compile_stmt(body)?;
                self.pending_label = None;
            }

            // ── Echo (PHP/debug print) ──────────────────────────────────
            StmtKind::Echo(exprs) => {
                let line = self.line;
                let log_idx = self.import("wasi:cli", "log");
                let php_echo = self.profile.name == "php";
                if self.profile.name == "cobol" {
                    if exprs.is_empty() {
                        self.emit_const(Value::String(Arc::from("")));
                    } else {
                        let tostring_global = self.str_const("__vybe_tostring");
                        self.emit_u16(Op::GLOBAL_GET, tostring_global);
                        self.compile_expr(&exprs[0])?;
                        self.emit_u8(Op::CALL_REF, 1);
                        for expr in exprs.iter().skip(1) {
                            self.emit_u16(Op::GLOBAL_GET, tostring_global);
                            self.compile_expr(expr)?;
                            self.emit_u8(Op::CALL_REF, 1);
                            self.emit(Op::DYN_ADD);
                        }
                    }
                    common::io::emit_print_with_import(self.chunk(), log_idx, 1, line);
                } else {
                    for expr in exprs {
                        self.compile_expr(expr)?;
                        if php_echo {
                        // PHP: when echoing an object with `__toString`,
                        // call the method and print its result. Other
                        // values pass through. The check is a runtime
                        // struct_get on the value's `__toString` slot;
                        // if non-null, invoke as a method.
                        //
                        // Also: PHP `echo null;` writes no bytes (vs.
                        // Vybe's normal flow which would log ""); skip
                        // the log call when the expression is null so
                        // test-runner output entries match PHP-stdout
                        // bytes.
                        let v_slot = self.define_local("__echo_v");
                        self.emit_u16(Op::LOCAL_SET, v_slot);
                        self.emit(Op::DROP);
                        // Skip echo entirely if value is null.
                        self.emit_u16(Op::LOCAL_GET, v_slot);
                        self.emit(Op::REF_IS_NULL);
                        let skip_log = self.emit_jump(Op::BR_IF_TRUE);
                        // Probe __toString.
                        self.emit_u16(Op::LOCAL_GET, v_slot);
                        let ts_key = self.str_const("__toString");
                        self.emit_u16(Op::STRUCT_GET, ts_key);
                        let fn_slot = self.define_local("__echo_ts_fn");
                        self.emit_u16(Op::LOCAL_SET, fn_slot);
                        self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit(Op::REF_IS_NULL);
                        let no_method = self.emit_jump(Op::BR_IF_TRUE);
                        // Has __toString — invoke (fn, this).
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, v_slot);
                        self.emit_u8(Op::CALL_REF, 1);
                        let done = self.emit_jump(Op::BR);
                        self.patch_jump(no_method);
                        self.emit_u16(Op::LOCAL_GET, v_slot);
                        self.patch_jump(done);
                        self.emit_common("php.echo_stringify", 1, line);
                        common::io::emit_print_with_import(self.chunk(), log_idx, 1, line);
                            self.patch_jump(skip_log);
                        } else {
                            common::io::emit_print_with_import(self.chunk(), log_idx, 1, line);
                        }
                    }
                }
            }

            // ── Delete ──────────────────────────────────────────────────
            StmtKind::Delete(exprs) => {
                for expr in exprs {
                    match &expr.kind {
                        ExprKind::Member { object, field, .. } => {
                            self.compile_expr(object)?;
                            self.emit(Op::NULL);
                            let field_name = self.canon(field);
                            let idx = self.str_const(&field_name);
                            self.emit_u16(Op::STRUCT_SET, idx);
                            self.emit(Op::DROP);
                        }
                        ExprKind::Index { object, index, .. } => {
                            let line = self.line;
                            if self.is_python_profile() {
                                if let ExprKind::Slice { lower, upper, step } = &index.kind {
                                    if step.is_none() {
                                        self.compile_expr(object)?;
                                        let obj_tmp = self.define_local("__delete_slice_obj");
                                        self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);

                                        if let Some(lower) = lower {
                                            self.compile_expr(lower)?;
                                        } else {
                                            self.emit(Op::I32_CONST_0);
                                        }
                                        let start_tmp = self.define_local("__delete_slice_start");
                                        self.emit_u16(Op::LOCAL_SET, start_tmp); self.emit(Op::DROP);

                                        if let Some(upper) = upper {
                                            self.compile_expr(upper)?;
                                        } else {
                                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                            common::collections::emit_len(&mut self.chunks, self.current, line);
                                        }
                                        let end_tmp = self.define_local("__delete_slice_end");
                                        self.emit_u16(Op::LOCAL_SET, end_tmp); self.emit(Op::DROP);

                                        self.emit_u16(Op::LOCAL_GET, end_tmp);
                                        self.emit_u16(Op::LOCAL_GET, start_tmp);
                                        self.emit(Op::I32_SUB);
                                        let count_tmp = self.define_local("__delete_slice_count");
                                        self.emit_u16(Op::LOCAL_SET, count_tmp); self.emit(Op::DROP);

                                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                        self.emit_u16(Op::LOCAL_GET, start_tmp);
                                        self.emit_u16(Op::LOCAL_GET, count_tmp);
                                        common::collections::emit_remove_range(&mut self.chunks, self.current, line);
                                        self.emit(Op::DROP);
                                        continue;
                                    }
                                }
                            }

                            self.compile_expr(object)?;
                            let obj_tmp = self.define_local("__delete_obj");
                            self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);
                            self.compile_expr(index)?;
                            let key_tmp = self.define_local("__delete_key");
                            self.emit_u16(Op::LOCAL_SET, key_tmp); self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            let is_array_idx = self.import("ecma:array", "isArray");
                            self.chunk().emit_op_u16(Op::CALL_IMPORT, is_array_idx, line);
                            self.chunk().emit(1, line);
                            self.emit(Op::I32_CONST_0);
                            self.emit(Op::DYN_NE);
                            let array_path = self.emit_jump(Op::BR_IF_TRUE);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                            common::dict::emit_method_delete(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);
                            let end = self.emit_jump(Op::BR);

                            self.patch_jump(array_path);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                            common::collections::emit_remove_at(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);
                            self.patch_jump(end);
                        }
                        _ => {
                            // Delete on non-member is a no-op
                        }
                    }
                }
            }

            // ── Assert ──────────────────────────────────────────────────
            StmtKind::Assert { test, msg } => {
                self.compile_expr(test)?;
                self.emit(Op::DYN_TO_BOOL);
                let ok = self.emit_jump(Op::BR_IF_TRUE);
                if let Some(m) = msg {
                    self.compile_expr(m)?;
                } else {
                    self.emit_const(Value::String(Arc::from("Assertion failed")));
                }
                let line = self.line;
                common::errors::emit_throw(&mut self.chunks[self.current], line);
                self.patch_jump(ok);
            }

            // ── Scope declarations (Python global/nonlocal) ─────────────
            StmtKind::ScopeDecl { .. } => {
                // Handled at parse time for variable resolution — no-op at compile time
            }

            // ── Match statement (Python) ────────────────────────────────
            StmtKind::MatchStatement { subject, cases } => {
                self.compile_expr(subject)?;
                let subject_slot = self.define_local("__match_subject");
                self.emit_u16(Op::LOCAL_SET, subject_slot); self.emit(Op::DROP);
                let mut end_patches = Vec::new();
                for case in cases {
                    let skip = self.emit_match_pattern_checks(&case.pattern, subject_slot)?;
                    self.emit_match_pattern_bindings(&case.pattern, subject_slot)?;
                    if let Some(guard) = &case.guard {
                        self.compile_expr(guard)?;
                        self.emit(Op::DYN_TO_BOOL);
                        let guard_skip = self.emit_jump(Op::BR_IF_FALSE);
                        for s in &case.body { self.compile_stmt(s)?; }
                        end_patches.push(self.emit_jump(Op::BR));
                        self.patch_jump(guard_skip);
                    } else {
                        for s in &case.body { self.compile_stmt(s)?; }
                        end_patches.push(self.emit_jump(Op::BR));
                    }
                    for s in skip { self.patch_jump(s); }
                }
                for p in end_patches { self.patch_jump(p); }
            }

            // ── Empty ───────────────────────────────────────────────────
            StmtKind::Empty => {}
        }
        Ok(())
    }

    fn compile_enum_decl_as_class(
        &mut self,
        name: &str,
        parent: Option<&str>,
        interfaces: &[String],
        members: Vec<ClassMember>,
        span: Span,
    ) -> Result<(), String> {
        let parents: Vec<String> = parent.into_iter().map(|value| value.to_string()).collect();
        let cname = self.canon(name);
        self.defined_globals.insert(cname.clone());
        self.defined_classes.insert(cname.clone());
        crate::common::classes::emit::emit_class_from_ast(
            self,
            span,
            &cname,
            &parents,
            interfaces,
            &members,
            &ClassModifiers::default(),
            false,
        )
    }

    fn compile_shared_enum_decl(
        &mut self,
        name: &str,
        interfaces: &[String],
        body_members: &[ClassMember],
        members: &[EnumMember],
        span: Span,
    ) -> Result<(), String> {
        let static_modifiers = {
            let mut modifiers = Modifiers::default();
            modifiers.is_static = true;
            modifiers
        };
        let mut synthetic_members = body_members.to_vec();
        let mut next_val = 0i64;

        for member in members {
            let value_expr = if let Some(value) = &member.value {
                if let ExprKind::Lit(Literal::Int(n)) = &value.kind {
                    next_val = *n;
                }
                value.clone()
            } else {
                Expression::new(ExprKind::Lit(Literal::Int(next_val)))
            };
            synthetic_members.push(ClassMember::Field {
                name: member.name.clone(),
                type_hint: Some(name.to_string()),
                init: Some(value_expr),
                modifiers: static_modifiers.clone(),
                with_events: false,
                array_bounds: None,
            });
            next_val += 1;
        }

        self.compile_enum_decl_as_class(name, None, interfaces, synthetic_members, span)
    }

    fn compile_dart_enum_decl(
        &mut self,
        name: &str,
        interfaces: &[String],
        body_members: &[ClassMember],
        members: &[EnumMember],
        span: Span,
    ) -> Result<(), String> {
        let mut synthetic_members = body_members.to_vec();
        let static_modifiers = {
            let mut modifiers = Modifiers::default();
            modifiers.is_static = true;
            modifiers
        };
        let mut values_array = Vec::new();

        for (index, member) in members.iter().enumerate() {
            let obj_expr = Expression::new(ExprKind::Object(vec![
                ObjectProperty::KeyValue {
                    key: Expression::string("index"),
                    value: Expression::new(ExprKind::Lit(Literal::Int(index as i64))),
                },
                ObjectProperty::KeyValue {
                    key: Expression::string("name"),
                    value: Expression::string(&member.name),
                },
                ObjectProperty::KeyValue {
                    key: Expression::string("__type"),
                    value: Expression::string(name),
                },
            ]));
            synthetic_members.push(ClassMember::Field {
                name: member.name.clone(),
                type_hint: None,
                init: Some(obj_expr.clone()),
                modifiers: static_modifiers.clone(),
                with_events: false,
                array_bounds: None,
            });
            values_array.push(ArrayElement {
                key: None,
                value: obj_expr,
                spread: false,
                by_ref: false,
            });
        }

        synthetic_members.push(ClassMember::Field {
            name: "values".into(),
            type_hint: None,
            init: Some(Expression::new(ExprKind::Array(values_array))),
            modifiers: static_modifiers,
            with_events: false,
            array_bounds: None,
        });

        self.compile_enum_decl_as_class(name, Some("Enum"), interfaces, synthetic_members, span)
    }

    fn compile_php_enum_decl(
        &mut self,
        name: &str,
        backing_type: Option<&String>,
        interfaces: &[String],
        body_members: &[ClassMember],
        members: &[EnumMember],
        span: Span,
    ) -> Result<(), String> {
        let mut synthetic_members = Vec::new();

        synthetic_members.extend(body_members.iter().cloned());

        if !members.is_empty() {
            let elements: Vec<ArrayElement> = members.iter().map(|member| ArrayElement {
                key: None,
                value: Expression::new(ExprKind::StaticAccess {
                    class: Box::new(Expression::ident(name)),
                    member: Box::new(Expression::ident(&member.name)),
                }),
                spread: false,
                by_ref: false,
            }).collect();
            let cases_method = Statement::new(StmtKind::FunctionDecl {
                name: "cases".to_string(),
                params: vec![],
                return_type: Some("array".to_string()),
                body: vec![Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Array(elements)))) )],
                modifiers: Modifiers { is_static: true, ..Modifiers::default() },
                handles: Vec::new(),
                is_async: false,
                is_generator: false,
                is_sub: false,
            });
            synthetic_members.push(ClassMember::Method(Box::new(cases_method)));
        }

        if backing_type.is_some() && members.iter().any(|member| member.value.is_some()) {
            let mk_param = |param_name: &str| Param {
                name: param_name.to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            };
            let case_ref = |case_name: &str| Expression::new(ExprKind::StaticAccess {
                class: Box::new(Expression::ident(name)),
                member: Box::new(Expression::ident(case_name)),
            });
            let build_match_chain = |fallback: Statement| -> Vec<Statement> {
                let mut body = Vec::new();
                for member in members {
                    if let Some(backing) = member.value.clone() {
                        let cond = Expression::new(ExprKind::Binary {
                            op: BinOp::StrictEq,
                            left: Box::new(Expression::ident("v")),
                            right: Box::new(backing),
                        });
                        body.push(Statement::new(StmtKind::If {
                            cond,
                            then_body: vec![Statement::new(StmtKind::Return(Some(case_ref(&member.name))))],
                            elifs: Vec::new(),
                            else_body: None,
                        }));
                    }
                }
                body.push(fallback);
                body
            };

            let try_from_method = Statement::new(StmtKind::FunctionDecl {
                name: "tryFrom".to_string(),
                params: vec![mk_param("v")],
                return_type: None,
                body: build_match_chain(Statement::new(StmtKind::Return(Some(Expression::null())))),
                modifiers: Modifiers { is_static: true, ..Modifiers::default() },
                handles: Vec::new(),
                is_async: false,
                is_generator: false,
                is_sub: false,
            });
            synthetic_members.push(ClassMember::Method(Box::new(try_from_method)));

            let from_method = Statement::new(StmtKind::FunctionDecl {
                name: "from".to_string(),
                params: vec![mk_param("v")],
                return_type: None,
                body: build_match_chain(Statement::new(StmtKind::Throw {
                    expr: Some(Expression::new(ExprKind::New {
                        class: Box::new(Expression::ident("Error")),
                        args: vec![Argument::positional(Expression::string(&format!("Invalid backing value for enum \"{}\"", name)))],
                    })),
                    cause: None,
                })),
                modifiers: Modifiers { is_static: true, ..Modifiers::default() },
                handles: Vec::new(),
                is_async: false,
                is_generator: false,
                is_sub: false,
            });
            synthetic_members.push(ClassMember::Method(Box::new(from_method)));
        }

        self.compile_enum_decl_as_class(name, None, interfaces, synthetic_members, span)?;

        let class_global = self.str_const(&self.canon(name));
        let name_key = self.str_const("name");
        let value_key = self.str_const("value");
        for member in members {
            let case_slot = self.define_local(&format!("__php_enum_case_{}_{}", self.canon(name), self.canon(&member.name)));
            let case_expr = Expression::new(ExprKind::New {
                class: Box::new(Expression::ident(name)),
                args: vec![],
            });
            self.compile_expr(&case_expr)?;
            self.emit_u16(Op::LOCAL_SET, case_slot);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, case_slot);
            self.emit_const(Value::String(Arc::from(member.name.as_str())));
            self.emit_u16(Op::STRUCT_SET, name_key);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, case_slot);
            if let Some(value_expr) = &member.value {
                self.compile_expr(value_expr)?;
            } else {
                self.emit_const(Value::String(Arc::from(member.name.as_str())));
            }
            self.emit_u16(Op::STRUCT_SET, value_key);
            self.emit(Op::DROP);

            self.emit_u16(Op::GLOBAL_GET, class_global);
            self.emit_u16(Op::LOCAL_GET, case_slot);
            let case_key = self.str_const(&member.name);
            self.emit_u16(Op::STRUCT_SET, case_key);
            self.emit(Op::DROP);
        }

        Ok(())
    }

    fn emit_match_pattern_checks(&mut self, pattern: &Pattern, value_slot: u16) -> Result<Vec<usize>, String> {
        let mut fail_patches = Vec::new();
        match pattern {
            Pattern::Value(expr) | Pattern::Singleton(expr) => {
                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.compile_expr(expr)?;
                self.emit(Op::DYN_EQ);
                fail_patches.push(self.emit_jump(Op::BR_IF_FALSE));
            }
            Pattern::Sequence(items) => {
                let star_index = items.iter().position(|item| matches!(item, Pattern::Star(_)));
                let suffix_count = star_index.map(|index| items.len() - index - 1).unwrap_or(0);
                let required_len = if star_index.is_some() { items.len().saturating_sub(1) } else { items.len() };
                let len_slot = self.define_local("__match_seq_len");
                self.emit_u16(Op::LOCAL_GET, value_slot);
                common::collections::emit_len(&mut self.chunks, self.current, self.line);
                self.emit_u16(Op::LOCAL_SET, len_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, len_slot);
                self.emit_const(Value::F64(required_len as f64));
                self.emit(if star_index.is_some() { Op::DYN_GE } else { Op::DYN_EQ });
                fail_patches.push(self.emit_jump(Op::BR_IF_FALSE));

                if suffix_count == 0 {
                    let prefix_len = star_index.unwrap_or(items.len());
                    for (index, item) in items.iter().take(prefix_len).enumerate() {
                        let elem_slot = self.define_local("__match_seq_item");
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.emit_const(Value::F64(index as f64));
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        self.emit_u16(Op::LOCAL_SET, elem_slot);
                        self.emit(Op::DROP);
                        fail_patches.extend(self.emit_match_pattern_checks(item, elem_slot)?);
                    }
                }
            }
            Pattern::As { pattern: Some(sub_pattern), .. } => {
                fail_patches.extend(self.emit_match_pattern_checks(sub_pattern, value_slot)?);
            }
            Pattern::Or(patterns) => {
                if let Some(first) = patterns.first() {
                    fail_patches.extend(self.emit_match_pattern_checks(first, value_slot)?);
                }
            }
            Pattern::Wildcard | Pattern::Star(_) | Pattern::As { pattern: None, .. } | Pattern::Mapping(_) | Pattern::Class { .. } => {}
        }
        Ok(fail_patches)
    }

    fn emit_match_pattern_bindings(&mut self, pattern: &Pattern, value_slot: u16) -> Result<(), String> {
        match pattern {
            Pattern::As { pattern, name } => {
                if let Some(sub_pattern) = pattern {
                    self.emit_match_pattern_bindings(sub_pattern, value_slot)?;
                }
                if let Some(name) = name {
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    let slot = self.scope().resolve(name).unwrap_or_else(|| self.define_local(name));
                    self.emit_u16(Op::LOCAL_SET, slot);
                    self.emit(Op::DROP);
                }
            }
            Pattern::Sequence(items) => {
                let star_index = items.iter().position(|item| matches!(item, Pattern::Star(_)));
                let suffix_count = star_index.map(|index| items.len() - index - 1).unwrap_or(0);

                if suffix_count == 0 {
                    let prefix_len = star_index.unwrap_or(items.len());
                    for (index, item) in items.iter().take(prefix_len).enumerate() {
                        let elem_slot = self.define_local("__match_bind_item");
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.emit_const(Value::F64(index as f64));
                        common::collections::emit_get(&mut self.chunks, self.current, self.line);
                        self.emit_u16(Op::LOCAL_SET, elem_slot);
                        self.emit(Op::DROP);
                        self.emit_match_pattern_bindings(item, elem_slot)?;
                    }

                    if let Some(star_pos) = star_index {
                        if let Pattern::Star(Some(name)) = &items[star_pos] {
                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            self.emit_const(Value::F64(star_pos as f64));
                            self.emit_u16(Op::LOCAL_GET, value_slot);
                            common::collections::emit_len(&mut self.chunks, self.current, self.line);
                            common::collections::emit_slice(&mut self.chunks, self.current, self.line);
                            let slot = self.scope().resolve(name).unwrap_or_else(|| self.define_local(name));
                            self.emit_u16(Op::LOCAL_SET, slot);
                            self.emit(Op::DROP);
                        }
                    }
                }
            }
            Pattern::Or(patterns) => {
                if let Some(first) = patterns.first() {
                    self.emit_match_pattern_bindings(first, value_slot)?;
                }
            }
            Pattern::Value(_) | Pattern::Singleton(_) | Pattern::Wildcard | Pattern::Star(_) | Pattern::Mapping(_) | Pattern::Class { .. } => {}
        }
        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Variable declarator compilation
    // ════════════════════════════════════════════════════════════════════════

    fn compile_var_declarator(&mut self, decl: &VarDeclarator, kind: &VarDeclKind) -> Result<(), String> {
        match &decl.pattern {
            BindingPattern::Ident(name) => {
                let reflection_binding = decl
                    .init
                    .as_ref()
                    .and_then(|expr| self.resolve_reflection_binding_expr(expr));
                let mut inferred_type_hint = decl.type_hint.clone().or_else(|| {
                    decl.init.as_ref().and_then(|expr| self.infer_expr_type_hint(expr))
                });
                let resolved_type_hint = inferred_type_hint
                    .as_deref()
                    .map(|type_hint| self.resolve_source_type_alias(type_hint));
                if decl.array_bounds.is_some() {
                    if let Some(type_hint) = inferred_type_hint.as_mut() {
                        if !type_hint.trim_end().ends_with("()") {
                            type_hint.push_str("()");
                        }
                    }
                }
                let is_pascal_type_alias_decl = self.profile.name == "pascal"
                    && *kind == VarDeclKind::Const
                    && decl.init.is_none()
                    && decl.array_bounds.is_none()
                    && self.scopes.len() == 1
                    && self.scope().depth == 0;
                if is_pascal_type_alias_decl {
                    if let Some(type_hint) = resolved_type_hint.as_deref().or(decl.type_hint.as_deref()) {
                        self.source_type_aliases.insert(self.canon(name), type_hint.to_string());
                    }
                    return Ok(());
                }
                if inferred_type_hint
                    .as_deref()
                    .is_some_and(|type_hint| type_hint.trim().ends_with("()"))
                    || decl.array_bounds.is_some()
                    || resolved_type_hint
                        .as_deref()
                        .is_some_and(|type_hint| self.pascal_array_type_hint_metadata(type_hint).is_some())
                {
                    let array_type_hint = resolved_type_hint.clone().or_else(|| inferred_type_hint.clone());
                    let pascal_bounds = array_type_hint
                        .as_deref()
                        .and_then(|type_hint| self.pascal_array_type_hint_metadata(type_hint));
                    let is_fixed = decl.array_bounds.as_ref().is_some_and(|bounds| !bounds.is_empty())
                        || resolved_type_hint
                            .as_deref()
                            .and_then(|type_hint| self.pascal_array_type_hint_metadata(type_hint))
                            .is_some_and(|metadata| metadata.is_fixed);
                    self.record_array_binding(name, ArrayBindingMetadata {
                        is_fixed,
                        type_hint: array_type_hint,
                        pascal_bounds,
                    });
                }
                // Top-level vars → globals.
                // `let`/`const` inside a block scope (depth > 0) are locals
                // even at the top level — they respect block scoping.
                // ECMA-262 §10.2.11: `var` inside a function is function-
                // scoped (a local), only script-level `var` is global.
                let is_toplevel = self.scopes.len() == 1 && self.scope().depth == 0;
                let is_hoisted = *kind == VarDeclKind::Var
                    && self.profile.hoist_var
                    && self.scopes.len() == 1;

                if *kind == VarDeclKind::Static {
                    let binding = self.ensure_static_local_binding(name, inferred_type_hint.clone())?;
                    let flag_idx = self.str_const(&binding.init_flag_name);
                    self.emit_u16(Op::GLOBAL_GET, flag_idx);
                    self.emit(Op::DYN_TO_BOOL);
                    let skip_init = self.emit_jump(Op::BR_IF_TRUE);

                    self.emit_var_decl_initializer_value(decl, resolved_type_hint.as_deref())?;
                    let value_idx = self.str_const(&binding.global_name);
                    self.emit_u16(Op::GLOBAL_SET, value_idx);
                    self.emit(Op::DROP);
                    self.emit(Op::TRUE);
                    self.emit_u16(Op::GLOBAL_SET, flag_idx);
                    self.emit(Op::DROP);
                    self.patch_jump(skip_init);

                    let binding_key = self.canon(name);
                    if let Some(binding) = reflection_binding {
                        self.reflection_bindings.insert(binding_key, binding);
                    } else {
                        self.reflection_bindings.remove(&binding_key);
                    }
                    return Ok(());
                }

                // Recursive local lambdas need their binding slot defined
                // before compiling the initializer so captures resolve to the
                // enclosing local rather than an unresolved global.
                let mut predeclared_local_slot: Option<u16> = None;
                if !is_toplevel && !is_hoisted {
                    if let Some(init_expr) = decl.init.as_ref() {
                        let recursive_lambda_init = matches!(init_expr.kind,
                            ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_));
                        if recursive_lambda_init {
                            let slot = if *kind == VarDeclKind::Var && self.profile.hoist_var {
                                self.scope_mut().define_at_function_scope(name, inferred_type_hint.clone())
                            } else {
                                self.define_local_typed(name, inferred_type_hint.clone())
                            };
                            self.emit(Op::NULL);
                            self.emit_u16(Op::LOCAL_SET, slot);
                            self.emit(Op::DROP);
                            predeclared_local_slot = Some(slot);
                        }
                    }
                }

                if let Some(ref init_expr) = decl.init {
                    self.compile_expr_with_value_copy(init_expr)?;
                    self.maybe_promote_pascal_array_literal_to_set(decl.type_hint.as_deref(), init_expr);
                    // ECMA-262 §10.2.9 SetFunctionName — anonymous
                    // function expressions assigned to a binding take
                    // the binding name as their `name` property.
                    // Covers `const f = () => x` / `const f = function() {}`.
                    if self.is_js_profile() {
                        let is_anon_fn = match &init_expr.kind {
                            ExprKind::Lambda { .. } => true,
                            ExprKind::FunctionExpr(stmt) => {
                                matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name.is_empty())
                            }
                            _ => false,
                        };
                        if is_anon_fn {
                            let line = self.line;
                            self.emit(Op::DUP);
                            self.emit_const(Value::String(Arc::from(name.as_str())));
                            let name_key = self.str_const("name");
                            self.chunk().emit_op_u16(Op::STRUCT_SET, name_key, line);
                            self.emit(Op::DROP);
                        }
                    }
                } else if decl.array_bounds.is_some() || decl.type_hint.is_some() {
                    self.emit_var_decl_initializer_value(decl, resolved_type_hint.as_deref())?;
                } else {
                    self.emit(Op::NULL);
                }

                if is_toplevel || is_hoisted {
                    let cn = self.canon(name);
                    let idx = self.str_const(&cn);
                    self.emit_u16(Op::GLOBAL_SET, idx);
                    self.emit(Op::DROP);
                    if let Some(type_hint) = inferred_type_hint.as_deref() {
                        self.global_type_hints.insert(cn.clone(), Self::normalize_type_hint(type_hint));
                    }
                    self.defined_globals.insert(cn);
                } else {
                    // ECMA-262 §10.2.11: `var` is function-scoped (must
                    // survive enclosing-block exits). `let` / `const`
                    // are block-scoped. The scope helper picks the right
                    // depth based on the kind.
                    let slot = if let Some(slot) = predeclared_local_slot {
                        slot
                    } else if *kind == VarDeclKind::Var && self.profile.hoist_var {
                        self.scope_mut().define_at_function_scope(name, inferred_type_hint.clone())
                    } else {
                        self.define_local_typed(name, inferred_type_hint.clone())
                    };
                    self.emit_u16(Op::LOCAL_SET, slot);
                    self.emit(Op::DROP);
                }

                let binding_key = self.canon(name);
                if let Some(binding) = reflection_binding {
                    self.reflection_bindings.insert(binding_key, binding);
                } else {
                    self.reflection_bindings.remove(&binding_key);
                }
            }
            BindingPattern::Object(_) | BindingPattern::Array(_) => {
                // Destructuring `let { a, b } = expr` / `let [a, b] = expr`.
                // Compile RHS, then recursively bind via the helper so
                // arbitrary nesting (`{ a: { b: { c } } }`) works.
                if let Some(ref init_expr) = decl.init {
                    self.compile_expr(init_expr)?;
                    self.compile_destructure_bind(&decl.pattern)?;
                }
            }
        }
        Ok(())
    }

    /// Recursively bind a `BindingPattern` from a value on TOS. Consumes
    /// the value. Used by `let { a: { b: { c } } } = ...` and friends.
    /// Defines locals at every leaf ident — call sites must already be
    /// in the right scope.
    fn compile_destructure_bind(&mut self, pattern: &BindingPattern) -> Result<(), String> {
        match pattern {
            BindingPattern::Ident(name) => {
                let slot = self.define_local(name);
                self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
            }
            BindingPattern::Object(props) => {
                let obj_slot = self.define_local("__destruct_obj");
                self.emit_u16(Op::LOCAL_SET, obj_slot); self.emit(Op::DROP);
                for prop in props {
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_const(Value::String(Arc::from(prop.key.as_str())));
                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                    if let Some(ref default) = prop.default {
                        self.emit(Op::DUP);
                        self.emit(Op::REF_IS_NULL);
                        let has_val = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit(Op::DROP);
                        self.compile_expr(default)?;
                        self.patch_jump(has_val);
                    }
                    let target = match &prop.value {
                        Some(p) => p.clone(),
                        None => BindingPattern::Ident(prop.key.clone()),
                    };
                    self.compile_destructure_bind(&target)?;
                }
            }
            BindingPattern::Array(elems) => {
                // JS profile: if the value is a generator (ObjectKind::
                // Continuation, e.g. `let [a,b] = gen()`), drain it via
                // the WASM stack-switching `__stdlib_drain_generator`
                // helper into a real Array first. ARRAY_GET on a
                // Continuation returns undefined otherwise.
                if self.is_js_profile() {
                    let raw_slot = self.define_local("__destruct_raw");
                    self.emit_u16(Op::LOCAL_SET, raw_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, raw_slot);
                    let is_gen_idx = self.import("ecma:value", "isGenerator");
                    self.emit_host_call(is_gen_idx, 1);
                    let not_gen = self.emit_jump(Op::BR_IF_FALSE);
                    let drain_key = self.str_const("__vybe_drain_generator");
                    self.emit_u16(Op::GLOBAL_GET, drain_key);
                    self.emit_u16(Op::LOCAL_GET, raw_slot);
                    self.emit_u8(Op::CALL_REF, 1);
                    let done = self.emit_jump(Op::BR);
                    self.patch_jump(not_gen);
                    self.emit_u16(Op::LOCAL_GET, raw_slot);
                    self.patch_jump(done);
                }
                let arr_slot = self.define_local("__destruct_arr");
                self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);
                for (i, elem) in elems.iter().enumerate() {
                    match elem {
                        ArrayPatternElem::Pattern(pat, default) => {
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_const(Value::F64(i as f64));
                            { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                            if let Some(def) = default {
                                self.emit(Op::DUP);
                                self.emit(Op::REF_IS_NULL);
                                let has_val = self.emit_jump(Op::BR_IF_FALSE);
                                self.emit(Op::DROP);
                                self.compile_expr(def)?;
                                self.patch_jump(has_val);
                            }
                            self.compile_destructure_bind(pat)?;
                        }
                        ArrayPatternElem::Rest(name) => {
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_const(Value::F64(i as f64));
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                            let line = self.line;
                            common::collections::emit_slice(&mut self.chunks, self.current, line);
                            let slot = self.define_local(name);
                            self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                        }
                        ArrayPatternElem::Hole => {}
                    }
                }
            }
        }
        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Assignment target
    // ════════════════════════════════════════════════════════════════════════

    fn compile_array_pattern_assignment_from_slot(
        &mut self,
        arr_slot: u16,
        elems: &[ArrayPatternElem],
    ) -> Result<(), String> {
        for (i, elem) in elems.iter().enumerate() {
            match elem {
                ArrayPatternElem::Pattern(BindingPattern::Ident(name), _) => {
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit_const(Value::F64(i as f64));
                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                    self.emit_var_set(name);
                }
                ArrayPatternElem::Pattern(BindingPattern::Array(items), _) => {
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit_const(Value::F64(i as f64));
                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                    let nested_slot = self.define_local("__destruct_nested_arr");
                    self.emit_u16(Op::LOCAL_SET, nested_slot); self.emit(Op::DROP);
                    self.compile_array_pattern_assignment_from_slot(nested_slot, items)?;
                }
                ArrayPatternElem::Rest(name) => {
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit_const(Value::F64(i as f64));
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                    let line = self.line;
                    common::collections::emit_slice(&mut self.chunks, self.current, line);
                    self.emit_var_set(name);
                }
                ArrayPatternElem::Hole | ArrayPatternElem::Pattern(BindingPattern::Object(_), _) => {}
            }
        }
        Ok(())
    }

    fn compile_assign_target(&mut self, target: &Expression) -> Result<(), String> {
        match &target.kind {
            ExprKind::Ident(name) => {
                // FuncName := value assigns to Result slot (Pascal/VB)
                if let Some(ref fn_name) = self.current_func_name.clone() {
                    let matches = if self.case_sensitive { name == fn_name } else { name.eq_ignore_ascii_case(fn_name) };
                    if matches {
                        if let Some(rs) = self.current_result_slot {
                            self.emit_u16(Op::LOCAL_SET, rs);
                            self.emit(Op::DROP);
                            return Ok(());
                        }
                    }
                }
                // Local variable / parameter takes priority over implicit self field
                let is_local = self.scope().resolve(name).is_some()
                    || (!self.case_sensitive && self.scope().resolve_ci(name).is_some())
                    || self.has_static_local_binding(name);

                // Implicit self field write (only if NOT a local)
                if !is_local && self.is_class_field(name) {
                    let self_kw = self.profile.self_keyword.clone();
                    if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                        let tmp = self.define_local("__field_tmp");
                        self.emit_u16(Op::LOCAL_SET, tmp); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, slot);
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        let field_name = self.canon(name);
                        let idx = self.str_const(&field_name);
                        self.emit_u16(Op::STRUCT_SET, idx);
                        self.emit(Op::DROP);
                        return Ok(());
                    }
                }
                self.emit_var_set(name);
            }
            ExprKind::StaticAccess { class, member } => {
                let value_tmp = self.define_local("__static_access_value");
                self.emit_u16(Op::LOCAL_SET, value_tmp);
                self.emit(Op::DROP);

                self.compile_expr(class)?;
                let class_tmp = self.define_local("__static_access_class");
                self.emit_u16(Op::LOCAL_SET, class_tmp);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, class_tmp);
                if let ExprKind::Ident(name) = &member.kind {
                    let field_name = self.canon(name);
                    self.emit_u16(Op::LOCAL_GET, value_tmp);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::STRUCT_SET, idx);
                    self.emit(Op::DROP);
                } else {
                    self.compile_expr(member)?;
                    self.emit_u16(Op::LOCAL_GET, value_tmp);
                    let line = self.line;
                    common::collections::emit_set(&mut self.chunks, self.current, line);
                    self.emit(Op::DROP);
                }
            }
            ExprKind::Member { object, field, .. } => {
                if let ExprKind::Ident(obj_name) = &object.kind {
                    if let Some(key) = self.generic_static_member_key(obj_name, field) {
                        let tmp = self.define_local("__tmp");
                        self.emit_u16(Op::LOCAL_SET, tmp); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        let idx = self.str_const(&key);
                        self.emit_u16(Op::GLOBAL_SET, idx);
                        self.emit(Op::DROP);
                        return Ok(());
                    }

                    let needs_value_type_writeback = self.expr_user_value_type_name(object).is_some()
                        || (self.profile.name == "fortran"
                            && self
                                .lookup_var_type_hint(obj_name)
                                .and_then(|type_hint| self.resolve_pending_class_name_for_type_hint(type_hint))
                                .is_some());
                    if needs_value_type_writeback {
                        let value_tmp = self.define_local("__tmp");
                        let obj_tmp = self.define_local("__value_type_member_obj");
                        self.emit_u16(Op::LOCAL_SET, value_tmp);
                        self.emit(Op::DROP);

                        self.compile_expr(object)?;
                        self.emit_u16(Op::LOCAL_SET, obj_tmp);
                        self.emit(Op::DROP);

                        let field_name = self.canon(field);
                        let idx = self.str_const(&field_name);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::LOCAL_GET, value_tmp);
                        self.emit_u16(Op::STRUCT_SET, idx);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_var_set(obj_name);
                        return Ok(());
                    }
                }

                // Proxy set-trap dispatch (JS profile, only when the
                // module references `Proxy`). Stack on entry is [value]
                // (caller pushed it); the dispatcher needs [obj, key,
                // value] so we re-stash, push obj + key string, reload
                // value, then call.
                if self.is_js_profile() && self.uses_proxy {
                    let tmp = self.define_local("__proxy_set_v");
                    self.emit_u16(Op::LOCAL_SET, tmp); self.emit(Op::DROP);
                    self.compile_expr(object)?;
                    self.emit_const(Value::String(Arc::from(field.as_str())));
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    let line = self.line;
                    crate::emitter::js::proxy_adapter::emit_proxy_set_dispatch(
                        &mut self.chunks, self.current, line,
                    );
                    self.emit(Op::DROP); // adapter leaves [value] on stack
                    return Ok(());
                }
                let tmp = self.define_local("__tmp");
                self.emit_u16(Op::LOCAL_SET, tmp); self.emit(Op::DROP);
                let field_name = self.canon(field);
                if self.profile.name == "fortran" {
                    if let ExprKind::Index { object: collection_owner, index, .. } = &object.kind {
                        let line = self.line;
                        let coll_tmp = self.define_local("__fortran_index_member_coll");
                        let key_tmp = self.define_local("__fortran_index_member_key");
                        let elem_tmp = self.define_local("__fortran_index_member_elem");
                        let field_idx = self.str_const(&field_name);

                        self.compile_expr(collection_owner)?;
                        self.emit_u16(Op::LOCAL_SET, coll_tmp);
                        self.emit(Op::DROP);

                        self.compile_array_index_operand_for_owner(collection_owner, index)?;
                        self.emit_u16(Op::LOCAL_SET, key_tmp);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, coll_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        common::collections::emit_get(&mut self.chunks, self.current, line);
                        self.emit_u16(Op::LOCAL_SET, elem_tmp);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, elem_tmp);
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        self.emit_u16(Op::STRUCT_SET, field_idx);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, coll_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        self.emit_u16(Op::LOCAL_GET, elem_tmp);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, coll_tmp);
                        self.compile_assign_target(collection_owner)?;
                        return Ok(());
                    }
                }
                if matches!(self.profile.name.as_str(), "csharp" | "vb") {
                    self.compile_expr(object)?;
                    self.emit_autoderef_pointer_cell();
                    let obj_tmp = self.define_local("__member_set_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_tmp);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let setter_key = self.str_const(&format!("__set_{}", field_name));
                    self.emit_u16(Op::STRUCT_GET, setter_key);
                    let setter_tmp = self.define_local("__member_setter");
                    self.emit_u16(Op::LOCAL_SET, setter_tmp);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, setter_tmp);
                    self.emit(Op::REF_IS_NULL);
                    let fallback = self.emit_jump(Op::BR_IF_TRUE);

                    self.emit_u16(Op::LOCAL_GET, setter_tmp);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    self.emit_u8(Op::CALL_REF, 2);
                    self.emit(Op::DROP);
                    let done = self.emit_jump(Op::BR);

                    self.patch_jump(fallback);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::STRUCT_SET, idx);
                    self.emit(Op::DROP);
                    self.patch_jump(done);
                    return Ok(());
                }

                self.compile_expr(object)?;
                self.emit_autoderef_pointer_cell();
                // JS `Object.keys` / `Object.entries` need insertion order
                // (ECMA-262 §7.3.22). The HashMap backing properties is
                // non-deterministic, so we mirror each direct write into
                // `__keys` via the host trackKey helper. Only fires for
                // JS — other languages don't promise insertion order or
                // pay the host-call overhead.
                if self.is_js_profile() && !field_name.starts_with("__") {
                    let line = self.line;
                    self.emit(Op::DUP);
                    self.emit_const(Value::String(Arc::from(field_name.as_str())));
                    let track_idx = self.import("ecma:object", "trackKey");
                    self.chunk().emit_op_u16(Op::CALL_IMPORT, track_idx, line);
                    self.chunk().emit(2, line);
                    self.emit(Op::DROP);
                }
                // JS profile member writes route through `ecma:object.set`
                // for ECMA-262 §10.1.5 OrdinarySet enforcement: frozen /
                // sealed / preventExtensions gates + `__set_<key>`
                // accessor dispatch in one place. Internal `__*` keys
                // bypass — VM bookkeeping (proxy, prototype, type stamps)
                // that the gates would block.
                if self.is_js_profile() && !field_name.starts_with("__") {
                    // Bind `__js_this = obj` so a setter installed by
                    // `Object.defineProperty` (arity-1 `set(val)`) sees
                    // the receiver via the JS method-call protocol.
                    // Stack on entry: [obj]. Stash, set __js_this,
                    // re-push, call, restore.
                    let line = self.line;
                    let obj_slot = self.define_local("__js_set_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot); self.emit(Op::DROP);
                    let saved_this = self.save_js_this("__js_prev_this_set");
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.set_js_this_from_stack();
                    self.emit_u16(Op::LOCAL_GET, obj_slot);
                    self.emit_const(Value::String(Arc::from(field_name.as_str())));
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    let set_idx = self.import("ecma:object", "set");
                    self.chunk().emit_op_u16(Op::CALL_IMPORT, set_idx, line);
                    self.chunk().emit(3, line);
                    self.emit(Op::DROP);
                    self.restore_js_this(saved_this);
                } else {
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::STRUCT_SET, idx);
                    self.emit(Op::DROP);
                }
            }
            ExprKind::Unary { op: UnaryOp::Deref, expr } | ExprKind::RefLoad(expr) => {
                let value_slot = self.define_local("__ref_store_value");
                self.emit_u16(Op::LOCAL_SET, value_slot);
                self.emit(Op::DROP);
                self.compile_expr(expr)?;
                common::references::emit_cell_store(&mut self.chunks, self.current, value_slot, self.line);
                self.emit(Op::DROP);
            }
            ExprKind::Index { object, index, .. } => {
                if self.profile.name == "fortran" {
                    if let ExprKind::Slice { lower, upper, step } = &index.kind {
                        if step.is_none() {
                            let line = self.line;
                            let value_tmp = self.define_local("__fortran_slice_value");
                            let obj_tmp = self.define_local("__fortran_slice_obj");
                            let start_tmp = self.define_local("__fortran_slice_start");
                            let end_tmp = self.define_local("__fortran_slice_end");
                            let count_tmp = self.define_local("__fortran_slice_count");
                            let replacement_tmp = self.define_local("__fortran_slice_replacement");
                            let string_value_tmp = self.define_local("__fortran_slice_string_value");

                            self.emit_u16(Op::LOCAL_SET, value_tmp); self.emit(Op::DROP);

                            self.compile_expr(object)?;
                            self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);

                            if let Some(lower) = lower {
                                self.compile_expr(lower)?;
                            } else {
                                self.emit(Op::I32_CONST_0);
                            }
                            self.emit_u16(Op::LOCAL_SET, start_tmp); self.emit(Op::DROP);

                            if let Some(upper) = upper {
                                self.compile_expr(upper)?;
                            } else {
                                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                common::collections::emit_len(&mut self.chunks, self.current, line);
                            }
                            self.emit_u16(Op::LOCAL_SET, end_tmp); self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, end_tmp);
                            self.emit_u16(Op::LOCAL_GET, start_tmp);
                            self.emit(Op::I32_SUB);
                            self.emit_u16(Op::LOCAL_SET, count_tmp); self.emit(Op::DROP);

                            let known_string_object = self
                                .infer_expr_type_hint(object)
                                .as_deref()
                                .is_some_and(Self::is_string_type_hint);
                            let string_path = if known_string_object {
                                self.emit_jump(Op::BR)
                            } else {
                                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                self.emit(Op::REF_IS_STRING);
                                self.emit_jump(Op::BR_IF_TRUE)
                            };

                            self.emit_u16(Op::LOCAL_GET, value_tmp);
                            self.emit(Op::REF_IS_ARRAY);
                            let value_is_array = self.emit_jump(Op::BR_IF_TRUE);

                            self.emit_u16(Op::LOCAL_GET, count_tmp);
                            common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
                            self.emit(Op::DUP);
                            self.emit_u16(Op::LOCAL_GET, value_tmp);
                            self.emit_const(Value::I32(0));
                            self.emit_u16(Op::LOCAL_GET, count_tmp);
                            common::collections::emit_fill(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_SET, replacement_tmp); self.emit(Op::DROP);
                            let replacement_done = self.emit_jump(Op::BR);

                            self.patch_jump(value_is_array);
                            self.emit_u16(Op::LOCAL_GET, value_tmp);
                            self.emit_u16(Op::LOCAL_SET, replacement_tmp); self.emit(Op::DROP);
                            self.patch_jump(replacement_done);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, start_tmp);
                            self.emit_u16(Op::LOCAL_GET, count_tmp);
                            common::collections::emit_remove_range(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, start_tmp);
                            self.emit_u16(Op::LOCAL_GET, replacement_tmp);
                            common::collections::emit_insert_range(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);
                            let slice_assign_done = self.emit_jump(Op::BR);

                            self.patch_jump(string_path);

                            let to_string = self.import("ecma:string", "String");
                            let pad_end = self.import("ecma:string", "padEnd");

                            self.emit_u16(Op::LOCAL_GET, value_tmp);
                            self.emit_host_call(to_string, 1);
                            self.emit_u16(Op::LOCAL_SET, string_value_tmp); self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, string_value_tmp);
                            common::strings::emit_length(self.chunk(), line);
                            self.emit_u16(Op::LOCAL_GET, count_tmp);
                            self.emit(Op::DYN_GT);
                            let no_truncate = self.emit_jump(Op::BR_IF_FALSE);

                            self.emit_u16(Op::LOCAL_GET, string_value_tmp);
                            self.emit(Op::I32_CONST_0);
                            self.emit_u16(Op::LOCAL_GET, count_tmp);
                            common::strings::emit_substring(self.chunk(), line);
                            self.emit_u16(Op::LOCAL_SET, string_value_tmp); self.emit(Op::DROP);

                            self.patch_jump(no_truncate);

                            self.emit_u16(Op::LOCAL_GET, string_value_tmp);
                            self.emit_u16(Op::LOCAL_GET, count_tmp);
                            self.emit_const(Value::String(Arc::from(" ")));
                            self.emit_host_call(pad_end, 3);
                            self.emit_u16(Op::LOCAL_SET, string_value_tmp); self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit(Op::I32_CONST_0);
                            self.emit_u16(Op::LOCAL_GET, start_tmp);
                            common::strings::emit_substring(self.chunk(), line);
                            self.emit_u16(Op::LOCAL_GET, string_value_tmp);
                            common::strings::emit_str_concat(self.chunk(), line);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, end_tmp);
                            self.emit_const(Value::I32(i32::MAX));
                            common::strings::emit_substring(self.chunk(), line);
                            common::strings::emit_str_concat(self.chunk(), line);
                            self.compile_assign_target(object)?;

                            self.patch_jump(slice_assign_done);
                            return Ok(());
                        }
                    }

                    if matches!(&object.kind, ExprKind::Index { .. }) {
                        let line = self.line;
                        let value_tmp = self.define_local("__fortran_nested_index_value");
                        let coll_tmp = self.define_local("__fortran_nested_index_coll");
                        let key_tmp = self.define_local("__fortran_nested_index_key");

                        self.emit_u16(Op::LOCAL_SET, value_tmp);
                        self.emit(Op::DROP);

                        self.compile_expr(object)?;
                        self.emit_u16(Op::LOCAL_SET, coll_tmp);
                        self.emit(Op::DROP);

                        self.compile_array_index_operand_for_owner(object, index)?;
                        self.emit_u16(Op::LOCAL_SET, key_tmp);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, coll_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        self.emit_u16(Op::LOCAL_GET, value_tmp);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, coll_tmp);
                        self.compile_assign_target(object)?;
                        return Ok(());
                    }
                }
                if self.is_python_profile() {
                    if let ExprKind::Slice { lower, upper, step } = &index.kind {
                        if step.is_none() {
                            let line = self.line;
                            let value_tmp = self.define_local("__py_slice_value");
                            let obj_tmp = self.define_local("__py_slice_obj");
                            let start_tmp = self.define_local("__py_slice_start");
                            let end_tmp = self.define_local("__py_slice_end");
                            let count_tmp = self.define_local("__py_slice_count");

                            self.emit_u16(Op::LOCAL_SET, value_tmp); self.emit(Op::DROP);

                            self.compile_expr(object)?;
                            self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);

                            if let Some(lower) = lower {
                                self.compile_expr(lower)?;
                            } else {
                                self.emit(Op::I32_CONST_0);
                            }
                            self.emit_u16(Op::LOCAL_SET, start_tmp); self.emit(Op::DROP);

                            if let Some(upper) = upper {
                                self.compile_expr(upper)?;
                            } else {
                                self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                common::collections::emit_len(&mut self.chunks, self.current, line);
                            }
                            self.emit_u16(Op::LOCAL_SET, end_tmp); self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, end_tmp);
                            self.emit_u16(Op::LOCAL_GET, start_tmp);
                            self.emit(Op::I32_SUB);
                            self.emit_u16(Op::LOCAL_SET, count_tmp); self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, start_tmp);
                            self.emit_u16(Op::LOCAL_GET, count_tmp);
                            common::collections::emit_remove_range(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, start_tmp);
                            self.emit_u16(Op::LOCAL_GET, value_tmp);
                            common::collections::emit_insert_range(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);
                            return Ok(());
                        }
                    }
                }

                // Proxy set-trap dispatch — same shape as Member assign
                // but the key is a runtime expression.
                if self.is_js_profile() && self.uses_proxy {
                    let tmp = self.define_local("__proxy_idx_set_v");
                    self.emit_u16(Op::LOCAL_SET, tmp); self.emit(Op::DROP);
                    self.compile_expr(object)?;
                    self.compile_expr(index)?;
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    let line = self.line;
                    crate::emitter::js::proxy_adapter::emit_proxy_set_dispatch(
                        &mut self.chunks, self.current, line,
                    );
                    self.emit(Op::DROP);
                    return Ok(());
                }
                // PHP `$arr[] = v` — empty bracket with null index is the
                // auto-append form; route through collections::emit_push.
                // Every emit here goes via common::collections so the
                // provider (ecma:array / vybe:array / polyfill) is
                // swappable in one place.
                let is_append = matches!(
                    &index.kind,
                    ExprKind::Lit(crate::ast::Literal::Null)
                );
                let line = self.line;
                let tmp = self.define_local("__tmp");
                self.emit_u16(Op::LOCAL_SET, tmp); self.emit(Op::DROP);
                if self.is_php_profile() {
                    if let ExprKind::Member { object: recv, field, null_safe } = &object.kind {
                        if !*null_safe {
                            let recv_tmp = self.define_local("__php_index_member_recv");
                            let coll_tmp = self.define_local("__php_index_member_coll");
                            let field_name = self.canon(field);

                            self.compile_expr(recv)?;
                            self.emit_u16(Op::LOCAL_SET, recv_tmp); self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, recv_tmp);
                            let field_idx = self.str_const(&field_name);
                            self.emit_u16(Op::STRUCT_GET, field_idx);
                            self.emit_u16(Op::LOCAL_SET, coll_tmp); self.emit(Op::DROP);

                            if is_append {
                                self.emit_u16(Op::LOCAL_GET, coll_tmp);
                                self.emit_u16(Op::LOCAL_GET, tmp);
                                common::collections::emit_push(&mut self.chunks, self.current, line);
                                self.emit(Op::DROP);
                            } else {
                                self.compile_expr(index)?;
                                let key_tmp = self.define_local("__php_index_member_key");
                                self.emit_u16(Op::LOCAL_SET, key_tmp); self.emit(Op::DROP);
                                self.emit_php_promote_empty_array_for_string_key(coll_tmp, key_tmp, line);
                                let needs_track_tmp = self.emit_php_array_assoc_key_is_missing(coll_tmp, key_tmp, line);
                                self.emit_u16(Op::LOCAL_GET, coll_tmp);
                                self.emit_u16(Op::LOCAL_GET, key_tmp);
                                self.emit_u16(Op::LOCAL_GET, tmp);
                                common::collections::emit_set(&mut self.chunks, self.current, line);
                                self.emit(Op::DROP);

                                self.emit_u16(Op::LOCAL_GET, recv_tmp);
                                self.emit_u16(Op::LOCAL_GET, coll_tmp);
                                self.emit_u16(Op::STRUCT_SET, field_idx);
                                self.emit(Op::DROP);

                                self.emit_u16(Op::LOCAL_GET, recv_tmp);
                                self.emit_u16(Op::STRUCT_GET, field_idx);
                                self.emit_u16(Op::LOCAL_SET, coll_tmp);
                                self.emit(Op::DROP);

                                self.emit_php_array_assoc_key_tracking_when_missing(coll_tmp, key_tmp, needs_track_tmp, line);
                            }

                            self.emit_u16(Op::LOCAL_GET, recv_tmp);
                            self.emit_u16(Op::LOCAL_GET, coll_tmp);
                            self.emit_u16(Op::STRUCT_SET, field_idx);
                            self.emit(Op::DROP);
                            return Ok(());
                        }
                    }
                }
                if self.profile.name == "fortran" {
                    if let ExprKind::Member { object: recv, field, null_safe } = &object.kind {
                        if !*null_safe {
                            let recv_tmp = self.define_local("__fortran_index_member_recv");
                            let coll_tmp = self.define_local("__fortran_index_member_coll");
                            let key_tmp = self.define_local("__fortran_index_member_key");
                            let field_name = self.canon(field);
                            let field_idx = self.str_const(&field_name);

                            self.compile_expr(recv)?;
                            self.emit_u16(Op::LOCAL_SET, recv_tmp);
                            self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, recv_tmp);
                            self.emit_u16(Op::STRUCT_GET, field_idx);
                            self.emit_u16(Op::LOCAL_SET, coll_tmp);
                            self.emit(Op::DROP);

                            self.compile_array_index_operand_for_owner(object, index)?;
                            self.emit_u16(Op::LOCAL_SET, key_tmp);
                            self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, coll_tmp);
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                            self.emit_u16(Op::LOCAL_GET, tmp);
                            common::collections::emit_set(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, recv_tmp);
                            self.emit_u16(Op::LOCAL_GET, coll_tmp);
                            self.emit_u16(Op::STRUCT_SET, field_idx);
                            self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, recv_tmp);
                            self.compile_assign_target(recv)?;

                            return Ok(());
                        }
                    }
                }
                if is_append {
                    self.compile_expr(object)?;
                    self.emit_autoderef_pointer_cell();
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    common::collections::emit_push(&mut self.chunks, self.current, line);
                    // ecma:array.push leaves [new_length]; drop it.
                    self.emit(Op::DROP);
                } else if matches!(self.profile.name.as_str(), "csharp" | "vb") {
                    self.compile_expr(object)?;
                    self.emit_autoderef_pointer_cell();
                    let obj_tmp = self.define_local("__index_set_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    let setter_key = self.str_const("__set___index__");
                    self.emit_u16(Op::STRUCT_GET, setter_key);
                    let setter_tmp = self.define_local("__index_setter");
                    self.emit_u16(Op::LOCAL_SET, setter_tmp); self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, setter_tmp);
                    self.emit(Op::REF_IS_NULL);
                    let fallback = self.emit_jump(Op::BR_IF_TRUE);

                    self.emit_u16(Op::LOCAL_GET, setter_tmp);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.compile_collection_key(object, index)?;
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    self.emit_u8(Op::CALL_REF, 3);
                    self.emit(Op::DROP);
                    let done = self.emit_jump(Op::BR);

                    self.patch_jump(fallback);
                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                    self.compile_collection_key(object, index)?;
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    common::collections::emit_set(&mut self.chunks, self.current, line);
                    self.emit(Op::DROP);
                    self.patch_jump(done);
                } else {
                    self.compile_expr(object)?;
                    self.emit_autoderef_pointer_cell();
                    self.compile_array_index_operand_for_owner(object, index)?;
                    if self.is_php_profile() {
                        let key_tmp = self.define_local("__php_idx_key");
                        let obj_tmp = self.define_local("__php_idx_obj");
                        self.emit_u16(Op::LOCAL_SET, key_tmp); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);
                        self.emit_php_promote_empty_array_for_string_key(obj_tmp, key_tmp, line);
                        let needs_track_tmp = self.emit_php_array_assoc_key_is_missing(obj_tmp, key_tmp, line);

                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);
                        self.emit_php_array_assoc_key_tracking_when_missing(obj_tmp, key_tmp, needs_track_tmp, line);
                        if let ExprKind::Ident(name) = &object.kind {
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_var_set(name);
                        }
                        return Ok(());
                    }
                    if self.is_python_profile() {
                        let key_tmp = self.define_local("__py_idx_key");
                        let obj_tmp = self.define_local("__py_idx_obj");
                        self.emit_u16(Op::LOCAL_SET, key_tmp); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        let is_array_idx = self.import("ecma:array", "isArray");
                        self.chunk().emit_op_u16(Op::CALL_IMPORT, is_array_idx, line);
                        self.chunk().emit(1, line);
                        self.emit(Op::I32_CONST_0);
                        self.emit(Op::DYN_NE);
                        let array_path = self.emit_jump(Op::BR_IF_TRUE);

                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        let keys_key = self.str_const("__keys");
                        self.emit_u16(Op::STRUCT_GET, keys_key);
                        self.emit(Op::DUP);
                        self.emit(Op::REF_IS_NULL);
                        let no_keys = self.emit_jump(Op::BR_IF_TRUE);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        common::collections::emit_index_of(&mut self.chunks, self.current, line);
                        self.emit(Op::I32_CONST_0);
                        self.emit(Op::DYN_LT);
                        let key_exists = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::STRUCT_GET, keys_key);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        common::collections::emit_push(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);
                        self.patch_jump(key_exists);
                        let after_track = self.emit_jump(Op::BR);
                        self.patch_jump(no_keys);
                        self.emit(Op::DROP);
                        self.patch_jump(after_track);

                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);
                        let end = self.emit_jump(Op::BR);

                        self.patch_jump(array_path);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        common::collections::emit_set(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);
                        self.patch_jump(end);
                        return Ok(());
                    } else {
                    if self.profile.name == "go" {
                        let go_map_type = match &object.kind {
                            ExprKind::Ident(name) => self.lookup_var_type_hint(name).map(str::to_string),
                            _ => self.infer_expr_type_hint(object),
                        };
                        if go_map_type.as_deref().is_some_and(|type_hint| type_hint.trim().starts_with("map[")) {
                            let key_tmp = self.define_local("__go_idx_key");
                            let obj_tmp = self.define_local("__go_idx_obj");
                            self.emit_u16(Op::LOCAL_SET, key_tmp); self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            let keys_key = self.str_const("__keys");
                            self.emit_u16(Op::STRUCT_GET, keys_key);
                            self.emit(Op::DUP);
                            self.emit(Op::REF_IS_NULL);
                            let no_keys = self.emit_jump(Op::BR_IF_TRUE);
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                            common::collections::emit_index_of(&mut self.chunks, self.current, line);
                            self.emit(Op::I32_CONST_0);
                            self.emit(Op::DYN_LT);
                            let key_exists = self.emit_jump(Op::BR_IF_FALSE);
                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::STRUCT_GET, keys_key);
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                            common::collections::emit_push(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);
                            self.patch_jump(key_exists);
                            let after_track = self.emit_jump(Op::BR);
                            self.patch_jump(no_keys);
                            self.emit(Op::DROP);
                            self.patch_jump(after_track);

                            self.emit_u16(Op::LOCAL_GET, obj_tmp);
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                            self.emit_u16(Op::LOCAL_GET, tmp);
                            common::collections::emit_set(&mut self.chunks, self.current, line);
                            self.emit(Op::DROP);
                            return Ok(());
                        }
                    }
                    // JS profile: track insertion order via the
                    // `__keys` side channel so `Object.keys` /
                    // `Object.entries` / `Object.values` see the
                    // correct order. The HashMap backing Ordinary
                    // PHP polyfills that build assoc results
                    // (`array_flip`, `array_diff_assoc`, etc.) and
                    // any JS code that relies on §7.3.22 ordering.
                    if self.is_js_profile() {
                        let key_tmp = self.define_local("__idx_key");
                        self.emit_u16(Op::LOCAL_SET, key_tmp); self.emit(Op::DROP);
                        self.emit(Op::DUP);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                        let track_idx = self.import("ecma:object", "trackKey");
                        self.chunk().emit_op_u16(Op::CALL_IMPORT, track_idx, line);
                        self.chunk().emit(2, line);
                        self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, key_tmp);
                    }
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    common::collections::emit_set(&mut self.chunks, self.current, line);
                    // ecma:array.set leaves [null]; drop it.
                    self.emit(Op::DROP);
                    }
                }
            }
            // VB/Pascal: arr(idx) = val — Call used as index because () can
            // represent indexed access in those frontends.
            ExprKind::Call { callee, args, .. } if args.len() == 1 => {
                // Route the subscript through the owner-aware normalization
                // path so Pascal char-bound arrays and other declaration-
                // relative indices match the read path.
                let tmp = self.define_local("__tmp");
                self.emit_u16(Op::LOCAL_SET, tmp); self.emit(Op::DROP);
                self.compile_expr(callee)?;
                self.compile_array_index_operand_for_owner(callee, &args[0].value)?;
                self.emit_u16(Op::LOCAL_GET, tmp);
                let l = self.line;
                common::collections::emit_set(&mut self.chunks, self.current, l);
                self.emit(Op::DROP); // drop returned null
            }
            ExprKind::Destructure(pattern) => {
                // Destructuring assignment
                match pattern {
                    DestructurePattern::Object(props) => {
                        let obj_slot = self.define_local("__destruct_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_slot); self.emit(Op::DROP);
                        for prop in props {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            self.emit_const(Value::String(Arc::from(prop.key.as_str())));
                            { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                            let bind_name = if let Some(BindingPattern::Ident(ref n)) = prop.value {
                                n.clone()
                            } else {
                                prop.key.clone()
                            };
                            self.emit_var_set(&bind_name);
                        }
                    }
                    DestructurePattern::Array(elems) => {
                        let arr_slot = self.define_local("__destruct_arr");
                        self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);
                        self.compile_array_pattern_assignment_from_slot(arr_slot, elems)?;
                    }
                }
            }
            // JS destructuring assignment shorthand `[a, b] = [b, a]` /
            // `({ x } = obj)` — the walker produces an Array/Object
            // literal for the LHS, but the assignment target re-uses
            // the same shape. Treat each element as a separate
            // assignment to mirror the desugar `let _t = rhs; a = _t[0]; b = _t[1]`.
            ExprKind::Array(elems) => {
                let arr_slot = self.define_local("__assign_destruct_arr");
                self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);
                for (i, elem) in elems.iter().enumerate() {
                    if elem.spread { continue; }
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit_const(Value::F64(i as f64));
                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                    let target = elem.value.clone();
                    self.compile_assign_target(&target)?;
                }
            }
            ExprKind::Tuple(elems) => {
                let arr_slot = self.define_local("__assign_tuple_arr");
                self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);
                for (i, elem) in elems.iter().enumerate() {
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit_const(Value::F64(i as f64));
                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                    self.compile_assign_target(elem)?;
                }
            }
            ExprKind::Object(props) => {
                let obj_slot = self.define_local("__assign_destruct_obj");
                self.emit_u16(Op::LOCAL_SET, obj_slot); self.emit(Op::DROP);
                for prop in props {
                    if let crate::ast::ObjectProperty::Shorthand(name) = prop {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        let key = self.str_const(name);
                        self.emit_u16(Op::STRUCT_GET, key);
                        self.emit_var_set(name);
                    } else if let crate::ast::ObjectProperty::KeyValue { key, value } = prop {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        if let ExprKind::Lit(crate::ast::Literal::Str(ref s)) = key.kind {
                            let k = self.str_const(s);
                            self.emit_u16(Op::STRUCT_GET, k);
                        } else {
                            self.compile_expr(key)?;
                            let l = self.line;
                            common::collections::emit_get(&mut self.chunks, self.current, l);
                        }
                        let target = value.clone();
                        self.compile_assign_target(&target)?;
                    }
                }
            }
            _ => {}
        }
        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Binary operator emission
    // ════════════════════════════════════════════════════════════════════════

    /// JS profile: emit a single `ToPrimitive(hint)` call on the
    /// top-of-stack value, routed through the `__vybe_to_primitive`
    /// JS-source polyfill (compiled to bytecode at vybex build time).
    /// Going through the polyfill — instead of a host fn that calls
    /// `dispatch` — keeps the JS method-call protocol intact, so
    /// `__js_this` is set when the user's `valueOf` / `toString`
    /// body executes.
    ///
    /// Critical fast path: only enter the polyfill when the operand
    /// is an `Object`. Primitives skip it. Without this guard, the
    /// polyfill's own `<` / `+` / etc. operators (which the JS
    /// compiler also routes through `emit_to_primitive`) recurse
    /// into the polyfill on every iteration → infinite loop.
    fn emit_to_primitive(&mut self, hint: &str) {
        self.emit(Op::DUP);
        self.emit(Op::REF_IS_OBJECT);
        let skip = self.emit_jump(Op::BR_IF_FALSE);
        let helper = self.str_const("__vybe_to_primitive");
        let val_slot = self.define_local("__top_v");
        self.emit_u16(Op::LOCAL_SET, val_slot); self.emit(Op::DROP);
        self.emit_u16(Op::GLOBAL_GET, helper);
        self.emit_u16(Op::LOCAL_GET, val_slot);
        self.emit_const(Value::String(Arc::from(hint)));
        self.emit_u8(Op::CALL_REF, 2);
        self.patch_jump(skip);
    }

    /// JS profile: coerce both top-of-stack operands via the
    /// ToPrimitive polyfill, then to_f64 via the VM's existing
    /// `Value::as_f64` once the operand is no longer an Object.
    /// Used for `-`, `*`, `/`. Passes hint="number" per ECMA §7.1.4
    /// step 1 (ToNumber unboxes Objects with hint=number first).
    fn coerce_top_two_to_number(&mut self) {
        let t_b = self.define_local("__binop_b");
        self.emit_u16(Op::LOCAL_SET, t_b); self.emit(Op::DROP);
        // a on top → coerce
        self.emit_to_primitive("number");
        self.emit_u16(Op::LOCAL_GET, t_b);
        self.emit_to_primitive("number");
    }

    /// JS profile: ToPrimitive(hint=number) on both operands. Used
    /// before DYN_LT / DYN_GT / DYN_LE / DYN_GE so string-string lex
    /// compare and Date/valueOf-overriding instances both work.
    fn coerce_top_two_to_primitive(&mut self) {
        let t_b = self.define_local("__cmpop_b");
        self.emit_u16(Op::LOCAL_SET, t_b); self.emit(Op::DROP);
        self.emit_to_primitive("number");
        self.emit_u16(Op::LOCAL_GET, t_b);
        self.emit_to_primitive("number");
    }

    fn maybe_unbox_php_datetime_slot(&mut self, slot: u16) {
        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit(Op::REF_IS_OBJECT);
        let not_object = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, slot);
        let time_key = self.str_const("__time");
        self.emit_u16(Op::STRUCT_GET, time_key);
        let time_slot = self.define_local("__php_cmp_time");
        self.emit_u16(Op::LOCAL_SET, time_slot); self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_GET, time_slot);
        self.emit(Op::REF_IS_NULL);
        let no_time = self.emit_jump(Op::BR_IF_TRUE);
        self.emit_u16(Op::LOCAL_GET, time_slot);
        self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
        self.patch_jump(no_time);
        self.patch_jump(not_object);
    }

    fn coerce_top_two_php_datetime_for_compare(&mut self) {
        let t_b = self.define_local("__php_cmp_b");
        let t_a = self.define_local("__php_cmp_a");
        self.emit_u16(Op::LOCAL_SET, t_b); self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_SET, t_a); self.emit(Op::DROP);
        self.maybe_unbox_php_datetime_slot(t_a);
        self.maybe_unbox_php_datetime_slot(t_b);
        self.emit_u16(Op::LOCAL_GET, t_a);
        self.emit_u16(Op::LOCAL_GET, t_b);
    }

    fn emit_php_relational_compare(&mut self, compare_op: Op) {
        let t_b = self.define_local("__php_rel_cmp_b");
        let t_a = self.define_local("__php_rel_cmp_a");
        self.emit_u16(Op::LOCAL_SET, t_b); self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_SET, t_a); self.emit(Op::DROP);

        self.maybe_unbox_php_datetime_slot(t_a);
        self.maybe_unbox_php_datetime_slot(t_b);

        self.emit_u16(Op::LOCAL_GET, t_a);
        self.emit(Op::REF_TYPEOF);
        self.emit_const(Value::String(Arc::from("string")));
        self.emit(Op::DYN_EQ);
        let fallback_a = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, t_b);
        self.emit(Op::REF_TYPEOF);
        self.emit_const(Value::String(Arc::from("string")));
        self.emit(Op::DYN_EQ);
        let fallback_b = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, t_a);
        self.emit_u16(Op::LOCAL_GET, t_b);
        self.emit(Op::STR_COMPARE);
        self.emit_const(Value::I32(0));
        self.emit(compare_op);
        let done = self.emit_jump(Op::BR);

        self.patch_jump(fallback_a);
        self.patch_jump(fallback_b);
        self.emit_u16(Op::LOCAL_GET, t_a);
        self.emit_u16(Op::LOCAL_GET, t_b);
        self.emit(compare_op);
        self.patch_jump(done);
    }

    fn emit_pascal_relational_compare(&mut self, compare_op: Op) {
        let t_b = self.define_local("__pas_cmp_b");
        let t_a = self.define_local("__pas_cmp_a");
        self.emit_u16(Op::LOCAL_SET, t_b); self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_SET, t_a); self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, t_a);
        self.emit(Op::REF_TYPEOF);
        self.emit_const(Value::String(Arc::from("string")));
        self.emit(Op::DYN_EQ);
        let fallback_a = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, t_b);
        self.emit(Op::REF_TYPEOF);
        self.emit_const(Value::String(Arc::from("string")));
        self.emit(Op::DYN_EQ);
        let fallback_b = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, t_a);
        self.emit_u16(Op::LOCAL_GET, t_b);
        self.emit(Op::STR_COMPARE);
        self.emit_const(Value::I32(0));
        self.emit(compare_op);
        let done = self.emit_jump(Op::BR);

        self.patch_jump(fallback_a);
        self.patch_jump(fallback_b);
        self.emit_u16(Op::LOCAL_GET, t_a);
        self.emit_u16(Op::LOCAL_GET, t_b);
        self.emit(compare_op);
        self.patch_jump(done);
    }

    /// JS profile: ToPrimitive(hint=default) on both operands. Used
    /// before DYN_ADD per ECMA §13.15.4 — the `+` operator picks the
    /// "default" hint, which gives valueOf the first shot and falls
    /// back to toString.
    fn coerce_top_two_to_default_primitive(&mut self) {
        let t_b = self.define_local("__addop_b");
        self.emit_u16(Op::LOCAL_SET, t_b); self.emit(Op::DROP);
        self.emit_to_primitive("default");
        self.emit_u16(Op::LOCAL_GET, t_b);
        self.emit_to_primitive("default");
    }

    fn emit_js_add(&mut self) {
        let rhs_slot = self.define_local("__js_add_rhs");
        let lhs_slot = self.define_local("__js_add_lhs");
        self.emit_u16(Op::LOCAL_SET, rhs_slot); self.emit(Op::DROP);
        self.emit_u16(Op::LOCAL_SET, lhs_slot); self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit(Op::REF_TYPEOF);
        self.emit_const(Value::String(Arc::from("string")));
        self.emit(Op::DYN_EQ);
        let check_rhs_string = self.emit_jump(Op::BR_IF_FALSE);
        let lhs_string = self.emit_jump(Op::BR);

        self.patch_jump(check_rhs_string);
        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        self.emit(Op::REF_TYPEOF);
        self.emit_const(Value::String(Arc::from("string")));
        self.emit(Op::DYN_EQ);
        let numeric_path = self.emit_jump(Op::BR_IF_FALSE);
        let rhs_string = self.emit_jump(Op::BR);

        self.patch_jump(lhs_string);
        self.patch_jump(rhs_string);
        let tostring_global = self.str_const("__vybe_tostring");
        self.emit_u16(Op::GLOBAL_GET, tostring_global);
        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit_u8(Op::CALL_REF, 1);
        self.emit_u16(Op::GLOBAL_GET, tostring_global);
        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        self.emit_u8(Op::CALL_REF, 1);
        self.emit(Op::STR_CONCAT);
        let done = self.emit_jump(Op::BR);

        self.patch_jump(numeric_path);
        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit(Op::REF_IS_UNDEFINED);
        let lhs_defined = self.emit_jump(Op::BR_IF_FALSE);
        self.emit_const(Value::F64(f64::NAN));
        let lhs_nan_done = self.emit_jump(Op::BR);

        self.patch_jump(lhs_defined);
        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        self.emit(Op::REF_IS_UNDEFINED);
        let rhs_defined = self.emit_jump(Op::BR_IF_FALSE);
        self.emit_const(Value::F64(f64::NAN));
        let rhs_nan_done = self.emit_jump(Op::BR);

        self.patch_jump(rhs_defined);
        self.emit_u16(Op::LOCAL_GET, lhs_slot);
        self.emit_u16(Op::LOCAL_GET, rhs_slot);
        self.emit(Op::DYN_ADD);

        self.patch_jump(lhs_nan_done);
        self.patch_jump(rhs_nan_done);
        self.patch_jump(done);
    }

    fn compile_binop(&mut self, op: &BinOp) {
        match op {
            BinOp::Add => {
                // `dynamic_add`: JS-style `+` — concatenates when either
                // operand is a string, otherwise adds numerically. PHP,
                // Python, Lua, etc. use `.` / `..` / other operators for
                // string concat, so `+` is purely numeric and coerces
                // string operands (`"2026" + 4 == 2030`). `F64_ADD`
                // coerces both sides via `Value::as_f64()`; `DYN_ADD`
                // has the JS-style string-concat special case.
                if self.profile.dynamic_add && self.profile.name != "cobol" {
                    // JS profile: ECMA §13.15.4 — call ToPrimitive on
                    // both operands with hint "default" before adding.
                    // The polyfill returns the operand unchanged for
                    // primitives (fast path) and unboxes Objects via
                    // their valueOf/toString chain (Date, custom
                    // valueOf, class instances).
                    if self.is_js_profile() {
                        self.coerce_top_two_to_default_primitive();
                        self.emit_js_add();
                        return;
                    }
                    self.emit(Op::DYN_ADD);
                } else {
                    self.emit(Op::F64_ADD);
                }
            }
            BinOp::Sub => {
                if self.is_js_profile() { self.coerce_top_two_to_number(); }
                self.emit(Op::F64_SUB);
            },
            BinOp::Mul => {
                if self.is_js_profile() { self.coerce_top_two_to_number(); }
                self.emit(Op::F64_MUL);
            },
            BinOp::Div => {
                if self.is_js_profile() { self.coerce_top_two_to_number(); }
                self.emit(Op::F64_DIV);
            },
            BinOp::IDiv => { self.emit(Op::F64_DIV); let l = self.line; common::math::emit_trunc(self.chunk(), l); }
            BinOp::FloorDiv => { self.emit(Op::F64_DIV); let l = self.line; common::math::emit_floor(self.chunk(), l); }
            BinOp::Mod => {
                let l = self.line;
                if self.is_python_profile() {
                    common::math::emit_python_floor_mod(self.chunk(), l);
                } else {
                    let idx = self.import("ecma:math", "fmod");
                    common::expressions::emit_f64_mod_with_import(self.chunk(), idx, l);
                }
            },
            BinOp::Pow => { let l = self.line; common::math::emit_pow(self.chunk(), l); }
            BinOp::Eq => self.emit(Op::DYN_EQ),
            BinOp::NotEq => self.emit(Op::DYN_NE),
            BinOp::StrictEq => {
                // JS ===: no type coercion. Emit ref_typeof compare first,
                // then dyn_eq only if types match.
                // Stack: [a, b] → check types, then value equality.
                // Simplest: dup both, compare typeof, if different → false,
                // else dyn_eq. Using temp locals to avoid deep stack ops.
                let a_slot = self.define_local("__seq_a");
                let b_slot = self.define_local("__seq_b");
                self.emit_u16(Op::LOCAL_SET, b_slot); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_SET, a_slot); self.emit(Op::DROP);
                // Compare types
                self.emit_u16(Op::LOCAL_GET, a_slot);
                self.emit(Op::REF_TYPEOF);
                self.emit_u16(Op::LOCAL_GET, b_slot);
                self.emit(Op::REF_TYPEOF);
                self.emit(Op::DYN_EQ);
                let types_match = self.emit_jump(Op::BR_IF_TRUE);
                // Types differ → false
                self.emit_const(Value::Bool(false));
                let done = self.emit_jump(Op::BR);
                // Types match → dyn_eq
                self.patch_jump(types_match);
                self.emit_u16(Op::LOCAL_GET, a_slot);
                self.emit_u16(Op::LOCAL_GET, b_slot);
                self.emit(Op::DYN_EQ);
                self.patch_jump(done);
            }
            BinOp::StrictNotEq => {
                // JS !==: same as !(===)
                let a_slot = self.define_local("__sne_a");
                let b_slot = self.define_local("__sne_b");
                self.emit_u16(Op::LOCAL_SET, b_slot); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_SET, a_slot); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, a_slot);
                self.emit(Op::REF_TYPEOF);
                self.emit_u16(Op::LOCAL_GET, b_slot);
                self.emit(Op::REF_TYPEOF);
                self.emit(Op::DYN_EQ);
                let types_match = self.emit_jump(Op::BR_IF_TRUE);
                // Types differ → true
                self.emit_const(Value::Bool(true));
                let done = self.emit_jump(Op::BR);
                // Types match → dyn_ne
                self.patch_jump(types_match);
                self.emit_u16(Op::LOCAL_GET, a_slot);
                self.emit_u16(Op::LOCAL_GET, b_slot);
                self.emit(Op::DYN_NE);
                self.patch_jump(done);
            }
            BinOp::Lt => {
                if self.is_js_profile() { self.coerce_top_two_to_primitive(); }
                else if self.is_php_profile() { self.emit_php_relational_compare(Op::DYN_LT); return; }
                if self.profile.name == "pascal" { self.emit_pascal_relational_compare(Op::DYN_LT); }
                else if self.is_js_profile() || self.is_php_profile() { self.emit(Op::DYN_LT); }
                else {
                    let right_slot = self.define_local("__rich_cmp_rhs");
                    let left_slot = self.define_local("__rich_cmp_lhs");
                    self.emit_u16(Op::LOCAL_SET, right_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_SET, left_slot); self.emit(Op::DROP);
                    let line = self.line;
                    common::expressions::emit_rich_compare_locals(self.chunk(), left_slot, right_slot, "__lt__", Op::DYN_LT, line);
                }
            },
            BinOp::Gt => {
                if self.is_js_profile() { self.coerce_top_two_to_primitive(); }
                else if self.is_php_profile() { self.emit_php_relational_compare(Op::DYN_GT); return; }
                if self.profile.name == "pascal" { self.emit_pascal_relational_compare(Op::DYN_GT); }
                else if self.is_js_profile() || self.is_php_profile() { self.emit(Op::DYN_GT); }
                else {
                    let right_slot = self.define_local("__rich_cmp_rhs");
                    let left_slot = self.define_local("__rich_cmp_lhs");
                    self.emit_u16(Op::LOCAL_SET, right_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_SET, left_slot); self.emit(Op::DROP);
                    let line = self.line;
                    common::expressions::emit_rich_compare_locals(self.chunk(), left_slot, right_slot, "__gt__", Op::DYN_GT, line);
                }
            },
            BinOp::LtEq => {
                if self.is_js_profile() { self.coerce_top_two_to_primitive(); }
                else if self.is_php_profile() { self.emit_php_relational_compare(Op::DYN_LE); return; }
                if self.profile.name == "pascal" { self.emit_pascal_relational_compare(Op::DYN_LE); }
                else if self.is_js_profile() || self.is_php_profile() { self.emit(Op::DYN_LE); }
                else {
                    let right_slot = self.define_local("__rich_cmp_rhs");
                    let left_slot = self.define_local("__rich_cmp_lhs");
                    self.emit_u16(Op::LOCAL_SET, right_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_SET, left_slot); self.emit(Op::DROP);
                    let line = self.line;
                    common::expressions::emit_rich_compare_locals(self.chunk(), left_slot, right_slot, "__le__", Op::DYN_LE, line);
                }
            },
            BinOp::GtEq => {
                if self.is_js_profile() { self.coerce_top_two_to_primitive(); }
                else if self.is_php_profile() { self.emit_php_relational_compare(Op::DYN_GE); return; }
                if self.profile.name == "pascal" { self.emit_pascal_relational_compare(Op::DYN_GE); }
                else if self.is_js_profile() || self.is_php_profile() { self.emit(Op::DYN_GE); }
                else {
                    let right_slot = self.define_local("__rich_cmp_rhs");
                    let left_slot = self.define_local("__rich_cmp_lhs");
                    self.emit_u16(Op::LOCAL_SET, right_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_SET, left_slot); self.emit(Op::DROP);
                    let line = self.line;
                    common::expressions::emit_rich_compare_locals(self.chunk(), left_slot, right_slot, "__ge__", Op::DYN_GE, line);
                }
            },
            BinOp::Spaceship => {
                let right_slot = self.define_local("__spaceship_rhs");
                let left_slot = self.define_local("__spaceship_lhs");
                self.emit_u16(Op::LOCAL_SET, right_slot); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_SET, left_slot); self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, left_slot);
                self.emit_u16(Op::LOCAL_GET, right_slot);
                if self.is_js_profile() { self.coerce_top_two_to_primitive(); }
                else if self.is_php_profile() { self.emit_php_relational_compare(Op::DYN_LT); }
                else if self.profile.name == "pascal" { self.emit_pascal_relational_compare(Op::DYN_LT); }
                else { self.emit(Op::DYN_LT); }
                let lt_true = self.emit_jump(Op::BR_IF_TRUE);

                self.emit_u16(Op::LOCAL_GET, left_slot);
                self.emit_u16(Op::LOCAL_GET, right_slot);
                if self.is_js_profile() { self.coerce_top_two_to_primitive(); }
                else if self.is_php_profile() { self.emit_php_relational_compare(Op::DYN_GT); }
                else if self.profile.name == "pascal" { self.emit_pascal_relational_compare(Op::DYN_GT); }
                else { self.emit(Op::DYN_GT); }
                let gt_true = self.emit_jump(Op::BR_IF_TRUE);

                self.emit_const(Value::I32(0));
                let done = self.emit_jump(Op::BR);

                self.patch_jump(lt_true);
                self.emit_const(Value::I32(-1));
                let done_lt = self.emit_jump(Op::BR);

                self.patch_jump(gt_true);
                self.emit_const(Value::I32(1));

                self.patch_jump(done);
                self.patch_jump(done_lt);
            }
            BinOp::And | BinOp::Or => unreachable!(), // handled with short-circuit
            BinOp::Xor => {
                self.emit(Op::I32_XOR);
                self.emit(Op::DYN_TO_BOOL);
            }
            BinOp::Eqv => {
                self.emit(Op::I32_XOR);
                self.emit(Op::DYN_TO_BOOL);
                self.emit(Op::DYN_NOT);
            }
            BinOp::Imp => {
                let rhs_slot = self.define_local("__imp_rhs");
                let lhs_slot = self.define_local("__imp_lhs");
                self.emit_u16(Op::LOCAL_SET, rhs_slot); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_SET, lhs_slot); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, lhs_slot);
                self.emit(Op::DYN_TO_BOOL);
                self.emit(Op::DYN_NOT);
                self.emit_u16(Op::LOCAL_GET, rhs_slot);
                self.emit(Op::DYN_TO_BOOL);
                self.emit(Op::I32_OR);
                self.emit(Op::DYN_TO_BOOL);
            }
            BinOp::BitAnd => self.emit(Op::I32_AND),
            BinOp::BitOr => self.emit(Op::I32_OR),
            BinOp::BitXor => self.emit(Op::I32_XOR),
            BinOp::Shl => self.emit(Op::I32_SHL),
            BinOp::Shr => self.emit(Op::I32_SHR_S),
            BinOp::UShr => {
                self.emit(Op::I32_SHR_U);
                if self.is_js_profile() {
                    // ECMA-262 §13.10.2: `>>>` produces an unsigned 32-bit
                    // integer (Number). I32_SHR_U leaves the bit pattern
                    // in i32, but if the high bit is set (e.g. `-1 >>> 0`)
                    // the i32 → Number coercion would render as negative.
                    // Reinterpret as u32 by adding 2^32 when the i32 is
                    // negative — keeps within the f64 53-bit mantissa.
                    self.emit(Op::F64_FROM_I32);
                    self.emit(Op::DUP);
                    self.emit_const(Value::F64(0.0));
                    self.emit(Op::F64_LT);
                    let skip = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit_const(Value::F64(4_294_967_296.0));
                    self.emit(Op::F64_ADD);
                    self.patch_jump(skip);
                }
            },
            BinOp::Concat => { let l = self.line; common::strings::emit_str_concat(self.chunk(), l); }
            BinOp::In => {
                if self.is_python_profile() {
                    let l = self.line;
                    let t_y = self.define_local("__py_in_y");
                    let t_x = self.define_local("__py_in_x");
                    self.emit_u16(Op::LOCAL_SET, t_y); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_SET, t_x); self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, t_y);
                    self.emit(Op::REF_IS_STRING);
                    let string_path = self.emit_jump(Op::BR_IF_TRUE);

                    self.emit_u16(Op::LOCAL_GET, t_y);
                    let is_array = self.import("ecma:array", "isArray");
                    self.chunk().emit_op_u16(Op::CALL_IMPORT, is_array, l);
                    self.chunk().emit(1, l);
                    self.emit(Op::I32_CONST_0);
                    self.emit(Op::DYN_NE);
                    let array_path = self.emit_jump(Op::BR_IF_TRUE);

                    self.emit_u16(Op::LOCAL_GET, t_y);
                    self.emit_u16(Op::LOCAL_GET, t_x);
                    let has_in = self.import("ecma:object", "hasIn");
                    self.chunk().emit_op_u16(Op::CALL_IMPORT, has_in, l);
                    self.chunk().emit(2, l);
                    let end = self.emit_jump(Op::BR);

                    self.patch_jump(array_path);
                    self.emit_u16(Op::LOCAL_GET, t_y);
                    self.emit_u16(Op::LOCAL_GET, t_x);
                    common::collections::emit_contains(&mut self.chunks, self.current, l);
                    let array_end = self.emit_jump(Op::BR);

                    self.patch_jump(string_path);
                    self.emit_u16(Op::LOCAL_GET, t_y);
                    self.emit_u16(Op::LOCAL_GET, t_x);
                    self.emit(Op::STR_CONTAINS);

                    self.patch_jump(array_end);
                    self.patch_jump(end);
                    return;
                }

                if self.profile.name == "pascal" {
                    let t_set = self.define_local("__pascal_in_set");
                    let t_value = self.define_local("__pascal_in_value");
                    self.emit_u16(Op::LOCAL_SET, t_set); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_SET, t_value); self.emit(Op::DROP);
                    let helper = self.str_const("__vybe_pascal_set_contains");
                    self.emit_u16(Op::GLOBAL_GET, helper);
                    self.emit_u16(Op::LOCAL_GET, t_value);
                    self.emit_u16(Op::LOCAL_GET, t_set);
                    self.emit_u8(Op::CALL_REF, 2);
                    return;
                }

                // `x in y` — JS: is `x` a property KEY of `y` (not a value).
                // ECMA-262 §13.10.1 walks the prototype chain. PHP
                // `in_array` / Python `key in dict` are own-only and
                // route through their language profiles separately.
                //
                // Walker stack: `[x, y]`. hasIn expects `[y, x]`.
                let l = self.line;
                let t_y = self.define_local("__in_y");
                let t_x = self.define_local("__in_x");
                self.emit_u16(Op::LOCAL_SET, t_y); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_SET, t_x); self.emit(Op::DROP);

                // Proxy has-trap dispatch on the JS profile when the
                // module references `Proxy`. Stack: [obj, key].
                if self.is_js_profile() && self.uses_proxy {
                    self.emit_u16(Op::LOCAL_GET, t_y);
                    self.emit_u16(Op::LOCAL_GET, t_x);
                    crate::emitter::js::proxy_adapter::emit_proxy_has_dispatch(
                        &mut self.chunks, self.current, l,
                    );
                    return;
                }

                self.emit_u16(Op::LOCAL_GET, t_y);
                self.emit_u16(Op::LOCAL_GET, t_x);
                // JS uses prototype-walking `hasIn`; other languages
                // (case-insensitive profiles or non-JS) keep own-only
                // `hasOwn` semantics for their `in`-shaped operators.
                let import = if self.is_js_profile() { "hasIn" } else { "hasOwn" };
                let idx = self.import("ecma:object", import);
                self.chunk().emit_op_u16(Op::CALL_IMPORT, idx, l);
                self.chunk().emit(2, l);
            }
            BinOp::NotIn => {
                if self.is_python_profile() {
                    let l = self.line;
                    let t_y = self.define_local("__py_nin_y");
                    let t_x = self.define_local("__py_nin_x");
                    self.emit_u16(Op::LOCAL_SET, t_y); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_SET, t_x); self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, t_y);
                    self.emit(Op::REF_IS_STRING);
                    let string_path = self.emit_jump(Op::BR_IF_TRUE);

                    self.emit_u16(Op::LOCAL_GET, t_y);
                    let is_array = self.import("ecma:array", "isArray");
                    self.chunk().emit_op_u16(Op::CALL_IMPORT, is_array, l);
                    self.chunk().emit(1, l);
                    self.emit(Op::I32_CONST_0);
                    self.emit(Op::DYN_NE);
                    let array_path = self.emit_jump(Op::BR_IF_TRUE);

                    self.emit_u16(Op::LOCAL_GET, t_y);
                    self.emit_u16(Op::LOCAL_GET, t_x);
                    let has_in = self.import("ecma:object", "hasIn");
                    self.chunk().emit_op_u16(Op::CALL_IMPORT, has_in, l);
                    self.chunk().emit(2, l);
                    let end = self.emit_jump(Op::BR);

                    self.patch_jump(array_path);
                    self.emit_u16(Op::LOCAL_GET, t_y);
                    self.emit_u16(Op::LOCAL_GET, t_x);
                    common::collections::emit_contains(&mut self.chunks, self.current, l);
                    let array_end = self.emit_jump(Op::BR);

                    self.patch_jump(string_path);
                    self.emit_u16(Op::LOCAL_GET, t_y);
                    self.emit_u16(Op::LOCAL_GET, t_x);
                    self.emit(Op::STR_CONTAINS);

                    self.patch_jump(array_end);
                    self.patch_jump(end);
                    self.emit(Op::DYN_NOT);
                    return;
                }

                if self.profile.name == "pascal" {
                    let t_set = self.define_local("__pascal_nin_set");
                    let t_value = self.define_local("__pascal_nin_value");
                    self.emit_u16(Op::LOCAL_SET, t_set); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_SET, t_value); self.emit(Op::DROP);
                    let helper = self.str_const("__vybe_pascal_set_contains");
                    self.emit_u16(Op::GLOBAL_GET, helper);
                    self.emit_u16(Op::LOCAL_GET, t_value);
                    self.emit_u16(Op::LOCAL_GET, t_set);
                    self.emit_u8(Op::CALL_REF, 2);
                    self.emit(Op::DYN_NOT);
                    return;
                }

                let l = self.line;
                let t_y = self.define_local("__nin_y");
                let t_x = self.define_local("__nin_x");
                self.emit_u16(Op::LOCAL_SET, t_y); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_SET, t_x); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, t_y);
                self.emit_u16(Op::LOCAL_GET, t_x);
                // Same key-check as `in` above — route through hasOwn.
                let idx = self.import("ecma:object", "hasOwn");
                self.chunk().emit_op_u16(Op::CALL_IMPORT, idx, l);
                self.chunk().emit(2, l);
                self.emit(Op::DYN_NOT);
            }
            BinOp::InstanceOf => {
                if self.is_js_profile() {
                    let rhs_slot = self.define_local("__js_instanceof_rhs");
                    let lhs_slot = self.define_local("__js_instanceof_lhs");
                    self.emit_u16(Op::LOCAL_SET, rhs_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_SET, lhs_slot); self.emit(Op::DROP);
                    // ECMA-262 §13.10.2: `a instanceof B` first checks for
                    // `B[Symbol.hasInstance]` (canonical name `hasinstance`)
                    // and calls it as `B[hasinstance](a)` if present.
                    // Compiler-side dispatch keeps the JS method-call
                    // protocol intact (`__js_this` bound to B) — host
                    // `ctx.invoke` can't do that, so we emit the
                    // method-call inline instead of going through the
                    // host fn for this case.
                    let has_inst_key = self.str_const("hasinstance");
                    self.emit_u16(Op::LOCAL_GET, rhs_slot);
                    self.emit_u16(Op::STRUCT_GET, has_inst_key);
                    let method_slot = self.define_local("__has_inst_method");
                    self.emit_u16(Op::LOCAL_SET, method_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, method_slot);
                    self.emit(Op::REF_IS_NULL);
                    let no_custom = self.emit_jump(Op::BR_IF_TRUE);
                    let saved_this = self.save_js_this("__js_prev_this_hasinst");
                    self.emit_u16(Op::LOCAL_GET, rhs_slot);
                    self.set_js_this_from_stack();
                    self.emit_u16(Op::LOCAL_GET, method_slot);
                    self.emit_u16(Op::LOCAL_GET, lhs_slot);
                    self.emit_u8(Op::CALL_REF, 1);
                    self.emit(Op::DYN_TO_BOOL);
                    let result_slot = self.define_local("__has_inst_result");
                    self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                    self.restore_js_this(saved_this);
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    let done = self.emit_jump(Op::BR);
                    self.patch_jump(no_custom);
                    let helper = self.import("ecma:value", "instanceOf");
                    self.emit_u16(Op::LOCAL_GET, lhs_slot);
                    self.emit_u16(Op::LOCAL_GET, rhs_slot);
                    self.emit_host_call(helper, 2);
                    self.patch_jump(done);
                } else {
                    // Dynamic-RHS fallback: the static `a instanceof TypeName`
                    // form is intercepted upstream in `expressions.rs` and
                    // emitted as `Op::REF_TEST` directly. This branch only
                    // fires for the rare `a instanceof <expression>` shape.
                    //
                    // Stack on entry: [val, ctor]. We string-compare
                    // `val.__type` against `ctor.name` — the same compile-time
                    // type-stamp the constructors install via `set_type_id`.
                    let l = self.line;
                    let t_ctor = self.define_local("__io_ctor");
                    self.emit_u16(Op::LOCAL_SET, t_ctor); self.emit(Op::DROP);
                    // val is on top — get its __type
                    let type_key = self.str_const("__type");
                    self.chunk().emit_op_u16(Op::STRUCT_GET, type_key, l);
                    // push ctor.name
                    self.emit_u16(Op::LOCAL_GET, t_ctor);
                    let name_key = self.str_const("name");
                    self.chunk().emit_op_u16(Op::STRUCT_GET, name_key, l);
                    self.emit(Op::STR_EQUALS);
                }
            }
            BinOp::NullCoalesce => unreachable!(), // handled in compile_expr
            BinOp::MatMul => {
                let i = self.import("ecma:math", "matmul");
                self.emit_host_call(i, 2);
            }
            BinOp::Like => {
                let idx = self.import("vybe:string", "like");
                self.emit_host_call(idx, 2);
            }
            BinOp::Is => {
                // Reference equality
                self.emit(Op::DYN_EQ);
            }
            BinOp::IsNot => {
                self.emit(Op::DYN_EQ);
                self.emit(Op::DYN_NOT);
            }
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Compound assignment operator emission
    // ════════════════════════════════════════════════════════════════════════

    fn compile_compound_op(&mut self, op: &CompoundOp) {
        match op {
            CompoundOp::Add => {
                if self.profile.dynamic_add && self.profile.name != "cobol" {
                    if self.is_js_profile() {
                        self.coerce_top_two_to_default_primitive();
                    }
                    self.emit(Op::DYN_ADD);
                } else {
                    self.emit(Op::F64_ADD);
                }
            }
            CompoundOp::Sub => {
                if self.is_js_profile() {
                    self.coerce_top_two_to_number();
                }
                self.emit(Op::F64_SUB);
            }
            CompoundOp::Mul => {
                if self.is_js_profile() {
                    self.coerce_top_two_to_number();
                }
                self.emit(Op::F64_MUL);
            }
            CompoundOp::Div => {
                if self.is_js_profile() {
                    self.coerce_top_two_to_number();
                }
                self.emit(Op::F64_DIV);
            }
            CompoundOp::IDiv => {
                if self.is_js_profile() {
                    self.coerce_top_two_to_number();
                }
                self.emit(Op::F64_DIV);
                let l = self.line;
                common::math::emit_trunc(self.chunk(), l);
            }
            CompoundOp::Mod => {
                if self.is_js_profile() {
                    self.coerce_top_two_to_number();
                }
                let idx = self.import("ecma:math", "fmod");
                let l = self.line;
                common::expressions::emit_f64_mod_with_import(self.chunk(), idx, l);
            }
            CompoundOp::Pow => { let l = self.line; common::math::emit_pow(self.chunk(), l); }
            CompoundOp::Concat => { let l = self.line; common::strings::emit_str_concat(self.chunk(), l); }
            CompoundOp::BitAnd => self.emit(Op::I32_AND),
            CompoundOp::BitOr => self.emit(Op::I32_OR),
            CompoundOp::BitXor => self.emit(Op::I32_XOR),
            CompoundOp::Shl => self.emit(Op::I32_SHL),
            CompoundOp::Shr => self.emit(Op::I32_SHR_S),
            CompoundOp::UShr => self.emit(Op::I32_SHR_U),
            CompoundOp::And => self.emit(Op::DYN_TO_BOOL), // simplified
            CompoundOp::Or => self.emit(Op::DYN_TO_BOOL),  // simplified
            CompoundOp::NullCoalesce => {
                // a ??= b → if a is null, a = b
                // At this point both are on stack already — no-op, the whole compound assign handles it
            }
        }
    }

    fn is_csharp_delegate_handler_expr(&self, expr: &Expression) -> bool {
        match &expr.kind {
            ExprKind::Lambda { .. } | ExprKind::AddressOf(_) => true,
            ExprKind::Ident(name) => {
                if self
                    .lookup_var_type_hint(name)
                    .is_some_and(Self::is_callable_type_hint)
                {
                    return true;
                }
                if self.scope().resolve(name).is_some()
                    || (!self.case_sensitive && self.scope().resolve_ci(name).is_some())
                {
                    return false;
                }
                let cname = self.canon(name);
                self.defined_functions.contains(&cname)
                    || self.defined_class_methods.contains(&cname)
            }
            ExprKind::Member { field, .. } => {
                let cname = self.canon(field);
                self.defined_functions.contains(&cname)
                    || self.defined_class_methods.contains(&cname)
            }
            ExprKind::New { args, .. } if args.len() == 1 => {
                self.is_csharp_delegate_handler_expr(&args[0].value)
            }
            _ => false,
        }
    }

    fn assign_target_matches_expr(&self, target: &Expression, expr: &Expression) -> bool {
        match (&target.kind, &expr.kind) {
            (ExprKind::Ident(a), ExprKind::Ident(b)) => a == b,
            (
                ExprKind::Member { object: to, field: tf, .. },
                ExprKind::Member { object: eo, field: ef, .. },
            ) => {
                if !tf.eq_ignore_ascii_case(ef) {
                    return false;
                }
                self.assign_target_matches_expr(to, eo)
            }
            (ExprKind::This, ExprKind::This) => true,
            _ => false,
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Builtins (profile-driven)
    // ════════════════════════════════════════════════════════════════════════

    fn try_compile_builtin(&mut self, name: &str, args: &[&Expression]) -> Result<bool, String> {
        let line = self.line;

        if self.is_python_profile() && name == "globals" && args.is_empty() {
            common::dict::emit_new(&mut self.chunks, self.current, line);

            self.emit(Op::DUP);
            self.emit_const(Value::String(Arc::from("__main__")));
            let name_key = self.str_const("__name__");
            self.emit_u16(Op::STRUCT_SET, name_key);
            self.emit(Op::DROP);
            self.emit(Op::DUP);
            let keys_key = self.str_const("__keys");
            self.emit_u16(Op::STRUCT_GET, keys_key);
            self.emit_const(Value::String(Arc::from("__name__")));
            common::collections::emit_push(&mut self.chunks, self.current, line);
            self.emit(Op::DROP);

            let mut globals: Vec<String> = self.defined_globals.iter().cloned().collect();
            globals.sort();
            globals.dedup();
            for global in globals {
                if global == "__name__" { continue; }
                self.emit(Op::DUP);
                self.emit_var_get(&global);
                let key = self.str_const(&global);
                self.emit_u16(Op::STRUCT_SET, key);
                self.emit(Op::DROP);

                self.emit(Op::DUP);
                let keys_key = self.str_const("__keys");
                self.emit_u16(Op::STRUCT_GET, keys_key);
                self.emit_const(Value::String(Arc::from(global.as_str())));
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);
            }
            return Ok(true);
        }

        if self.is_python_profile() && name == "frozenset" && args.len() <= 1 {
            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
                let idx = self.import("ecma:array", "from");
                self.emit_host_call(idx, 1);
                self.emit_const(Value::String(Arc::from("\u{1f}")));
                common::collections::emit_join(&mut self.chunks, self.current, line);
            } else {
                self.emit_const(Value::String(Arc::from("")));
            }
            return Ok(true);
        }

        if self.profile.name == "pascal" {
            let builtin_name = self.canon(name);
            if builtin_name == "write" || builtin_name == "writeln" {
                let helper = if builtin_name == "write" {
                    "__vybe_pascal_write"
                } else {
                    "__vybe_pascal_writeln"
                };
                let helper_idx = self.str_const(helper);
                self.emit_u16(Op::GLOBAL_GET, helper_idx);

                if args.is_empty() {
                    self.emit_const(Value::String(Arc::from("")));
                } else {
                    let tostring_global = self.str_const("__vybe_tostring");
                    self.emit_u16(Op::GLOBAL_GET, tostring_global);
                    self.compile_expr(args[0])?;
                    self.emit_u8(Op::CALL_REF, 1);
                    for arg in args.iter().skip(1) {
                        self.emit_const(Value::String(Arc::from(" ")));
                        self.emit(Op::DYN_ADD);
                        self.emit_u16(Op::GLOBAL_GET, tostring_global);
                        self.compile_expr(arg)?;
                        self.emit_u8(Op::CALL_REF, 1);
                        self.emit(Op::DYN_ADD);
                    }
                }

                self.emit_u8(Op::CALL_REF, 1);
                return Ok(true);
            }

            if (builtin_name == "integer" || builtin_name == "int" || builtin_name == "longint")
                && args.len() == 1
            {
                self.compile_expr(args[0])?;
                common::math::emit_trunc(self.chunk(), line);
                return Ok(true);
            }

            if builtin_name == "delete"
                && args.len() == 3
                && matches!(&args[0].kind, ExprKind::Ident(_))
            {
                let ExprKind::Ident(var_name) = &args[0].kind else {
                    unreachable!();
                };
                let helper_idx = self.str_const("__vybe_pascal_str_remove_range");
                self.emit_u16(Op::GLOBAL_GET, helper_idx);
                self.emit_var_get(var_name);
                self.compile_expr(args[1])?;
                self.compile_expr(args[2])?;
                self.emit_u8(Op::CALL_REF, 3);
                self.emit_var_set(var_name);
                self.emit(Op::NULL);
                return Ok(true);
            }

            if builtin_name == "insert"
                && args.len() == 3
                && matches!(&args[1].kind, ExprKind::Ident(_))
            {
                let ExprKind::Ident(var_name) = &args[1].kind else {
                    unreachable!();
                };
                let helper_idx = self.str_const("__vybe_pascal_str_insert");
                self.emit_u16(Op::GLOBAL_GET, helper_idx);
                self.compile_expr(args[0])?;
                self.emit_var_get(var_name);
                self.compile_expr(args[2])?;
                self.emit_u8(Op::CALL_REF, 3);
                self.emit_var_set(var_name);
                self.emit(Op::NULL);
                return Ok(true);
            }
        }

        if self.profile.name == "pascal"
            && args.len() == 2
            && matches!(&args[0].kind, ExprKind::Ident(_))
        {
            let builtin_name = self.canon(name);
            let ExprKind::Ident(var_name) = &args[0].kind else {
                unreachable!();
            };

            let is_set_var = self
                .lookup_var_type_hint(var_name)
                .is_some_and(Self::is_pascal_set_type_hint);
            if is_set_var && (builtin_name == "include" || builtin_name == "exclude") {
                let helper = if builtin_name == "include" {
                    "__vybe_pascal_set_include"
                } else {
                    "__vybe_pascal_set_exclude"
                };
                let helper_idx = self.str_const(helper);
                self.emit_u16(Op::GLOBAL_GET, helper_idx);
                self.emit_var_get(var_name);
                self.compile_expr(args[1])?;
                self.emit_u8(Op::CALL_REF, 2);
                self.emit(Op::DROP);
                self.emit(Op::NULL);
                return Ok(true);
            }
        }

        if name.eq_ignore_ascii_case("setlength") {
            if args.len() >= 2 {
                self.compile_setlength(args[0], args[1])?;
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }

        // ── Component Model host-call resolution (qualified name → host fn) ──
        //
        // A qualified identifier whose first segment matches the profile's
        // `host_packages` list resolves directly to a Component Model host
        // call. This is how `\Vybe\Http\Response\set_status(404)` in PHP
        // reaches the `vybe:http/response` host module with zero profile
        // builtin entries. The same convention is intended to apply to every
        // language with namespaces (Python `vybe.http.request.method`, C#
        // `Vybe.Http.Request.Method`, etc.) — walkers normalize their
        // separators to `\` before reaching here so this single resolver
        // handles them all.
        if let Some((module, func)) = self.resolve_component_model_call(name) {
            for a in args { self.compile_expr(a)?; }
            let idx = self.import(&module, &func);
            self.emit_host_call(idx, args.len() as u8);
            return Ok(true);
        }

        if self.profile.name == "vb" && name.eq_ignore_ascii_case("Array") {
            common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
            for arg in args {
                self.emit(Op::DUP);
                self.compile_expr(arg)?;
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);
            }
            return Ok(true);
        }

        // ── Phase D1 pilot: Array(count, init) → ecma:array.newWithLength + fill ──
        //
        // COBOL's OCCURS walker emits `Call { callee: Array,
        // args: [count, element_init] }` in the high-level IR. This
        // intercept routes the pattern through the spec-conformant
        // `ecma:array.*` imports instead of the legacy VM-internal
        // opcodes. See `dynamicruntime_support.md` Phase D1 and the
        // reasoning in `project_dynamic_runtime_phase_state.md`.
        //
        // Narrow match: only intercept when we see `Array(count, init)`
        // specifically — 2 positional args, callee identifier "Array".
        // This avoids colliding with C#/VB `Array` namespace access
        // (`Array.Empty()`, `Array.IsArray()`, etc.) which hits
        // different code paths (namespace + member access).
        if name == "Array" && args.len() == 2 {
            // COBOL's OCCURS walker emits `Array(count, init)`. Emit:
            //   newWithLength(count)  — via common::collections
            //   fill(arr, init, 0, MAX)  — via common::collections
            // All emits route through compiler_common so the provider
            // (ecma:array / vybe:array / polyfill) is swappable in
            // one place.
            self.compile_expr(args[0])?;  // push count
            common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
            // Array is now on TOS. If the init is null-ish, we're done
            // (newWithLength already null-fills).
            let init_is_null = matches!(
                &args[1].kind,
                ExprKind::Lit(crate::ast::Literal::Null)
                    | ExprKind::Lit(crate::ast::Literal::Undefined)
            );
            if init_is_null {
                return Ok(true);
            }
            let init_is_nested_array_factory = matches!(
                &args[1].kind,
                ExprKind::Call { callee, .. }
                    if matches!(callee.kind, ExprKind::Ident(ref name) if name == "Array")
            );
            if init_is_nested_array_factory {
                let arr_slot = self.define_local("__array_ctor_result");
                self.emit_u16(Op::LOCAL_SET, arr_slot);
                self.emit(Op::DROP);

                let idx_slot = self.define_local("__array_ctor_idx");
                self.emit_const(Value::I32(0));
                self.emit_u16(Op::LOCAL_SET, idx_slot);
                self.emit(Op::DROP);

                let loop_start = self.chunks[self.current].current_offset();
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_u16(Op::LOCAL_GET, arr_slot);
                common::collections::emit_len(&mut self.chunks, self.current, line);
                self.emit(Op::DYN_LT);
                let done = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, arr_slot);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.compile_expr(args[1])?;
                self.emit(Op::ARRAY_SET);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_const(Value::I32(1));
                self.emit(Op::DYN_ADD);
                self.emit_u16(Op::LOCAL_SET, idx_slot);
                self.emit(Op::DROP);
                self.emit_loop(loop_start);
                self.patch_jump(done);
                self.emit_u16(Op::LOCAL_GET, arr_slot);
                return Ok(true);
            }
            // Stack: [arr]. Dup first so we still have the result.
            self.emit(Op::DUP);
            self.compile_expr(args[1])?;
            let zero_k = self.chunks[self.current].add_constant(vybe_bytecode::Value::I32(0));
            self.emit_u16(Op::CONST, zero_k);
            let max_k = self.chunks[self.current].add_constant(vybe_bytecode::Value::I32(i32::MAX));
            self.emit_u16(Op::CONST, max_k);
            common::collections::emit_fill(&mut self.chunks, self.current, line);
            // fill returns the array; drop the dup'd copy — the pre-dup
            // copy stays on TOS as the expression's value.
            self.emit(Op::DROP);
            return Ok(true);
        }

        if name.eq_ignore_ascii_case("__fortran_max") {
            if args.len() >= 2 {
                self.compile_expr(args[0])?;
                self.compile_expr(args[1])?;
                common::math::emit_max(self.chunk(), line);
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("__fortran_min") {
            if args.len() >= 2 {
                self.compile_expr(args[0])?;
                self.compile_expr(args[1])?;
                common::math::emit_min(self.chunk(), line);
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("__fortran_emit") || name.eq_ignore_ascii_case("__fortran_emitln") {
            let flush = name.eq_ignore_ascii_case("__fortran_emitln");
            let text_slot = self.define_local(if flush { "__fortran_emitln_text" } else { "__fortran_emit_text" });
            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
            } else {
                self.compile_expr(&Expression::string(""))?;
            }
            self.emit_u16(Op::LOCAL_SET, text_slot);
            self.emit(Op::DROP);

            self.emit_var_get("__vybe_fortran_io_buffer");
            self.emit_u16(Op::LOCAL_GET, text_slot);
            common::strings::emit_str_concat(self.chunk(), line);

            if flush {
                let message_slot = self.define_local("__fortran_emitln_message");
                self.emit_u16(Op::LOCAL_SET, message_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, message_slot);
                let idx = self.import("wasi:cli", "log");
                common::io::emit_print_with_import(self.chunk(), idx, 1, line);

                self.compile_expr(&Expression::string(""))?;
                self.emit_var_set("__vybe_fortran_io_buffer");
            } else {
                self.emit_var_set("__vybe_fortran_io_buffer");
                self.emit(Op::NULL);
            }
            return Ok(true);
        }

        if name.eq_ignore_ascii_case("__fortran_rewind") {
            let file_slot = self.define_local("__fortran_rewind_file");
            let path_slot = self.define_local("__fortran_rewind_path");

            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
            } else {
                self.emit_const(Value::I32(0));
            }
            self.emit_u16(Op::LOCAL_SET, file_slot);
            self.emit(Op::DROP);

            self.emit_global_map_get_into_local("__vb_file_path_by_handle", file_slot, path_slot);

            self.emit_u16(Op::LOCAL_GET, file_slot);
            let close_idx = self.import("wasi:filesystem", "closeFile");
            self.emit_host_call(close_idx, 1);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, path_slot);
            self.emit_const(Value::String(Arc::from("Input")));
            self.emit_u16(Op::LOCAL_GET, file_slot);
            let open_idx = self.import("wasi:filesystem", "openFile");
            self.emit_host_call(open_idx, 3);
            self.emit(Op::DROP);

            self.emit_global_map_set_const("__vb_file_eof_by_handle", file_slot, Value::Bool(false));
            self.emit_global_map_set_null("__vb_record_rows_by_handle", file_slot);
            self.emit_global_map_set_null("__vb_record_next_index_by_handle", file_slot);
            self.emit_global_map_set_null("__vb_record_current_index_by_handle", file_slot);
            self.emit(Op::NULL);
            return Ok(true);
        }

        if name.eq_ignore_ascii_case("__fortran_namelist_decl") {
            self.emit(Op::NULL);
            return Ok(true);
        }

        if name.eq_ignore_ascii_case("kind") {
            self.emit_const(Value::I32(8));
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("allocate") {
            for arg in args {
                match &arg.kind {
                    ExprKind::Call { callee, args: dims, .. } if !dims.is_empty() => {
                        if self.profile.name == "fortran" {
                            let mut dim_slots = Vec::with_capacity(dims.len());
                            for (index, dim) in dims.iter().enumerate() {
                                self.compile_expr(&dim.value)?;
                                let slot = self.define_local(&format!("__fortran_alloc_dim_{index}"));
                                self.emit_u16(Op::LOCAL_SET, slot);
                                self.emit(Op::DROP);
                                dim_slots.push(slot);
                            }

                            let ctor_name = self.fortran_allocate_ctor_name(callee);
                            self.emit_fortran_allocated_array(&dim_slots, ctor_name.as_deref());
                            self.compile_assign_target(callee)?;
                        } else {
                            self.compile_expr(&dims[0].value)?;
                            common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
                            self.compile_assign_target(callee)?;
                        }
                    }
                    _ => {
                        if self.profile.name == "fortran" {
                            if let Some(ctor_name) = self.fortran_allocate_ctor_name(arg) {
                                self.emit_fortran_ctor_call(&ctor_name);
                            } else {
                                self.emit(Op::NULL);
                            }
                            self.compile_assign_target(arg)?;
                        }
                    }
                }
            }
            self.emit(Op::NULL);
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("deallocate") {
            for arg in args {
                match &arg.kind {
                    ExprKind::Call { callee, .. } => {
                        self.emit(Op::NULL);
                        self.compile_assign_target(callee)?;
                    }
                    _ => {
                        self.emit(Op::NULL);
                        self.compile_assign_target(arg)?;
                    }
                }
            }
            self.emit(Op::NULL);
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("present") {
            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
                self.emit(Op::REF_IS_NULL);
                self.emit(Op::DYN_NOT);
            } else {
                self.emit(Op::FALSE);
            }
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("sum") {
            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
                common::collections::emit_sum(&mut self.chunks, self.current, line);
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("minval") {
            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
                common::collections::emit_pymin(&mut self.chunks, self.current, line);
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("maxval") {
            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
                common::collections::emit_pymax(&mut self.chunks, self.current, line);
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("nint") {
            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
                common::math::emit_round(self.chunk(), line);
                common::convert::emit_to_int(self.chunk(), line);
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }
        if name.eq_ignore_ascii_case("size") {
            if let Some(arg) = args.first() {
                self.compile_expr(arg)?;
                common::collections::emit_len(&mut self.chunks, self.current, line);
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }
        if self.profile.name == "fortran" && name.eq_ignore_ascii_case("matmul") {
            for arg in args {
                self.compile_expr(arg)?;
            }
            self.emit_common("fortran.matmul", args.len() as u8, line);
            return Ok(true);
        }
        if self.profile.name == "fortran" && name.eq_ignore_ascii_case("array_join") {
            if args.len() >= 2 {
                self.compile_expr(args[0])?;
                self.compile_expr(args[1])?;
                common::collections::emit_join(&mut self.chunks, self.current, line);
            } else {
                self.emit(Op::NULL);
            }
            return Ok(true);
        }
        if self.profile.name == "fortran" && name.eq_ignore_ascii_case("str_getcsv") {
            for arg in args {
                self.compile_expr(arg)?;
            }
            crate::languages::php::emitter::string_adapter::emit_str_getcsv(
                &mut self.chunks,
                self.current,
                args.len() as u8,
                line,
            );
            return Ok(true);
        }
        if self.profile.name == "fortran" && name.eq_ignore_ascii_case("this_image") {
            self.emit_const(Value::I32(1));
            return Ok(true);
        }
        if self.profile.name == "fortran" && name.eq_ignore_ascii_case("num_images") {
            self.emit_const(Value::I32(1));
            return Ok(true);
        }
        if self.profile.name == "fortran" && name.eq_ignore_ascii_case("co_sum") {
            self.emit(Op::NULL);
            return Ok(true);
        }

        // Canonical builtins — language-agnostic dispatch via compiler_common::canonical.
        // Walkers normalize language-specific syntax (arr.Length, len(arr), Length(arr),
        // arr.size, etc.) to canonical dunder names (__len__, __str__, etc.).
        // The compiler doesn't know about language-specific names — it just looks up
        // the canonical name in compiler_common's registry.
        if let Some(canonical_op) = common::canonical::CanonicalOp::from_name(name) {
            // Special case: __str__ uses stdlib via global, not host import
            if matches!(canonical_op, common::canonical::CanonicalOp::Str) {
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    let arg_slot = self.define_local("__canonical_str_arg");
                    self.emit_u16(Op::LOCAL_SET, arg_slot);
                    self.emit(Op::DROP);

                    let tostring_global = self.str_const("__vybe_tostring");
                    self.emit_u16(Op::GLOBAL_GET, tostring_global);
                    self.emit_u16(Op::LOCAL_GET, arg_slot);
                    self.emit_u8(Op::CALL_REF, 1);
                    return Ok(true);
                }
            } else {
                // Compile args, then dispatch to canonical emitter
                for a in args { self.compile_expr(a)?; }
                common::canonical::emit_canonical(canonical_op, &mut self.chunks, self.current, line);
                return Ok(true);
            }
        }

        // Look up in language profile FIRST — language profiles can
        // override the common import defaults (e.g. Dart `print` needs
        // toString conversion before logging, which is different from
        // generic `wasi:cli.log`).
        if self.profile.name == "vb"
            && args.len() == 1
        {
            if let Some(type_hint) = self.infer_expr_type_hint(&args[0]) {
                let normalized = Self::normalize_type_hint(&type_hint);
                if normalized == "datetime" || normalized.ends_with(".datetime") {
                    let field_name = if name.eq_ignore_ascii_case("year") {
                        Some("Year")
                    } else if name.eq_ignore_ascii_case("month") {
                        Some("Month")
                    } else if name.eq_ignore_ascii_case("day") {
                        Some("Day")
                    } else if name.eq_ignore_ascii_case("hour") {
                        Some("Hour")
                    } else if name.eq_ignore_ascii_case("minute") {
                        Some("Minute")
                    } else if name.eq_ignore_ascii_case("second") {
                        Some("Second")
                    } else {
                        None
                    };

                    if let Some(field_name) = field_name {
                        self.compile_expr(&args[0])?;
                        let idx = self.str_const(field_name);
                        self.emit_u16(Op::STRUCT_GET, idx);
                        return Ok(true);
                    }
                }
            }
        }

        let builtin = self.profile.lookup_builtin(name).cloned();
        // Check common import table only if the profile didn't bind it.
        if builtin.is_none() {
            if let Some(resolved) = common::imports::resolve_common_import(name) {
                match resolved {
                    common::imports::CommonImport::Host(module, func) => {
                        for a in args { self.compile_expr(a)?; }
                        let idx = self.import(module, func);
                        self.emit_host_call(idx, args.len() as u8);
                    }
                    common::imports::CommonImport::Intrinsic(intrinsic_name) => {
                        self.emit_intrinsic(intrinsic_name, args)?;
                    }
                }
                return Ok(true);
            }
        }

        if let Some(def) = builtin {
            match &def.emit {
                BuiltinEmit::Print => {
                    for a in args {
                        if let Some(enum_type) = self.console_enum_type_from_expr(a) {
                            self.emit_enum_value_to_string(&enum_type, a)?;
                        } else {
                            self.compile_expr(a)?;
                        }
                    }
                    let idx = self.import("wasi:cli", "log");
                    common::io::emit_print_with_import(self.chunk(), idx, args.len() as u8, line);
                }
                BuiltinEmit::StrLength => {
                    if !args.is_empty() {
                        self.compile_expr(args[0])?;
                        common::strings::emit_length(self.chunk(), line);
                    } else {
                        self.emit(Op::NULL);
                    }
                }
                BuiltinEmit::HostCall(module, func) => {
                    // Iterator-consuming host fns (e.g. `Array.from`,
                    // `Promise.all`) accept any iterable. JS generators
                    // (Continuation) need WASM stack-switching to
                    // drain — a host fn can't drive coroutine resume,
                    // so we drain via the `__stdlib_drain_generator`
                    // bytecode helper before the host call.
                    let drain_first_arg = self.is_js_profile()
                        && (module == "ecma:array" && func == "from");
                    if drain_first_arg && !args.is_empty() {
                        self.compile_expr(args[0])?;
                        let v_slot = self.define_local("__hc_iter_v");
                        self.emit_u16(Op::LOCAL_SET, v_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, v_slot);
                        let is_gen_idx = self.import("ecma:value", "isGenerator");
                        self.emit_host_call(is_gen_idx, 1);
                        let not_gen = self.emit_jump(Op::BR_IF_FALSE);
                        let drain_key = self.str_const("__vybe_drain_generator");
                        self.emit_u16(Op::GLOBAL_GET, drain_key);
                        self.emit_u16(Op::LOCAL_GET, v_slot);
                        self.emit_u8(Op::CALL_REF, 1);
                        let done = self.emit_jump(Op::BR);
                        self.patch_jump(not_gen);
                        self.emit_u16(Op::LOCAL_GET, v_slot);
                        self.patch_jump(done);
                        for a in args.iter().skip(1) { self.compile_expr(a)?; }
                    } else {
                        for a in args { self.compile_expr(a)?; }
                    }
                    let idx = self.import(module, func);
                    self.emit_host_call(idx, args.len() as u8);
                }
                BuiltinEmit::Opcode(op_name) => {
                    self.emit_builtin_opcode(op_name, args)?;
                }
                BuiltinEmit::MutateVar(op) => {
                    if let Some(first) = args.first() {
                        if let ExprKind::Ident(var) = &first.kind {
                            let var = var.clone();
                            self.emit_var_get(&var);
                            if args.len() > 1 { self.compile_expr(args[1])?; } else { self.emit_const(Value::F64(1.0)); }
                            match op.as_str() {
                                "add" => self.emit(Op::DYN_ADD),
                                "sub" => self.emit(Op::F64_SUB),
                                _ => self.emit(Op::DYN_ADD),
                            }
                            self.emit_var_set(&var);
                        }
                    }
                    self.emit(Op::NULL);
                }
                BuiltinEmit::Intrinsic(intrinsic_name) => {
                    self.emit_intrinsic(intrinsic_name, args)?;
                }
                BuiltinEmit::Common(name) => {
                    // Compile args, then dispatch to compiler_common emitter.
                    // Console.WriteLine/Write should preserve enum names instead
                    // of logging raw ordinals.
                    if name.eq_ignore_ascii_case("dotnet.console_writeline") && args.len() == 1 {
                        self.emit_dotnet_console_arg(args[0])?;
                    } else {
                        for a in args { self.compile_expr(a)?; }
                    }
                    let line = self.line;
                    self.emit_common(name.as_str(), args.len() as u8, line);
                }
                BuiltinEmit::Stdlib(stdlib_name) => {
                    let global_name = Self::stdlib_global_name(stdlib_name);
                    let name_idx = self.str_const(&global_name);
                    let synthetic_args: Vec<Argument> = args
                        .iter()
                        .map(|arg| Argument::positional((**arg).clone()))
                        .collect();
                    let rest_signature = crate::emitter::stdlib::rest_fixed_arity(stdlib_name)
                        .map(|fixed_count| CallSignature {
                            param_names: vec![String::new(); fixed_count as usize + 1],
                            min_arity: fixed_count as usize,
                            has_rest: true,
                        });

                    if let Some(signature) = rest_signature.as_ref() {
                        let callee_slot = self.define_local("__stdlib_fn");
                        self.emit_u16(Op::GLOBAL_GET, name_idx);
                        self.emit_u16(Op::LOCAL_SET, callee_slot);
                        self.emit(Op::DROP);
                        self.emit_known_rest_call_from_local(callee_slot, None, &synthetic_args, signature)?;
                    } else {
                        // Push func ref FIRST, then args, then call_ref.
                        self.emit_u16(Op::GLOBAL_GET, name_idx);
                        for a in args { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL_REF, args.len() as u8);
                    }
                }
                BuiltinEmit::Noop => {
                    self.emit(Op::NULL);
                }
                BuiltinEmit::Invoke(_) => {
                    // `invoke:` is only meaningful for value-method calls
                    // (receiver in hand). In the free-function path the
                    // profile shouldn't use it — emit null so misconfigured
                    // profiles fail loudly via type checks rather than
                    // silent wrong behaviour.
                    self.emit(Op::NULL);
                }
            }
            return Ok(true);
        }

        Ok(false)
    }

    /// Emit a compiler_common operation by namespaced name.
    /// Used by both `BuiltinEmit::Common` paths.
    ///
    /// `argc` is how many caller-supplied arguments are currently on
    /// the stack at the emit site. Multi-arity emits (e.g. .NET
    /// constructors with overloaded shapes) branch on it; most emits
    /// ignore it because their stack contract is fixed.
    fn emit_common(&mut self, name: &str, argc: u8, line: u32) {
        // First try the import-needing dispatch (sleep, etc.). It needs a
        // closure into the compiler to resolve imports against chunk[0].
        // We use a raw pointer to break the borrow of self.
        {
            let self_ptr = self as *mut Self;
            let chunk = self.chunk();
            let handled = common::dispatch::emit_common_with_imports(
                name,
                chunk,
                argc,
                line,
                |module, fname| unsafe { (*self_ptr).import(module, fname) },
            );
            if handled {
                self.sync_scope_slots_with_chunk();
                return;
            }
        }
        // Then the pure (chunk + line) common ops.
        let line2 = line;
        let handled = common::dispatch::emit_common(name, &mut self.chunks, self.current, argc, line2);
        if handled {
            self.sync_scope_slots_with_chunk();
        }
        if !handled {
            eprintln!("Unknown common emit: {}", name);
        }
    }

    fn sync_scope_slots_with_chunk(&mut self) {
        let chunk_slots = self.chunks[self.current].local_count;
        if let Some(scope) = self.scopes.last_mut() {
            if scope.next_slot < chunk_slots {
                scope.next_slot = chunk_slots;
            }
        }
    }

    /// Emit a named opcode sequence for a builtin.
    /// Emit a single opcode by name. Used for value methods where args are already on stack.
    fn emit_named_opcode(&mut self, op_name: &str) {
        let _line = self.line;
        match op_name {
            "f64_abs" => self.emit(Op::F64_ABS),
            "f64_floor" => self.emit(Op::F64_FLOOR),
            "f64_ceil" => self.emit(Op::F64_CEIL),
            "f64_sqrt" => self.emit(Op::F64_SQRT),
            "f64_trunc" => self.emit(Op::F64_TRUNC),
            "f64_nearest" => self.emit(Op::F64_NEAREST),
            "f64_min" => self.emit(Op::F64_MIN),
            "f64_max" => self.emit(Op::F64_MAX),
            "i32_from_f64" => self.emit(Op::I32_FROM_F64),
            "f64_from_i32" => self.emit(Op::F64_FROM_I32),
            "dyn_eq" => self.emit(Op::DYN_EQ),
            "dyn_to_bool" => self.emit(Op::DYN_TO_BOOL),
            "dyn_not" => self.emit(Op::DYN_NOT),
            "ref_is_null" => self.emit(Op::REF_IS_NULL),
            "ref_is_array" => self.emit(Op::REF_IS_ARRAY),
            "ref_typeof" => self.emit(Op::REF_TYPEOF),
            "str_length" => self.emit(Op::STR_LENGTH),
            "str_to_upper" => self.emit(Op::STR_TO_UPPER),
            "str_to_lower" => self.emit(Op::STR_TO_LOWER),
            "str_trim" => self.emit(Op::STR_TRIM),
            "str_trim_start" => self.emit(Op::STR_TRIM_START),
            "str_trim_end" => self.emit(Op::STR_TRIM_END),
            "str_reverse" => self.emit(Op::STR_REVERSE),
            "str_from_char_code" => self.emit(Op::STR_FROM_CHAR_CODE),
            "str_char_at" => self.emit(Op::STR_CHAR_AT),
            "str_char_code_at" => self.emit(Op::STR_CHAR_CODE_AT),
            "str_starts_with" => self.emit(Op::STR_STARTS_WITH),
            "str_ends_with" => self.emit(Op::STR_ENDS_WITH),
            "str_index_of" => self.emit(Op::STR_INDEX_OF),
            "str_last_index_of" => self.emit(Op::STR_LAST_INDEX_OF),
            "str_includes" => {
                // includes → indexOf then check >= 0
                self.emit(Op::STR_INDEX_OF);
                self.emit(Op::I32_CONST_0);
                self.emit(Op::DYN_GE);
            }
            "str_contains" => self.emit(Op::STR_CONTAINS),
            "str_substring" => self.emit(Op::STR_SUBSTRING),
            "str_split" => self.emit(Op::STR_SPLIT),
            "str_replace" => self.emit(Op::STR_REPLACE),
            "str_repeat" => self.emit(Op::STR_REPEAT),
            "str_pad_start" => self.emit(Op::STR_PAD_START),
            "str_pad_end" => self.emit(Op::STR_PAD_END),
            "str_compare" => self.emit(Op::STR_COMPARE),
            "str_concat" => self.emit(Op::STR_CONCAT),
            // Array primitives — every emit flows through
            // `common::collections::*` so the emitted bytecode uses
            // `ecma:array.*` imports. One-place-to-change: flip the
            // provider in collections.rs and every array op in every
            // language re-routes.
            "array_push" => { let l = self.line; common::collections::emit_push(&mut self.chunks, self.current, l); }
            "array_pop" => { let l = self.line; common::collections::emit_pop(&mut self.chunks, self.current, l); }
            "array_shift" => { let l = self.line; common::collections::emit_shift(&mut self.chunks, self.current, l); }
            "array_reverse" => { let l = self.line; common::collections::emit_reverse(&mut self.chunks, self.current, l); }
            "array_join" => { let l = self.line; common::collections::emit_join(&mut self.chunks, self.current, l); }
            "array_concat" => { let l = self.line; common::collections::emit_concat(&mut self.chunks, self.current, l); }
            "array_fill" => { let l = self.line; common::collections::emit_fill(&mut self.chunks, self.current, l); }
            "array_length" => { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
            "array_slice" => { let l = self.line; common::collections::emit_slice(&mut self.chunks, self.current, l); }
            "array_get" => { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
            "array_set" => { let l = self.line; common::collections::emit_set(&mut self.chunks, self.current, l); }
            "array_contains" => { let l = self.line; common::collections::emit_contains(&mut self.chunks, self.current, l); }
            "array_index_of" => { let l = self.line; common::collections::emit_index_of(&mut self.chunks, self.current, l); }
            _ => { let c = self.str_const(op_name); self.emit_u16(Op::GLOBAL_GET, c); }
        }
    }

    fn emit_builtin_opcode(&mut self, op_name: &str, args: &[&Expression]) -> Result<(), String> {
        let line = self.line;
        match op_name {
            "abs" => { self.compile_expr(args[0])?; common::math::emit_abs(self.chunk(), line); }
            "sqrt" => { self.compile_expr(args[0])?; common::math::emit_sqrt(self.chunk(), line); }
            "round" => {
                if args.len() >= 2 {
                    let number = self.import("ecma:number", "Number");
                    let scale_slot = self.define_local("__round_scale");
                    self.emit_const(Value::F64(10.0));
                    self.compile_expr(args[1])?;
                    common::math::emit_pow(self.chunk(), line);
                    self.emit_host_call(number, 1);
                    self.emit_u16(Op::LOCAL_SET, scale_slot); self.emit(Op::DROP);

                    self.compile_expr(args[0])?;
                    self.emit_host_call(number, 1);
                    self.emit_u16(Op::LOCAL_GET, scale_slot);
                    self.emit(Op::F64_MUL);
                    common::math::emit_round(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_GET, scale_slot);
                    self.emit(Op::F64_DIV);
                } else {
                    self.compile_expr(args[0])?;
                    common::math::emit_round(self.chunk(), line);
                }
            }
            "trunc" => { self.compile_expr(args[0])?; common::math::emit_trunc(self.chunk(), line); }
            "floor" => { self.compile_expr(args[0])?; common::math::emit_floor(self.chunk(), line); }
            "ceil" => { self.compile_expr(args[0])?; common::math::emit_ceil(self.chunk(), line); }
            "min" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; common::math::emit_min(self.chunk(), line); } else { self.emit(Op::NULL); } }
            "max" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; common::math::emit_max(self.chunk(), line); } else { self.emit(Op::NULL); } }
            "sqr" => { self.compile_expr(args[0])?; self.emit(Op::DUP); self.emit(Op::F64_MUL); }
            "succ" => { self.compile_expr(args[0])?; self.emit_const(Value::F64(1.0)); self.emit(Op::DYN_ADD); }
            "pred" => { self.compile_expr(args[0])?; self.emit_const(Value::F64(1.0)); self.emit(Op::F64_SUB); }
            "to_upper" => { self.compile_expr(args[0])?; common::strings::emit_to_upper(self.chunk(), line); }
            "to_lower" => { self.compile_expr(args[0])?; common::strings::emit_to_lower(self.chunk(), line); }
            "trim" => { self.compile_expr(args[0])?; common::strings::emit_trim(self.chunk(), line); }
            "str_contains" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    self.emit(Op::STR_CONTAINS);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "str_starts_with" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    self.emit(Op::STR_STARTS_WITH);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "str_ends_with" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    self.emit(Op::STR_ENDS_WITH);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "concat" => { for a in args { self.compile_expr(a)?; } common::strings::emit_concat(self.chunk(), args.len(), line); }
            "replace" => { if args.len() >= 3 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; self.compile_expr(args[2])?; common::strings::emit_replace(self.chunk(), line); } }
            "repeat" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; common::strings::emit_repeat(self.chunk(), line); } }
            "leftstr" => { self.compile_expr(args[0])?; self.emit_const(Value::F64(0.0)); self.compile_expr(args[1])?; common::strings::emit_substring(self.chunk(), line); }
            "high" => { self.compile_expr(args[0])?; common::strings::emit_length(self.chunk(), line); self.emit_const(Value::F64(1.0)); self.emit(Op::F64_SUB); }
            "low" => { self.emit_const(Value::F64(0.0)); }
            "setlength" => {
                if args.len() >= 2 {
                    self.compile_setlength(args[0], args[1])?;
                } else {
                    self.emit(Op::NULL);
                }
            }
            "trim_start" => { self.compile_expr(args[0])?; common::strings::emit_trim_start(self.chunk(), line); }
            "trim_end" => { self.compile_expr(args[0])?; common::strings::emit_trim_end(self.chunk(), line); }
            "pow" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; common::math::emit_pow(self.chunk(), line); } }
            "log" => { self.compile_expr(args[0])?; common::math::emit_log(self.chunk(), line); }
            "sin" => { self.compile_expr(args[0])?; common::math::emit_sin(self.chunk(), line); }
            "cos" => { self.compile_expr(args[0])?; common::math::emit_cos(self.chunk(), line); }
            "tan" => { self.compile_expr(args[0])?; common::math::emit_tan(self.chunk(), line); }
            "exp" => { self.compile_expr(args[0])?; common::math::emit_exp(self.chunk(), line); }
            "is_null" => { self.compile_expr(args[0])?; self.emit(Op::REF_IS_NULL); }
            "space" => { self.emit_const(Value::String(Arc::from(" "))); self.compile_expr(args[0])?; common::strings::emit_repeat(self.chunk(), line); }
            "assigned" => { self.compile_expr(args[0])?; self.emit(Op::REF_IS_NULL); self.emit(Op::DYN_NOT); }
            "freeandnil" => {
                if let Some(first) = args.first() {
                    if let ExprKind::Ident(var) = &first.kind {
                        let var = var.clone();
                        self.emit(Op::NULL);
                        self.emit_var_set(&var);
                    }
                }
                self.emit(Op::NULL);
            }
            // Direct WASM opcode names
            "f64_abs" => { self.compile_expr(args[0])?; self.emit(Op::F64_ABS); }
            "f64_floor" => { self.compile_expr(args[0])?; self.emit(Op::F64_FLOOR); }
            "f64_ceil" => { self.compile_expr(args[0])?; self.emit(Op::F64_CEIL); }
            "f64_sqrt" => { self.compile_expr(args[0])?; self.emit(Op::F64_SQRT); }
            "f64_trunc" => { self.compile_expr(args[0])?; self.emit(Op::F64_TRUNC); }
            "f64_nearest" => { self.compile_expr(args[0])?; self.emit(Op::F64_NEAREST); }
            "f64_min" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; self.emit(Op::F64_MIN); } }
            "f64_max" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; self.emit(Op::F64_MAX); } }
            "i32_from_f64" | "to_int" => {
                self.compile_expr(args[0])?;
                let line = self.line;
                common::convert::emit_to_int(self.chunk(), line);
            }
            "f64_from_i32" => { self.compile_expr(args[0])?; self.emit(Op::F64_FROM_I32); }
            "dyn_to_bool" => { self.compile_expr(args[0])?; self.emit(Op::DYN_TO_BOOL); }
            "ref_is_null" => { self.compile_expr(args[0])?; self.emit(Op::REF_IS_NULL); }
            "ref_is_array" => { self.compile_expr(args[0])?; self.emit(Op::REF_IS_ARRAY); }
            "ref_typeof" => { self.compile_expr(args[0])?; self.emit(Op::REF_TYPEOF); }
            "str_length" => { self.compile_expr(args[0])?; self.emit(Op::STR_LENGTH); }
            "str_to_upper" => { self.compile_expr(args[0])?; self.emit(Op::STR_TO_UPPER); }
            "str_to_lower" => { self.compile_expr(args[0])?; self.emit(Op::STR_TO_LOWER); }
            "str_trim" => { self.compile_expr(args[0])?; self.emit(Op::STR_TRIM); }
            "str_trim_start" => { self.compile_expr(args[0])?; self.emit(Op::STR_TRIM_START); }
            "str_trim_end" => { self.compile_expr(args[0])?; self.emit(Op::STR_TRIM_END); }
            "str_reverse" => { self.compile_expr(args[0])?; self.emit(Op::STR_REVERSE); }
            "str_last_index_of" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    self.emit(Op::STR_LAST_INDEX_OF);
                }
            }
            "str_from_char_code" => {
                // String.fromCharCode(72, 105) → "Hi"
                self.compile_expr(args[0])?;
                common::convert::emit_to_int(self.chunk(), line);
                self.emit(Op::STR_FROM_CHAR_CODE);
                for a in &args[1..] {
                    self.compile_expr(a)?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit(Op::STR_FROM_CHAR_CODE);
                    self.emit(Op::STR_CONCAT);
                }
            }
            "str_compare" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; self.emit(Op::STR_COMPARE); } }
            "str_split" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; self.emit(Op::STR_SPLIT); } }
            "str_getcsv" => {
                for arg in args {
                    self.compile_expr(arg)?;
                }
                crate::languages::php::emitter::string_adapter::emit_str_getcsv(
                    &mut self.chunks,
                    self.current,
                    args.len() as u8,
                    line,
                );
            }
            "str_repeat" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; self.emit(Op::STR_REPEAT); } }
            "array_join" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    { let l = self.line; common::collections::emit_join(&mut self.chunks, self.current, l); }
                }
            }
            // setTimeout/setInterval — emit Op::SET_TIMER directly. Old JS
            // compiler did this inline; the profile now routes through
            // `opcode:set_timer` so the dispatch lives here.
            // Stack: [callback, ms] → [timer_id]
            "set_timer" => {
                if let Some(cb) = args.first() {
                    self.compile_expr(cb)?;
                } else {
                    self.emit(Op::NULL);
                }
                if args.len() >= 2 {
                    self.compile_expr(args[1])?;
                } else {
                    self.emit(Op::I32_CONST_0);
                }
                self.emit(Op::SET_TIMER);
            }
            // Array primitives — every caller dispatches through
            // `common::collections::*`, which now routes to `ecma:array.*`
            // imports (Phase D). Keep the arg-evaluation and stack shape
            // details here; the emit itself lives in compiler_common so
            // the identical surface is used by every language.
            "array_length" => {
                if let Some(first) = args.first() {
                    self.compile_expr(first)?;
                    common::collections::emit_len(&mut self.chunks, self.current, line);
                } else {
                    self.emit_const(Value::I32(0));
                }
            }
            "array_push" => {
                // PHP `array_push($a, v1, v2, ...)` — push each value.
                // Returns the new length (of the last push).
                if let Some(arr) = args.first() {
                    if args.len() == 1 {
                        self.compile_expr(arr)?;
                        common::collections::emit_len(&mut self.chunks, self.current, line);
                    } else {
                        let tail = args.len() - 1;
                        for (i, v) in args[1..].iter().enumerate() {
                            self.compile_expr(arr)?;
                            self.compile_expr(v)?;
                            common::collections::emit_push(&mut self.chunks, self.current, line);
                            // Drop intermediate lengths; the final one
                            // is the expression's value.
                            if i != tail - 1 { self.emit(Op::DROP); }
                        }
                    }
                } else {
                    self.emit_const(Value::I32(0));
                }
            }
            "array_pop" => {
                if let Some(first) = args.first() {
                    self.compile_expr(first)?;
                    common::collections::emit_pop(&mut self.chunks, self.current, line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "array_shift" => {
                if let Some(first) = args.first() {
                    self.compile_expr(first)?;
                    common::collections::emit_shift(&mut self.chunks, self.current, line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "array_reverse" => {
                if let Some(first) = args.first() {
                    self.compile_expr(first)?;
                    common::collections::emit_reverse(&mut self.chunks, self.current, line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "array_concat" => {
                if args.is_empty() {
                    common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                } else {
                    self.compile_expr(args[0])?;
                    for v in &args[1..] {
                        self.compile_expr(v)?;
                        common::collections::emit_concat(&mut self.chunks, self.current, line);
                    }
                }
            }
            "array_index_of" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::collections::emit_index_of(&mut self.chunks, self.current, line);
                } else {
                    self.emit_const(Value::I32(-1));
                }
            }
            // PHP `in_array($needle, $haystack)` — walker already normalized
            // arg order to [haystack, needle, strict?] matching JS's
            // `arr.includes(needle, fromIndex?)`. emit_contains calls
            // `ecma:array.includes` which is polymorphic over Array,
            // Map, and Ordinary, so PHP's `in_array` works uniformly on
            // assoc arrays, indexed arrays, and superglobals.
            "array_contains" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::collections::emit_contains(&mut self.chunks, self.current, line);
                } else {
                    self.emit(Op::FALSE);
                }
            }
            _ => { self.emit(Op::NULL); }
        }
        Ok(())
    }

    /// Emit a multi-opcode intrinsic sequence.
    fn emit_fortran_scan_like(&mut self, args: &[&Expression], invert_match: bool) -> Result<(), String> {
        let line = self.line;
        if args.len() < 2 {
            self.emit(Op::NULL);
            return Ok(());
        }

        let source_slot = self.define_local("__fortran_scan_source");
        let set_slot = self.define_local("__fortran_scan_set");
        let back_slot = self.define_local("__fortran_scan_back");
        let len_slot = self.define_local("__fortran_scan_len");
        let index_slot = self.define_local("__fortran_scan_index");

        self.compile_expr(args[0])?;
        self.emit_u16(Op::LOCAL_SET, source_slot);
        self.emit(Op::DROP);

        self.compile_expr(args[1])?;
        self.emit_u16(Op::LOCAL_SET, set_slot);
        self.emit(Op::DROP);

        if let Some(back_arg) = args.get(2) {
            self.compile_expr(back_arg)?;
            self.emit(Op::DYN_TO_BOOL);
        } else {
            self.emit(Op::FALSE);
        }
        self.emit_u16(Op::LOCAL_SET, back_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, source_slot);
        common::strings::emit_length(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, len_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, back_slot);
        let forward_path = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, len_slot);
        self.emit(Op::I32_CONST_1);
        self.emit(Op::I32_SUB);
        self.emit_u16(Op::LOCAL_SET, index_slot);
        self.emit(Op::DROP);

        let back_loop = self.current_offset();
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit(Op::I32_CONST_0);
        self.emit(Op::DYN_LT);
        let back_not_found = self.emit_jump(Op::BR_IF_TRUE);

        self.emit_u16(Op::LOCAL_GET, set_slot);
        self.emit_u16(Op::LOCAL_GET, source_slot);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit(Op::STR_CHAR_AT);
        self.emit(Op::STR_CONTAINS);
        if invert_match {
            self.emit(Op::DYN_NOT);
        }
        let back_continue = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_const(Value::I32(1));
        self.emit(Op::I32_ADD);
        let back_done = self.emit_jump(Op::BR);

        self.patch_jump(back_continue);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit(Op::I32_CONST_1);
        self.emit(Op::I32_SUB);
        self.emit_u16(Op::LOCAL_SET, index_slot);
        self.emit(Op::DROP);
        self.emit_loop(back_loop);

        self.patch_jump(back_not_found);
        self.emit(Op::I32_CONST_0);
        let back_zero_done = self.emit_jump(Op::BR);

        self.patch_jump(forward_path);
        self.emit(Op::I32_CONST_0);
        self.emit_u16(Op::LOCAL_SET, index_slot);
        self.emit(Op::DROP);

        let forward_loop = self.current_offset();
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_u16(Op::LOCAL_GET, len_slot);
        self.emit(Op::DYN_GE);
        let forward_not_found = self.emit_jump(Op::BR_IF_TRUE);

        self.emit_u16(Op::LOCAL_GET, set_slot);
        self.emit_u16(Op::LOCAL_GET, source_slot);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit(Op::STR_CHAR_AT);
        self.emit(Op::STR_CONTAINS);
        if invert_match {
            self.emit(Op::DYN_NOT);
        }
        let forward_continue = self.emit_jump(Op::BR_IF_FALSE);

        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit_const(Value::I32(1));
        self.emit(Op::I32_ADD);
        let forward_done = self.emit_jump(Op::BR);

        self.patch_jump(forward_continue);
        self.emit_u16(Op::LOCAL_GET, index_slot);
        self.emit(Op::I32_CONST_1);
        self.emit(Op::I32_ADD);
        self.emit_u16(Op::LOCAL_SET, index_slot);
        self.emit(Op::DROP);
        self.emit_loop(forward_loop);

        self.patch_jump(forward_not_found);
        self.emit(Op::I32_CONST_0);

        self.patch_jump(back_done);
        self.patch_jump(back_zero_done);
        self.patch_jump(forward_done);
        Ok(())
    }

    fn emit_intrinsic(&mut self, name: &str, args: &[&Expression]) -> Result<(), String> {
        let line = self.line;
        match name {
            "cstr" => {
                self.compile_expr(args[0])?;
                let value_slot = self.define_local("__vb_cstr_value");
                self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.emit(Op::REF_TYPEOF);
                self.emit_const(Value::String(Arc::from("boolean")));
                self.emit(Op::DYN_EQ);
                let not_bool = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.emit(Op::DYN_TO_BOOL);
                let false_path = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_const(Value::String(Arc::from("True")));
                let done = self.emit_jump(Op::BR);
                self.patch_jump(false_path);
                self.emit_const(Value::String(Arc::from("False")));
                self.patch_jump(done);

                let end = self.emit_jump(Op::BR);
                self.patch_jump(not_bool);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                let string_idx = self.import("ecma:string", "String");
                self.emit_host_call(string_idx, 1);
                self.patch_jump(end);
            }
            "cbyte" => {
                self.compile_expr(args[0])?;
                common::convert::emit_to_int(self.chunk(), line);
                self.emit_const(Value::I32(0xFF));
                self.emit(Op::I32_AND);
            }
            "ubound" => {
                self.compile_expr(args[0])?;
                common::collections::emit_len(&mut self.chunks, self.current, line);
                self.emit_const(Value::I32(1));
                self.emit(Op::I32_SUB);
            }
            "lbound" => {
                self.emit(Op::I32_CONST_0);
            }
            "erase" => {
                // VB `Erase arr` — releases / clears the array contents. For
                // dynamic arrays, real VB frees the storage and leaves the
                // variable referring to an uninitialised array; for
                // fixed-size arrays, it re-zeros each element. We return a
                // fresh empty array, which satisfies both reads (`.Length`
                // works, yields 0) and assignment (`arr = Erase(arr)`).
                //
                // The arg is still compiled for any side effects and then
                // dropped — matches the VB semantic that the old binding is
                // released.
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    self.emit(Op::DROP);
                }
                let l = self.line;
                common::collections::emit_array_new(&mut self.chunks, self.current, 0, l);
            }
            "asc" => {
                self.compile_expr(args[0])?;
                self.emit(Op::I32_CONST_0);
                self.emit(Op::STR_CHAR_CODE_AT);
            }
            "space" => {
                self.emit_const(Value::String(Arc::from(" ")));
                self.compile_expr(args[0])?;
                common::convert::emit_to_int(self.chunk(), line);
                common::strings::emit_repeat(self.chunk(), line);
            }
            "isobject" => {
                if let Some(arg) = args.first() {
                    if let Some(result) = self.vb_is_object_expr(arg) {
                        self.emit_const(Value::Bool(result));
                    } else {
                        self.compile_expr(arg)?;
                        let value_slot = self.define_local("__vb_isobject_value");
                        self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.emit(Op::REF_IS_ARRAY);
                        let is_array = self.emit_jump(Op::BR_IF_TRUE);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.emit(Op::REF_IS_OBJECT);
                        let done = self.emit_jump(Op::BR);
                        self.patch_jump(is_array);
                        self.emit_const(Value::Bool(true));
                        self.patch_jump(done);
                    }
                } else {
                    self.emit_const(Value::Bool(false));
                }
            }
            "isreference" => {
                if let Some(arg) = args.first() {
                    if let Some(result) = self.vb_is_reference_expr(arg) {
                        self.emit_const(Value::Bool(result));
                    } else {
                        self.compile_expr(arg)?;
                        let value_slot = self.define_local("__vb_isref_value");
                        self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.emit(Op::REF_IS_STRING);
                        let is_string = self.emit_jump(Op::BR_IF_TRUE);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.emit(Op::REF_IS_ARRAY);
                        let is_array = self.emit_jump(Op::BR_IF_TRUE);
                        self.emit_u16(Op::LOCAL_GET, value_slot);
                        self.emit(Op::REF_IS_OBJECT);
                        let done = self.emit_jump(Op::BR);
                        self.patch_jump(is_string);
                        self.emit_const(Value::Bool(true));
                        let end_true = self.emit_jump(Op::BR);
                        self.patch_jump(is_array);
                        self.emit_const(Value::Bool(true));
                        self.patch_jump(end_true);
                        self.patch_jump(done);
                    }
                } else {
                    self.emit_const(Value::Bool(false));
                }
            }
            "typename" => {
                if let Some(arg) = args.first() {
                    if let Some(name) = self.vb_typename_from_expr(arg) {
                        self.emit_const(Value::String(Arc::from(name)));
                    } else {
                        self.compile_expr(arg)?;
                        self.emit(Op::REF_TYPEOF);
                    }
                } else {
                    self.emit_const(Value::String(Arc::from("Nothing")));
                }
            }
            "command" => {
                let args_idx = self.import("wasi:cli", "args");
                self.emit_host_call(args_idx, 0);
                self.emit_const(Value::String(Arc::from(" ")));
                common::collections::emit_join(&mut self.chunks, self.current, line);
            }
            "environ" => {
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    let env_idx = self.import("wasi:cli", "getEnv");
                    self.emit_host_call(env_idx, 1);
                } else {
                    self.emit_const(Value::String(Arc::from("")));
                }
            }
            "timer" => {
                let timer_idx = self.import("wasi:clocks", "vbTimer");
                self.emit_host_call(timer_idx, 0);
                let value_slot = self.define_local("__vb_timer_value");
                self.emit_u16(Op::LOCAL_SET, value_slot); self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.emit(Op::REF_IS_NUMBER);
                let not_number = self.emit_jump(Op::BR_IF_FALSE);

                self.emit_u16(Op::LOCAL_GET, value_slot);
                self.emit_const(Value::I32(0));
                self.emit(Op::DYN_LT);
                let non_negative = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_const(Value::I32(0));
                let done = self.emit_jump(Op::BR);

                self.patch_jump(non_negative);
                self.emit_u16(Op::LOCAL_GET, value_slot);
                let end = self.emit_jump(Op::BR);

                self.patch_jump(not_number);
                self.emit_const(Value::I32(0));
                self.patch_jump(done);
                self.patch_jump(end);
            }
            "switch" => {
                if args.len() < 2 {
                    self.emit(Op::NULL);
                } else {
                    let mut slots = Vec::with_capacity(args.len());
                    for (index, arg) in args.iter().enumerate() {
                        self.compile_expr(arg)?;
                        let slot = self.define_local(&format!("__vb_switch_{index}"));
                        self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                        slots.push(slot);
                    }

                    let mut done_jumps = Vec::new();
                    for pair in slots.chunks(2) {
                        if pair.len() < 2 {
                            break;
                        }
                        self.emit_u16(Op::LOCAL_GET, pair[0]);
                        self.emit(Op::DYN_TO_BOOL);
                        let next = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit_u16(Op::LOCAL_GET, pair[1]);
                        done_jumps.push(self.emit_jump(Op::BR));
                        self.patch_jump(next);
                    }
                    self.emit(Op::NULL);
                    for jump in done_jumps {
                        self.patch_jump(jump);
                    }
                }
            }
            "string_repeat" => {
                // String(n, char): VB arg order reversed
                if args.len() >= 2 {
                    self.compile_expr(args[1])?;
                    self.compile_expr(args[0])?;
                    common::strings::emit_repeat(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "left" => {
                // Left(s, n) → substring(s, 0, n)
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.emit(Op::I32_CONST_0);
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "php_is_float" => {
                // PHP `is_float` — true only for non-integer numbers.
                // Composes ecma:number.isInteger + boolean negation
                // with a leading `typeof v === "number"` guard so
                // strings / objects don't match (REF_IS_NUMBER opcode
                // covers the typeof-number predicate).
                if !args.is_empty() {
                    self.compile_expr(args[0])?;
                    let v_slot = self.define_local("__php_isf_v");
                    self.emit_u16(Op::LOCAL_SET, v_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, v_slot);
                    self.emit(Op::REF_IS_NUMBER);
                    let not_num = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit_u16(Op::LOCAL_GET, v_slot);
                    let is_int_idx = self.import("ecma:number", "isInteger");
                    self.emit_host_call(is_int_idx, 1);
                    self.emit(Op::DYN_NOT);
                    let done = self.emit_jump(Op::BR);
                    self.patch_jump(not_num);
                    self.emit_const(Value::Bool(false));
                    self.patch_jump(done);
                } else {
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_is_string" => {
                if !args.is_empty() {
                    self.compile_expr(args[0])?;
                    self.emit(Op::REF_IS_STRING);
                } else {
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_is_array" => {
                // PHP `is_array` matches any of: ObjectKind::Array,
                // ObjectKind::Map, ObjectKind::Ordinary (plain assoc
                // object). REF_IS_ARRAY only checks Array; we layer
                // an Object check via REF_IS_OBJECT (covers Map and
                // Ordinary too — both are Object-kind values).
                if !args.is_empty() {
                    self.compile_expr(args[0])?;
                    self.emit(Op::REF_IS_OBJECT);
                } else {
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_is_bool" => {
                if !args.is_empty() {
                    self.compile_expr(args[0])?;
                    self.emit(Op::REF_IS_BOOL);
                } else {
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_is_null" => {
                if !args.is_empty() {
                    self.compile_expr(args[0])?;
                    self.emit(Op::REF_IS_NULL);
                } else {
                    self.emit_const(Value::Bool(true));
                }
            }
            "php_is_object" => {
                // PHP `is_object` matches user objects but NOT plain
                // arrays. Approximated as REF_IS_OBJECT && !is_array.
                // For Phase-1 simplicity the same predicate as is_array
                // — distinction requires a class-instance vs assoc-array
                // tag which Vybe doesn't track yet.
                if !args.is_empty() {
                    self.compile_expr(args[0])?;
                    self.emit(Op::REF_IS_OBJECT);
                } else {
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_defined" => {
                if let Some(Expression { kind: ExprKind::Lit(Literal::Str(name)), .. }) = args.first() {
                    let global_name = self.canon(name);
                    let idx = self.str_const(&global_name);
                    self.emit_u16(Op::GLOBAL_GET, idx);
                    self.emit(Op::REF_TYPEOF);
                    self.emit_const(Value::String(Arc::from("undefined")));
                    self.emit(Op::DYN_EQ);
                    self.emit(Op::DYN_NOT);
                } else {
                    if let Some(arg) = args.first() {
                        self.compile_expr(arg)?;
                        self.emit(Op::DROP);
                    }
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_function_exists" => {
                if let Some(Expression { kind: ExprKind::Lit(Literal::Str(name)), .. }) = args.first() {
                    let builtin_exists = self.profile.lookup_builtin(name).is_some()
                        || common::imports::resolve_common_import(name).is_some();
                    if builtin_exists {
                        self.emit_const(Value::Bool(true));
                    } else {
                        let global_name = self.canon(name);
                        let idx = self.str_const(&global_name);
                        self.emit_u16(Op::GLOBAL_GET, idx);
                        self.emit_const(Value::Undefined);
                        self.emit(Op::DYN_EQ);
                        self.emit(Op::DYN_NOT);
                    }
                } else {
                    if let Some(arg) = args.first() {
                        self.compile_expr(arg)?;
                        let name_slot = self.define_local("__php_function_exists_name");
                        self.emit_u16(Op::LOCAL_SET, name_slot);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, name_slot);
                        self.emit(Op::REF_TYPEOF);
                        self.emit_const(Value::String(Arc::from("string")));
                        self.emit(Op::DYN_EQ);
                        let not_string = self.emit_jump(Op::BR_IF_FALSE);

                        self.emit_u16(Op::LOCAL_GET, name_slot);
                        self.emit(Op::STR_TO_LOWER);
                        let lowered_slot = self.define_local("__php_function_exists_lowered");
                        self.emit_u16(Op::LOCAL_SET, lowered_slot);
                        self.emit(Op::DROP);

                        let mut known_functions: Vec<String> = self.defined_functions.iter().cloned().collect();
                        known_functions.sort();
                        let mut done_jumps = Vec::new();
                        for function_name in known_functions {
                            self.emit_u16(Op::LOCAL_GET, lowered_slot);
                            self.emit_const(Value::String(Arc::from(function_name.to_ascii_lowercase())));
                            self.emit(Op::DYN_EQ);
                            let next = self.emit_jump(Op::BR_IF_FALSE);
                            self.emit_const(Value::Bool(true));
                            done_jumps.push(self.emit_jump(Op::BR));
                            self.patch_jump(next);
                        }

                        self.patch_jump(not_string);
                        self.emit_const(Value::Bool(false));
                        for done in done_jumps {
                            self.patch_jump(done);
                        }
                    } else {
                        self.emit_const(Value::Bool(false));
                    }
                }
            }
            "php_class_exists" => {
                if let Some(Expression { kind: ExprKind::Lit(Literal::Str(name)), .. }) = args.first() {
                    for arg in args.iter().skip(1) {
                        self.compile_expr(arg)?;
                        self.emit(Op::DROP);
                    }
                    let global_name = self.canon(name);
                    let idx = self.str_const(&global_name);
                    self.emit_u16(Op::GLOBAL_GET, idx);
                    self.emit(Op::REF_TYPEOF);
                    self.emit_const(Value::String(Arc::from("undefined")));
                    self.emit(Op::DYN_EQ);
                    self.emit(Op::DYN_NOT);
                } else {
                    if let Some(arg) = args.first() {
                        self.compile_expr(arg)?;
                        self.emit(Op::DROP);
                    }
                    for arg in args.iter().skip(1) {
                        self.compile_expr(arg)?;
                        self.emit(Op::DROP);
                    }
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_define" => {
                if args.len() < 2 {
                    self.emit_const(Value::Bool(false));
                } else if let ExprKind::Lit(Literal::Str(name)) = &args[0].kind {
                    if let Some(ignore_case) = args.get(2) {
                        self.compile_expr(ignore_case)?;
                        self.emit(Op::DROP);
                    }
                    self.compile_expr(args[1])?;
                    let global_name = self.canon(name);
                    let idx = self.str_const(&global_name);
                    self.emit_u16(Op::GLOBAL_SET, idx);
                    self.emit(Op::DROP);
                    self.defined_globals.insert(global_name);
                    self.emit_const(Value::Bool(true));
                } else {
                    self.compile_expr(args[0])?;
                    self.emit(Op::DROP);
                    self.compile_expr(args[1])?;
                    self.emit(Op::DROP);
                    if let Some(ignore_case) = args.get(2) {
                        self.compile_expr(ignore_case)?;
                        self.emit(Op::DROP);
                    }
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_is_callable" => {
                // PHP `is_callable` matches functions and Closure
                // instances. ref_typeof on Function / HostFunction
                // returns "function" — compare via DYN_EQ.
                if !args.is_empty() {
                    self.compile_expr(args[0])?;
                    self.emit(Op::REF_TYPEOF);
                    self.emit_const(Value::String(Arc::from("function")));
                    self.emit(Op::DYN_EQ);
                } else {
                    self.emit_const(Value::Bool(false));
                }
            }
            "php_version_compare" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.emit_common("dotnet.version_parse", 1, line);
                    self.compile_expr(args[1])?;
                    self.emit_common("dotnet.version_parse", 1, line);
                    self.emit_common("dotnet.version_compare", 2, line);

                    let cmp_slot = self.define_local("__php_version_compare_cmp");
                    self.emit_u16(Op::LOCAL_SET, cmp_slot);
                    self.emit(Op::DROP);

                    if let Some(operator) = args.get(2) {
                        let op_slot = self.define_local("__php_version_compare_op");
                        self.compile_expr(operator)?;
                        self.emit_u16(Op::LOCAL_SET, op_slot);
                        self.emit(Op::DROP);

                        let mut done_jumps = Vec::new();
                        for (op_text, compare_op) in [
                            ("<", Op::DYN_LT),
                            ("lt", Op::DYN_LT),
                            ("<=", Op::DYN_LE),
                            ("le", Op::DYN_LE),
                            (">", Op::DYN_GT),
                            ("gt", Op::DYN_GT),
                            (">=", Op::DYN_GE),
                            ("ge", Op::DYN_GE),
                            ("==", Op::DYN_EQ),
                            ("=", Op::DYN_EQ),
                            ("eq", Op::DYN_EQ),
                            ("!=", Op::DYN_NE),
                            ("<>", Op::DYN_NE),
                            ("ne", Op::DYN_NE),
                        ] {
                            self.emit_u16(Op::LOCAL_GET, op_slot);
                            self.emit_const(Value::String(Arc::from(op_text)));
                            self.emit(Op::DYN_EQ);
                            let next_check = self.emit_jump(Op::BR_IF_FALSE);
                            self.emit_u16(Op::LOCAL_GET, cmp_slot);
                            self.emit_const(Value::F64(0.0));
                            self.emit(compare_op);
                            done_jumps.push(self.emit_jump(Op::BR));
                            self.patch_jump(next_check);
                        }

                        self.emit_const(Value::Bool(false));
                        for jump in done_jumps {
                            self.patch_jump(jump);
                        }
                    } else {
                        self.emit_u16(Op::LOCAL_GET, cmp_slot);
                    }
                } else {
                    self.emit_const(Value::I32(0));
                }
            }
            "php_printf" => {
                if args.is_empty() {
                    self.emit_const(Value::I32(0));
                } else {
                    let global_name = Self::stdlib_global_name("sprintf");
                    let name_idx = self.str_const(&global_name);
                    let result_slot = self.define_local("__php_printf_result");
                    let log_idx = self.import("wasi:cli", "log");
                    let callee_slot = self.define_local("__php_printf_fn");
                    let synthetic_args: Vec<Argument> = args
                        .iter()
                        .map(|arg| Argument::positional((**arg).clone()))
                        .collect();
                    let signature = CallSignature {
                        param_names: vec![String::new(), String::new()],
                        min_arity: 1,
                        has_rest: true,
                    };

                    self.emit_u16(Op::GLOBAL_GET, name_idx);
                    self.emit_u16(Op::LOCAL_SET, callee_slot);
                    self.emit(Op::DROP);
                    self.emit_known_rest_call_from_local(callee_slot, None, &synthetic_args, &signature)?;
                    self.emit_u16(Op::LOCAL_SET, result_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.emit_common("php.echo_stringify", 1, line);
                    common::io::emit_print_with_import(self.chunk(), log_idx, 1, line);

                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    common::strings::emit_length(self.chunk(), line);
                }
            }
            "php_register_shutdown_function" => {
                for arg in args {
                    self.compile_expr(arg)?;
                    self.emit(Op::DROP);
                }
                self.emit(Op::NULL);
            }
            "php_date_default_timezone_set" => {
                for arg in args {
                    self.compile_expr(arg)?;
                    self.emit(Op::DROP);
                }
                self.emit_const(Value::Bool(true));
            }
            "php_rsort" => {
                // PHP `rsort($arr)` — descending in-place sort. Compose
                // from existing stdlib: `sort_in_place(arr)` for the
                // ascending sort, then `array_reverse` for descending.
                // PHP arrays are JS arrays in our model, so the sort +
                // reverse mutate the same backing storage the caller's
                // variable points to.
                if !args.is_empty() {
                    self.compile_expr(args[0])?;
                    let arr_slot = self.define_local("__php_rsort_arr");
                    self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);
                    let helper = self.str_const("__vybe_sort_in_place");
                    self.emit_u16(Op::GLOBAL_GET, helper);
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit_u8(Op::CALL_REF, 1);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    common::collections::emit_reverse(&mut self.chunks, self.current, line);
                    self.emit(Op::DROP);
                    self.emit(Op::NULL);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "right" => {
                // Right(s, n) → substring(s, len(s) - n, len(s))
                // Direct opcodes — no host call. Mirrors the `left`
                // intrinsic shape; goes through `common::strings`
                // emitters so the underlying provider (str_substring
                // opcode) stays the single source of truth.
                if args.len() >= 2 {
                    // Stash s and n in scratch slots so we can use len(s)
                    // and n twice (compute start = len - n, end = len).
                    let s_slot = self.define_local("__right_s");
                    let n_slot = self.define_local("__right_n");
                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, s_slot); self.emit(Op::DROP);
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, n_slot); self.emit(Op::DROP);
                    // substring(s, len(s) - n, len(s))
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    // start = len(s) - n
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    common::strings::emit_length(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_GET, n_slot);
                    self.emit(Op::I32_SUB);
                    let start_slot = self.define_local("__right_start");
                    self.emit_u16(Op::LOCAL_SET, start_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    self.emit(Op::I32_CONST_0);
                    self.emit(Op::DYN_LT);
                    let keep_start = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit(Op::I32_CONST_0);
                    let after_start = self.emit_jump(Op::BR);
                    self.patch_jump(keep_start);
                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    self.patch_jump(after_start);
                    // end = len(s)
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    common::strings::emit_length(self.chunk(), line);
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "php_substr" => {
                if args.len() >= 2 {
                    let str_slot = self.define_local("__php_substr_s");
                    let start_slot = self.define_local("__php_substr_start");

                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, str_slot); self.emit(Op::DROP);

                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, start_slot); self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, str_slot);
                    self.emit_u16(Op::LOCAL_GET, start_slot);

                    if args.len() >= 3 {
                        self.emit_u16(Op::LOCAL_GET, start_slot);
                        self.compile_expr(args[2])?;
                        common::convert::emit_to_int(self.chunk(), line);
                        self.emit(Op::I32_ADD);
                    } else {
                        self.emit_const(Value::I32(i32::MAX));
                    }

                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "php_strpos" => {
                if args.len() >= 2 {
                    let haystack_slot = self.define_local("__php_strpos_haystack");
                    let needle_slot = self.define_local("__php_strpos_needle");
                    let offset_slot = self.define_local("__php_strpos_offset");
                    let idx_slot = self.define_local("__php_strpos_idx");

                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, haystack_slot); self.emit(Op::DROP);

                    self.compile_expr(args[1])?;
                    self.emit_u16(Op::LOCAL_SET, needle_slot); self.emit(Op::DROP);

                    if args.len() >= 3 {
                        self.compile_expr(args[2])?;
                        common::convert::emit_to_int(self.chunk(), line);
                    } else {
                        self.emit(Op::I32_CONST_0);
                    }
                    self.emit_u16(Op::LOCAL_SET, offset_slot); self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, haystack_slot);
                    self.emit_u16(Op::LOCAL_GET, offset_slot);
                    self.emit_const(Value::I32(i32::MAX));
                    common::strings::emit_substring(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_GET, needle_slot);
                    common::strings::emit_index_of(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit(Op::I32_CONST_0);
                    self.emit(Op::DYN_LT);
                    let found = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit(Op::FALSE);
                    let done = self.emit_jump(Op::BR);
                    self.patch_jump(found);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit_u16(Op::LOCAL_GET, offset_slot);
                    self.emit(Op::I32_ADD);
                    self.patch_jump(done);
                } else {
                    self.emit(Op::FALSE);
                }
            }
            "php_str_contains" => {
                if args.len() >= 2 {
                    let tostring_global = self.str_const("__vybe_tostring");
                    self.emit_u16(Op::GLOBAL_GET, tostring_global);
                    self.compile_expr(args[0])?;
                    self.emit_u8(Op::CALL_REF, 1);

                    self.emit_u16(Op::GLOBAL_GET, tostring_global);
                    self.compile_expr(args[1])?;
                    self.emit_u8(Op::CALL_REF, 1);

                    self.emit(Op::STR_CONTAINS);
                } else {
                    self.emit(Op::FALSE);
                }
            }
            "php_str_starts_with" => {
                if args.len() >= 2 {
                    let tostring_global = self.str_const("__vybe_tostring");
                    self.emit_u16(Op::GLOBAL_GET, tostring_global);
                    self.compile_expr(args[0])?;
                    self.emit_u8(Op::CALL_REF, 1);

                    self.emit_u16(Op::GLOBAL_GET, tostring_global);
                    self.compile_expr(args[1])?;
                    self.emit_u8(Op::CALL_REF, 1);

                    self.emit(Op::STR_STARTS_WITH);
                } else {
                    self.emit(Op::FALSE);
                }
            }
            "php_str_ends_with" => {
                if args.len() >= 2 {
                    let tostring_global = self.str_const("__vybe_tostring");
                    self.emit_u16(Op::GLOBAL_GET, tostring_global);
                    self.compile_expr(args[0])?;
                    self.emit_u8(Op::CALL_REF, 1);

                    self.emit_u16(Op::GLOBAL_GET, tostring_global);
                    self.compile_expr(args[1])?;
                    self.emit_u8(Op::CALL_REF, 1);

                    self.emit(Op::STR_ENDS_WITH);
                } else {
                    self.emit(Op::FALSE);
                }
            }
            "php_array_search" => {
                if args.len() >= 2 {
                    let idx_slot = self.define_local("__php_array_search_idx");
                    self.compile_expr(args[1])?;
                    self.compile_expr(args[0])?;
                    common::collections::emit_index_of(&mut self.chunks, self.current, line);
                    self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit(Op::I32_CONST_0);
                    self.emit(Op::DYN_LT);
                    let found = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit(Op::FALSE);
                    let done = self.emit_jump(Op::BR);
                    self.patch_jump(found);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.patch_jump(done);
                } else {
                    self.emit(Op::FALSE);
                }
            }
            "php_array_slice" => {
                if args.len() >= 2 {
                    let arr_slot = self.define_local("__php_array_slice_arr");
                    let start_slot = self.define_local("__php_array_slice_start");

                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);

                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, start_slot); self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit_u16(Op::LOCAL_GET, start_slot);

                    if args.len() >= 3 {
                        self.emit_u16(Op::LOCAL_GET, start_slot);
                        self.compile_expr(args[2])?;
                        common::convert::emit_to_int(self.chunk(), line);
                        self.emit(Op::I32_ADD);
                    } else {
                        self.emit_const(Value::I32(i32::MAX));
                    }

                    common::collections::emit_slice(&mut self.chunks, self.current, line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "php_range" => {
                if args.len() >= 2 {
                    let start_slot = self.define_local("__php_range_start");
                    let end_slot = self.define_local("__php_range_end");
                    let step_slot = self.define_local("__php_range_step");
                    let stop_slot = self.define_local("__php_range_stop");

                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, start_slot); self.emit(Op::DROP);

                    self.compile_expr(args[1])?;
                    self.emit_u16(Op::LOCAL_SET, end_slot); self.emit(Op::DROP);

                    if args.len() >= 3 {
                        self.compile_expr(args[2])?;
                    } else {
                        self.emit(Op::I32_CONST_1);
                    }
                    self.emit_u16(Op::LOCAL_SET, step_slot); self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, step_slot);
                    self.emit(Op::I32_CONST_0);
                    self.emit(Op::DYN_LT);
                    let positive_step = self.emit_jump(Op::BR_IF_FALSE);

                    self.emit_u16(Op::LOCAL_GET, end_slot);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_SUB);
                    let stop_done = self.emit_jump(Op::BR);

                    self.patch_jump(positive_step);
                    self.emit_u16(Op::LOCAL_GET, end_slot);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_ADD);

                    self.patch_jump(stop_done);
                    self.emit_u16(Op::LOCAL_SET, stop_slot); self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    self.emit_u16(Op::LOCAL_GET, stop_slot);
                    self.emit_u16(Op::LOCAL_GET, step_slot);
                    common::collections::emit_range(&mut self.chunks, self.current, 3, line);
                } else {
                    common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                }
            }
            "php_print_expr" => {
                if let Some(arg) = args.first() {
                    let log_idx = self.import("wasi:cli", "log");
                    self.compile_expr(arg)?;
                    self.emit_common("php.echo_stringify", 1, line);
                    common::io::emit_print_with_import(self.chunk(), log_idx, 1, line);
                    self.emit(Op::I32_CONST_1);
                } else {
                    self.emit(Op::I32_CONST_1);
                }
            }
            "string_isnullorempty" => {
                // String.IsNullOrEmpty(s) → s is null OR str_length(s) == 0.
                // Compile s, dup, ref_is_null → if true return true, else
                // str_length == 0.
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    // [s]
                    self.emit(Op::DUP);
                    // [s, s]
                    self.emit(Op::REF_IS_NULL);
                    // [s, is_null]
                    let if_null = self.emit_jump(Op::BR_IF_TRUE);
                    // not null branch: [s] → str_length → cmp 0
                    common::strings::emit_length(self.chunk(), line);
                    self.emit(Op::I32_CONST_0);
                    self.emit(Op::DYN_EQ);
                    let end = self.emit_jump(Op::BR);
                    // null branch: drop [s], push true
                    self.patch_jump(if_null);
                    self.emit(Op::DROP);
                    self.emit(Op::TRUE);
                    self.patch_jump(end);
                } else {
                    self.emit(Op::TRUE);
                }
            }
            "mid" | "mid_1based" => {
                // Mid(s, start[, len]) — 1-based
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_SUB); // start0
                    if args.len() >= 3 {
                        self.emit(Op::DUP);
                        self.compile_expr(args[2])?;
                        common::convert::emit_to_int(self.chunk(), line);
                        self.emit(Op::I32_ADD); // start0 + length
                    } else {
                        self.emit_const(Value::I32(0x7FFF_FFFF));
                    }
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "number_isnan" => {
                self.compile_expr(args[0])?;
                self.emit(Op::DUP);
                self.emit(Op::DYN_NE);
            }
            "number_isfinite" => {
                self.compile_expr(args[0])?;
                common::math::emit_abs(self.chunk(), line);
                self.emit_const(Value::F64(f64::MAX));
                self.emit(Op::DYN_LE);
            }
            "number_isinteger" => {
                self.compile_expr(args[0])?;
                self.emit(Op::DUP);
                self.emit(Op::F64_TRUNC);
                self.emit(Op::DYN_EQ);
            }
            "map_size" => {
                self.compile_expr(args[0])?;
                common::dict::emit_keys(&mut self.chunks, self.current, line);
                common::collections::emit_len(&mut self.chunks, self.current, line);
            }
            "array_at" => {
                // .at() supports negative indices for both arrays and strings.
                // Receiver is already on stack from value method dispatch.
                // `Array.prototype.at` per ECMA-262 §23.1.3.1.
                if args.len() >= 1 {
                    self.compile_expr(args[0])?;
                    let idx = self.import("ecma:array", "at");
                    self.emit_host_call(idx, 2);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "instr" => {
                if args.len() == 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::strings::emit_index_of(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_ADD);
                } else if args.len() == 3 {
                    let start_slot = self.define_local("__instr_start");
                    self.compile_expr(args[0])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, start_slot);
                    self.emit(Op::DROP);

                    self.compile_expr(args[1])?;
                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    self.emit_const(Value::I32(0x7FFF_FFFF));
                    common::strings::emit_substring(self.chunk(), line);
                    self.compile_expr(args[2])?;
                    common::strings::emit_index_of(self.chunk(), line);
                    let idx_slot = self.define_local("__instr_idx");
                    self.emit_u16(Op::LOCAL_SET, idx_slot);
                    self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit(Op::I32_CONST_0);
                    self.emit(Op::DYN_LT);
                    let found = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit(Op::I32_CONST_0);
                    let done = self.emit_jump(Op::BR);
                    self.patch_jump(found);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    self.emit(Op::I32_ADD);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_ADD);
                    self.patch_jump(done);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "strcomp" => {
                if args.len() >= 2 {
                    let left_slot = self.define_local("__strcomp_left");
                    let right_slot = self.define_local("__strcomp_right");
                    let text_slot = self.define_local("__strcomp_text");

                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, left_slot);
                    self.emit(Op::DROP);
                    self.compile_expr(args[1])?;
                    self.emit_u16(Op::LOCAL_SET, right_slot);
                    self.emit(Op::DROP);
                    if let Some(compare_arg) = args.get(2) {
                        self.compile_expr(compare_arg)?;
                        self.emit(Op::DYN_TO_BOOL);
                    } else {
                        self.emit(Op::FALSE);
                    }
                    self.emit_u16(Op::LOCAL_SET, text_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, text_slot);
                    let binary_compare = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit_u16(Op::LOCAL_GET, left_slot);
                    common::strings::emit_to_lower(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_GET, right_slot);
                    common::strings::emit_to_lower(self.chunk(), line);
                    self.emit(Op::STR_COMPARE);
                    let done = self.emit_jump(Op::BR);
                    self.patch_jump(binary_compare);
                    self.emit_u16(Op::LOCAL_GET, left_slot);
                    self.emit_u16(Op::LOCAL_GET, right_slot);
                    self.emit(Op::STR_COMPARE);
                    self.patch_jump(done);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "instrrev" => {
                if args.len() >= 2 {
                    if args.len() >= 3 {
                        let source_slot = self.define_local("__instrrev_source");
                        let start_slot = self.define_local("__instrrev_start");
                        self.compile_expr(args[0])?;
                        self.emit_u16(Op::LOCAL_SET, source_slot);
                        self.emit(Op::DROP);
                        self.compile_expr(args[2])?;
                        common::convert::emit_to_int(self.chunk(), line);
                        self.emit_u16(Op::LOCAL_SET, start_slot);
                        self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, source_slot);
                        self.emit(Op::I32_CONST_0);
                        self.emit_u16(Op::LOCAL_GET, start_slot);
                        common::strings::emit_substring(self.chunk(), line);
                        self.compile_expr(args[1])?;
                        common::strings::emit_last_index_of(self.chunk(), line);
                        let idx_slot = self.define_local("__instrrev_idx");
                        self.emit_u16(Op::LOCAL_SET, idx_slot);
                        self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit(Op::I32_CONST_0);
                        self.emit(Op::DYN_LT);
                        let found = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit(Op::I32_CONST_0);
                        let done = self.emit_jump(Op::BR);
                        self.patch_jump(found);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::I32_ADD);
                        self.patch_jump(done);
                    } else {
                        self.compile_expr(args[0])?;
                        self.compile_expr(args[1])?;
                        common::strings::emit_last_index_of(self.chunk(), line);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::I32_ADD);
                    }
                } else {
                    self.emit(Op::NULL);
                }
            }
            "fortran_index" => {
                if args.len() >= 2 {
                    let source_slot = self.define_local("__fortran_index_source");
                    let search_slot = self.define_local("__fortran_index_search");
                    let back_slot = self.define_local("__fortran_index_back");

                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, source_slot);
                    self.emit(Op::DROP);

                    self.compile_expr(args[1])?;
                    self.emit_u16(Op::LOCAL_SET, search_slot);
                    self.emit(Op::DROP);

                    if let Some(back_arg) = args.get(2) {
                        self.compile_expr(back_arg)?;
                        self.emit(Op::DYN_TO_BOOL);
                    } else {
                        self.emit(Op::FALSE);
                    }
                    self.emit_u16(Op::LOCAL_SET, back_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, back_slot);
                    let forward_path = self.emit_jump(Op::BR_IF_FALSE);

                    self.emit_u16(Op::LOCAL_GET, source_slot);
                    self.emit_u16(Op::LOCAL_GET, search_slot);
                    common::strings::emit_last_index_of(self.chunk(), line);
                    let done = self.emit_jump(Op::BR);

                    self.patch_jump(forward_path);
                    self.emit_u16(Op::LOCAL_GET, source_slot);
                    self.emit_u16(Op::LOCAL_GET, search_slot);
                    common::strings::emit_index_of(self.chunk(), line);
                    self.patch_jump(done);

                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_ADD);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "fortran_scan" => {
                self.emit_fortran_scan_like(args, false)?;
            }
            "fortran_verify" => {
                self.emit_fortran_scan_like(args, true)?;
            }
            "replace" => {
                if args.len() >= 3 {
                    let source_slot = self.define_local("__vb_replace_source");
                    let find_slot = self.define_local("__vb_replace_find");
                    let repl_slot = self.define_local("__vb_replace_repl");
                    let start_slot = self.define_local("__vb_replace_start");
                    let count_slot = self.define_local("__vb_replace_count");
                    let text_slot = self.define_local("__vb_replace_text");
                    let result_slot = self.define_local("__vb_replace_result");
                    let remaining_slot = self.define_local("__vb_replace_remaining");
                    let find_cmp_slot = self.define_local("__vb_replace_find_cmp");
                    let current_cmp_slot = self.define_local("__vb_replace_current_cmp");
                    let find_len_slot = self.define_local("__vb_replace_find_len");
                    let idx_slot = self.define_local("__vb_replace_idx");
                    let replaced_slot = self.define_local("__vb_replace_done");
                    let prefix_slot = self.define_local("__vb_replace_prefix_end");

                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, source_slot);
                    self.emit(Op::DROP);

                    self.compile_expr(args[1])?;
                    self.emit_u16(Op::LOCAL_SET, find_slot);
                    self.emit(Op::DROP);

                    self.compile_expr(args[2])?;
                    self.emit_u16(Op::LOCAL_SET, repl_slot);
                    self.emit(Op::DROP);

                    if let Some(start_arg) = args.get(3) {
                        self.compile_expr(start_arg)?;
                        common::convert::emit_to_int(self.chunk(), line);
                    } else {
                        self.emit_const(Value::I32(0));
                    }
                    self.emit_u16(Op::LOCAL_SET, start_slot);
                    self.emit(Op::DROP);

                    if let Some(count_arg) = args.get(4) {
                        self.compile_expr(count_arg)?;
                        common::convert::emit_to_int(self.chunk(), line);
                    } else if args.get(3).is_some() {
                        self.emit_const(Value::I32(1));
                    } else {
                        self.emit_const(Value::I32(-1));
                    }
                    self.emit_u16(Op::LOCAL_SET, count_slot);
                    self.emit(Op::DROP);

                    if let Some(compare_arg) = args.get(5) {
                        self.compile_expr(compare_arg)?;
                        self.emit(Op::DYN_TO_BOOL);
                    } else {
                        self.emit(Op::FALSE);
                    }
                    self.emit_u16(Op::LOCAL_SET, text_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, find_slot);
                    common::strings::emit_length(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, find_len_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, text_slot);
                    let use_binary_find = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit_u16(Op::LOCAL_GET, find_slot);
                    common::strings::emit_to_lower(self.chunk(), line);
                    let find_cmp_done = self.emit_jump(Op::BR);
                    self.patch_jump(use_binary_find);
                    self.emit_u16(Op::LOCAL_GET, find_slot);
                    self.patch_jump(find_cmp_done);
                    self.emit_u16(Op::LOCAL_SET, find_cmp_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, start_slot);
                    self.emit_u16(Op::LOCAL_SET, prefix_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, prefix_slot);
                    self.emit(Op::I32_CONST_0);
                    self.emit(Op::DYN_GE);
                    let prefix_ok = self.emit_jump(Op::BR_IF_TRUE);
                    self.emit(Op::I32_CONST_0);
                    self.emit_u16(Op::LOCAL_SET, prefix_slot);
                    self.emit(Op::DROP);
                    self.patch_jump(prefix_ok);

                    self.emit_u16(Op::LOCAL_GET, source_slot);
                    self.emit(Op::I32_CONST_0);
                    self.emit_u16(Op::LOCAL_GET, prefix_slot);
                    common::strings::emit_substring(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, result_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, source_slot);
                    if args.get(3).is_some() && args.get(4).is_none() {
                        self.emit_u16(Op::LOCAL_GET, start_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::I32_SUB);
                    } else {
                        self.emit_u16(Op::LOCAL_GET, prefix_slot);
                    }
                    self.emit_u16(Op::LOCAL_GET, source_slot);
                    common::strings::emit_length(self.chunk(), line);
                    common::strings::emit_substring(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, remaining_slot);
                    self.emit(Op::DROP);

                    self.emit(Op::I32_CONST_0);
                    self.emit_u16(Op::LOCAL_SET, replaced_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, find_len_slot);
                    self.emit(Op::I32_CONST_0);
                    self.emit(Op::DYN_EQ);
                    let nonempty_find = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit_u16(Op::LOCAL_GET, source_slot);
                    let done_empty_find = self.emit_jump(Op::BR);
                    self.patch_jump(nonempty_find);

                    let loop_top = self.current_offset();
                    let mut exit_patches = Vec::new();

                    self.emit_u16(Op::LOCAL_GET, count_slot);
                    self.emit_const(Value::I32(0));
                    self.emit(Op::DYN_GE);
                    let unlimited_count = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit_u16(Op::LOCAL_GET, replaced_slot);
                    self.emit_u16(Op::LOCAL_GET, count_slot);
                    self.emit(Op::DYN_GE);
                    exit_patches.push(self.emit_jump(Op::BR_IF_TRUE));
                    self.patch_jump(unlimited_count);

                    self.emit_u16(Op::LOCAL_GET, text_slot);
                    let binary_current = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit_u16(Op::LOCAL_GET, remaining_slot);
                    common::strings::emit_to_lower(self.chunk(), line);
                    let current_cmp_done = self.emit_jump(Op::BR);
                    self.patch_jump(binary_current);
                    self.emit_u16(Op::LOCAL_GET, remaining_slot);
                    self.patch_jump(current_cmp_done);
                    self.emit_u16(Op::LOCAL_SET, current_cmp_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, current_cmp_slot);
                    self.emit_u16(Op::LOCAL_GET, find_cmp_slot);
                    common::strings::emit_index_of(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, idx_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit(Op::I32_CONST_0);
                    self.emit(Op::DYN_LT);
                    exit_patches.push(self.emit_jump(Op::BR_IF_TRUE));

                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.emit_u16(Op::LOCAL_GET, remaining_slot);
                    self.emit(Op::I32_CONST_0);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    common::strings::emit_substring(self.chunk(), line);
                    self.emit(Op::DYN_ADD);
                    self.emit_u16(Op::LOCAL_GET, repl_slot);
                    self.emit(Op::DYN_ADD);
                    self.emit_u16(Op::LOCAL_SET, result_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, remaining_slot);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    self.emit_u16(Op::LOCAL_GET, find_len_slot);
                    self.emit(Op::I32_ADD);
                    self.emit_u16(Op::LOCAL_GET, remaining_slot);
                    common::strings::emit_length(self.chunk(), line);
                    common::strings::emit_substring(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, remaining_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, replaced_slot);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_ADD);
                    self.emit_u16(Op::LOCAL_SET, replaced_slot);
                    self.emit(Op::DROP);

                    self.emit_loop(loop_top);
                    for exit in exit_patches {
                        self.patch_jump(exit);
                    }

                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.emit_u16(Op::LOCAL_GET, remaining_slot);
                    self.emit(Op::DYN_ADD);
                    self.patch_jump(done_empty_find);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "filter" => {
                if args.len() >= 2 {
                    let arr_slot = self.define_local("__vb_filter_arr");
                    let match_slot = self.define_local("__vb_filter_match");
                    let include_slot = self.define_local("__vb_filter_include");
                    let text_slot = self.define_local("__vb_filter_text");
                    let result_slot = self.define_local("__vb_filter_result");
                    let idx_slot = self.define_local("__vb_filter_idx");
                    let elem_slot = self.define_local("__vb_filter_elem");

                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, arr_slot);
                    self.emit(Op::DROP);

                    self.compile_expr(args[1])?;
                    self.emit_u16(Op::LOCAL_SET, match_slot);
                    self.emit(Op::DROP);

                    if let Some(include_arg) = args.get(2) {
                        self.compile_expr(include_arg)?;
                    } else {
                        self.emit(Op::TRUE);
                    }
                    self.emit(Op::DYN_TO_BOOL);
                    self.emit_u16(Op::LOCAL_SET, include_slot);
                    self.emit(Op::DROP);

                    if let Some(compare_arg) = args.get(3) {
                        self.compile_expr(compare_arg)?;
                        self.emit(Op::DYN_TO_BOOL);
                    } else {
                        self.emit(Op::FALSE);
                    }
                    self.emit_u16(Op::LOCAL_SET, text_slot);
                    self.emit(Op::DROP);

                    common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                    self.emit_u16(Op::LOCAL_SET, result_slot);
                    self.emit(Op::DROP);

                    let state = common::loops::emit_for_in_start(&mut self.chunks, self.current, arr_slot, idx_slot, line);
                    self.emit_u16(Op::LOCAL_SET, elem_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, text_slot);
                    let binary_compare = self.emit_jump(Op::BR_IF_FALSE);

                    self.emit_u16(Op::LOCAL_GET, elem_slot);
                    common::strings::emit_to_lower(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_GET, match_slot);
                    common::strings::emit_to_lower(self.chunk(), line);
                    common::strings::emit_index_of(self.chunk(), line);
                    let after_compare = self.emit_jump(Op::BR);

                    self.patch_jump(binary_compare);
                    self.emit_u16(Op::LOCAL_GET, elem_slot);
                    self.emit_u16(Op::LOCAL_GET, match_slot);
                    common::strings::emit_index_of(self.chunk(), line);
                    self.patch_jump(after_compare);

                    self.emit(Op::I32_CONST_0);
                    self.emit(Op::DYN_GE);

                    self.emit_u16(Op::LOCAL_GET, include_slot);
                    self.emit(Op::DYN_EQ);

                    let if_block = self.chunks[self.current].emit_block(line);
                    self.emit(Op::DYN_NOT);
                    self.chunks[self.current].emit_br_if(0, line);

                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.emit_u16(Op::LOCAL_GET, elem_slot);
                    common::collections::emit_push(&mut self.chunks, self.current, line);
                    self.emit(Op::DROP);

                    self.chunks[self.current].emit_end(line);
                    self.chunks[self.current].patch_block(if_block);

                    common::loops::emit_for_in_end(&mut self.chunks, self.current, idx_slot, state, line);
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "split" => {
                let source_slot = self.define_local("__vb_split_source");
                let delim_slot = self.define_local("__vb_split_delim");
                let delim_cmp_slot = self.define_local("__vb_split_delim_cmp");
                let limit_slot = self.define_local("__vb_split_limit");
                let text_slot = self.define_local("__vb_split_text");
                let result_slot = self.define_local("__vb_split_result");
                let count_slot = self.define_local("__vb_split_count");
                let remaining_slot = self.define_local("__vb_split_remaining");
                let cmp_slot = self.define_local("__vb_split_cmp");
                let delim_len_slot = self.define_local("__vb_split_delim_len");
                let idx_slot = self.define_local("__vb_split_idx");

                self.compile_expr(args[0])?;
                self.emit_u16(Op::LOCAL_SET, source_slot);
                self.emit(Op::DROP);

                if let Some(delim_arg) = args.get(1) {
                    self.compile_expr(delim_arg)?;
                } else {
                    self.emit_const(Value::String(Arc::from(" ")));
                }
                self.emit_u16(Op::LOCAL_SET, delim_slot);
                self.emit(Op::DROP);

                if let Some(limit_arg) = args.get(2) {
                    self.compile_expr(limit_arg)?;
                    common::convert::emit_to_int(self.chunk(), line);
                } else {
                    self.emit_const(Value::I32(-1));
                }
                self.emit_u16(Op::LOCAL_SET, limit_slot);
                self.emit(Op::DROP);

                if let Some(compare_arg) = args.get(3) {
                    self.compile_expr(compare_arg)?;
                    self.emit(Op::DYN_TO_BOOL);
                } else {
                    self.emit(Op::FALSE);
                }
                self.emit_u16(Op::LOCAL_SET, text_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, delim_slot);
                common::strings::emit_length(self.chunk(), line);
                self.emit_u16(Op::LOCAL_SET, delim_len_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, text_slot);
                let split_binary_delim = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_u16(Op::LOCAL_GET, delim_slot);
                common::strings::emit_to_lower(self.chunk(), line);
                let split_delim_done = self.emit_jump(Op::BR);
                self.patch_jump(split_binary_delim);
                self.emit_u16(Op::LOCAL_GET, delim_slot);
                self.patch_jump(split_delim_done);
                self.emit_u16(Op::LOCAL_SET, delim_cmp_slot);
                self.emit(Op::DROP);

                common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                self.emit_u16(Op::LOCAL_SET, result_slot);
                self.emit(Op::DROP);

                self.emit(Op::I32_CONST_0);
                self.emit_u16(Op::LOCAL_SET, count_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, source_slot);
                self.emit_u16(Op::LOCAL_SET, remaining_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, delim_len_slot);
                self.emit(Op::I32_CONST_0);
                self.emit(Op::DYN_EQ);
                let nonempty_delim = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                self.emit_u16(Op::LOCAL_GET, source_slot);
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);
                let done_empty_delim = self.emit_jump(Op::BR);
                self.patch_jump(nonempty_delim);

                let loop_top = self.current_offset();
                let mut exit_patches = Vec::new();

                self.emit_u16(Op::LOCAL_GET, limit_slot);
                self.emit_const(Value::I32(0));
                self.emit(Op::DYN_GT);
                let unlimited_limit = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_u16(Op::LOCAL_GET, count_slot);
                self.emit_u16(Op::LOCAL_GET, limit_slot);
                self.emit_const(Value::I32(1));
                self.emit(Op::I32_SUB);
                self.emit(Op::DYN_GE);
                exit_patches.push(self.emit_jump(Op::BR_IF_TRUE));
                self.patch_jump(unlimited_limit);

                self.emit_u16(Op::LOCAL_GET, text_slot);
                let split_binary_current = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_u16(Op::LOCAL_GET, remaining_slot);
                common::strings::emit_to_lower(self.chunk(), line);
                let split_cmp_done = self.emit_jump(Op::BR);
                self.patch_jump(split_binary_current);
                self.emit_u16(Op::LOCAL_GET, remaining_slot);
                self.patch_jump(split_cmp_done);
                self.emit_u16(Op::LOCAL_SET, cmp_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, cmp_slot);
                self.emit_u16(Op::LOCAL_GET, delim_cmp_slot);
                common::strings::emit_index_of(self.chunk(), line);
                self.emit_u16(Op::LOCAL_SET, idx_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit(Op::I32_CONST_0);
                self.emit(Op::DYN_LT);
                exit_patches.push(self.emit_jump(Op::BR_IF_TRUE));

                self.emit_u16(Op::LOCAL_GET, result_slot);
                self.emit_u16(Op::LOCAL_GET, remaining_slot);
                self.emit(Op::I32_CONST_0);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                common::strings::emit_substring(self.chunk(), line);
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, count_slot);
                self.emit_const(Value::I32(1));
                self.emit(Op::I32_ADD);
                self.emit_u16(Op::LOCAL_SET, count_slot);
                self.emit(Op::DROP);

                self.emit_u16(Op::LOCAL_GET, remaining_slot);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_u16(Op::LOCAL_GET, delim_len_slot);
                self.emit(Op::I32_ADD);
                self.emit_u16(Op::LOCAL_GET, remaining_slot);
                common::strings::emit_length(self.chunk(), line);
                common::strings::emit_substring(self.chunk(), line);
                self.emit_u16(Op::LOCAL_SET, remaining_slot);
                self.emit(Op::DROP);

                self.emit_loop(loop_top);
                for exit in exit_patches {
                    self.patch_jump(exit);
                }

                self.emit_u16(Op::LOCAL_GET, result_slot);
                self.emit_u16(Op::LOCAL_GET, remaining_slot);
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, result_slot);
                self.patch_jump(done_empty_delim);
            }
            "join" => {
                // Two callers:
                //   - Intrinsic (`Join(arr, sep)`): args = [arr, sep],
                //     no receiver pre-pushed.
                //   - Value-method (`arr.join(sep)`): receiver `arr`
                //     already on stack, args = [sep].
                // Disambiguate by argc: 2 args → intrinsic shape; 1 arg
                // → value-method shape (only sep to push).
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                } else if args.len() == 1 {
                    self.compile_expr(args[0])?;
                } else {
                    self.emit_const(Value::String(Arc::from(",")));
                }
                { let l = self.line; common::collections::emit_join(&mut self.chunks, self.current, l); }
            }

            // ── Pascal ordinal/array intrinsics (canonical compiler_common ops) ──

            "high" => {
                // High(arr) → __len__(arr) - 1
                self.compile_expr(args[0])?;
                common::collections::emit_len(&mut self.chunks, self.current, line);
                self.emit_const(Value::I32(1));
                self.emit(Op::I32_SUB);
            }
            "low" => {
                // Low(arr) → 0 (always 0 for dynamic arrays in our VM)
                self.emit(Op::I32_CONST_0);
            }
            "succ" => {
                // Succ(x) → x + 1
                self.compile_expr(args[0])?;
                self.emit_const(Value::F64(1.0));
                self.emit(Op::DYN_ADD);
            }
            "pred" => {
                // Pred(x) → x - 1
                self.compile_expr(args[0])?;
                self.emit_const(Value::F64(1.0));
                self.emit(Op::F64_SUB);
            }
            "sqr" => {
                // Sqr(x) → x * x (square, NOT square root)
                self.compile_expr(args[0])?;
                self.emit(Op::DUP);
                self.emit(Op::F64_MUL);
            }
            "assigned" => {
                // Assigned(x) → x is not null
                self.compile_expr(args[0])?;
                self.emit(Op::NULL);
                self.emit(Op::DYN_NE);
            }
            "sizeof" => {
                // SizeOf(x) → 4 (boxed value)
                self.compile_expr(args[0])?;
                self.emit(Op::DROP);
                self.emit_const(Value::I32(4));
            }
            "classname" => {
                // ClassName(obj) → obj.__type
                self.compile_expr(args[0])?;
                let idx = self.str_const("__type");
                self.emit_u16(Op::STRUCT_GET, idx);
            }
            "pos" => {
                // Pos(substr, s) → IndexOf(s, substr) + 1 (Pascal 1-based)
                if args.len() == 2 {
                    self.compile_expr(args[1])?;
                    self.compile_expr(args[0])?;
                    common::strings::emit_index_of(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_ADD);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "copy" => {
                // Copy(s, start, len) → substring(s, start-1, start-1+len) — Pascal 1-based
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_SUB);
                    if args.len() >= 3 {
                        self.emit(Op::DUP);
                        self.compile_expr(args[2])?;
                        common::convert::emit_to_int(self.chunk(), line);
                        self.emit(Op::I32_ADD);
                    } else {
                        self.emit_const(Value::I32(0x7FFF_FFFF));
                    }
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "leftstr" => {
                // LeftStr(s, n) → substring(s, 0, n)
                if args.len() == 2 {
                    self.compile_expr(args[0])?;
                    self.emit(Op::I32_CONST_0);
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "str_concat" => {
                // Concat(a, b, c, ...) → a + b + c + ... using compiler_common::strings
                if args.is_empty() {
                    self.emit_const(Value::String(Arc::from("")));
                } else {
                    self.compile_expr(args[0])?;
                    for a in &args[1..] {
                        self.compile_expr(a)?;
                        common::strings::emit_str_concat(self.chunk(), line);
                    }
                }
            }
            "rightstr" => {
                // RightStr(s, n) → substring(s, len(s)-n, len(s))
                if args.len() == 2 {
                    self.compile_expr(args[0])?;
                    let s_slot = self.define_local("__rs_s");
                    self.emit_u16(Op::LOCAL_SET, s_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    common::strings::emit_length(self.chunk(), line);
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit(Op::I32_SUB);
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    common::strings::emit_length(self.chunk(), line);
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }

            "strtr" => {
                // PHP strtr two-arg form: strtr($str, $array) replaces every
                // occurrence of each array key with its value, applied in
                // insertion order. Implemented as a real loop:
                //
                //   entries = ecma:object.entries(array)
                //   str_slot  = $str
                //   for [k, v] in entries:
                //       str_slot = STR_REPLACE(str_slot, k, v)
                //   push str_slot
                //
                // Three-arg form `strtr($str, $from, $to)` (single-char swap
                // by position) is not yet covered — falls through to NULL
                // until a test demands it.
                if args.len() == 2 {
                    self.compile_expr(args[0])?;
                    let str_slot = self.define_local("__strtr_str");
                    self.emit_u16(Op::LOCAL_SET, str_slot); self.emit(Op::DROP);

                    self.compile_expr(args[1])?;
                    common::collections::emit_iter_entries(&mut self.chunks, self.current, line);
                    let entries_slot = self.define_local("__strtr_entries");
                    self.emit_u16(Op::LOCAL_SET, entries_slot); self.emit(Op::DROP);

                    let idx_slot = self.define_local("__strtr_idx");
                    let state = common::loops::emit_for_in_start(
                        &mut self.chunks, self.current, entries_slot, idx_slot, line,
                    );
                    // `emit_for_in_start` leaves the current entry on the
                    // stack — drop it; we re-fetch `[k, v]` via the index
                    // so we can pull both fields without a swap.
                    self.emit(Op::DROP);

                    // str_slot = STR_REPLACE(str_slot, entry[0], entry[1])
                    self.emit_u16(Op::LOCAL_GET, str_slot);
                    self.emit_u16(Op::LOCAL_GET, entries_slot);
                    self.emit_u16(Op::LOCAL_GET, idx_slot);
                    common::collections::emit_get(&mut self.chunks, self.current, line);
                    let pair_slot = self.define_local("__strtr_pair");
                    self.emit_u16(Op::LOCAL_SET, pair_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, pair_slot);
                    self.emit_const(Value::I32(0));
                    common::collections::emit_get(&mut self.chunks, self.current, line);
                    self.emit_u16(Op::LOCAL_GET, pair_slot);
                    self.emit_const(Value::I32(1));
                    common::collections::emit_get(&mut self.chunks, self.current, line);
                    self.emit(Op::STR_REPLACE);
                    self.emit_u16(Op::LOCAL_SET, str_slot); self.emit(Op::DROP);

                    common::loops::emit_for_in_end(
                        &mut self.chunks, self.current, idx_slot, state, line,
                    );

                    self.emit_u16(Op::LOCAL_GET, str_slot);
                } else {
                    self.emit(Op::NULL);
                }
            }

            // ── String compositions of ecma:string primitives ──────────
            //
            // Each of these used to live as a separate `vybe:string.*`
            // host fn; now compiled inline so the underlying providers
            // (ecma:string.padStart, ecma:string.toUpperCase, etc.) are
            // the single source of truth for semantics. The compositions
            // are well-known JS idioms — see comments per arm.

            "zfill" => {
                // Python str.zfill(width) → padStart(width, "0").
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_const(Value::String(Arc::from("0")));
                    let idx = self.import("ecma:string", "padStart");
                    self.emit_host_call(idx, 3);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "capitalize" => {
                // Python/Ruby `s.capitalize()` → s[0].toUpperCase() +
                // s.slice(1).toLowerCase(). Compose via ecma:string.
                if let Some(arg) = args.first() {
                    let s_slot = self.define_local("__cap_s");
                    self.compile_expr(arg)?;
                    self.emit_u16(Op::LOCAL_SET, s_slot); self.emit(Op::DROP);
                    // first char upper
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    self.emit(Op::I32_CONST_0);
                    self.emit_const(Value::I32(1));
                    common::strings::emit_substring(self.chunk(), line);
                    let upper_idx = self.import("ecma:string", "toUpperCase");
                    self.emit_host_call(upper_idx, 1);
                    // rest lower
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    self.emit_const(Value::I32(1));
                    self.emit_const(Value::I32(0x7FFF_FFFF));
                    common::strings::emit_substring(self.chunk(), line);
                    let lower_idx = self.import("ecma:string", "toLowerCase");
                    self.emit_host_call(lower_idx, 1);
                    // concat
                    self.emit(Op::DYN_ADD);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "center" => {
                // Python str.center(width, fill?) — pad symmetrically.
                // Compose: padStart(ceil((w + len)/2), fill).padEnd(w, fill).
                if args.len() >= 2 {
                    let s_slot = self.define_local("__cen_s");
                    let w_slot = self.define_local("__cen_w");
                    let pad_slot = self.define_local("__cen_pad");
                    self.compile_expr(args[0])?;
                    self.emit_u16(Op::LOCAL_SET, s_slot); self.emit(Op::DROP);
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_u16(Op::LOCAL_SET, w_slot); self.emit(Op::DROP);
                    if args.len() >= 3 {
                        self.compile_expr(args[2])?;
                    } else {
                        self.emit_const(Value::String(Arc::from(" ")));
                    }
                    self.emit_u16(Op::LOCAL_SET, pad_slot); self.emit(Op::DROP);
                    // Step 1: padStart with target = (w + len) / 2 + len_remainder
                    // For simplicity: padStart with (w + len + 1)/2.
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    // target = (w + len + 1) / 2
                    self.emit_u16(Op::LOCAL_GET, w_slot);
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    common::strings::emit_length(self.chunk(), line);
                    self.emit(Op::I32_ADD);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_ADD);
                    self.emit_const(Value::I32(2));
                    self.emit(Op::I32_DIV_S);
                    self.emit_u16(Op::LOCAL_GET, pad_slot);
                    let pad_start = self.import("ecma:string", "padStart");
                    self.emit_host_call(pad_start, 3);
                    // Step 2: padEnd to full width.
                    self.emit_u16(Op::LOCAL_GET, w_slot);
                    self.emit_u16(Op::LOCAL_GET, pad_slot);
                    let pad_end = self.import("ecma:string", "padEnd");
                    self.emit_host_call(pad_end, 3);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "count" => {
                // Python `s.count(sub)` / PHP `substr_count($s, $sub)` —
                // count non-overlapping occurrences. Compose:
                // s.split(sub).length - 1.
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    let split_idx = self.import("ecma:string", "split");
                    self.emit_host_call(split_idx, 2);
                    common::collections::emit_len(&mut self.chunks, self.current, line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_SUB);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "chop" => {
                // Ruby `s.chop` — drop last char. Compose: s.slice(0, len(s)-1).
                if let Some(arg) = args.first() {
                    let s_slot = self.define_local("__chop_s");
                    self.compile_expr(arg)?;
                    self.emit_u16(Op::LOCAL_SET, s_slot); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    self.emit(Op::I32_CONST_0);
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    common::strings::emit_length(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_SUB);
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "chars" => {
                // Ruby/PHP `s.chars` — array of single-char strings.
                // Compose: s.split("").
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    self.emit_const(Value::String(Arc::from("")));
                    let split_idx = self.import("ecma:string", "split");
                    self.emit_host_call(split_idx, 2);
                } else {
                    let empty_arr = self.import("ecma:array", "new");
                    self.emit_host_call(empty_arr, 0);
                }
            }

            // ── Numeric conversion intrinsics ─────────────────────────
            //
            // VB / Pascal / Python `cint` / `int(x)` / `clng` — coerce
            // to a number and round to nearest-even so midpoint cases
            // line up with VB's Round semantics.
            "cint" | "clng" => {
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    let value_slot = self.define_local("__cint_value");
                    self.emit_u16(Op::LOCAL_SET, value_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.emit(Op::REF_TYPEOF);
                    self.emit_const(Value::String(Arc::from("string")));
                    self.emit(Op::DYN_EQ);
                    let generic_numeric = self.emit_jump(Op::BR_IF_FALSE);

                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    common::strings::emit_length(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::DYN_EQ);
                    let not_single_char = self.emit_jump(Op::BR_IF_FALSE);

                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    self.emit(Op::I32_CONST_0);
                    self.emit(Op::STR_CHAR_CODE_AT);
                    let done = self.emit_jump(Op::BR);

                    self.patch_jump(not_single_char);
                    self.patch_jump(generic_numeric);
                    self.emit_u16(Op::LOCAL_GET, value_slot);
                    let num = self.import("ecma:number", "Number");
                    self.emit_host_call(num, 1);
                    let rounded_value_slot = self.define_local("__cint_rounded_value");
                    let floor_slot = self.define_local("__cint_floor");
                    let ceil_slot = self.define_local("__cint_ceil");
                    let frac_slot = self.define_local("__cint_frac");

                    self.emit_u16(Op::LOCAL_SET, rounded_value_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, rounded_value_slot);
                    self.emit(Op::F64_FLOOR);
                    self.emit_u16(Op::LOCAL_SET, floor_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, rounded_value_slot);
                    self.emit(Op::F64_CEIL);
                    self.emit_u16(Op::LOCAL_SET, ceil_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, rounded_value_slot);
                    self.emit_u16(Op::LOCAL_GET, floor_slot);
                    self.emit(Op::F64_SUB);
                    self.emit_u16(Op::LOCAL_SET, frac_slot);
                    self.emit(Op::DROP);

                    self.emit_u16(Op::LOCAL_GET, frac_slot);
                    self.emit_const(Value::F64(0.5));
                    self.emit(Op::DYN_LT);
                    let not_less_than_half = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit_u16(Op::LOCAL_GET, floor_slot);
                    let rounded_done = self.emit_jump(Op::BR);

                    self.patch_jump(not_less_than_half);
                    self.emit_u16(Op::LOCAL_GET, frac_slot);
                    self.emit_const(Value::F64(0.5));
                    self.emit(Op::DYN_GT);
                    let not_greater_than_half = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit_u16(Op::LOCAL_GET, ceil_slot);
                    let ceil_done = self.emit_jump(Op::BR);

                    self.patch_jump(not_greater_than_half);
                    self.emit_u16(Op::LOCAL_GET, floor_slot);
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_AND);
                    self.emit_const(Value::I32(0));
                    self.emit(Op::DYN_EQ);
                    let floor_is_odd = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit_u16(Op::LOCAL_GET, floor_slot);
                    let tie_done = self.emit_jump(Op::BR);

                    self.patch_jump(floor_is_odd);
                    self.emit_u16(Op::LOCAL_GET, ceil_slot);
                    self.patch_jump(tie_done);
                    self.patch_jump(ceil_done);
                    self.patch_jump(rounded_done);

                    common::convert::emit_to_int(self.chunk(), line);
                    self.patch_jump(done);
                } else {
                    self.emit_const(Value::F64(0.0));
                }
            }

            // VB `hex(n)` / `Hex$` — uppercase hex string.
            // ECMA composition: `Number(n).toString(16).toUpperCase()`.
            // `Number.prototype.toString` is called via a method
            // dispatch on the numeric receiver; `String.prototype.
            // toUpperCase` likewise.
            "hex" => {
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    let num = self.import("ecma:number", "Number");
                    self.emit_host_call(num, 1);
                    // Number(n).toString(16)
                    self.emit_const(Value::F64(16.0));
                    let to_str = self.import("ecma:number", "toString");
                    self.emit_host_call(to_str, 2);
                    // .toUpperCase()
                    let upper = self.import("ecma:string", "toUpperCase");
                    self.emit_host_call(upper, 1);
                } else {
                    self.emit_const(Value::String(Arc::from("0")));
                }
            }

            // VB `oct(n)` / `Oct$` — octal string.
            "oct" => {
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    let num = self.import("ecma:number", "Number");
                    self.emit_host_call(num, 1);
                    self.emit_const(Value::F64(8.0));
                    let to_str = self.import("ecma:number", "toString");
                    self.emit_host_call(to_str, 2);
                } else {
                    self.emit_const(Value::String(Arc::from("0")));
                }
            }

            _ => { self.emit(Op::NULL); }
        }
        Ok(())
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
                        if n < 2 || n > 255 { return false; }
                        match arity {
                            None => *arity = Some(n as u8),
                            Some(a) if *a as usize == n => {}
                            _ => return false,
                        }
                    } else {
                        return false;
                    }
                }
                StmtKind::Return(None) => { *saw_any = true; return false; }
                StmtKind::If { then_body, elifs, else_body, .. } => {
                    if !walk(then_body, arity, saw_any) { return false; }
                    for (_, b) in elifs {
                        if !walk(b, arity, saw_any) { return false; }
                    }
                    if let Some(b) = else_body {
                        if !walk(b, arity, saw_any) { return false; }
                    }
                }
                StmtKind::While { body, else_body, .. }
                | StmtKind::ForIn { body, else_body, .. } => {
                    if !walk(body, arity, saw_any) { return false; }
                    if let Some(b) = else_body {
                        if !walk(b, arity, saw_any) { return false; }
                    }
                }
                StmtKind::For { body, .. } | StmtKind::DoWhile { body, .. }
                | StmtKind::With { body, .. } | StmtKind::Using { body, .. } => {
                    if !walk(body, arity, saw_any) { return false; }
                }
                StmtKind::Try { body, catches, else_body, finally } => {
                    if !walk(body, arity, saw_any) { return false; }
                    for c in catches {
                        if !walk(&c.body, arity, saw_any) { return false; }
                    }
                    if let Some(b) = else_body {
                        if !walk(b, arity, saw_any) { return false; }
                    }
                    if let Some(b) = finally {
                        if !walk(b, arity, saw_any) { return false; }
                    }
                }
                StmtKind::Block(b) => {
                    if !walk(b, arity, saw_any) { return false; }
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
    if !walk(body, &mut arity, &mut saw_any) { return None; }
    if saw_any { arity } else { None }
}

fn is_hoisted_deconstruction_block(stmts: &[Statement]) -> bool {
    let call_stmt = match stmts {
        [Statement { kind: StmtKind::Expr(expr), .. }] => expr,
        [Statement { kind: StmtKind::VarDecl { .. }, .. }, Statement { kind: StmtKind::Expr(expr), .. }] => expr,
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
                let obj_is_self = matches!(
                    &object.kind,
                    ExprKind::This | ExprKind::Super
                ) || matches!(
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
        if case_sensitive { name.to_string() } else { name.to_lowercase() }
    };

    // First pass: collect (name → first_index)
    let mut first_index: std::collections::HashMap<String, usize> = std::collections::HashMap::new();
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
                } = &mut merged.kind {
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
                    let mut has_ctor = m.iter().any(|mb| matches!(mb, ClassMember::Constructor { .. }));
                    for later in body.iter().skip(i + 1) {
                        if let StmtKind::ClassDecl {
                            name: ln, members: lm, parents: lp, interfaces: li, ..
                        } = &later.kind {
                            if key(ln) == k {
                                for lmem in lm {
                                    if matches!(lmem, ClassMember::Constructor { .. }) {
                                        if has_ctor { continue; }
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
