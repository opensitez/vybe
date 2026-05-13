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

use std::sync::Arc;
use std::collections::{HashSet, HashMap};
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
    fields: Vec<String>,
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
    /// Nested type names attached to this class constructor object.
    nested_types: Vec<String>,
    /// Static methods: (name, chunk_idx) — tracked for inheritance
    statics: Vec<(String, usize)>,
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
}

#[derive(Debug, Clone, Default)]
pub(crate) struct ReflectionMemberMetadata {
    pub decorators: Vec<Expression>,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct ReflectionTypeMetadata {
    pub parents: Vec<String>,
    pub decorators: Vec<Expression>,
    pub methods: HashMap<String, ReflectionMethodMetadata>,
    pub properties: HashMap<String, ReflectionMemberMetadata>,
    pub fields: HashMap<String, ReflectionMemberMetadata>,
}

#[derive(Debug, Clone)]
pub(crate) enum ReflectionBinding {
    Type(String),
    Method { type_name: String, method_name: String },
    Property { type_name: String, property_name: String },
    Field { type_name: String, field_name: String },
    Parameter { type_name: String, method_name: String, index: usize },
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
    defined_functions: HashSet<String>,
    function_param_modes: HashMap<String, Vec<PassBy>>,
    function_min_arity: HashMap<String, usize>,
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
    /// Mirrors `NormalClass.implicit_self_fields` for the class the
    /// compiler is currently inside. Saved/restored by `compile_class`
    /// alongside `current_class`. Expression + call-site resolution
    /// consults this instead of `profile.implicit_self_fields`, so the
    /// walker stays the single source of truth for per-language class
    /// semantics.
    pub(crate) current_class_implicit_self: bool,
    /// Label for the next loop to be pushed (set by StmtKind::Labeled).
    pending_label: Option<String>,
    with_targets: Vec<u16>,


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
            defined_functions: HashSet::new(),
            function_param_modes: HashMap::new(),
            function_min_arity: HashMap::new(),
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
            current_class_implicit_self: false,
            pending_label: None,
            with_targets: Vec::new(),


            multi_return_functions: HashMap::new(),
            generator_functions: HashSet::new(),
            host_import_bindings: HashMap::new(),
            host_namespace_aliases: HashMap::new(),
            host_package_roots: HashMap::new(),
            module_exports: HashMap::new(),
            active_finally_blocks: Vec::new(),
            uses_proxy: false,
        }
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
                members,
                decorators,
                ..
            } => {
                let runtime_name = self.reflection_runtime_type_name(name, parent_runtime_name);
                self.record_reflection_type(
                    &runtime_name,
                    parents,
                    decorators,
                    members,
                );
            }
            StmtKind::StructDecl {
                name,
                interfaces,
                members,
                decorators,
                ..
            } => {
                let runtime_name = self.reflection_runtime_type_name(name, parent_runtime_name);
                self.record_reflection_type(
                    &runtime_name,
                    interfaces,
                    decorators,
                    members,
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
                    interfaces,
                    decorators,
                    body_members,
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
        decorators: &[Expression],
        members: &[ClassMember],
    ) {
        let mut metadata = ReflectionTypeMetadata {
            parents: parents
            .iter()
            .map(|parent| self.reflection_runtime_type_name(parent, None))
            .collect(),
            decorators: decorators.to_vec(),
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
                    modifiers,
                    ..
                } => {
                    metadata.properties.insert(
                        name.clone(),
                        ReflectionMemberMetadata {
                            decorators: modifiers.decorators.clone(),
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
                        },
                    );
                }
                ClassMember::NestedType(stmt) => {
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
        let trimmed = type_name.trim().trim_end_matches('?').trim();
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
        let normalized = without_generics.trim();
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
        self.emit(Op::DROP);
        true
    }


    /// Emit a `for v in gen():` loop that drives the generator via
    /// `GEN_NEXT`. Layout:
    ///   <cont> = compile(iter)
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

        self.compile_generator_for_in_cont(var, cont_slot, body, else_body)
    }

    fn compile_generator_for_in_cont(
        &mut self,
        var: &str,
        cont_slot: u16,
        body: &[Statement],
        else_body: Option<&[Statement]>,
    ) -> Result<(), String> {

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
            let started_key = self.str_const("__php_gen_started");
            let current_key = self.str_const("__php_gen_current");
            let done_key = self.str_const("__php_gen_done");
            let return_key = self.str_const("__php_gen_return");

            self.emit_u16(Op::LOCAL_GET, cont_slot);
            self.emit_const(Value::Bool(true));
            self.emit_u16(Op::STRUCT_SET, started_key);
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, has_more_slot);
            self.emit(Op::DYN_TO_BOOL);
            let exhausted = self.emit_jump(Op::BR_IF_FALSE);

            self.emit_u16(Op::LOCAL_GET, cont_slot);
            self.emit_const(Value::Bool(false));
            self.emit_u16(Op::STRUCT_SET, done_key);
            self.emit(Op::DROP);
            self.emit_u16(Op::LOCAL_GET, cont_slot);
            self.emit_u16(Op::LOCAL_GET, value_slot);
            self.emit_u16(Op::STRUCT_SET, current_key);
            self.emit(Op::DROP);
            let loop_ready = self.emit_jump(Op::BR);

            self.patch_jump(exhausted);
            self.emit_u16(Op::LOCAL_GET, cont_slot);
            self.emit_const(Value::Bool(true));
            self.emit_u16(Op::STRUCT_SET, done_key);
            self.emit(Op::DROP);
            self.emit_u16(Op::LOCAL_GET, cont_slot);
            self.emit_u16(Op::LOCAL_GET, value_slot);
            self.emit_u16(Op::STRUCT_SET, return_key);
            self.emit(Op::DROP);
            self.emit_u16(Op::LOCAL_GET, cont_slot);
            self.emit_const(Value::Bool(false));
            self.emit_u16(Op::STRUCT_SET, current_key);
            self.emit(Op::DROP);
            self.emit_u8(Op::BR_LABEL, 1);

            self.patch_jump(loop_ready);
        } else {
            self.emit_u16(Op::LOCAL_GET, has_more_slot);
            self.emit(Op::DYN_TO_BOOL);
            self.emit(Op::DYN_NOT);
            // br_if_label 1 → jump to $exit when has_more was 0.
            self.emit_u8(Op::BR_IF_LABEL, 1);
        }

        // Pop the value into `var`.
        let var_slot = self.define_local(var);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_u16(Op::LOCAL_SET, var_slot); self.emit(Op::DROP);

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
        let (callee_name, _args) = match &value.kind {
            ExprKind::Call { callee, args, .. } => {
                match &callee.kind {
                    ExprKind::Ident(n) => (self.canon(n), args),
                    _ => return None,
                }
            }
            _ => return None,
        };
        let n = *self.multi_return_functions.get(&callee_name)?;
        if n as usize != idents.len() { return None; }
        Some((n, idents))
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
        self.emit_return();
        Ok(())
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
            fields: Vec::new(),
            static_fields: Vec::new(),
            static_field_types: HashMap::new(),
            static_method_names: Vec::new(),
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
    /// Static cases (push a string constant):
    ///   - `Ident("btn")`        → "btn"
    ///   - `Me` / `This`         → current class name (lowercased)
    ///   - `Member { Me, "btn" }` → "btn"
    ///
    /// Dynamic fallback (runtime lookup):
    ///   - any other expression  → compile expr, struct_get __control_name
    fn emit_event_control_key(&mut self, control: &Expression, line: u32) -> Result<(), String> {
        let is_self_ident = |c: &Compiler, n: &str| {
            let cn = c.canon(n);
            cn == c.profile.self_keyword || cn == "me" || cn == "this" || cn == "mybase"
        };
        let key: Option<String> = match &control.kind {
            // `Me` / `This` / `MyBase` as identifier or as keyword node →
            // the enclosing class is the control. Used for `Handles Me.Load`,
            // `Handles MyBase.Load`, `this.Click += h`, etc.
            ExprKind::This | ExprKind::Super => self.current_class.clone().map(|c| self.canon(&c)),
            ExprKind::Ident(name) if is_self_ident(self, name) =>
                self.current_class.clone().map(|c| self.canon(&c)),
            // Plain identifier. If it's a **class field** on the enclosing
            // type (designer-style `Me.btn1.Click += h` where the walker
            // stripped the `Me.`), the identifier name IS the key. But if
            // it's a **local variable** holding a freshly-constructed
            // control with a user-assigned `.Name`, the compile-time
            // variable name ("btn") and the runtime widget name ("b1")
            // diverge — so fall through to the runtime extraction path
            // which reads `__control_name` off the object.
            ExprKind::Ident(name) => {
                let is_class_field = if let Some(ref cn) = self.current_class {
                    self.pending_classes.get(cn.as_str())
                        .map(|pc| pc.fields.iter().any(|f| f.eq_ignore_ascii_case(name)))
                        .unwrap_or(false)
                } else { false };
                if is_class_field { Some(self.canon(name)) } else { None }
            }
            // `Me.btn` / `this.btn` → the field name on the form/class.
            ExprKind::Member { object, field, .. } => {
                let is_self = matches!(&object.kind, ExprKind::This | ExprKind::Super)
                    || matches!(&object.kind, ExprKind::Ident(n) if is_self_ident(self, n));
                if is_self {
                    Some(self.canon(field))
                } else {
                    None
                }
            }
            _ => None,
        };
        if let Some(k) = key {
            self.emit_const(Value::String(Arc::from(k.as_str())));
        } else {
            self.compile_expr(control)?;
            common::gui::emit_get_control_name(self.chunk(), line);
        }
        Ok(())
    }

    pub(crate) fn canon(&self, name: &str) -> String {
        if self.case_sensitive { name.to_string() } else { name.to_lowercase() }
    }

    fn normalize_type_hint(type_hint: &str) -> String {
        type_hint.trim().to_lowercase()
    }

    fn is_string_type_hint(type_hint: &str) -> bool {
        let normalized = Self::normalize_type_hint(type_hint);
        normalized == "string"
            || normalized == "system.string"
            || normalized.ends_with(".string")
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
        self.compile_expr(key)?;
        if self.expr_uses_case_insensitive_string_keys(owner) {
            let line = self.line;
            common::strings::emit_to_lower(self.chunk(), line);
        }
        Ok(())
    }

    pub(super) fn is_callable_type_hint(type_hint: &str) -> bool {
        let normalized = Self::normalize_type_hint(type_hint);
        normalized.starts_with("func")
            || normalized.starts_with("action")
            || normalized.contains(".func")
            || normalized.contains(".action")
            || normalized.contains(" delegate")
    }

    pub(super) fn is_pascal_set_type_hint(type_hint: &str) -> bool {
        Self::normalize_type_hint(type_hint).starts_with("set of ")
    }

    fn lookup_var_type_hint(&self, name: &str) -> Option<&str> {
        if let Some(type_hint) = self.scope().resolve_type(name) {
            return Some(type_hint);
        }
        if !self.case_sensitive {
            if let Some(type_hint) = self.scope().resolve_type_ci(name) {
                return Some(type_hint);
            }
        }
        let cname = self.canon(name);
        self.global_type_hints.get(&cname).map(|s| s.as_str())
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

    fn infer_expr_type_hint(&self, expr: &Expression) -> Option<String> {
        match &expr.kind {
            ExprKind::New { class, .. } => Self::expr_terminal_type_name(class),
            ExprKind::Call { callee, .. } => self.infer_dotnet_factory_return_type(callee),
            ExprKind::Member { object, .. } => {
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
            return;
        }
        if !self.case_sensitive {
            if let Some(slot) = self.scope().resolve_ci(name) {
                self.emit_u16(Op::LOCAL_GET, slot);
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
            let class_idx = self.str_const(&class_name);
            self.emit_u16(Op::GLOBAL_GET, class_idx);
            let field_idx = self.str_const(&self.canon(name));
            self.emit_u16(Op::STRUCT_GET, field_idx);
            return;
        }
        // Bare static method in class scope — `Double(x)` inside
        // `class Converter` resolves to `Converter.Double`.
        if let Some(class_name) = self.is_class_static_method(name) {
            let class_idx = self.str_const(&class_name);
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
        let idx = self.str_const(&cname);
        self.emit_u16(Op::GLOBAL_GET, idx);
    }

    fn emit_var_set(&mut self, name: &str) {
        // Local
        if let Some(slot) = self.scope().resolve(name) {
            self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
            return;
        }
        if !self.case_sensitive {
            if let Some(slot) = self.scope().resolve_ci(name) {
                self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
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
            let class_idx = self.str_const(&class_name);
            self.emit_u16(Op::GLOBAL_GET, class_idx);
            self.emit_u16(Op::LOCAL_GET, value_slot);
            let field_idx = self.str_const(&self.canon(name));
            self.emit_u16(Op::STRUCT_SET, field_idx);
            self.emit(Op::DROP);
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
        let idx = self.str_const(&cname); self.emit_u16(Op::GLOBAL_SET, idx); self.emit(Op::DROP);
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
            let mut current = Some(class_name.as_str());
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
        }
        None
    }

    fn is_class_static_field_type_hint(&self, name: &str) -> Option<String> {
        if let Some(ref class_name) = self.current_class {
            let mut current = Some(class_name.as_str());
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
        }
        None
    }

    fn is_class_nested_type(&self, name: &str) -> Option<String> {
        if let Some(ref class_name) = self.current_class {
            let mut current = Some(class_name.as_str());
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
            let mut current = Some(class_name.as_str());
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
        }
        None
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

    fn emit_python_truthiness_from_stack(&mut self) {
        if !self.is_python_profile() {
            self.emit(Op::DYN_TO_BOOL);
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
            ExprKind::Ident(name) => vec![name.clone()],
            ExprKind::This => vec![self.profile.self_keyword.clone()],
            ExprKind::Super => vec![self.profile.base_keyword.clone().unwrap_or_else(|| "super".into())],
            ExprKind::Member { object, field, .. } => {
                let mut parts = self.flatten_member_chain(object);
                parts.push(field.clone());
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
                if !all_decls { self.scope_mut().begin_scope(); }
                for s in stmts { self.compile_stmt(s)?; }
                if !all_decls { self.scope_mut().end_scope(); }
            }

            // ── Variable declarations ───────────────────────────────────
            StmtKind::VarDecl { declarations, kind } => {
                for decl in declarations {
                    self.compile_var_declarator(decl, kind)?;
                }
            }

            // ── Assignment ──────────────────────────────────────────────
            StmtKind::Assign { targets, value } => {
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
                    self.compile_expr(value)?;
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
                self.compile_expr(value)?;
                self.compile_compound_op(op);
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
                self.emit_python_truthiness_from_stack();
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
                    self.emit_python_truthiness_from_stack();
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
                for s in body { self.compile_stmt(s)?; }
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
                    self.compile_generator_for_in(var, iter, body, else_body.as_deref())?;
                } else {
                    let line = self.line;
                    self.compile_expr(iter)?;
                    let iter_slot = self.define_local("__forin_iter");
                    self.emit_u16(Op::LOCAL_SET, iter_slot); self.emit(Op::DROP);

                    let runtime_generator_done = if *of && key.is_none() {
                        self.emit_u16(Op::LOCAL_GET, iter_slot);
                        let is_gen_idx = self.import("ecma:value", "isGenerator");
                        self.emit_host_call(is_gen_idx, 1);
                        let not_gen = self.emit_jump(Op::BR_IF_FALSE);
                        self.compile_generator_for_in_cont(var, iter_slot, body, else_body.as_deref())?;
                        Some((not_gen, self.emit_jump(Op::BR)))
                    } else {
                        None
                    };

                    if let Some((not_gen, _)) = runtime_generator_done {
                        self.patch_jump(not_gen);
                    }

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
                        let var_slot = self.define_local(var);
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

                    if let Some((_, done)) = runtime_generator_done {
                        self.patch_jump(done);
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
                        for s in &c.body { self.compile_stmt(s)?; }
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
                } else {
                    self.emit(Op::NULL);
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
                // Every language's profile has `uses_normalize_class = true`
                // after Phase 3. ClassDecl always goes through
                // walker → normalize_class → emit_class → compile_class.
                // If a new language is added that hasn't written its
                // normalizer yet, `emit_class_from_ast` returns an error
                // loudly rather than silently picking a legacy path.
                let span = stmt.span.clone();
                crate::common::classes::emit::emit_class_from_ast(
                    self, span, &cname, parents, interfaces, members, modifiers,
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
                let span = stmt.span.clone();
                crate::common::classes::emit::emit_class_from_ast(
                    self, span, &cn, &[], &[], members, &crate::ast::ClassModifiers::default(),
                )?;
            }

            // ── Module declaration (VB) ─────────────────────────────────
            // Models WASM Component Model: members are exports of the module.
            // - Members compile as globals (so call_ref works)
            // - Bare member names register in enum_members map → resolve to Module.Member
            // - A namespace struct is built so qualified `Module.Member` works too
            StmtKind::ModuleDecl { name, members, .. } => {
                let module_name = self.canon(name);
                let mut member_names: Vec<String> = Vec::new();

                // First pass: compile all members as globals + collect names
                for m in members {
                    match m {
                        ClassMember::Method(stmt) => {
                            if let StmtKind::FunctionDecl { name: mname, .. } = &stmt.kind {
                                let mn = self.canon(mname);
                                self.compile_stmt(stmt)?;
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
                            // Nested class — gets its own global; also attach to module
                            if let StmtKind::ClassDecl { name: cname, .. } = &stmt.kind {
                                let cn = self.canon(cname);
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
                            self.compile_stmt(&ctor_stmt)?;
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
                let ns_name = self.canon(name);
                let mut member_names: Vec<(String, String, bool)> = Vec::new();
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
                self.defined_globals.insert(ns_name);

                let namespace_parts: Vec<&str> = name.split('.').map(|part| part.trim()).filter(|part| !part.is_empty()).collect();
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

                        // old = arr
                        self.emit_var_get(array);
                        self.emit_u16(Op::LOCAL_SET, old_slot); self.emit(Op::DROP);
                        // new_len = N + 1
                        self.compile_expr(size_expr)?;
                        self.emit_const(Value::F64(1.0));
                        self.emit(Op::DYN_ADD);
                        self.emit_u16(Op::LOCAL_SET, new_len_slot); self.emit(Op::DROP);
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
            // `__control_name` for general expressions). This decouples the
            // registry key from the runtime `.Name` property — renaming a
            // control via `btn.Name = "x"` doesn't break wired-up handlers.
            StmtKind::AddHandler { control, event, handler } => {
                let line = self.line;
                let bind_idx = self.import("vybe:gui", common::gui::HOST_FN_BIND_EVENT);
                // Stack: [control_name, event_name, handler_fn]
                self.emit_event_control_key(control, line)?;
                self.emit_const(Value::String(Arc::from(event.as_str())));
                self.compile_expr(handler)?;
                common::gui::emit_bind_event(self.chunk(), bind_idx, line);
                self.emit(Op::DROP); // statement: discard host call result
            }

            StmtKind::RemoveHandler { control, event, handler } => {
                let line = self.line;
                let unbind_idx = self.import("vybe:gui", common::gui::HOST_FN_UNBIND_EVENT);
                self.emit_event_control_key(control, line)?;
                self.emit_const(Value::String(Arc::from(event.as_str())));
                self.compile_expr(handler)?;
                common::gui::emit_unbind_event(self.chunk(), unbind_idx, line);
                self.emit(Op::DROP); // statement: discard host call result
            }

            StmtKind::RaiseEvent { event_name, args } => {
                let line = self.line;
                let raise_idx = self.import("vybe:gui", common::gui::HOST_FN_RAISE_EVENT);
                for a in args { self.compile_expr(a)?; }
                self.emit_const(Value::String(Arc::from(event_name.as_str())));
                common::gui::emit_raise_event(self.chunk(), raise_idx, (args.len() + 1) as u8, line);
                self.emit(Op::DROP); // statement: discard host call result
            }

            // ── VB legacy error handling ────────────────────────────────
            StmtKind::OnErrorResumeNext => { /* no-op in bytecode VM */ }
            StmtKind::OnErrorGoTo(_) => { /* no-op */ }
            StmtKind::GoTo(_) => { /* no-op — structured bytecode doesn't support arbitrary gotos */ }
            StmtKind::Label(_) => { /* no-op */ }

            // ── VB legacy file I/O ──────────────────────────────────────
            StmtKind::OpenFile { path, mode: _, file_number } => {
                self.compile_expr(path)?;
                self.compile_expr(file_number)?;
                let idx = self.import("wasi:filesystem", "openFile");
                self.emit_host_call(idx, 2);
                self.emit(Op::DROP);
            }
            StmtKind::CloseFile(file_num) => {
                if let Some(fnum) = file_num {
                    self.compile_expr(fnum)?;
                } else {
                    self.emit(Op::NULL);
                }
                let idx = self.import("wasi:filesystem", "closeFile");
                self.emit_host_call(idx, 1);
                self.emit(Op::DROP);
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
                let idx = self.import("wasi:filesystem", "writeFile");
                self.emit_host_call(idx, (items.len() + 1) as u8);
                self.emit(Op::DROP);
            }
            StmtKind::InputFile { file_number, variables } => {
                self.compile_expr(file_number)?;
                let idx = self.import("wasi:filesystem", "inputFile");
                self.emit_host_call(idx, 1);
                if let Some(first) = variables.first() {
                    self.emit_var_set(first);
                } else {
                    self.emit(Op::DROP);
                }
            }
            StmtKind::LineInput { file_number, variable } => {
                self.compile_expr(file_number)?;
                let idx = self.import("wasi:filesystem", "lineInput");
                self.emit_host_call(idx, 1);
                self.emit_var_set(variable);
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

            let try_from_lambda = Expression::new(ExprKind::Lambda {
                params: vec![mk_param("_self"), mk_param("v")],
                body: LambdaBody::Block(build_match_chain(Statement::new(StmtKind::Return(Some(Expression::null()))))),
                is_async: false,
                captures: vec![],
            });
            synthetic_members.push(ClassMember::Const {
                name: "tryFrom".to_string(),
                type_hint: None,
                value: try_from_lambda,
                visibility: Visibility::Public,
            });

            let from_lambda = Expression::new(ExprKind::Lambda {
                params: vec![mk_param("_self"), mk_param("v")],
                body: LambdaBody::Block(build_match_chain(Statement::new(StmtKind::Throw {
                    expr: Some(Expression::new(ExprKind::New {
                        class: Box::new(Expression::ident("Error")),
                        args: vec![Argument::positional(Expression::string(&format!("Invalid backing value for enum \"{}\"", name)))],
                    })),
                    cause: None,
                }))),
                is_async: false,
                captures: vec![],
            });
            synthetic_members.push(ClassMember::Const {
                name: "from".to_string(),
                type_hint: None,
                value: from_lambda,
                visibility: Visibility::Public,
            });
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
                let inferred_type_hint = decl.type_hint.clone().or_else(|| {
                    decl.init.as_ref().and_then(|expr| self.infer_expr_type_hint(expr))
                });
                // Top-level vars → globals.
                // `let`/`const` inside a block scope (depth > 0) are locals
                // even at the top level — they respect block scoping.
                // ECMA-262 §10.2.11: `var` inside a function is function-
                // scoped (a local), only script-level `var` is global.
                let is_toplevel = self.scopes.len() == 1 && self.scope().depth == 0;
                let is_hoisted = *kind == VarDeclKind::Var
                    && self.profile.hoist_var
                    && self.scopes.len() == 1;

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
                    self.compile_expr(init_expr)?;
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
                } else if let Some(ref bounds) = decl.array_bounds {
                    // Array with bounds: Dim arr(N) — N is the UPPER bound,
                    // so the array length is N+1. Emit through
                    // `common::collections` so the provider is swappable
                    // in one place (Phase D2).
                    if let Some(size_expr) = bounds.first() {
                        let line = self.line;
                        self.compile_expr(size_expr)?;
                        self.emit_const(Value::F64(1.0));
                        self.emit(Op::DYN_ADD);
                        common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
                    } else {
                        self.emit(Op::NULL);
                    }
                } else {
                    // Default values based on type hint
                    match decl.type_hint.as_deref().map(|s| s.to_lowercase()).as_deref() {
                        Some("integer") | Some("int") | Some("longint") | Some("real") | Some("double") | Some("float") => {
                            self.emit(Op::F64_CONST_0);
                        }
                        Some("boolean") | Some("bool") => self.emit(Op::FALSE),
                        Some("string") => self.emit_const(Value::String(Arc::from(""))),
                        _ => self.emit(Op::NULL),
                    }
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
                    || (!self.case_sensitive && self.scope().resolve_ci(name).is_some());

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
                self.compile_expr(object)?;
                let field_name = self.canon(field);
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
            ExprKind::Index { object, index, .. } => {
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
                if is_append {
                    self.compile_expr(object)?;
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    common::collections::emit_push(&mut self.chunks, self.current, line);
                    // ecma:array.push leaves [new_length]; drop it.
                    self.emit(Op::DROP);
                } else if matches!(self.profile.name.as_str(), "csharp" | "vb") {
                    self.compile_expr(object)?;
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
                    self.compile_expr(index)?;
                    if self.is_python_profile() {
                        let key_tmp = self.define_local("__py_idx_key");
                        let obj_tmp = self.define_local("__py_idx_obj");
                        self.emit_u16(Op::LOCAL_SET, key_tmp); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);

                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        let is_array_idx = self.import("ecma:array", "isArray");
                        self.chunk().emit_op_u16(Op::CALL_IMPORT, is_array_idx, line);
                        self.chunk().emit(1, line);
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
            // VB: arr(idx) = val — Call used as index because () is both call and index
            ExprKind::Call { callee, args, .. } if args.len() == 1 => {
                // VB `arr(idx) = val` — Call used as index-set because () is
                // both call and index in VB syntax. Route through
                // ecma:array.set per Phase D.
                let tmp = self.define_local("__tmp");
                self.emit_u16(Op::LOCAL_SET, tmp); self.emit(Op::DROP);
                self.compile_expr(callee)?;
                self.compile_expr(&args[0].value)?;
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
                        for (i, elem) in elems.iter().enumerate() {
                            match elem {
                                ArrayPatternElem::Pattern(BindingPattern::Ident(name), _) => {
                                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                                    self.emit_const(Value::F64(i as f64));
                                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                                    self.emit_var_set(name);
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
                                _ => {}
                            }
                        }
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
                if self.profile.dynamic_add {
                    // JS profile: ECMA §13.15.4 — call ToPrimitive on
                    // both operands with hint "default" before adding.
                    // The polyfill returns the operand unchanged for
                    // primitives (fast path) and unboxes Objects via
                    // their valueOf/toString chain (Date, custom
                    // valueOf, class instances).
                    if self.is_js_profile() {
                        self.coerce_top_two_to_default_primitive();
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
            BinOp::Mod => { let idx = self.import("ecma:math", "fmod"); let l = self.line; common::expressions::emit_f64_mod_with_import(self.chunk(), idx, l); },
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
                else if self.is_php_profile() { self.coerce_top_two_php_datetime_for_compare(); }
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
                else if self.is_php_profile() { self.coerce_top_two_php_datetime_for_compare(); }
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
                else if self.is_php_profile() { self.coerce_top_two_php_datetime_for_compare(); }
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
                else if self.is_php_profile() { self.coerce_top_two_php_datetime_for_compare(); }
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
                // a <=> b: returns -1, 0, or 1
                let i = self.import("ecma:math", "spaceship");
                self.emit_host_call(i, 2);
            }
            BinOp::And | BinOp::Or => unreachable!(), // handled with short-circuit
            BinOp::Xor => self.emit(Op::I32_XOR),
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
            CompoundOp::Add => self.emit(Op::DYN_ADD),
            CompoundOp::Sub => self.emit(Op::F64_SUB),
            CompoundOp::Mul => self.emit(Op::F64_MUL),
            CompoundOp::Div => self.emit(Op::F64_DIV),
            CompoundOp::IDiv => { self.emit(Op::F64_DIV); let l = self.line; common::math::emit_trunc(self.chunk(), l); }
            CompoundOp::Mod => { let idx = self.import("ecma:math", "fmod"); let l = self.line; common::expressions::emit_f64_mod_with_import(self.chunk(), idx, l); },
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
            let Some(var_name) = (match &args[0].kind {
                ExprKind::Ident(var_name) => Some(var_name.as_str()),
                _ => None,
            }) else {
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

        // Canonical builtins — language-agnostic dispatch via compiler_common::canonical.
        // Walkers normalize language-specific syntax (arr.Length, len(arr), Length(arr),
        // arr.size, etc.) to canonical dunder names (__len__, __str__, etc.).
        // The compiler doesn't know about language-specific names — it just looks up
        // the canonical name in compiler_common's registry.
        if let Some(canonical_op) = common::canonical::CanonicalOp::from_name(name) {
            // Special case: __str__ uses stdlib via global, not host import
            if matches!(canonical_op, common::canonical::CanonicalOp::Str) {
                if let Some(arg) = args.first() {
                    let tostring_global = self.str_const("__vybe_tostring");
                    self.emit_u16(Op::GLOBAL_GET, tostring_global);
                    self.compile_expr(arg)?;
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
                        if let Some(enum_type) = self.console_enum_type_from_expr(args[0]) {
                            self.emit_enum_value_to_string(&enum_type, args[0])?;
                        } else {
                            self.compile_expr(args[0])?;
                        }
                    } else {
                        for a in args { self.compile_expr(a)?; }
                    }
                    let line = self.line;
                    self.emit_common(name.as_str(), args.len() as u8, line);
                }
                BuiltinEmit::Stdlib(stdlib_name) => {
                    // Push func ref FIRST, then args, then call_ref
                    let global_name = format!("__vybe_{}", stdlib_name);
                    let name_idx = self.str_const(&global_name);
                    self.emit_u16(Op::GLOBAL_GET, name_idx);
                    for a in args { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, args.len() as u8);
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
            if handled { return; }
        }
        // Then the pure (chunk + line) common ops.
        let line2 = line;
        let handled = common::dispatch::emit_common(name, &mut self.chunks, self.current, argc, line2);
        if !handled {
            eprintln!("Unknown common emit: {}", name);
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
            "round" => { self.compile_expr(args[0])?; common::math::emit_round(self.chunk(), line); }
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
            "concat" => { for a in args { self.compile_expr(a)?; } common::strings::emit_concat(self.chunk(), args.len(), line); }
            "replace" => { if args.len() >= 3 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; self.compile_expr(args[2])?; common::strings::emit_replace(self.chunk(), line); } }
            "repeat" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; common::strings::emit_repeat(self.chunk(), line); } }
            "leftstr" => { self.compile_expr(args[0])?; self.emit_const(Value::F64(0.0)); self.compile_expr(args[1])?; common::strings::emit_substring(self.chunk(), line); }
            "high" => { self.compile_expr(args[0])?; common::strings::emit_length(self.chunk(), line); self.emit_const(Value::F64(1.0)); self.emit(Op::F64_SUB); }
            "low" => { self.emit_const(Value::F64(0.0)); }
            "setlength" => {
                if let Some(first) = args.first() {
                    if let ExprKind::Ident(var) = &first.kind {
                        let var = var.clone();
                        if args.len() > 1 { self.compile_expr(args[1])?; }
                        let idx = self.import("ecma:array", "newWithLength");
                        self.emit_host_call(idx, 1);
                        self.emit_var_set(&var);
                    }
                }
                self.emit(Op::NULL);
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
            "str_repeat" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; self.emit(Op::STR_REPEAT); } }
            "array_join" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; { let l = self.line; common::collections::emit_join(&mut self.chunks, self.current, l); } } }
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
    fn emit_intrinsic(&mut self, name: &str, args: &[&Expression]) -> Result<(), String> {
        let line = self.line;
        match name {
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
                    // end = len(s)
                    self.emit_u16(Op::LOCAL_GET, s_slot);
                    common::strings::emit_length(self.chunk(), line);
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
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
                    self.compile_expr(args[1])?;
                    self.compile_expr(args[0])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_SUB);
                    self.emit_const(Value::I32(0x7FFF_FFFF));
                    common::strings::emit_substring(self.chunk(), line);
                    self.compile_expr(args[2])?;
                    common::strings::emit_index_of(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_ADD);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "instrrev" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::strings::emit_last_index_of(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::I32_ADD);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "replace" => {
                if args.len() >= 3 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    self.compile_expr(args[2])?;
                    common::strings::emit_replace(self.chunk(), line);
                } else {
                    self.emit(Op::NULL);
                }
            }
            "split" => {
                self.compile_expr(args[0])?;
                if args.len() >= 2 {
                    self.compile_expr(args[1])?;
                } else {
                    self.emit_const(Value::String(Arc::from(" ")));
                }
                self.emit(Op::STR_SPLIT);
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
            // to a number then floor. Matches the legacy
            // `vybe:convert.cint` semantics (floor, not banker's
            // rounding — VB6 `cint` uses banker's, but the legacy host
            // fn used floor, so we preserve that here. A separate
            // banker's rounding intrinsic would be a behavior change.)
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
                    self.emit(Op::F64_FLOOR);
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
