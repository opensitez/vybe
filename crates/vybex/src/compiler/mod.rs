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
}

// ════════════════════════════════════════════════════════════════════════════
// Pending class bookkeeping
// ════════════════════════════════════════════════════════════════════════════

struct PendingClass {
    parent: Option<String>,
    fields: Vec<String>,
    /// Static methods: (name, chunk_idx) — tracked for inheritance
    statics: Vec<(String, usize)>,
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
    line: u32,
    defined_globals: HashSet<String>,
    defined_functions: HashSet<String>,
    defined_classes: HashSet<String>,
    /// Names of methods defined on any user class — used to avoid value method
    /// hijacking (e.g. user class `Calc.Add()` shouldn't match array `add`).
    defined_class_methods: HashSet<String>,
    /// Map from member name → containing namespace name.
    /// Used for bare-name resolution within modules/namespaces/enums.
    /// E.g. `Main` inside `Module Program` resolves to `Program.Main`.
    /// `Green` inside `enum TColor` resolves to `TColor.Green`.
    /// Models the WASM Component Model's namespace-scoped imports.
    enum_members: HashMap<String, String>,
    case_sensitive: bool,
    profile: LanguageProfile,
    current_func_name: Option<String>,
    current_result_slot: Option<u16>,
    pending_classes: HashMap<String, PendingClass>,
    current_class: Option<String>,
    /// Label for the next loop to be pushed (set by StmtKind::Labeled).
    pending_label: Option<String>,
    /// Functions whose every explicit `Return` carries an `ExprKind::Tuple`
    /// of the same arity. Populated by a pre-pass before any function is
    /// compiled so both callee (set `chunk.result_arity`, push N values
    /// without packing) and caller (destructure directly off the stack)
    /// can agree on the multi-value ABI at emit time.
    multi_return_functions: HashMap<String, u8>,
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
            line: 1,
            defined_globals: HashSet::new(),
            defined_functions: HashSet::new(),
            defined_classes: HashSet::new(),
            defined_class_methods: HashSet::new(),
            enum_members: HashMap::new(),
            case_sensitive: profile.case_sensitive,
            profile,
            current_func_name: None,
            current_result_slot: None,
            pending_classes: HashMap::new(),
            current_class: None,
            pending_label: None,
            multi_return_functions: HashMap::new(),
        }
    }

    /// Compile a module to bytecode chunks.
    pub fn compile(mut self, module: &Module) -> Result<Vec<Chunk>, String> {
        self.case_sensitive = self.profile.case_sensitive;

        // Register the .NET BCL class wrappers (Object → … → Form, Button, …)
        // before walking the user body, so user code that writes
        // `Inherits Form` finds a real `Form` class with a real ctor chain.
        // Gated on `profile.namespaces.use_dotnet` so non-.NET languages
        // don't get the names installed in their global scope.
        if self.profile.namespaces.use_dotnet {
            self.register_dotnet_classes()?;
        }

        // Pre-pass: merge `Partial Class` declarations sharing the same name
        // when the language profile enables it. After merging, the body has
        // exactly one ClassDecl per class name with all fields/methods/etc.
        // pooled together. This is a language-agnostic transform — every
        // language that sets `partial_classes = true` (VB, C#) gets it.
        let merged_body = if self.profile.partial_classes {
            merge_partial_classes(&module.body, self.case_sensitive)
        } else {
            module.body.clone()
        };

        // Multi-value pre-scan: any function whose every explicit `Return`
        // is a same-arity tuple literal is a candidate for the WASM
        // multi-value ABI. We only opt in when the language profile
        // requests it — other languages keep tuple-as-heap-object semantics.
        if self.profile.multi_value_tuple_returns {
            self.collect_multi_return_functions(&merged_body);
        }

        for stmt in &merged_body {
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
        common::bundle::finalize_with_stdlib(&mut self.chunks);
        Ok(self.chunks)
    }

    // ════════════════════════════════════════════════════════════════════════
    // Multi-value tuple returns (opt-in via `multi_value_tuple_returns`)
    // ════════════════════════════════════════════════════════════════════════

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
            let s = self.scope_mut().define("__mv_pack");
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
    fn chunk(&mut self) -> &mut Chunk { &mut self.chunks[self.current] }

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
            // Plain identifier → it's a field name (class field) or local
            // variable name. The source-stable identifier IS the key.
            ExprKind::Ident(name) => Some(self.canon(name)),
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

    fn canon(&self, name: &str) -> String {
        if self.case_sensitive { name.to_string() } else { name.to_lowercase() }
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
        if self.profile.implicit_self_fields && self.is_class_field(name) {
            let self_kw = self.profile.self_keyword.clone();
            if let Some(self_slot) = self.scope().resolve(&self_kw)
                .or_else(|| self.scope().resolve_ci(&self_kw))
            {
                self.emit_u16(Op::LOCAL_GET, self_slot);
                let cname = self.canon(name);
                let idx = self.str_const(&cname);
                self.emit_u16(Op::STRUCT_GET, idx);
                return;
            }
        }
        // Known type used as a value (e.g. `e instanceof RangeError`) — emit
        // the type name as a string so vybe:object:instanceOf can look it up
        // via its String fallback. Without this, `RangeError` would become
        // `global_get` of a nonexistent global → null.
        // Only do this when the name isn't shadowed by an actual global
        // (e.g. `Dim list As New List(Of String)` shadows the `list` type name).
        if self.profile.known_types.contains_key(name)
            && !self.defined_globals.contains(name)
            && !self.defined_globals.contains(&self.canon(name))
        {
            self.emit_const(Value::String(Arc::from(name)));
            return;
        }
        // Global — canonicalize name for case-insensitive languages
        let cname = self.canon(name);
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
        // Global — canonicalize name for case-insensitive languages
        let cname = self.canon(name);
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

    /// Check if a name is a field of the current class (for implicit self resolution).
    fn is_class_field(&self, name: &str) -> bool {
        if !self.profile.implicit_self_fields { return false; }
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
                        self.emit_var_get(name);
                        self.emit_u8(Op::CALL_REF, 0);
                        self.emit(Op::DROP);
                    }
                    // obj.method as statement → method call with 0 args
                    ExprKind::Member { object, field, .. } => {
                        self.compile_expr(object)?;
                        let field_name = self.canon(field);
                        let prop = self.str_const(&field_name);
                        self.emit(Op::DUP);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        let fn_tmp = self.scope_mut().define("__fn");
                        self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                        let obj_tmp = self.scope_mut().define("__obj");
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
                            self.scope_mut().define(name);
                        }
                        self.emit_var_set(name);
                    }
                } else {
                    self.compile_expr(value)?;
                    for (i, target) in targets.iter().enumerate() {
                        if i < targets.len() - 1 { self.emit(Op::DUP); }
                        self.compile_assign_target(target)?;
                    }
                }
            }

            StmtKind::CompoundAssign { target, op, value } => {
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
                self.emit(Op::DYN_TO_BOOL);
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
                    self.emit(Op::DYN_TO_BOOL);
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
                self.loops.push(LoopCtx { label: self.pending_label.take(), break_label_depth: break_depth, continue_label_depth: continue_depth });
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
                self.loops.push(LoopCtx { label: self.pending_label.take(), break_label_depth: break_depth, continue_label_depth: continue_depth });
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
            StmtKind::ForIn { var, iter, body, else_body, of, .. } => {
                self.compile_expr(iter)?;
                if !of {
                    let idx = self.import("vybe:object", "keys");
                    self.emit_host_call(idx, 1);
                }
                let arr_slot = self.scope_mut().define("__forin_arr");
                self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);
                let idx_slot = self.scope_mut().define("__forin_idx");
                let line = self.line;
                let lp = common::loops::emit_for_in_start(
                    &mut self.chunks, self.current, arr_slot, idx_slot, line,
                );
                // for_in_start emits: block + loop + cond + block $body = 3 labels
                let break_depth = self.label_depth + 1; // outer block
                let continue_depth = self.label_depth + 3; // body block (innermost)
                self.label_depth += 3;
                let var_slot = self.scope_mut().define(var);
                self.emit_u16(Op::LOCAL_SET, var_slot); self.emit(Op::DROP);
                self.loop_states.push(lp);
                self.loops.push(LoopCtx { label: self.pending_label.take(), break_label_depth: break_depth, continue_label_depth: continue_depth });
                for s in body { self.compile_stmt(s)?; }
                self.loops.pop();
                let lp = self.loop_states.pop().unwrap();
                common::loops::emit_for_in_end(
                    &mut self.chunks, self.current, idx_slot, lp, line,
                );
                self.label_depth -= 3;
                if let Some(else_stmts) = else_body {
                    for s in else_stmts { self.compile_stmt(s)?; }
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
                self.loops.push(LoopCtx { label: self.pending_label.take(), break_label_depth: break_depth, continue_label_depth: continue_depth });
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
                let sw_slot = self.scope_mut().define("__sw_expr");
                self.emit_u16(Op::LOCAL_SET, sw_slot); self.emit(Op::DROP);

                // Switch uses a BLOCK for break — push onto loop stack so break can find it
                let line = self.line;
                let switch_block = self.chunk().emit_block(line);
                self.label_depth += 1;
                let switch_lp = common::loops::LoopState { block_patch: switch_block, loop_patch: 0, body_block_patch: None };
                self.loop_states.push(switch_lp);
                self.loops.push(LoopCtx { label: self.pending_label.take(), break_label_depth: self.label_depth, continue_label_depth: self.label_depth });

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
                let catch_jump = common::errors::emit_try_start(&mut self.chunks[self.current], line);
                for s in body { self.compile_stmt(s)?; }
                common::errors::emit_try_end(&mut self.chunks[self.current], line);
                // Python else: runs if no exception
                if let Some(else_stmts) = else_body {
                    for s in else_stmts { self.compile_stmt(s)?; }
                }
                let skip_to_finally = self.emit_jump(Op::BR);
                common::errors::patch_catch(&mut self.chunks[self.current], catch_jump);
                if catches.is_empty() {
                    self.emit(Op::DROP);
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

                        let mut skip_arm: Option<usize> = None;
                        if !is_catch_all {
                            let mut to_body: Vec<usize> = Vec::new();
                            for ty in &types {
                                self.emit(Op::DUP);
                                let line = self.line;
                                let key = self.str_const("__exception_type");
                                self.chunks[self.current]
                                    .emit_op_u16(Op::STRUCT_GET, key, line);
                                let v = self.str_const(ty);
                                self.chunks[self.current]
                                    .emit_op_u16(Op::CONST, v, line);
                                self.emit(Op::DYN_EQ);
                                to_body.push(self.emit_jump(Op::BR_IF_TRUE));
                            }
                            skip_arm = Some(self.emit_jump(Op::BR));
                            for p in to_body { self.patch_jump(p); }
                        }

                        if let Some(ref var) = c.var_name {
                            let slot = self.scope_mut().define(var);
                            self.emit_u16(Op::LOCAL_SET, slot);
                            self.emit(Op::DROP);
                        } else {
                            self.emit(Op::DROP);
                        }
                        for s in &c.body { self.compile_stmt(s)?; }
                        end_patches.push(self.emit_jump(Op::BR));

                        if let Some(p) = skip_arm { self.patch_jump(p); }
                    }
                    // Fallthrough = no arm matched. Re-throw the exception.
                    let line = self.line;
                    common::errors::emit_throw(self.chunk(), line);
                    for p in end_patches { self.patch_jump(p); }
                }
                self.patch_jump(skip_to_finally);
                if let Some(fin) = finally {
                    for s in fin { self.compile_stmt(s)?; }
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
                            self.emit(Op::RETURN);
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
                self.emit(Op::RETURN);
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
                        self.emit(Op::RETURN);
                    }
                    BreakTarget::Implicit | BreakTarget::Kind(_) | BreakTarget::Level(_) => {
                        if let Some(depth) = self.break_depth(None) {
                            self.chunk().emit_br(depth, line);
                        }
                    }
                    BreakTarget::Label(label) => {
                        if let Some(depth) = self.break_depth(Some(label)) {
                            self.chunk().emit_br(depth, line);
                        }
                    }
                    BreakTarget::Value(expr) => {
                        self.compile_expr(expr)?;
                        self.emit(Op::RETURN);
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
            StmtKind::ClassDecl { name, parents, members, .. } => {
                let cname = self.canon(name);
                self.defined_globals.insert(cname.clone());
                self.defined_classes.insert(cname.clone());
                let parent = parents.first().map(|p| self.canon(p));
                self.compile_class(&cname, &parent, members)?;
            }

            // ── Interface declaration ───────────────────────────────────
            StmtKind::InterfaceDecl { .. } => {
                // No-op — interfaces are type-level only
            }

            // ── Enum declaration ────────────────────────────────────────
            // Compiles to a namespace object: Color = { Red: 0, Green: 1, Blue: 2 }
            // Bare member references (e.g. Pascal `c := Green`) are resolved at
            // compile time via the enum_members map.
            StmtKind::EnumDecl { name, members, .. } => {
                let cname = self.canon(name);
                self.emit_u16(Op::STRUCT_NEW, 0);
                let mut next_val = 0i64;
                for m in members {
                    self.emit(Op::DUP);
                    if let Some(ref v) = m.value {
                        if let ExprKind::Lit(Literal::Int(n)) = &v.kind {
                            next_val = *n;
                        }
                        self.compile_expr(v)?;
                    } else {
                        self.emit_const(Value::F64(next_val as f64));
                    }
                    next_val += 1;
                    let mname = self.canon(&m.name);
                    let key = self.str_const(&mname);
                    self.emit_u16(Op::STRUCT_SET, key);
                    self.emit(Op::DROP);
                    // Register member → enum type for bare-name resolution
                    self.enum_members.insert(mname, cname.clone());
                }
                let gidx = self.str_const(&cname);
                self.emit_u16(Op::GLOBAL_SET, gidx);
                self.emit(Op::DROP);
                self.defined_globals.insert(cname);
            }

            // ── Struct declaration (same as class) ──────────────────────
            StmtKind::StructDecl { name, interfaces: _, members, .. } => {
                let cn = self.canon(name);
                self.defined_globals.insert(cn.clone());
                self.compile_class(&cn, &None, members)?;
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
                            self.compile_expr(value)?;
                            let cn = self.canon(cname);
                            let idx = self.str_const(&cn);
                            self.emit_u16(Op::GLOBAL_SET, idx);
                            self.emit(Op::DROP);
                            self.defined_globals.insert(cn.clone());
                            member_names.push(cn);
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
                let mut member_names: Vec<String> = Vec::new();
                for s in body {
                    // Track top-level type/function names declared in this namespace
                    match &s.kind {
                        StmtKind::ClassDecl { name: cn, .. }
                        | StmtKind::StructDecl { name: cn, .. }
                        | StmtKind::EnumDecl { name: cn, .. }
                        | StmtKind::InterfaceDecl { name: cn, .. }
                        | StmtKind::ModuleDecl { name: cn, .. }
                        | StmtKind::FunctionDecl { name: cn, .. } => {
                            member_names.push(self.canon(cn));
                        }
                        _ => {}
                    }
                    self.compile_stmt(s)?;
                }

                // Build namespace struct
                self.emit_u16(Op::STRUCT_NEW, 0);
                for mn in &member_names {
                    self.emit(Op::DUP);
                    let gidx = self.str_const(mn);
                    self.emit_u16(Op::GLOBAL_GET, gidx);
                    let key = self.str_const(mn);
                    self.emit_u16(Op::STRUCT_SET, key);
                    self.emit(Op::DROP);
                }
                let ns_idx = self.str_const(&ns_name);
                self.emit_u16(Op::GLOBAL_SET, ns_idx);
                self.emit(Op::DROP);
                self.defined_globals.insert(ns_name);
            }

            // ── Delegate declaration ────────────────────────────────────
            StmtKind::DelegateDecl { .. } => {
                // No-op — delegates are type-level
            }

            // ── With ────────────────────────────────────────────────────
            StmtKind::With { items, body, .. } => {
                // Simplified: compile the first item (if any) and just run the body
                if let Some(first) = items.first() {
                    self.compile_expr(&first.expr)?;
                    if let Some(ref var) = first.var {
                        let slot = self.scope_mut().define(var);
                        self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                    } else {
                        self.emit(Op::DROP);
                    }
                }
                for s in body { self.compile_stmt(s)?; }
            }

            // ── Using ───────────────────────────────────────────────────
            StmtKind::Using { var, resource, body } => {
                self.compile_expr(resource)?;
                let slot = self.scope_mut().define(var);
                self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                for s in body { self.compile_stmt(s)?; }
                // Dispose is a no-op in our VM
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
                        let old_slot = self.scope_mut().define("__redim_old");
                        let new_slot = self.scope_mut().define("__redim_new");
                        let new_len_slot = self.scope_mut().define("__redim_nlen");
                        let idx_slot = self.scope_mut().define("__redim_idx");

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
                        let elem_slot = self.scope_mut().define("__redim_el");
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
                let idx = self.import("wasi:cli", "log");
                for expr in exprs {
                    self.compile_expr(expr)?;
                    common::io::emit_print_with_import(self.chunk(), idx, 1, line);
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
                let subject_slot = self.scope_mut().define("__match_subject");
                self.emit_u16(Op::LOCAL_SET, subject_slot); self.emit(Op::DROP);
                let mut end_patches = Vec::new();
                for case in cases {
                    // Simplified: match on value patterns only, wildcard always matches
                    let skip = match &case.pattern {
                        Pattern::Value(val) => {
                            self.emit_u16(Op::LOCAL_GET, subject_slot);
                            self.compile_expr(val)?;
                            self.emit(Op::DYN_EQ);
                            Some(self.emit_jump(Op::BR_IF_FALSE))
                        }
                        Pattern::Wildcard => None,
                        Pattern::As { name: Some(name), .. } => {
                            // Bind subject to name
                            self.emit_u16(Op::LOCAL_GET, subject_slot);
                            let slot = self.scope_mut().define(name);
                            self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                            None
                        }
                        _ => None, // Other patterns: always match (simplified)
                    };
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
                    if let Some(s) = skip { self.patch_jump(s); }
                }
                for p in end_patches { self.patch_jump(p); }
            }

            // ── Empty ───────────────────────────────────────────────────
            StmtKind::Empty => {}
        }
        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Variable declarator compilation
    // ════════════════════════════════════════════════════════════════════════

    fn compile_var_declarator(&mut self, decl: &VarDeclarator, kind: &VarDeclKind) -> Result<(), String> {
        match &decl.pattern {
            BindingPattern::Ident(name) => {
                if let Some(ref init_expr) = decl.init {
                    self.compile_expr(init_expr)?;
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
                // Top-level / hoisted vars → globals.
                // `let`/`const` inside a block scope (depth > 0) are locals
                // even at the top level — they respect block scoping.
                let is_toplevel = self.scopes.len() == 1 && self.scope().depth == 0;
                let is_hoisted = *kind == VarDeclKind::Var && self.profile.hoist_var;
                if is_toplevel || (is_hoisted && self.scopes.len() <= 2) {
                    let cn = self.canon(name);
                    let idx = self.str_const(&cn);
                    self.emit_u16(Op::GLOBAL_SET, idx);
                    self.emit(Op::DROP);
                    self.defined_globals.insert(cn);
                } else {
                    let slot = self.scope_mut().define(name);
                    self.emit_u16(Op::LOCAL_SET, slot);
                    self.emit(Op::DROP);
                }
            }
            BindingPattern::Object(props) => {
                // Destructuring: let { a, b } = expr
                if let Some(ref init_expr) = decl.init {
                    self.compile_expr(init_expr)?;
                    let obj_slot = self.scope_mut().define("__destruct_obj");
                    self.emit_u16(Op::LOCAL_SET, obj_slot); self.emit(Op::DROP);
                    for prop in props {
                        self.emit_u16(Op::LOCAL_GET, obj_slot);
                        let key = self.str_const(&prop.key);
                        self.emit_u16(Op::STRUCT_GET, key);
                        if let Some(ref default) = prop.default {
                            // If value is null, use default
                            self.emit(Op::DUP);
                            self.emit(Op::REF_IS_NULL);
                            let has_val = self.emit_jump(Op::BR_IF_FALSE);
                            self.emit(Op::DROP);
                            self.compile_expr(default)?;
                            self.patch_jump(has_val);
                        }
                        match &prop.value {
                            Some(BindingPattern::Ident(n)) => {
                                let slot = self.scope_mut().define(n);
                                self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                            }
                            Some(BindingPattern::Object(nested_props)) => {
                                // Nested destructuring: { nested: { b } }
                                // Value from struct_get is the nested object
                                let nested_slot = self.scope_mut().define("__nested");
                                self.emit_u16(Op::LOCAL_SET, nested_slot); self.emit(Op::DROP);
                                for np in nested_props {
                                    self.emit_u16(Op::LOCAL_GET, nested_slot);
                                    let nk = self.str_const(&np.key);
                                    self.emit_u16(Op::STRUCT_GET, nk);
                                    let bind = if let Some(BindingPattern::Ident(ref n)) = np.value {
                                        n.as_str()
                                    } else {
                                        &np.key
                                    };
                                    let slot = self.scope_mut().define(bind);
                                    self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                                }
                            }
                            _ => {
                                let slot = self.scope_mut().define(&prop.key);
                                self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                            }
                        }
                    }
                }
            }
            BindingPattern::Array(elems) => {
                // Destructuring: let [a, b] = expr
                if let Some(ref init_expr) = decl.init {
                    self.compile_expr(init_expr)?;
                    let arr_slot = self.scope_mut().define("__destruct_arr");
                    self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);
                    for (i, elem) in elems.iter().enumerate() {
                        match elem {
                            ArrayPatternElem::Pattern(BindingPattern::Ident(name), default) => {
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
                                let slot = self.scope_mut().define(name);
                                self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                            }
                            ArrayPatternElem::Rest(name) => {
                                // ...rest: slice from current index
                                self.emit_u16(Op::LOCAL_GET, arr_slot);
                                self.emit_const(Value::F64(i as f64));
                                // end = arr.length
                                self.emit_u16(Op::LOCAL_GET, arr_slot);
                                { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                                let line = self.line;
                                common::collections::emit_slice(&mut self.chunks, self.current, line);
                                let slot = self.scope_mut().define(name);
                                self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                            }
                            ArrayPatternElem::Hole => { /* skip */ }
                            _ => { /* nested patterns — simplified as no-op */ }
                        }
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
                        let tmp = self.scope_mut().define("__field_tmp");
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
                let tmp = self.scope_mut().define("__tmp");
                self.emit_u16(Op::LOCAL_SET, tmp); self.emit(Op::DROP);
                self.compile_expr(object)?;
                self.emit_u16(Op::LOCAL_GET, tmp);
                let field_name = self.canon(field);
                let idx = self.str_const(&field_name);
                self.emit_u16(Op::STRUCT_SET, idx);
                self.emit(Op::DROP);
            }
            ExprKind::Index { object, index } => {
                // PHP `$arr[] = v` — empty bracket with null index is the
                // auto-append form; route through collections::emit_push.
                // Every emit here goes via common::collections so the
                // provider (wasm:js-array / vybe:array / polyfill) is
                // swappable in one place.
                let is_append = matches!(
                    &index.kind,
                    ExprKind::Lit(crate::ast::Literal::Null)
                );
                let line = self.line;
                let tmp = self.scope_mut().define("__tmp");
                self.emit_u16(Op::LOCAL_SET, tmp); self.emit(Op::DROP);
                if is_append {
                    self.compile_expr(object)?;
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    common::collections::emit_push(&mut self.chunks, self.current, line);
                    // wasm:js-array.push leaves [new_length]; drop it.
                    self.emit(Op::DROP);
                } else {
                    self.compile_expr(object)?;
                    self.compile_expr(index)?;
                    self.emit_u16(Op::LOCAL_GET, tmp);
                    common::collections::emit_set(&mut self.chunks, self.current, line);
                    // wasm:js-array.set leaves [null]; drop it.
                    self.emit(Op::DROP);
                }
            }
            // VB: arr(idx) = val — Call used as index because () is both call and index
            ExprKind::Call { callee, args, .. } if args.len() == 1 => {
                // VB `arr(idx) = val` — Call used as index-set because () is
                // both call and index in VB syntax. Route through
                // wasm:js-array.set per Phase D.
                let tmp = self.scope_mut().define("__tmp");
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
                        let obj_slot = self.scope_mut().define("__destruct_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_slot); self.emit(Op::DROP);
                        for prop in props {
                            self.emit_u16(Op::LOCAL_GET, obj_slot);
                            let key = self.str_const(&prop.key);
                            self.emit_u16(Op::STRUCT_GET, key);
                            let bind_name = if let Some(BindingPattern::Ident(ref n)) = prop.value {
                                n.clone()
                            } else {
                                prop.key.clone()
                            };
                            self.emit_var_set(&bind_name);
                        }
                    }
                    DestructurePattern::Array(elems) => {
                        let arr_slot = self.scope_mut().define("__destruct_arr");
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
            _ => {}
        }
        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Binary operator emission
    // ════════════════════════════════════════════════════════════════════════

    fn compile_binop(&mut self, op: &BinOp) {
        match op {
            BinOp::Add => { if self.profile.dynamic_add { self.emit(Op::DYN_ADD); } else { self.emit(Op::DYN_ADD); } }
            BinOp::Sub => self.emit(Op::F64_SUB),
            BinOp::Mul => self.emit(Op::F64_MUL),
            BinOp::Div => self.emit(Op::F64_DIV),
            BinOp::IDiv => { self.emit(Op::F64_DIV); let l = self.line; common::math::emit_trunc(self.chunk(), l); }
            BinOp::FloorDiv => { self.emit(Op::F64_DIV); let l = self.line; common::math::emit_floor(self.chunk(), l); }
            BinOp::Mod => { let idx = self.import("vybe:math", "fmod"); let l = self.line; common::expressions::emit_f64_mod_with_import(self.chunk(), idx, l); },
            BinOp::Pow => { let l = self.line; common::math::emit_pow(self.chunk(), l); }
            BinOp::Eq => self.emit(Op::DYN_EQ),
            BinOp::NotEq => self.emit(Op::DYN_NE),
            BinOp::StrictEq => {
                // JS ===: no type coercion. Emit ref_typeof compare first,
                // then dyn_eq only if types match.
                // Stack: [a, b] → check types, then value equality.
                // Simplest: dup both, compare typeof, if different → false,
                // else dyn_eq. Using temp locals to avoid deep stack ops.
                let a_slot = self.scope_mut().define("__seq_a");
                let b_slot = self.scope_mut().define("__seq_b");
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
                let a_slot = self.scope_mut().define("__sne_a");
                let b_slot = self.scope_mut().define("__sne_b");
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
            BinOp::Lt => self.emit(Op::DYN_LT),
            BinOp::Gt => self.emit(Op::DYN_GT),
            BinOp::LtEq => self.emit(Op::DYN_LE),
            BinOp::GtEq => self.emit(Op::DYN_GE),
            BinOp::Spaceship => {
                // a <=> b: returns -1, 0, or 1
                let i = self.import("vybe:math", "spaceship");
                self.emit_host_call(i, 2);
            }
            BinOp::And | BinOp::Or => unreachable!(), // handled with short-circuit
            BinOp::Xor => self.emit(Op::I32_XOR),
            BinOp::BitAnd => self.emit(Op::I32_AND),
            BinOp::BitOr => self.emit(Op::I32_OR),
            BinOp::BitXor => self.emit(Op::I32_XOR),
            BinOp::Shl => self.emit(Op::I32_SHL),
            BinOp::Shr => self.emit(Op::I32_SHR_S),
            BinOp::UShr => self.emit(Op::I32_SHR_U),
            BinOp::Concat => { let l = self.line; common::strings::emit_str_concat(self.chunk(), l); }
            BinOp::In => {
                // `x in y` — JS: is `x` a property key / array index of `y`.
                // Walker pushes `[x, y]`; `wasm:js-array.includes(y, x)` wants
                // `[y, x]`. No SWAP opcode, so stash through local slots.
                // The import is polymorphic (strings, arrays, plain objects).
                let l = self.line;
                let t_y = self.scope_mut().define("__in_y");
                let t_x = self.scope_mut().define("__in_x");
                self.emit_u16(Op::LOCAL_SET, t_y); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_SET, t_x); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, t_y);
                self.emit_u16(Op::LOCAL_GET, t_x);
                common::collections::emit_contains(&mut self.chunks, self.current, l);
            }
            BinOp::NotIn => {
                let l = self.line;
                let t_y = self.scope_mut().define("__nin_y");
                let t_x = self.scope_mut().define("__nin_x");
                self.emit_u16(Op::LOCAL_SET, t_y); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_SET, t_x); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, t_y);
                self.emit_u16(Op::LOCAL_GET, t_x);
                common::collections::emit_contains(&mut self.chunks, self.current, l);
                self.emit(Op::DYN_NOT);
            }
            BinOp::InstanceOf => {
                // a instanceof B → check __type chain via host fn
                let idx = self.import("vybe:object", "instanceOf");
                self.emit_host_call(idx, 2);
            }
            BinOp::NullCoalesce => unreachable!(), // handled in compile_expr
            BinOp::MatMul => {
                let i = self.import("vybe:math", "matmul");
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
            CompoundOp::Mod => { let idx = self.import("vybe:math", "fmod"); let l = self.line; common::expressions::emit_f64_mod_with_import(self.chunk(), idx, l); },
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

    // ════════════════════════════════════════════════════════════════════════
    // Builtins (profile-driven)
    // ════════════════════════════════════════════════════════════════════════

    fn try_compile_builtin(&mut self, name: &str, args: &[&Expression]) -> Result<bool, String> {
        let line = self.line;

        // ── Phase D1 pilot: Array(count, init) → wasm:js-array.newWithLength + fill ──
        //
        // COBOL's OCCURS walker emits `Call { callee: Array,
        // args: [count, element_init] }` in the high-level IR. This
        // intercept routes the pattern through the spec-conformant
        // `wasm:js-array.*` imports instead of the legacy VM-internal
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
            // (wasm:js-array / vybe:array / polyfill) is swappable in
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

        // Check common import table first
        if let Some((module, func)) = common::imports::resolve_common_import(name) {
            for a in args { self.compile_expr(a)?; }
            let idx = self.import(module, func);
            self.emit_host_call(idx, args.len() as u8);
            return Ok(true);
        }

        // Look up in language profile
        let builtin = self.profile.lookup_builtin(name).cloned();
        if let Some(def) = builtin {
            match &def.emit {
                BuiltinEmit::Print => {
                    for a in args { self.compile_expr(a)?; }
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
                    for a in args { self.compile_expr(a)?; }
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
                    // Compile args, then dispatch to compiler_common emitter
                    for a in args { self.compile_expr(a)?; }
                    let line = self.line;
                    self.emit_common(name.as_str(), line);
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
    fn emit_common(&mut self, name: &str, line: u32) {
        // First try the import-needing dispatch (sleep, etc.). It needs a
        // closure into the compiler to resolve imports against chunk[0].
        // We use a raw pointer to break the borrow of self.
        {
            let self_ptr = self as *mut Self;
            let chunk = self.chunk();
            let handled = common::dispatch::emit_common_with_imports(
                name,
                chunk,
                line,
                |module, fname| unsafe { (*self_ptr).import(module, fname) },
            );
            if handled { return; }
        }
        // Then the pure (chunk + line) common ops.
        let line2 = line;
        let handled = common::dispatch::emit_common(name, &mut self.chunks, self.current, line2);
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
            // `wasm:js-array.*` imports. One-place-to-change: flip the
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
                        let idx = self.import("vybe:array", "newWithLength");
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
            "i32_from_f64" => { self.compile_expr(args[0])?; self.emit(Op::I32_FROM_F64); }
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
            "str_from_char_code" => {
                // String.fromCharCode(72, 105) → "Hi"
                self.compile_expr(args[0])?;
                self.emit(Op::STR_FROM_CHAR_CODE);
                for a in &args[1..] {
                    self.compile_expr(a)?;
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
            // `common::collections::*`, which now routes to `wasm:js-array.*`
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
                // Route through host fn vybe:array:at (handles arrays AND strings via negative idx).
                if args.len() >= 1 {
                    self.compile_expr(args[0])?;
                    let idx = self.import("vybe:array", "at");
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
                self.compile_expr(args[0])?;
                if args.len() >= 2 {
                    self.compile_expr(args[1])?;
                } else {
                    self.emit_const(Value::String(Arc::from("")));
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
                    let s_slot = self.scope_mut().define("__rs_s");
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
                    members: ref mut m,
                    parents: ref mut p,
                    interfaces: ref mut iface,
                    ..
                } = &mut merged.kind {
                    // Append members from every later declaration of this name.
                    for later in body.iter().skip(i + 1) {
                        if let StmtKind::ClassDecl {
                            name: ln, members: lm, parents: lp, interfaces: li, ..
                        } = &later.kind {
                            if key(ln) == k {
                                m.extend(lm.iter().cloned());
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
