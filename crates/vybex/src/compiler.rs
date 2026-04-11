use std::sync::Arc;
use std::collections::{HashSet, HashMap};
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use crate::profile::*;
use vybe_compiler_common as common;
use crate::ast::*;
use crate::scope::Scope;

// ════════════════════════════════════════════════════════════════════════════
// Loop context for break/continue patching
// ════════════════════════════════════════════════════════════════════════════

struct LoopCtx {
    break_patches: Vec<usize>,
    continue_target: usize,
    /// Forward jumps for continue in for-in loops where continue must skip to increment.
    /// If empty, continue uses continue_target directly (backward jump).
    /// If non-empty, continue emits a forward jump that gets patched later.
    continue_patches: Vec<usize>,
    /// When true, continue emits a forward jump (patched later) instead of backward loop.
    continue_needs_patch: bool,
    label: Option<String>,
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
}

impl Compiler {
    pub fn with_profile(profile: LanguageProfile) -> Self {
        Self {
            chunks: vec![Chunk::new("<script>")],
            scopes: vec![Scope::new()],
            current: 0,
            loops: Vec::new(),
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

        for stmt in &merged_body {
            self.compile_stmt(stmt)?;
        }

        // Auto-call entry point if defined
        if let Some(ref ep) = self.profile.entry_point.clone() {
            let has_ep = self.defined_globals.contains(ep)
                || (!self.case_sensitive && self.defined_globals.iter().any(|g| g.eq_ignore_ascii_case(ep)));
            if has_ep {
                self.emit_var_get(ep);
                self.emit_u8(Op::call_ref, 0);
                self.emit(Op::drop);
            }
        }

        self.emit(Op::null);
        self.emit(Op::halt);
        let locals = self.scope().next_slot;
        self.chunks[0].local_count = locals;
        common::bundle::finalize_with_stdlib(&mut self.chunks);
        Ok(self.chunks)
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
    fn emit_const(&mut self, val: Value) { let idx = self.chunks[self.current].add_constant(val); self.emit_u16(Op::r#const, idx); }
    fn emit_jump(&mut self, op: Op) -> usize { let l = self.line; self.chunks[self.current].emit_jump(op, l) }
    fn patch_jump(&mut self, o: usize) { self.chunks[self.current].patch_jump(o); }
    fn emit_loop(&mut self, t: usize) { let l = self.line; self.chunks[self.current].emit_loop(t, l); }
    fn current_offset(&self) -> usize { self.chunks[self.current].current_offset() }
    fn str_const(&mut self, s: &str) -> u16 { self.chunks[self.current].add_constant(Value::String(Arc::from(s))) }

    fn import(&mut self, module: &str, name: &str) -> u16 { self.chunks[0].add_import(module, name) }
    fn emit_host_call(&mut self, idx: u16, argc: u8) {
        let l = self.line;
        self.chunks[self.current].emit_op_u16(Op::call_import, idx, l);
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
            self.emit_u16(Op::local_get, slot);
            return;
        }
        if !self.case_sensitive {
            if let Some(slot) = self.scope().resolve_ci(name) {
                self.emit_u16(Op::local_get, slot);
                return;
            }
        }
        // Upvalue (closure capture)
        if self.scopes.len() > 1 {
            if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, name) {
                self.emit_u8(Op::upvalue_get, uv);
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
                self.emit_u16(Op::local_get, self_slot);
                let cname = self.canon(name);
                let idx = self.str_const(&cname);
                self.emit_u16(Op::struct_get, idx);
                return;
            }
        }
        // Global — canonicalize name for case-insensitive languages
        let cname = self.canon(name);
        let idx = self.str_const(&cname);
        self.emit_u16(Op::global_get, idx);
    }

    fn emit_var_set(&mut self, name: &str) {
        // Local
        if let Some(slot) = self.scope().resolve(name) {
            self.emit_u16(Op::local_set, slot); self.emit(Op::drop);
            return;
        }
        if !self.case_sensitive {
            if let Some(slot) = self.scope().resolve_ci(name) {
                self.emit_u16(Op::local_set, slot); self.emit(Op::drop);
                return;
            }
        }
        // Upvalue (closure capture)
        if self.scopes.len() > 1 {
            if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, name) {
                self.emit_u8(Op::upvalue_set, uv);
                return;
            }
        }
        // Global — canonicalize name for case-insensitive languages
        let cname = self.canon(name);
        let idx = self.str_const(&cname); self.emit_u16(Op::global_set, idx); self.emit(Op::drop);
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
                        self.emit_u8(Op::call_ref, 0);
                        self.emit(Op::drop);
                    }
                    // obj.method as statement → method call with 0 args
                    ExprKind::Member { object, field, .. } => {
                        self.compile_expr(object)?;
                        let field_name = self.canon(field);
                        let prop = self.str_const(&field_name);
                        self.emit(Op::dup);
                        self.emit_u16(Op::struct_get, prop);
                        let fn_tmp = self.scope_mut().define("__fn");
                        self.emit_u16(Op::local_set, fn_tmp); self.emit(Op::drop);
                        let obj_tmp = self.scope_mut().define("__obj");
                        self.emit_u16(Op::local_set, obj_tmp); self.emit(Op::drop);
                        self.emit_u16(Op::local_get, fn_tmp);
                        self.emit_u16(Op::local_get, obj_tmp);
                        self.emit_u8(Op::call_ref, 1);
                        self.emit(Op::drop);
                    }
                    _ => {
                        self.compile_expr(expr)?;
                        self.emit(Op::drop);
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
                self.compile_expr(value)?;
                for (i, target) in targets.iter().enumerate() {
                    if i < targets.len() - 1 { self.emit(Op::dup); }
                    self.compile_assign_target(target)?;
                }
            }

            StmtKind::CompoundAssign { target, op, value } => {
                // Load current value
                self.compile_expr(target)?;
                self.compile_expr(value)?;
                self.compile_compound_op(op);
                self.compile_assign_target(target)?;
            }

            // ── If / Elif / Else ────────────────────────────────────────
            StmtKind::If { cond, then_body, elifs, else_body } => {
                self.compile_expr(cond)?;
                self.emit(Op::dyn_to_bool);
                let else_j = self.emit_jump(Op::br_if_false);
                for s in then_body { self.compile_stmt(s)?; }
                let mut end_jumps = vec![];
                if !elifs.is_empty() || else_body.is_some() {
                    end_jumps.push(self.emit_jump(Op::br));
                }
                self.patch_jump(else_j);
                for (elif_cond, elif_body) in elifs {
                    self.compile_expr(elif_cond)?;
                    self.emit(Op::dyn_to_bool);
                    let skip = self.emit_jump(Op::br_if_false);
                    for s in elif_body { self.compile_stmt(s)?; }
                    end_jumps.push(self.emit_jump(Op::br));
                    self.patch_jump(skip);
                }
                if let Some(else_stmts) = else_body {
                    for s in else_stmts { self.compile_stmt(s)?; }
                }
                for j in end_jumps { self.patch_jump(j); }
            }

            // ── While (compiler_common::loops) ─────────────────────────
            StmtKind::While { cond, body, else_body } => {
                let start = common::loops::emit_loop_start(self.chunk());
                self.loops.push(LoopCtx { break_patches: vec![], continue_target: start, continue_patches: vec![], continue_needs_patch: false, label: None });
                self.compile_expr(cond)?;
                let line = self.line;
                let exit = common::loops::emit_loop_cond(self.chunk(), line);
                for s in body { self.compile_stmt(s)?; }
                let line = self.line;
                common::loops::emit_loop_end(self.chunk(), start, exit, line);
                if let Some(else_stmts) = else_body {
                    for s in else_stmts { self.compile_stmt(s)?; }
                }
                let ctx = self.loops.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
            }

            // ── For C-style (compiler_common::loops) ────────────────────
            StmtKind::For { init, cond, update, body } => {
                self.scope_mut().begin_scope();
                if let Some(init_stmt) = init { self.compile_stmt(init_stmt)?; }
                let start = common::loops::emit_loop_start(self.chunk());
                let has_update = update.is_some();
                self.loops.push(LoopCtx { break_patches: vec![], continue_target: start, continue_patches: vec![], continue_needs_patch: has_update, label: None });
                if let Some(c) = cond {
                    self.compile_expr(c)?;
                } else {
                    self.emit(Op::r#true);
                }
                let line = self.line;
                let exit = common::loops::emit_loop_cond(self.chunk(), line);
                for s in body { self.compile_stmt(s)?; }
                let ctx = self.loops.pop().unwrap();
                // Patch continue jumps to land here (at the update expression)
                for p in ctx.continue_patches { self.patch_jump(p); }
                if let Some(u) = update { self.compile_expr(u)?; self.emit(Op::drop); }
                let line = self.line;
                common::loops::emit_loop_end(self.chunk(), start, exit, line);
                for p in ctx.break_patches { self.patch_jump(p); }
                self.scope_mut().end_scope();
            }

            // ── ForIn / ForOf ───────────────────────────────────────────
            StmtKind::ForIn { var, iter, body, else_body, .. } => {
                self.compile_expr(iter)?;
                let arr_slot = self.scope_mut().define("__forin_arr");
                self.emit_u16(Op::local_set, arr_slot); self.emit(Op::drop);
                let idx_slot = self.scope_mut().define("__forin_idx");
                let line = self.line;
                let (loop_start, exit_jump) = common::loops::emit_for_in_start(
                    &mut self.chunks[self.current], arr_slot, idx_slot, line,
                );
                // Define loop variable and set it
                let var_slot = self.scope_mut().define(var);
                self.emit_u16(Op::local_set, var_slot); self.emit(Op::drop);
                // continue must jump to the increment, not the condition check.
                // Use a placeholder — we'll set it after the body to the increment location.
                self.loops.push(LoopCtx { break_patches: vec![], continue_target: loop_start, continue_patches: vec![], continue_needs_patch: true, label: None });
                for s in body { self.compile_stmt(s)?; }
                let ctx = self.loops.pop().unwrap();
                // Patch continue forward jumps to land here (at the increment)
                for p in ctx.continue_patches {
                    self.patch_jump(p);
                }
                common::loops::emit_for_in_end(
                    &mut self.chunks[self.current], idx_slot, loop_start, exit_jump, line,
                );
                if let Some(else_stmts) = else_body {
                    for s in else_stmts { self.compile_stmt(s)?; }
                }
                for p in ctx.break_patches { self.patch_jump(p); }
            }

            // ── DoWhile / Do Until ──────────────────────────────────────
            // ── DoWhile (compiler_common::loops) ────────────────────────
            StmtKind::DoWhile { body, cond, until } => {
                let start = common::loops::emit_do_loop_start(self.chunk());
                self.loops.push(LoopCtx { break_patches: vec![], continue_target: start, continue_patches: vec![], continue_needs_patch: false, label: None });
                for s in body { self.compile_stmt(s)?; }
                self.compile_expr(cond)?;
                let line = self.line;
                common::loops::emit_do_loop_end(self.chunk(), start, *until, line);
                let ctx = self.loops.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
            }

            // ── Switch / Select Case ────────────────────────────────────
            StmtKind::Switch { expr, cases, default } => {
                self.compile_expr(expr)?;
                // Push a loop context so `break;` inside case bodies collects its
                // jump patch and we can resolve it to the end of the switch.
                self.loops.push(LoopCtx { break_patches: vec![], continue_target: 0, continue_patches: vec![], continue_needs_patch: false, label: None });
                let mut end_patches = Vec::new();
                for case in cases {
                    let mut match_patches = Vec::new();
                    for cond in &case.conditions {
                        match cond {
                            CaseCondition::Value(val) => {
                                self.emit(Op::dup);
                                self.compile_expr(val)?;
                                self.emit(Op::dyn_eq);
                                match_patches.push(self.emit_jump(Op::br_if_true));
                            }
                            CaseCondition::Range { from, to } => {
                                // val >= from && val <= to
                                self.emit(Op::dup);
                                self.compile_expr(from)?;
                                self.emit(Op::dyn_ge);
                                let first = self.emit_jump(Op::br_if_false);
                                self.emit(Op::dup);
                                self.compile_expr(to)?;
                                self.emit(Op::dyn_le);
                                match_patches.push(self.emit_jump(Op::br_if_true));
                                self.patch_jump(first);
                            }
                            CaseCondition::Comparison { op, expr: cmp_expr } => {
                                self.emit(Op::dup);
                                self.compile_expr(cmp_expr)?;
                                match op {
                                    ComparisonOp::Eq => self.emit(Op::dyn_eq),
                                    ComparisonOp::NotEq => self.emit(Op::dyn_ne),
                                    ComparisonOp::Lt => self.emit(Op::dyn_lt),
                                    ComparisonOp::LtEq => self.emit(Op::dyn_le),
                                    ComparisonOp::Gt => self.emit(Op::dyn_gt),
                                    ComparisonOp::GtEq => self.emit(Op::dyn_ge),
                                }
                                match_patches.push(self.emit_jump(Op::br_if_true));
                            }
                        }
                    }
                    let skip = self.emit_jump(Op::br);
                    for p in match_patches { self.patch_jump(p); }
                    for s in &case.body { self.compile_stmt(s)?; }
                    end_patches.push(self.emit_jump(Op::br));
                    self.patch_jump(skip);
                }
                if let Some(def) = default {
                    for s in def { self.compile_stmt(s)?; }
                }
                for p in end_patches { self.patch_jump(p); }
                let ctx = self.loops.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
                self.emit(Op::drop); // drop the switch expression
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
                let skip_to_finally = self.emit_jump(Op::br);
                common::errors::patch_catch(&mut self.chunks[self.current], catch_jump);
                if catches.is_empty() {
                    self.emit(Op::drop);
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
                                self.emit(Op::dup);
                                let line = self.line;
                                let key = self.str_const("__exception_type");
                                self.chunks[self.current]
                                    .emit_op_u16(Op::struct_get, key, line);
                                let v = self.str_const(ty);
                                self.chunks[self.current]
                                    .emit_op_u16(Op::r#const, v, line);
                                self.emit(Op::dyn_eq);
                                to_body.push(self.emit_jump(Op::br_if_true));
                            }
                            skip_arm = Some(self.emit_jump(Op::br));
                            for p in to_body { self.patch_jump(p); }
                        }

                        if let Some(ref var) = c.var_name {
                            let slot = self.scope_mut().define(var);
                            self.emit_u16(Op::local_set, slot);
                            self.emit(Op::drop);
                        } else {
                            self.emit(Op::drop);
                        }
                        for s in &c.body { self.compile_stmt(s)?; }
                        end_patches.push(self.emit_jump(Op::br));

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
                if let Some(v) = val {
                    self.compile_expr(v)?;
                } else if let Some(rs) = self.current_result_slot {
                    // ResultSlot return: return the result slot value
                    self.emit_u16(Op::local_get, rs);
                } else {
                    self.emit(Op::null);
                }
                self.emit(Op::r#return);
            }

            // ── Break ───────────────────────────────────────────────────
            StmtKind::Break(target) => {
                match target {
                    BreakTarget::Implicit | BreakTarget::Kind(_) | BreakTarget::Level(_) => {
                        let p = self.emit_jump(Op::br);
                        if let Some(ctx) = self.loops.last_mut() { ctx.break_patches.push(p); }
                    }
                    BreakTarget::Label(label) => {
                        let p = self.emit_jump(Op::br);
                        // Find labeled loop
                        for ctx in self.loops.iter_mut().rev() {
                            if ctx.label.as_deref() == Some(label.as_str()) {
                                ctx.break_patches.push(p);
                                break;
                            }
                        }
                    }
                    BreakTarget::Value(expr) => {
                        // Ruby: break value — emit value then break
                        self.compile_expr(expr)?;
                        self.emit(Op::r#return);
                    }
                }
            }

            // ── Continue ────────────────────────────────────────────────
            StmtKind::Continue(target) => {
                match target {
                    ContinueTarget::Implicit | ContinueTarget::Kind(_) | ContinueTarget::Level(_) => {
                        let needs_patch = self.loops.last().map(|c| c.continue_needs_patch).unwrap_or(false);
                        let target = self.loops.last().map(|c| c.continue_target).unwrap_or(0);
                        if needs_patch {
                            let jump = self.emit_jump(Op::br);
                            if let Some(ctx) = self.loops.last_mut() {
                                ctx.continue_patches.push(jump);
                            }
                        } else if self.loops.last().is_some() {
                            self.emit_loop(target);
                        }
                    }
                    ContinueTarget::Label(label) => {
                        let mut found_idx = None;
                        let mut needs_patch = false;
                        let mut target = 0;
                        for (i, ctx) in self.loops.iter().enumerate().rev() {
                            if ctx.label.as_deref() == Some(label.as_str()) {
                                found_idx = Some(i);
                                needs_patch = ctx.continue_needs_patch;
                                target = ctx.continue_target;
                                break;
                            }
                        }
                        if let Some(idx) = found_idx {
                            if needs_patch {
                                let jump = self.emit_jump(Op::br);
                                self.loops[idx].continue_patches.push(jump);
                            } else {
                                self.emit_loop(target);
                            }
                        }
                    }
                }
            }

            // ── Throw ───────────────────────────────────────────────────
            StmtKind::Throw { expr, cause: _ } => {
                if let Some(v) = expr { self.compile_expr(v)?; } else { self.emit(Op::null); }
                let line = self.line;
                common::errors::emit_throw(&mut self.chunks[self.current], line);
            }

            // ── Function declaration ────────────────────────────────────
            StmtKind::FunctionDecl { name, params, return_type, body, modifiers: _, handles, is_async: _, is_generator: _, is_sub } => {
                self.compile_function_decl(name, params, return_type, body, *is_sub, handles)?;
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
                self.emit_u16(Op::struct_new, 0);
                let mut next_val = 0i64;
                for m in members {
                    self.emit(Op::dup);
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
                    self.emit_u16(Op::struct_set, key);
                    self.emit(Op::drop);
                    // Register member → enum type for bare-name resolution
                    self.enum_members.insert(mname, cname.clone());
                }
                let gidx = self.str_const(&cname);
                self.emit_u16(Op::global_set, gidx);
                self.emit(Op::drop);
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
                                self.emit(Op::null);
                            }
                            let cname = self.canon(fname);
                            let idx = self.str_const(&cname);
                            self.emit_u16(Op::global_set, idx);
                            self.emit(Op::drop);
                            self.defined_globals.insert(cname.clone());
                            member_names.push(cname);
                        }
                        ClassMember::Const { name: cname, value, .. } => {
                            self.compile_expr(value)?;
                            let cn = self.canon(cname);
                            let idx = self.str_const(&cn);
                            self.emit_u16(Op::global_set, idx);
                            self.emit(Op::drop);
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
                self.emit_u16(Op::struct_new, 0);
                for mn in &member_names {
                    self.emit(Op::dup);
                    let gidx = self.str_const(mn);
                    self.emit_u16(Op::global_get, gidx);
                    let key = self.str_const(mn);
                    self.emit_u16(Op::struct_set, key);
                    self.emit(Op::drop);
                    // Register bare member → module name for qualified resolution
                    self.enum_members.insert(mn.clone(), module_name.clone());
                }
                let mod_idx = self.str_const(&module_name);
                self.emit_u16(Op::global_set, mod_idx);
                self.emit(Op::drop);
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
                self.emit_u16(Op::struct_new, 0);
                for mn in &member_names {
                    self.emit(Op::dup);
                    let gidx = self.str_const(mn);
                    self.emit_u16(Op::global_get, gidx);
                    let key = self.str_const(mn);
                    self.emit_u16(Op::struct_set, key);
                    self.emit(Op::drop);
                }
                let ns_idx = self.str_const(&ns_name);
                self.emit_u16(Op::global_set, ns_idx);
                self.emit(Op::drop);
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
                        self.emit_u16(Op::local_set, slot); self.emit(Op::drop);
                    } else {
                        self.emit(Op::drop);
                    }
                }
                for s in body { self.compile_stmt(s)?; }
            }

            // ── Using ───────────────────────────────────────────────────
            StmtKind::Using { var, resource, body } => {
                self.compile_expr(resource)?;
                let slot = self.scope_mut().define(var);
                self.emit_u16(Op::local_set, slot); self.emit(Op::drop);
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
                        self.emit_u16(Op::local_set, old_slot); self.emit(Op::drop);
                        // new_len = N + 1
                        self.compile_expr(size_expr)?;
                        self.emit_const(Value::F64(1.0));
                        self.emit(Op::dyn_add);
                        self.emit_u16(Op::local_set, new_len_slot); self.emit(Op::drop);
                        // new = array_new_default(new_len)
                        self.emit_u16(Op::local_get, new_len_slot);
                        self.emit(Op::array_new_default);
                        self.emit_u16(Op::local_set, new_slot); self.emit(Op::drop);

                        // Iterate old array with the canonical for-in helper.
                        // The helper leaves [element] on the stack each pass
                        // and exposes the index in `idx_slot`.
                        let (loop_start, exit_jump) = common::loops::emit_for_in_start(
                            self.chunk(), old_slot, idx_slot, line);
                        // Stack: [element]. If idx >= new_len, drop and break
                        // (don't write past the new array). Otherwise
                        // new[idx] = element.
                        self.emit_u16(Op::local_get, idx_slot);
                        self.emit_u16(Op::local_get, new_len_slot);
                        self.emit(Op::dyn_lt);
                        let in_bounds = self.emit_jump(Op::br_if_true);
                        // out of bounds: drop the element from for_in_start
                        self.emit(Op::drop);
                        let after = self.emit_jump(Op::br);
                        self.patch_jump(in_bounds);
                        // in bounds: new[idx] = element
                        // Stack currently has [element]. Build [new, idx, element].
                        let elem_slot = self.scope_mut().define("__redim_el");
                        self.emit_u16(Op::local_set, elem_slot); self.emit(Op::drop);
                        self.emit_u16(Op::local_get, new_slot);
                        self.emit_u16(Op::local_get, idx_slot);
                        self.emit_u16(Op::local_get, elem_slot);
                        self.emit(Op::array_set);
                        self.emit(Op::drop);
                        self.patch_jump(after);

                        common::loops::emit_for_in_end(
                            self.chunk(), idx_slot, loop_start, exit_jump, line);

                        // arr = new
                        self.emit_u16(Op::local_get, new_slot);
                        self.emit_var_set(array);
                    } else {
                        self.compile_expr(size_expr)?;
                        self.emit_const(Value::F64(1.0));
                        self.emit(Op::dyn_add);
                        self.emit(Op::array_new_default);
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
                self.emit(Op::drop); // statement: discard host call result
            }

            StmtKind::RemoveHandler { control, event, handler } => {
                let line = self.line;
                let unbind_idx = self.import("vybe:gui", common::gui::HOST_FN_UNBIND_EVENT);
                self.emit_event_control_key(control, line)?;
                self.emit_const(Value::String(Arc::from(event.as_str())));
                self.compile_expr(handler)?;
                common::gui::emit_unbind_event(self.chunk(), unbind_idx, line);
                self.emit(Op::drop); // statement: discard host call result
            }

            StmtKind::RaiseEvent { event_name, args } => {
                let line = self.line;
                let raise_idx = self.import("vybe:gui", common::gui::HOST_FN_RAISE_EVENT);
                for a in args { self.compile_expr(a)?; }
                self.emit_const(Value::String(Arc::from(event_name.as_str())));
                common::gui::emit_raise_event(self.chunk(), raise_idx, (args.len() + 1) as u8, line);
                self.emit(Op::drop); // statement: discard host call result
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
                self.emit(Op::drop);
            }
            StmtKind::CloseFile(file_num) => {
                if let Some(fnum) = file_num {
                    self.compile_expr(fnum)?;
                } else {
                    self.emit(Op::null);
                }
                let idx = self.import("wasi:filesystem", "closeFile");
                self.emit_host_call(idx, 1);
                self.emit(Op::drop);
            }
            StmtKind::PrintFile { file_number, items } => {
                self.compile_expr(file_number)?;
                for item in items { self.compile_expr(item)?; }
                let idx = self.import("wasi:filesystem", "printFile");
                self.emit_host_call(idx, (items.len() + 1) as u8);
                self.emit(Op::drop);
            }
            StmtKind::WriteFile { file_number, items } => {
                self.compile_expr(file_number)?;
                for item in items { self.compile_expr(item)?; }
                let idx = self.import("wasi:filesystem", "writeFile");
                self.emit_host_call(idx, (items.len() + 1) as u8);
                self.emit(Op::drop);
            }
            StmtKind::InputFile { file_number, variables } => {
                self.compile_expr(file_number)?;
                let idx = self.import("wasi:filesystem", "inputFile");
                self.emit_host_call(idx, 1);
                if let Some(first) = variables.first() {
                    self.emit_var_set(first);
                } else {
                    self.emit(Op::drop);
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
                    self.emit_u16(Op::global_set, idx);
                    self.emit(Op::drop);
                }
            }

            // ── Labeled statement ───────────────────────────────────────
            StmtKind::Labeled { label, body } => {
                // Push label onto the next loop context if the body is a loop
                // For non-loop bodies, just compile
                if let Some(ctx) = self.loops.last_mut() {
                    ctx.label = Some(label.clone());
                }
                self.compile_stmt(body)?;
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
                            self.emit(Op::null);
                            let field_name = self.canon(field);
                            let idx = self.str_const(&field_name);
                            self.emit_u16(Op::struct_set, idx);
                            self.emit(Op::drop);
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
                self.emit(Op::dyn_to_bool);
                let ok = self.emit_jump(Op::br_if_true);
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
                self.emit_u16(Op::local_set, subject_slot); self.emit(Op::drop);
                let mut end_patches = Vec::new();
                for case in cases {
                    // Simplified: match on value patterns only, wildcard always matches
                    let skip = match &case.pattern {
                        Pattern::Value(val) => {
                            self.emit_u16(Op::local_get, subject_slot);
                            self.compile_expr(val)?;
                            self.emit(Op::dyn_eq);
                            Some(self.emit_jump(Op::br_if_false))
                        }
                        Pattern::Wildcard => None,
                        Pattern::As { name: Some(name), .. } => {
                            // Bind subject to name
                            self.emit_u16(Op::local_get, subject_slot);
                            let slot = self.scope_mut().define(name);
                            self.emit_u16(Op::local_set, slot); self.emit(Op::drop);
                            None
                        }
                        _ => None, // Other patterns: always match (simplified)
                    };
                    if let Some(guard) = &case.guard {
                        self.compile_expr(guard)?;
                        self.emit(Op::dyn_to_bool);
                        let guard_skip = self.emit_jump(Op::br_if_false);
                        for s in &case.body { self.compile_stmt(s)?; }
                        end_patches.push(self.emit_jump(Op::br));
                        self.patch_jump(guard_skip);
                    } else {
                        for s in &case.body { self.compile_stmt(s)?; }
                        end_patches.push(self.emit_jump(Op::br));
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
                    // Array with bounds: Dim arr(N)
                    if let Some(size_expr) = bounds.first() {
                        self.compile_expr(size_expr)?;
                        self.emit_const(Value::F64(1.0));
                        self.emit(Op::dyn_add);
                        self.emit(Op::array_new_default);
                    } else {
                        self.emit(Op::null);
                    }
                } else {
                    // Default values based on type hint
                    match decl.type_hint.as_deref().map(|s| s.to_lowercase()).as_deref() {
                        Some("integer") | Some("int") | Some("longint") | Some("real") | Some("double") | Some("float") => {
                            self.emit(Op::f64_const_0);
                        }
                        Some("boolean") | Some("bool") => self.emit(Op::r#false),
                        Some("string") => self.emit_const(Value::String(Arc::from(""))),
                        _ => self.emit(Op::null),
                    }
                }
                // Top-level / hoisted vars → globals
                let is_toplevel = self.scopes.len() == 1;
                let is_hoisted = *kind == VarDeclKind::Var && self.profile.hoist_var;
                if is_toplevel || (is_hoisted && self.scopes.len() <= 2) {
                    let cn = self.canon(name);
                    let idx = self.str_const(&cn);
                    self.emit_u16(Op::global_set, idx);
                    self.emit(Op::drop);
                    self.defined_globals.insert(cn);
                } else {
                    let slot = self.scope_mut().define(name);
                    self.emit_u16(Op::local_set, slot);
                    self.emit(Op::drop);
                }
            }
            BindingPattern::Object(props) => {
                // Destructuring: let { a, b } = expr
                if let Some(ref init_expr) = decl.init {
                    self.compile_expr(init_expr)?;
                    let obj_slot = self.scope_mut().define("__destruct_obj");
                    self.emit_u16(Op::local_set, obj_slot); self.emit(Op::drop);
                    for prop in props {
                        self.emit_u16(Op::local_get, obj_slot);
                        let key = self.str_const(&prop.key);
                        self.emit_u16(Op::struct_get, key);
                        if let Some(ref default) = prop.default {
                            // If value is null, use default
                            self.emit(Op::dup);
                            self.emit(Op::ref_is_null);
                            let has_val = self.emit_jump(Op::br_if_false);
                            self.emit(Op::drop);
                            self.compile_expr(default)?;
                            self.patch_jump(has_val);
                        }
                        let bind_name = if let Some(BindingPattern::Ident(ref n)) = prop.value {
                            n.as_str()
                        } else {
                            &prop.key
                        };
                        let slot = self.scope_mut().define(bind_name);
                        self.emit_u16(Op::local_set, slot); self.emit(Op::drop);
                    }
                }
            }
            BindingPattern::Array(elems) => {
                // Destructuring: let [a, b] = expr
                if let Some(ref init_expr) = decl.init {
                    self.compile_expr(init_expr)?;
                    let arr_slot = self.scope_mut().define("__destruct_arr");
                    self.emit_u16(Op::local_set, arr_slot); self.emit(Op::drop);
                    for (i, elem) in elems.iter().enumerate() {
                        match elem {
                            ArrayPatternElem::Pattern(BindingPattern::Ident(name), default) => {
                                self.emit_u16(Op::local_get, arr_slot);
                                self.emit_const(Value::F64(i as f64));
                                self.emit(Op::array_get);
                                if let Some(def) = default {
                                    self.emit(Op::dup);
                                    self.emit(Op::ref_is_null);
                                    let has_val = self.emit_jump(Op::br_if_false);
                                    self.emit(Op::drop);
                                    self.compile_expr(def)?;
                                    self.patch_jump(has_val);
                                }
                                let slot = self.scope_mut().define(name);
                                self.emit_u16(Op::local_set, slot); self.emit(Op::drop);
                            }
                            ArrayPatternElem::Rest(name) => {
                                // ...rest: slice from current index
                                self.emit_u16(Op::local_get, arr_slot);
                                self.emit_const(Value::F64(i as f64));
                                // end = arr.length
                                self.emit_u16(Op::local_get, arr_slot);
                                self.emit(Op::array_length);
                                let line = self.line;
                                common::collections::emit_slice(self.chunk(), line);
                                let slot = self.scope_mut().define(name);
                                self.emit_u16(Op::local_set, slot); self.emit(Op::drop);
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
                            self.emit_u16(Op::local_set, rs);
                            self.emit(Op::drop);
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
                        self.emit_u16(Op::local_set, tmp); self.emit(Op::drop);
                        self.emit_u16(Op::local_get, slot);
                        self.emit_u16(Op::local_get, tmp);
                        let field_name = self.canon(name);
                        let idx = self.str_const(&field_name);
                        self.emit_u16(Op::struct_set, idx);
                        self.emit(Op::drop);
                        return Ok(());
                    }
                }
                self.emit_var_set(name);
            }
            ExprKind::Member { object, field, .. } => {
                let tmp = self.scope_mut().define("__tmp");
                self.emit_u16(Op::local_set, tmp); self.emit(Op::drop);
                self.compile_expr(object)?;
                self.emit_u16(Op::local_get, tmp);
                let field_name = self.canon(field);
                let idx = self.str_const(&field_name);
                self.emit_u16(Op::struct_set, idx);
                self.emit(Op::drop);
            }
            ExprKind::Index { object, index } => {
                let tmp = self.scope_mut().define("__tmp");
                self.emit_u16(Op::local_set, tmp); self.emit(Op::drop);
                self.compile_expr(object)?;
                self.compile_expr(index)?;
                self.emit_u16(Op::local_get, tmp);
                self.emit(Op::array_set);
                self.emit(Op::drop);
            }
            // VB: arr(idx) = val — Call used as index because () is both call and index
            ExprKind::Call { callee, args, .. } if args.len() == 1 => {
                let tmp = self.scope_mut().define("__tmp");
                self.emit_u16(Op::local_set, tmp); self.emit(Op::drop);
                self.compile_expr(callee)?;
                self.compile_expr(&args[0].value)?;
                self.emit_u16(Op::local_get, tmp);
                self.emit(Op::array_set);
                self.emit(Op::drop);
            }
            ExprKind::Destructure(pattern) => {
                // Destructuring assignment
                match pattern {
                    DestructurePattern::Object(props) => {
                        let obj_slot = self.scope_mut().define("__destruct_obj");
                        self.emit_u16(Op::local_set, obj_slot); self.emit(Op::drop);
                        for prop in props {
                            self.emit_u16(Op::local_get, obj_slot);
                            let key = self.str_const(&prop.key);
                            self.emit_u16(Op::struct_get, key);
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
                        self.emit_u16(Op::local_set, arr_slot); self.emit(Op::drop);
                        for (i, elem) in elems.iter().enumerate() {
                            match elem {
                                ArrayPatternElem::Pattern(BindingPattern::Ident(name), _) => {
                                    self.emit_u16(Op::local_get, arr_slot);
                                    self.emit_const(Value::F64(i as f64));
                                    self.emit(Op::array_get);
                                    self.emit_var_set(name);
                                }
                                ArrayPatternElem::Rest(name) => {
                                    self.emit_u16(Op::local_get, arr_slot);
                                    self.emit_const(Value::F64(i as f64));
                                    self.emit_u16(Op::local_get, arr_slot);
                                    self.emit(Op::array_length);
                                    let line = self.line;
                                    common::collections::emit_slice(self.chunk(), line);
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
    // Function declaration compilation
    // ════════════════════════════════════════════════════════════════════════

    fn compile_function_decl(
        &mut self, name: &str, params: &[Param], return_type: &Option<String>,
        body: &[Statement], is_sub: bool, handles: &[String],
    ) -> Result<(), String> {
        let cname = self.canon(name);
        self.defined_globals.insert(cname.clone());
        self.defined_functions.insert(cname.clone());
        let name = &cname;

        let arity: u8 = params.len() as u8;
        let func_idx = self.chunks.len();
        let chunk = common::functions::create_function_chunk(name, arity);
        self.chunks.push(chunk);
        self.scopes.push(Scope::new_function());
        let saved = self.current;
        self.current = func_idx;

        // Define params
        for p in params {
            self.scope_mut().define(&p.name);
            // Default parameters
            if let Some(ref default) = p.default {
                let slot = self.scope().resolve(&p.name).unwrap();
                self.emit_u16(Op::local_get, slot);
                self.emit(Op::ref_is_null);
                let has_val = self.emit_jump(Op::br_if_false);
                self.compile_expr(default)?;
                self.emit_u16(Op::local_set, slot); self.emit(Op::drop);
                self.patch_jump(has_val);
            }
        }

        // Result slot for functions with return type (Pascal/VB Function).
        // The slot name is profile-driven so VB can keep it internal
        // (`__result__`) and avoid shadowing user classes named `Result`,
        // while Pascal keeps it as `Result` (user-visible per Pascal idiom).
        let result_slot = if return_type.is_some() && self.profile.function_return == ReturnStyle::ResultSlot {
            let slot_name = self.profile.result_slot_name.clone();
            let rs = self.scope_mut().define(&slot_name);
            self.emit(Op::null); self.emit_u16(Op::local_set, rs); self.emit(Op::drop);
            Some(rs)
        } else {
            None
        };

        let saved_fn = self.current_func_name.take();
        let saved_rs = self.current_result_slot.take();
        self.current_func_name = Some(name.to_string());
        self.current_result_slot = result_slot;

        for s in body { self.compile_stmt(s)?; }

        self.current_func_name = saved_fn;
        self.current_result_slot = saved_rs;

        if let Some(rs) = result_slot {
            self.emit_u16(Op::local_get, rs);
            self.emit(Op::r#return);
        } else {
            let line = self.line;
            common::functions::emit_function_epilogue(&mut self.chunks[func_idx], line);
        }

        let locals = self.scope().next_slot;
        self.chunks[func_idx].local_count = locals;
        let uvs = self.scopes.last().unwrap().upvalues.clone();
        self.scopes.pop();
        self.current = saved;

        let line = self.line;
        common::functions::emit_ref_func(&mut self.chunks[self.current], func_idx, uvs.len() as u8, line);
        for uv in &uvs {
            self.chunks[self.current].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current].emit(uv.index, line);
        }
        let idx = self.str_const(name);
        self.emit_u16(Op::global_set, idx);
        self.emit(Op::drop);

        // VB `Handles ctrl.Event` clause on a top-level Sub: register the
        // event handler with the canonical GUI binding. The same canonical
        // emit path serves C# `+=`, JS `addEventListener`, etc.
        for handle in handles {
            let parts: Vec<&str> = handle.splitn(2, '.').collect();
            if parts.len() == 2 {
                let line = self.line;
                let bind_idx = self.import("vybe:gui", common::gui::HOST_FN_BIND_EVENT);
                self.emit_var_get(parts[0]);
                common::gui::emit_get_control_name(self.chunk(), line);
                let ev = parts[1].to_lowercase();
                self.emit_const(Value::String(Arc::from(ev.as_str())));
                self.emit_var_get(name);
                common::gui::emit_bind_event(self.chunk(), bind_idx, line);
                self.emit(Op::drop); // statement: discard host call result
            }
        }

        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Class compilation
    // ════════════════════════════════════════════════════════════════════════

    fn compile_class(&mut self, name: &str, parent: &Option<String>, members: &[ClassMember]) -> Result<(), String> {
        let self_kw = self.profile.self_keyword.clone();
        let ctor_name = self.profile.constructor_name.clone();
        let result_style = self.profile.function_return.clone();

        // Collect fields and initializers (separate instance vs static)
        // Auto-properties are treated as plain fields (matches old C# compiler).
        let mut fields = Vec::new();
        let mut field_inits: Vec<(String, Option<Expression>)> = Vec::new();
        let mut static_field_inits: Vec<(String, Option<Expression>)> = Vec::new();
        for m in members {
            if let ClassMember::Field { name: fname, init, modifiers, .. } = m {
                let fname = self.canon(fname);
                if modifiers.is_static {
                    static_field_inits.push((fname, init.clone()));
                } else {
                    fields.push(fname.clone());
                    field_inits.push((fname, init.clone()));
                }
            }
            if let ClassMember::Property { name: pname, is_auto, modifiers, .. } = m {
                if *is_auto {
                    let pname_canon = self.canon(pname);
                    if modifiers.is_static {
                        if !static_field_inits.iter().any(|(n, _)| n == &pname_canon) {
                            static_field_inits.push((pname_canon, None));
                        }
                    } else if !fields.contains(&pname_canon) {
                        fields.push(pname_canon.clone());
                        field_inits.push((pname_canon, None));
                    }
                }
            }
        }

        // Store field list for implicit self resolution
        self.pending_classes.insert(name.to_string(), PendingClass {
            parent: parent.clone(),
            fields: fields.clone(),
            statics: Vec::new(), // filled after methods are compiled
        });

        // Compile methods (including constructor body)
        // (name, chunk_idx, is_ctor, is_static)
        let mut method_chunks: Vec<(String, usize, bool, bool)> = Vec::new();
        let saved_class = self.current_class.take();
        self.current_class = Some(name.to_string());

        // Pre-register all method names to avoid value-method hijacking
        for m in members {
            if let ClassMember::Method(stmt) = m {
                if let StmtKind::FunctionDecl { name: mname, .. } = &stmt.kind {
                    self.defined_class_methods.insert(self.canon(mname));
                }
            }
            if let ClassMember::Property { name: pname, .. } = m {
                self.defined_class_methods.insert(self.canon(pname));
            }
        }


        for m in members {
            match m {
                ClassMember::Method(stmt) => {
                    if let StmtKind::FunctionDecl { name: mname, params, return_type, body, modifiers, is_sub, .. } = &stmt.kind {
                        // NOTE: do NOT skip empty-body methods. They still need
                        // a chunk + binding so that callers (e.g. an explicit
                        // constructor calling `InitializeComponent()`) can
                        // dispatch through `me.<method>`. Skipping here is what
                        // caused VB Forms tests to fail with "null is not
                        // callable" — the empty `Sub InitializeComponent` was
                        // never bound on `me`.

                        let is_ctor = if self.case_sensitive {
                            mname == &ctor_name || (modifiers.is_static && mname == "new")
                        } else {
                            mname.eq_ignore_ascii_case(&ctor_name)
                            || modifiers.is_static && mname.eq_ignore_ascii_case("new")
                        };

                        let user_params: Vec<&Param> = if self.profile.explicit_self_param {
                            params.iter().skip(1).collect()
                        } else {
                            params.iter().collect()
                        };
                        let arity = (user_params.len() + 1) as u8; // +1 for self

                        let ci = self.chunks.len();
                        let chunk = common::functions::create_function_chunk(mname, arity);
                        self.chunks.push(chunk);
                        self.scopes.push(Scope::new_function());
                        let saved = self.current;
                        self.current = ci;

                        self.scope_mut().define(&self_kw);
                        for p in &user_params { self.scope_mut().define(&p.name); }

                        if is_ctor {
                            for s in body { self.compile_stmt(s)?; }
                            if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                                self.emit_u16(Op::local_get, slot);
                                self.emit(Op::r#return);
                            }
                        } else if return_type.is_some() && result_style == ReturnStyle::ResultSlot {
                            let slot_name = self.profile.result_slot_name.clone();
                            let rs = self.scope_mut().define(&slot_name);
                            self.emit(Op::null); self.emit_u16(Op::local_set, rs); self.emit(Op::drop);
                            let saved_fn = self.current_func_name.take();
                            let saved_rs = self.current_result_slot.take();
                            self.current_func_name = Some(mname.clone());
                            self.current_result_slot = Some(rs);
                            for s in body { self.compile_stmt(s)?; }
                            self.current_func_name = saved_fn;
                            self.current_result_slot = saved_rs;
                            self.emit_u16(Op::local_get, rs);
                            self.emit(Op::r#return);
                        } else {
                            for s in body { self.compile_stmt(s)?; }
                            let line = self.line;
                            common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
                        }

                        let locals = self.scope().next_slot;
                        self.chunks[ci].local_count = locals;
                        self.scopes.pop();
                        self.current = saved;

                        let bound_name = self.canon(mname);
                        method_chunks.push((bound_name, ci, is_ctor, modifiers.is_static));
                    }
                }
                ClassMember::Constructor { .. } => {
                    // Constructor body is handled by the main constructor flow below
                    // (extracted via ctor_body). No separate chunk needed.
                }
                ClassMember::Property { name: pname, getter, setter, is_auto, .. } => {
                    // Auto-properties are handled as plain fields above — skip getter/setter compilation
                    if *is_auto { continue; }
                    let pname_canon = self.canon(pname);

                    // Getter → __get_<prop>
                    if let Some(getter_body) = getter {
                        let get_name = format!("__get_{}", pname_canon);
                        let ci = self.chunks.len();
                        let chunk = common::functions::create_function_chunk(&get_name, 1); // self
                        self.chunks.push(chunk);
                        self.scopes.push(Scope::new_function());
                        let saved = self.current;
                        self.current = ci;
                        self.scope_mut().define(&self_kw);

                        if getter_body.is_empty() {
                            // Auto-property getter: return backing field
                            if let Some(slot) = self.scope().resolve(&self_kw) {
                                self.emit_u16(Op::local_get, slot);
                                let backing = self.str_const(&format!("__{}", pname_canon));
                                self.emit_u16(Op::struct_get, backing);
                                self.emit(Op::r#return);
                            }
                        } else {
                            let slot_name = self.profile.result_slot_name.clone();
                            let rs = self.scope_mut().define(&slot_name);
                            self.emit(Op::null); self.emit_u16(Op::local_set, rs); self.emit(Op::drop);
                            let saved_fn = self.current_func_name.take();
                            let saved_rs = self.current_result_slot.take();
                            self.current_func_name = Some(pname.clone());
                            self.current_result_slot = Some(rs);
                            for s in getter_body { self.compile_stmt(s)?; }
                            self.current_func_name = saved_fn;
                            self.current_result_slot = saved_rs;
                            self.emit_u16(Op::local_get, rs);
                            self.emit(Op::r#return);
                        }

                        let locals = self.scope().next_slot;
                        self.chunks[ci].local_count = locals;
                        self.scopes.pop();
                        self.current = saved;
                        method_chunks.push((get_name, ci, false, false));
                    }

                    // Setter → __set_<prop>
                    if let Some(setter_info) = setter {
                        let set_name = format!("__set_{}", pname_canon);
                        let ci = self.chunks.len();
                        let chunk = common::functions::create_function_chunk(&set_name, 2); // self, value
                        self.chunks.push(chunk);
                        self.scopes.push(Scope::new_function());
                        let saved = self.current;
                        self.current = ci;
                        self.scope_mut().define(&self_kw);
                        self.scope_mut().define(&setter_info.param.name);

                        if setter_info.body.is_empty() {
                            // Auto-property setter: set backing field
                            if let Some(self_slot) = self.scope().resolve(&self_kw) {
                                self.emit_u16(Op::local_get, self_slot);
                                if let Some(val_slot) = self.scope().resolve(&setter_info.param.name) {
                                    self.emit_u16(Op::local_get, val_slot);
                                }
                                let backing = self.str_const(&format!("__{}", pname_canon));
                                self.emit_u16(Op::struct_set, backing);
                                self.emit(Op::drop);
                            }
                        } else {
                            for s in &setter_info.body { self.compile_stmt(s)?; }
                        }

                        let line = self.line;
                        common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
                        let locals = self.scope().next_slot;
                        self.chunks[ci].local_count = locals;
                        self.scopes.pop();
                        self.current = saved;
                        method_chunks.push((set_name, ci, false, false));
                    }
                }
                ClassMember::Const { name: cname, value, .. } => {
                    // Class-level constant → global
                    self.compile_expr(value)?;
                    let global_name = self.canon(&format!("{}.{}", name, cname));
                    let idx = self.str_const(&global_name);
                    self.emit_u16(Op::global_set, idx);
                    self.emit(Op::drop);
                    self.defined_globals.insert(global_name);
                }
                ClassMember::Event { .. } => { /* type-level only */ }
                ClassMember::NestedType(stmt) => { self.compile_stmt(stmt)?; }
                _ => {}
            }
        }

        self.current_class = saved_class;

        // Find constructor body and its user arity
        let ctor = method_chunks.iter().find(|(_, _, is_ctor, _)| *is_ctor);
        let ctor_body: Option<(&Vec<Statement>, &Vec<Param>, Option<&Vec<Expression>>)> = members.iter().find_map(|m| {
            match m {
                ClassMember::Method(stmt) => {
                    if let StmtKind::FunctionDecl { name: mname, params, body, modifiers, .. } = &stmt.kind {
                        let is_ctor = if self.case_sensitive {
                            mname == &ctor_name || (modifiers.is_static && mname == "new")
                        } else {
                            mname.eq_ignore_ascii_case(&ctor_name)
                            || modifiers.is_static && mname.eq_ignore_ascii_case("new")
                        };
                        if is_ctor && !body.is_empty() { return Some((body, params, None)); }
                    }
                    None
                }
                ClassMember::Constructor { params, body, base_args, .. } => Some((body, params, base_args.as_ref())),
                _ => None,
            }
        });

        let user_params: Vec<String> = ctor_body.map(|(_, params, _)| {
            if self.profile.explicit_self_param {
                params.iter().skip(1).map(|p| p.name.clone()).collect()
            } else {
                params.iter().map(|p| p.name.clone()).collect()
            }
        }).unwrap_or_default();
        let user_arity = user_params.len() as u8;

        // ── Single constructor function (not split wrapper + body) ──────
        // This is the ONLY function that `new ClassName(args)` calls.
        // It creates the object, initializes fields, binds methods, runs
        // user constructor body, and returns this.
        let ctor_idx = self.chunks.len();
        let ctor_chunk = common::functions::create_function_chunk(name, user_arity);
        self.chunks.push(ctor_chunk);
        self.scopes.push(Scope::new_function());
        let saved_cur = self.current;
        let saved_class2 = self.current_class.take();
        self.current = ctor_idx;
        self.current_class = Some(name.to_string());

        // Define user params (slot 1..N), then this (slot N+1)
        for p in &user_params { self.scope_mut().define(p); }
        self.scope_mut().define(&self_kw); // this_slot = user_arity + 1
        let this_slot = (user_arity as u16) + 1;

        let is_child = parent.is_some();
        let line = self.line;

        // Separate instance methods from static methods
        let instance_methods: Vec<&(String, usize, bool, bool)> = method_chunks.iter()
            .filter(|(_, _, ic, is_static)| !*ic && !*is_static)
            .collect();
        let static_methods: Vec<&(String, usize, bool, bool)> = method_chunks.iter()
            .filter(|(_, _, ic, is_static)| !*ic && *is_static)
            .collect();
        let instance_method_names: Vec<String> = instance_methods.iter()
            .map(|(n, _, _, _)| n.clone())
            .collect();

        if is_child {
            // ── Child class ─────────────────────────────────────────────
            // For child classes, the constructor body calls super() which
            // creates the object. We run the body FIRST, then bind methods.
            // This works for both explicit super (JS) and implicit (VB/C#)
            // because super() stores the result in this_slot.

            // ── Step 1: Call parent constructor to get the object ────────
            self.emit(Op::null);
            self.emit_u16(Op::local_set, this_slot);
            self.emit(Op::drop);

            if let Some((_, _, base_args)) = &ctor_body {
                // Explicit constructor with : base(args) (C#/VB)
                // Only auto-call parent if base_args is explicitly provided.
                // If base_args is None, the constructor body handles super() itself (JS pattern).
                if let Some(bargs) = base_args {
                    if let Some(parent_name) = parent {
                        let pname = self.canon(parent_name);
                        let pidx = self.str_const(&pname);
                        self.emit_u16(Op::global_get, pidx);
                        for a in *bargs { self.compile_expr(a)?; }
                        self.emit_u8(Op::call, bargs.len() as u8);
                        self.emit_u16(Op::local_set, this_slot);
                        self.emit(Op::drop);
                    }
                }
                // If base_args is None, body will call super() which sets this_slot
            } else {
                // No explicit constructor — auto-call parent with user args
                if let Some(parent_name) = parent {
                    let pname = self.canon(parent_name);
                    let pidx = self.str_const(&pname);
                    self.emit_u16(Op::global_get, pidx);
                    for i in 0..user_arity {
                        self.emit_u16(Op::local_get, (i as u16) + 1);
                    }
                    self.emit_u8(Op::call, user_arity);
                    self.emit_u16(Op::local_set, this_slot);
                    self.emit(Op::drop);
                }
            }

            // Check if this is a C#-style constructor (base_args provided explicitly)
            // vs JS/VB-style (super() called inside body)
            let has_explicit_base = ctor_body.as_ref().map_or(false, |(_, _, ba)| ba.is_some());

            if has_explicit_base || ctor_body.is_none() {
                // C#-style: base call already done above, or no-ctor auto-call done above.
                // Order: re-stamp __type → fields → save base → bind methods → body
                //
                // The parent ctor stamped __type with the parent name. Re-stamp with
                // the child name so `obj is ChildType` returns true.
                self.emit_u16(Op::local_get, this_slot);
                self.emit_const(Value::String(Arc::from(name)));
                let type_key = self.str_const("__type");
                self.emit_u16(Op::struct_set, type_key);
                self.emit(Op::drop);

                for (fname, init) in &field_inits {
                    if let Some(init_expr) = init {
                        common::classes::emit_init_field_start(self.chunk(), this_slot, line);
                        self.compile_expr(init_expr)?;
                        common::classes::emit_init_field_end(self.chunk(), fname, line);
                    } else {
                        common::classes::emit_init_field_null(self.chunk(), this_slot, fname, line);
                    }
                }

                if let Some(parent_name) = parent {
                    let pname = self.canon(parent_name);
                    for method_name in &instance_method_names {
                        common::classes::emit_save_base_method(self.chunk(), this_slot, method_name, line);
                    }
                    common::classes::emit_store_super(self.chunk(), this_slot, &pname, line);
                }

                for (mname, mci, _, _) in &instance_methods {
                    if mname.starts_with("__get_") {
                        let prop = mname.strip_prefix("__get_").unwrap_or(mname);
                        common::classes::emit_bind_getter(self.chunk(), this_slot, prop, *mci, line);
                    } else if mname.starts_with("__set_") {
                        let prop = mname.strip_prefix("__set_").unwrap_or(mname);
                        common::classes::emit_bind_setter(self.chunk(), this_slot, prop, *mci, line);
                    } else {
                        common::classes::emit_bind_method_with_aliases(self.chunk(), this_slot, mname, *mci, line);
                    }
                }

                if let Some((body, _, _)) = ctor_body {
                    for s in body { self.compile_stmt(s)?; }
                }
            } else {
                // JS/VB/Pascal-style: constructor body contains the super() call which
                // sets this_slot.
                //
                // Order matters for VB.NET / Pascal correctness: real .NET binds the
                // class's instance methods on `this` BEFORE the user-visible body
                // runs, so that user body code can call its own methods (e.g.
                // `InitializeComponent()` inside `New()`). The catch is that
                // method binding can only happen AFTER `this` exists, which in
                // this branch means after the super/inherited call.
                //
                // We split the body at the super call:
                //   body[..=super_idx]   — "preamble": runs to set up `this`
                //   bind methods + save base + field inits + re-stamp __type
                //   body[super_idx+1..]  — "main": user code that can now call
                //                          methods on `this`
                //
                // The walker normalization for each language is responsible for
                // putting a super call in the body for `Inherits` classes:
                //   - VB: walker injects `MyBase.New()` (and an
                //     `Me.__control_name = "<lower class name>"` stamp) at the
                //     top of every ctor body for `Inherits` classes — real
                //     VB.NET semantics where the runtime implicitly calls the
                //     parameterless parent ctor.
                //   - C#: walker sets `base_args = Some(_)` (handled by the
                //     C#-style branch above, not this one).
                //   - Pascal: user writes `inherited Create(...)`.
                //   - JS: user writes `super(...)`.
                //
                // We skip null-init for fields with no explicit initializer because
                // the body may have already assigned them (Pascal pattern:
                // `inherited Create(X); FY := Y;`) — and a no-op null-init would
                // clobber that assignment. Fields default to null on dynamic
                // structs anyway, so this is safe.

                let body_stmts: &[Statement] = ctor_body
                    .as_ref()
                    .map(|(b, _, _)| b.as_slice())
                    .unwrap_or(&[]);

                // Find the index of the first super-call statement in the body.
                // This is the "this exists" boundary. Different walkers emit
                // the super call as different node shapes:
                //   - VB walker: `Expr(SuperCall { method: Some("New"), args })`
                //   - C# walker (when not using : base): same shape
                //   - JS walker: `Expr(Call { callee: Super, args })`
                //   - Pascal walker: `Expr(SuperCall { method: Some("Create"), args })`
                //     OR an `inherited Create(...)` that lowers to a Super call.
                // We match all of them to keep this branch language-agnostic.
                //
                // The walker normalization for VB also injects
                // `Me.__control_name = "..."` immediately after; we include it
                // in the preamble so methods bind onto a fully-stamped `this`.
                let is_super_call = |s: &Statement| -> bool {
                    if let StmtKind::Expr(e) = &s.kind {
                        match &e.kind {
                            ExprKind::SuperCall { .. } => true,
                            ExprKind::Call { callee, .. } => matches!(callee.kind, ExprKind::Super),
                            _ => false,
                        }
                    } else {
                        false
                    }
                };
                let super_idx = body_stmts.iter().position(is_super_call);
                let preamble_end = match super_idx {
                    Some(i) => {
                        // Extend through any immediately-following identity stamps
                        // (Me.__control_name = ..., Me.__type = ..., etc.) so the
                        // method binding sees the canonical control name.
                        let mut end = i + 1;
                        while end < body_stmts.len() && is_identity_stamp(&body_stmts[end]) {
                            end += 1;
                        }
                        end
                    }
                    None => 0,
                };

                // Compile preamble (super call + any identity stamps).
                for s in &body_stmts[..preamble_end] {
                    self.compile_stmt(s)?;
                }

                // Re-stamp __type with the child name (the body's super call
                // stamped it with the parent name).
                self.emit_u16(Op::local_get, this_slot);
                self.emit_const(Value::String(Arc::from(name)));
                let type_key2 = self.str_const("__type");
                self.emit_u16(Op::struct_set, type_key2);
                self.emit(Op::drop);

                if let Some(parent_name) = parent {
                    let pname = self.canon(parent_name);
                    for method_name in &instance_method_names {
                        common::classes::emit_save_base_method(self.chunk(), this_slot, method_name, line);
                    }
                    common::classes::emit_store_super(self.chunk(), this_slot, &pname, line);
                }

                for (fname, init) in &field_inits {
                    if let Some(init_expr) = init {
                        common::classes::emit_init_field_start(self.chunk(), this_slot, line);
                        self.compile_expr(init_expr)?;
                        common::classes::emit_init_field_end(self.chunk(), fname, line);
                    }
                }

                for (mname, mci, _, _) in &instance_methods {
                    if mname.starts_with("__get_") {
                        let prop = mname.strip_prefix("__get_").unwrap_or(mname);
                        common::classes::emit_bind_getter(self.chunk(), this_slot, prop, *mci, line);
                    } else if mname.starts_with("__set_") {
                        let prop = mname.strip_prefix("__set_").unwrap_or(mname);
                        common::classes::emit_bind_setter(self.chunk(), this_slot, prop, *mci, line);
                    } else {
                        common::classes::emit_bind_method_with_aliases(self.chunk(), this_slot, mname, *mci, line);
                    }
                }

                // Compile the main body (everything after the preamble).
                for s in &body_stmts[preamble_end..] {
                    self.compile_stmt(s)?;
                }
            }
        } else {
            // ── Base class ──────────────────────────────────────────────
            common::classes::emit_new_typed_object(self.chunk(), this_slot, name, line);

            // Initialize fields
            for (fname, init) in &field_inits {
                if let Some(init_expr) = init {
                    common::classes::emit_init_field_start(self.chunk(), this_slot, line);
                    self.compile_expr(init_expr)?;
                    common::classes::emit_init_field_end(self.chunk(), fname, line);
                } else {
                    common::classes::emit_init_field_null(self.chunk(), this_slot, fname, line);
                }
            }

            // Bind instance methods
            for (mname, mci, _, _) in &instance_methods {
                if mname.starts_with("__get_") {
                    let prop = mname.strip_prefix("__get_").unwrap_or(mname);
                    common::classes::emit_bind_getter(self.chunk(), this_slot, prop, *mci, line);
                } else if mname.starts_with("__set_") {
                    let prop = mname.strip_prefix("__set_").unwrap_or(mname);
                    common::classes::emit_bind_setter(self.chunk(), this_slot, prop, *mci, line);
                } else {
                    common::classes::emit_bind_method_with_aliases(self.chunk(), this_slot, mname, *mci, line);
                }
            }

            // Run user constructor body
            if let Some((body, _, _)) = ctor_body {
                for s in body { self.compile_stmt(s)?; }
            }
        }

        // Check for auto InitializeComponent (.NET pattern)
        let has_init_component = instance_methods.iter()
            .any(|(n, _, _, _)| n.eq_ignore_ascii_case("initializecomponent"));
        let has_explicit_ctor = ctor_body.is_some();
        if has_init_component && !has_explicit_ctor {
            common::classes::emit_auto_init_component(self.chunk(), this_slot, line);
        }

        // Finalize: instanceof chain
        common::classes::emit_instanceof_chain(self.chunk(), this_slot, name, line);
        common::classes::emit_constructor_return(self.chunk(), this_slot, line);

        let locals = self.scope().next_slot;
        self.chunks[ctor_idx].local_count = locals;
        self.scopes.pop();
        self.current = saved_cur;
        self.current_class = saved_class2;

        // Store constructor globally and register type
        let ctor_local = self.scope_mut().define(&format!("__{}_ctor", name));
        common::classes::emit_store_constructor(self.chunk(), name, ctor_idx, ctor_local, line);

        // Initialize static fields on the constructor object
        for (fname, init) in &static_field_inits {
            self.emit_u16(Op::local_get, ctor_local);
            if let Some(init_expr) = init {
                self.compile_expr(init_expr)?;
            } else {
                self.emit(Op::null);
            }
            let fk = self.str_const(fname);
            self.emit_u16(Op::struct_set, fk);
            self.emit(Op::drop);
        }

        // Attach static methods to the constructor object
        let mut all_statics: Vec<(String, usize)> = Vec::new();
        for (mname, mci, _, _) in &static_methods {
            common::classes::emit_attach_static_method(self.chunk(), ctor_local, mname, *mci, line);
            all_statics.push((mname.clone(), *mci));
        }

        // Inherit parent's static methods — walk up the chain via PendingClass
        if let Some(parent_name) = parent {
            let mut current_parent = Some(self.canon(parent_name));
            while let Some(ref pname) = current_parent {
                let parent_statics = self.pending_classes.get(pname.as_str())
                    .map(|pc| pc.statics.clone())
                    .unwrap_or_default();
                let next_parent = self.pending_classes.get(pname.as_str())
                    .and_then(|pc| pc.parent.clone());
                for (sname, sci) in &parent_statics {
                    // Only inherit if child doesn't already define it
                    if !all_statics.iter().any(|(n, _)| n == sname) {
                        common::classes::emit_attach_static_method(self.chunk(), ctor_local, sname, *sci, line);
                        all_statics.push((sname.clone(), *sci));
                    }
                }
                current_parent = next_parent;
            }
        }

        // Store statics in PendingClass for grandchildren to inherit
        if let Some(pc) = self.pending_classes.get_mut(name) {
            pc.statics = all_statics;
        }

        let all_methods: Vec<(String, usize)> = method_chunks.iter().map(|(n, c, _, _)| (n.clone(), *c)).collect();
        let parent_str = parent.clone().unwrap_or_default();
        common::classes::register_type(&mut self.chunks, name, &parent_str, fields, all_methods, false, Vec::new(), Some(ctor_idx));

        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Expression compilation
    // ════════════════════════════════════════════════════════════════════════

    fn compile_expr(&mut self, expr: &Expression) -> Result<(), String> {
        match &expr.kind {
            // ── Literals ────────────────────────────────────────────────
            ExprKind::Lit(lit) => {
                match lit {
                    Literal::Int(n) => self.emit_const(Value::F64(*n as f64)),
                    Literal::Float(n) => self.emit_const(Value::F64(*n)),
                    Literal::Str(s) => self.emit_const(Value::String(Arc::from(s.as_str()))),
                    Literal::Char(c) => self.emit_const(Value::String(Arc::from(c.to_string().as_str()))),
                    Literal::Bool(b) => if *b { self.emit(Op::r#true) } else { self.emit(Op::r#false) },
                    Literal::Null => self.emit(Op::null),
                    Literal::Undefined => self.emit(Op::null),
                    Literal::Ellipsis => self.emit(Op::null),
                }
            }

            // ── Identifier ──────────────────────────────────────────────
            ExprKind::Ident(name) => {
                // Local variable / parameter takes priority over implicit self field
                let is_local = self.scope().resolve(name).is_some()
                    || (!self.case_sensitive && self.scope().resolve_ci(name).is_some());

                // Implicit self field access (only if NOT a local)
                if !is_local && self.is_class_field(name) {
                    let self_kw = self.profile.self_keyword.clone();
                    if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                        self.emit_u16(Op::local_get, slot);
                        let field_name = self.canon(name);
                        let idx = self.str_const(&field_name);
                        self.emit_u16(Op::struct_get, idx);
                        return Ok(());
                    }
                }

                // Bare enum member: `Green` → `TColor.Green`
                if !is_local {
                    let canon_name = self.canon(name);
                    if let Some(enum_type) = self.enum_members.get(&canon_name).cloned() {
                        let type_idx = self.str_const(&enum_type);
                        self.emit_u16(Op::global_get, type_idx);
                        let mem_idx = self.str_const(&canon_name);
                        self.emit_u16(Op::struct_get, mem_idx);
                        return Ok(());
                    }
                }

                // Bare profile namespace constant (e.g. Pascal `MaxInt`, `Pi`)
                if !is_local && !self.defined_globals.contains(&self.canon(name)) {
                    if let Some(cv) = self.profile.lookup_constant(name) {
                        match cv {
                            ConstantValue::Float(f) => self.emit_const(Value::F64(*f)),
                            ConstantValue::Str(s) => self.emit_const(Value::String(Arc::from(s.as_str()))),
                        }
                        return Ok(());
                    }
                }

                self.emit_var_get(name);
            }

            // ── This / Super ────────────────────────────────────────────
            ExprKind::This => {
                let self_kw = &self.profile.self_keyword;
                if let Some(slot) = self.scope().resolve(self_kw)
                    .or_else(|| self.scope().resolve_ci(self_kw))
                    .or_else(|| self.scope().resolve("Self"))
                    .or_else(|| self.scope().resolve("self"))
                    .or_else(|| self.scope().resolve("this"))
                {
                    self.emit_u16(Op::local_get, slot);
                } else {
                    self.emit(Op::null);
                }
            }

            ExprKind::Super => {
                // super refers to the parent class constructor.
                // Look up the parent from the current class's PendingClass info.
                if let Some(ref class_name) = self.current_class.clone() {
                    if let Some(parent_name) = self.pending_classes.get(class_name.as_str()).and_then(|pc| pc.parent.clone()) {
                        let pname = self.canon(&parent_name);
                        let idx = self.str_const(&pname);
                        self.emit_u16(Op::global_get, idx);
                    } else {
                        self.emit(Op::null);
                    }
                } else {
                    self.emit(Op::null);
                }
            }

            // ── Binary ──────────────────────────────────────────────────
            ExprKind::Binary { op, left, right } => {
                // Short-circuit for And/Or
                if *op == BinOp::And {
                    self.compile_expr(left)?;
                    let line = self.line;
                    let skip = common::expressions::emit_and_start(&mut self.chunks[self.current], line);
                    self.compile_expr(right)?;
                    common::expressions::emit_short_circuit_end(&mut self.chunks[self.current], skip);
                    return Ok(());
                }
                if *op == BinOp::Or {
                    self.compile_expr(left)?;
                    let line = self.line;
                    let skip = common::expressions::emit_or_start(&mut self.chunks[self.current], line);
                    self.compile_expr(right)?;
                    common::expressions::emit_short_circuit_end(&mut self.chunks[self.current], skip);
                    return Ok(());
                }
                // NullCoalesce as binary op
                if *op == BinOp::NullCoalesce {
                    self.compile_expr(left)?;
                    self.emit(Op::dup);
                    self.emit(Op::ref_is_null);
                    let skip = self.emit_jump(Op::br_if_false);
                    self.emit(Op::drop);
                    self.compile_expr(right)?;
                    self.patch_jump(skip);
                    return Ok(());
                }
                // Pow → canonical stdlib path: push func ref BEFORE operands
                // so [func, base, exponent] is on the stack for call_ref.
                if *op == BinOp::Pow {
                    let line = self.line;
                    common::math::emit_pow_push_func(self.chunk(), line);
                    self.compile_expr(left)?;
                    self.compile_expr(right)?;
                    common::math::emit_pow_invoke(self.chunk(), line);
                    return Ok(());
                }
                self.compile_expr(left)?;
                self.compile_expr(right)?;
                self.compile_binop(op);
            }

            // ── Unary ───────────────────────────────────────────────────
            ExprKind::Unary { op, expr: inner } => {
                match op {
                    UnaryOp::PreInc | UnaryOp::PostInc => {
                        // ++x / x++ : load, add 1, store
                        self.compile_expr(inner)?;
                        if *op == UnaryOp::PostInc { self.emit(Op::dup); }
                        self.emit_const(Value::F64(1.0));
                        self.emit(Op::dyn_add);
                        if *op == UnaryOp::PreInc { self.emit(Op::dup); }
                        self.compile_assign_target(inner)?;
                    }
                    UnaryOp::PreDec | UnaryOp::PostDec => {
                        self.compile_expr(inner)?;
                        if *op == UnaryOp::PostDec { self.emit(Op::dup); }
                        self.emit_const(Value::F64(1.0));
                        self.emit(Op::f64_sub);
                        if *op == UnaryOp::PreDec { self.emit(Op::dup); }
                        self.compile_assign_target(inner)?;
                    }
                    _ => {
                        self.compile_expr(inner)?;
                        match op {
                            UnaryOp::Neg => { let l = self.line; common::math::emit_neg(self.chunk(), l); }
                            UnaryOp::Pos => {} // no-op
                            UnaryOp::Not => self.emit(Op::dyn_not),
                            UnaryOp::BitNot => self.emit(Op::i32_not),
                            UnaryOp::Typeof => self.emit(Op::ref_typeof),
                            UnaryOp::Void => { self.emit(Op::drop); self.emit(Op::null); }
                            UnaryOp::Delete => { self.emit(Op::drop); self.emit(Op::r#true); }
                            UnaryOp::Deref => { let idx = self.str_const("__value"); self.emit_u16(Op::struct_get, idx); }
                            UnaryOp::AddrOf => {} // no-op in VM
                            UnaryOp::Await => {} // handled below in ExprKind::Await
                            _ => {} // PreInc etc handled above
                        }
                    }
                }
            }

            // ── Ternary ─────────────────────────────────────────────────
            ExprKind::Ternary { cond, then, else_ } => {
                self.compile_expr(cond)?;
                self.emit(Op::dyn_to_bool);
                let else_j = self.emit_jump(Op::br_if_false);
                self.compile_expr(then)?;
                let end_j = self.emit_jump(Op::br);
                self.patch_jump(else_j);
                self.compile_expr(else_)?;
                self.patch_jump(end_j);
            }

            // ── Call ────────────────────────────────────────────────────
            ExprKind::Call { callee, args, .. } => {
                self.compile_call(callee, args)?;
            }

            // ── Member access ───────────────────────────────────────────
            ExprKind::Member { object, field, null_safe } => {
                // Namespace constant check (Math.PI, etc.)
                if let ExprKind::Ident(obj_name) = &object.kind {
                    let compound = format!("{}.{}", obj_name, field);
                    if let Some(cv) = self.profile.lookup_constant(&compound) {
                        match cv {
                            ConstantValue::Float(f) => self.emit_const(Value::F64(*f)),
                            ConstantValue::Str(s) => self.emit_const(Value::String(Arc::from(s.as_str()))),
                        }
                        return Ok(());
                    }
                    // Constructor call with 0 args: ClassName.Create
                    let ctor_nm = &self.profile.constructor_name;
                    let is_ctor = if self.case_sensitive { field == ctor_nm } else { field.eq_ignore_ascii_case(ctor_nm) };
                    if is_ctor && self.defined_globals.contains(obj_name.as_str()) {
                        self.emit_var_get(obj_name);
                        self.emit_u8(Op::call_ref, 0);
                        return Ok(());
                    }
                }

                self.compile_expr(object)?;

                if *null_safe {
                    // ?. — check null before accessing
                    self.emit(Op::dup);
                    self.emit(Op::ref_is_null);
                    let skip = self.emit_jump(Op::br_if_false);
                    // Object is null — result is null
                    let end = self.emit_jump(Op::br);
                    self.patch_jump(skip);
                    let field_name = self.canon(field);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::struct_get, idx);
                    self.patch_jump(end);
                } else {
                    let field_name = self.canon(field);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::struct_get, idx);
                }
            }

            // ── Index access ────────────────────────────────────────────
            ExprKind::Index { object, index } => {
                // A Range used as the index is a slice operation
                // (C# `arr[1..3]` / `s[0..5]`, Python `arr[1:3]` / `s[0:5]`).
                // Route through compiler_common's polymorphic slice helper so
                // strings and arrays both work uniformly.
                if let ExprKind::Range { start, end, .. } = &index.kind {
                    let line = self.line;
                    common::collections::emit_slice_push_func(self.chunk(), line);
                    self.compile_expr(object)?;
                    self.compile_expr(start)?;
                    self.compile_expr(end)?;
                    common::collections::emit_slice_invoke(self.chunk(), line);
                } else {
                    self.compile_expr(object)?;
                    self.compile_expr(index)?;
                    self.emit(Op::array_get);
                }
            }

            // ── New ─────────────────────────────────────────────────────
            ExprKind::New { class, args } => {
                if let ExprKind::Ident(type_name) = &class.kind {
                    // User-defined classes take priority over all built-in type mappings.
                    // This ensures `class Point { ... }` followed by `new Point()` calls
                    // the user constructor, not vybe:drawing::pointNew.
                    let canon_type = self.canon(type_name);
                    if self.defined_classes.contains(&canon_type) {
                        // Bypass compile_expr to avoid the implicit-self-field
                        // shadowing path: in case-insensitive languages a field
                        // named `inner` and a class named `Inner` both
                        // canonicalize to "inner", and the implicit-self-field
                        // check would mis-route to `me.inner` instead of the
                        // class global. Type names always come from globals.
                        let idx = self.str_const(&canon_type);
                        self.emit_u16(Op::global_get, idx);
                        for a in args { self.compile_expr(&a.value)?; }
                        self.emit_u8(Op::call_ref, args.len() as u8);
                        return Ok(());
                    }

                    let bare = type_name.to_lowercase();
                    let bare = bare.split('(').next().unwrap_or(&bare).trim();
                    let bare_str = bare.rsplit('.').next().unwrap_or(bare);

                    // WASM threading/async — use compiler_common, NOT host calls
                    match bare_str {
                        "thread" => {
                            // New Thread(callback) → cont_new only (Start resumes)
                            if let Some(a) = args.first() {
                                self.compile_expr(&a.value)?;
                            }
                            let line = self.line;
                            common::threading::emit_thread_new(self.chunk(), line);
                            return Ok(());
                        }
                        "task" => {
                            // New Task(callback) → cont_new only
                            if let Some(a) = args.first() {
                                self.compile_expr(&a.value)?;
                            }
                            let line = self.line;
                            common::threading::emit_thread_new(self.chunk(), line);
                            return Ok(());
                        }
                        "mutex" | "semaphore" => {
                            // New Mutex() → allocate atomic address for lock
                            self.emit_const(Value::I32(0)); // initial lock value
                            return Ok(());
                        }
                        _ => {}
                    }

                    // Built-in exception types — route through compiler_common
                    // so that every language produces the canonical 4-field
                    // shape and the type name is normalized. PHP `RuntimeException`,
                    // Python `RuntimeError`, JS `Error`, etc. all produce identical
                    // bytecode and can catch each other cross-language.
                    if common::errors::is_exception_type(bare_str) {
                        self.emit_u16(Op::struct_new, 0);
                        self.emit(Op::dup);
                        if let Some(msg_arg) = args.first() {
                            self.compile_expr(&msg_arg.value)?;
                        } else {
                            self.emit_const(Value::String(Arc::from("")));
                        }
                        let line = self.line;
                        common::errors::emit_exception_new_finalize(
                            self.chunk(),
                            bare_str,
                            line,
                        );
                        return Ok(());
                    }

                    // Profile known types (collections, GUI controls, etc.)
                    if let Some((module, func)) = self.profile.lookup_known_type(type_name).map(|(m, f)| (m.to_string(), f.to_string())) {
                        for a in args { self.compile_expr(&a.value)?; }
                        // Special module "common" → use compiler_common emitter (no host call)
                        if module == "common" {
                            let line = self.line;
                            self.emit_common(&func, line);
                        } else {
                            let idx = self.import(&module, &func);
                            self.emit_host_call(idx, args.len() as u8);
                        }
                        return Ok(());
                    }
                    // GUI control: Button, TextBox, Label, Timer, etc.
                    // Checked BEFORE dotnet known_types so GUI controls always
                    // route through the canonical gui emitter regardless of
                    // whether they overlap with .NET BCL types (Timer is both
                    // a GUI control and a System.Threading.Timer — the GUI
                    // form takes priority because we're in `New X()` syntax).
                    let canonical = common::gui::canonical_control_name(bare_str);
                    if !canonical.is_empty() {
                        let host_name = common::gui::host_fn_new_control(&canonical);
                        let new_idx = self.import("vybe:gui", &host_name);
                        for a in args { self.compile_expr(&a.value)?; }
                        let line = self.line;
                        common::gui::emit_new_control(self.chunk(), new_idx, args.len() as u8, line);
                        return Ok(());
                    }
                    // Dotnet known types (collections, etc.) — fallback after
                    // GUI so .NET-only types like Dictionary still work.
                    let known = common::dotnet::known_types();
                    if let Some(&(module, func)) = known.get(bare_str) {
                        for a in args { self.compile_expr(&a.value)?; }
                        let idx = self.import(module, func);
                        self.emit_host_call(idx, args.len() as u8);
                        return Ok(());
                    }
                }
                // User-defined class constructor
                self.compile_expr(class)?;
                for a in args { self.compile_expr(&a.value)?; }
                self.emit_u8(Op::call_ref, args.len() as u8);
            }

            // ── Assignment as expression ────────────────────────────────
            ExprKind::Assign { target, value } => {
                self.compile_expr(value)?;
                self.emit(Op::dup);
                self.compile_assign_target(target)?;
            }

            // ── Lambda ──────────────────────────────────────────────────
            ExprKind::Lambda { params, body, .. } => {
                self.compile_lambda(params, body)?;
            }

            // ── Array literal ───────────────────────────────────────────
            ExprKind::Array(elements) => {
                for elem in elements {
                    if elem.spread {
                        self.compile_expr(&elem.value)?;
                        self.emit(Op::spread);
                    } else {
                        self.compile_expr(&elem.value)?;
                    }
                }
                let line = self.line;
                self.chunks[self.current].emit_op_u16(Op::array_new, elements.len() as u16, line);
            }

            // ── Tuple (Python) ──────────────────────────────────────────
            ExprKind::Tuple(elements) => {
                for elem in elements { self.compile_expr(elem)?; }
                let line = self.line;
                self.chunks[self.current].emit_op_u16(Op::array_new, elements.len() as u16, line);
            }

            // ── Set (Python) ────────────────────────────────────────────
            ExprKind::Set(elements) => {
                for elem in elements { self.compile_expr(elem)?; }
                let line = self.line;
                self.chunks[self.current].emit_op_u16(Op::array_new, elements.len() as u16, line);
                // Convert to set via host call
                let idx = self.import("vybe:collections", "arrayToSet");
                self.emit_host_call(idx, 1);
            }

            // ── Object literal ──────────────────────────────────────────
            ExprKind::Object(props) => {
                let line = self.line;
                common::dict::emit_new(&mut self.chunks[self.current], line);
                for prop in props {
                    match prop {
                        ObjectProperty::KeyValue { key, value } => {
                            self.emit(Op::dup);
                            self.compile_expr(value)?;
                            if let ExprKind::Lit(Literal::Str(k)) = &key.kind {
                                let idx = self.str_const(k);
                                self.emit_u16(Op::struct_set, idx);
                            } else {
                                self.compile_expr(key)?;
                                self.emit(Op::array_set);
                            }
                            self.emit(Op::drop);
                        }
                        ObjectProperty::Shorthand(name) => {
                            self.emit(Op::dup);
                            self.emit_var_get(name);
                            let idx = self.str_const(name);
                            self.emit_u16(Op::struct_set, idx);
                            self.emit(Op::drop);
                        }
                        ObjectProperty::Spread(expr) => {
                            // Object spread: merge properties from expr into current object
                            self.compile_expr(expr)?;
                            let idx = self.import("vybe:object", "assign");
                            self.emit_host_call(idx, 2);
                        }
                        ObjectProperty::Method { key, value } => {
                            self.emit(Op::dup);
                            if let StmtKind::FunctionDecl { params, body, .. } = &value.kind {
                                // Object methods receive `this` as implicit first arg
                                let mut method_params = vec![Param {
                                    name: self.profile.self_keyword.clone(),
                                    type_hint: None, default: None,
                                    pass_by: PassBy::Value, is_rest: false,
                                    is_kwargs: false, is_optional: false, is_nullable: false,
                                }];
                                method_params.extend(params.iter().cloned());
                                self.compile_lambda(&method_params, &LambdaBody::Block(body.clone()))?;
                            } else {
                                self.emit(Op::null);
                            }
                            let idx = self.str_const(key);
                            self.emit_u16(Op::struct_set, idx);
                            self.emit(Op::drop);
                        }
                        ObjectProperty::Accessor { kind, key, value } => {
                            self.emit(Op::dup);
                            if let StmtKind::FunctionDecl { params, body, .. } = &value.kind {
                                // Accessors receive `this` as first arg
                                let mut accessor_params = vec![Param {
                                    name: self.profile.self_keyword.clone(),
                                    type_hint: None, default: None,
                                    pass_by: PassBy::Value, is_rest: false,
                                    is_kwargs: false, is_optional: false, is_nullable: false,
                                }];
                                accessor_params.extend(params.iter().cloned());
                                self.compile_lambda(&accessor_params, &LambdaBody::Block(body.clone()))?;
                            } else {
                                self.emit(Op::null);
                            }
                            let accessor_name = match kind {
                                AccessorKind::Get => format!("__get_{}", key),
                                AccessorKind::Set => format!("__set_{}", key),
                            };
                            let idx = self.str_const(&accessor_name);
                            self.emit_u16(Op::struct_set, idx);
                            self.emit(Op::drop);
                        }
                        ObjectProperty::Computed { key, value } => {
                            self.emit(Op::dup);
                            self.compile_expr(value)?;
                            self.compile_expr(key)?;
                            self.emit(Op::array_set);
                            self.emit(Op::drop);
                        }
                    }
                }
            }

            // ── String interpolation ────────────────────────────────────
            ExprKind::Interpolation(parts) => {
                if parts.is_empty() {
                    self.emit_const(Value::String(Arc::from("")));
                    return Ok(());
                }
                // Use stdlib __vybe_tostring (pure WASM, populated by bundle::finalize_with_stdlib)
                let tostring_global = self.str_const("__vybe_tostring");
                for (i, part) in parts.iter().enumerate() {
                    match part {
                        InterpolPart::Text(s) => self.emit_const(Value::String(Arc::from(s.as_str()))),
                        InterpolPart::Expr(e) | InterpolPart::Formatted(e, _) => {
                            // Push func ref FIRST, then the value, then call_ref
                            self.emit_u16(Op::global_get, tostring_global);
                            self.compile_expr(e)?;
                            self.emit_u8(Op::call_ref, 1);
                        }
                    }
                    if i > 0 {
                        let line = self.line;
                        common::strings::emit_str_concat(self.chunk(), line);
                    }
                }
            }

            // ── Type operations ─────────────────────────────────────────
            ExprKind::IsType { expr: inner, type_name } => {
                // Compare against canonicalized class name (case-insensitive
                // languages like VB/Pascal store class __type lowercased).
                let canon_type = self.canon(type_name);
                self.compile_expr(inner)?;
                let key = self.str_const("__type");
                self.emit_u16(Op::struct_get, key);
                self.emit_const(Value::String(Arc::from(canon_type.as_str())));
                self.emit(Op::dyn_eq);
            }

            ExprKind::Cast { expr: inner, .. } => {
                // Cast is a no-op in our dynamic VM
                self.compile_expr(inner)?;
            }

            ExprKind::TypeOf(inner) => {
                self.compile_expr(inner)?;
                self.emit(Op::ref_typeof);
            }

            // ── NullCoalesce ────────────────────────────────────────────
            ExprKind::NullCoalesce { left, right } => {
                self.compile_expr(left)?;
                self.emit(Op::dup);
                self.emit(Op::ref_is_null);
                let skip = self.emit_jump(Op::br_if_false);
                self.emit(Op::drop);
                self.compile_expr(right)?;
                self.patch_jump(skip);
            }

            // ── Spread ──────────────────────────────────────────────────
            ExprKind::Spread(inner) => {
                self.compile_expr(inner)?;
                self.emit(Op::spread);
            }

            // ── Await ───────────────────────────────────────────────────
            ExprKind::Await(inner) => {
                self.compile_expr(inner)?;
                // Await is a no-op in sync VM (JSPI handles transparent async)
            }

            // ── Yield ───────────────────────────────────────────────────
            ExprKind::Yield(val) => {
                if let Some(v) = val { self.compile_expr(v)?; } else { self.emit(Op::null); }
                self.emit_u16(Op::suspend, 0);
            }

            ExprKind::YieldFrom(inner) => {
                self.compile_expr(inner)?;
                // Simplified: yield from → just pass through
            }

            // ── AddressOf (VB) ──────────────────────────────────────────
            ExprKind::AddressOf(name) => {
                self.emit_var_get(name);
            }

            // ── SuperCall (VB/Python) ───────────────────────────────────
            ExprKind::SuperCall { method, args } => {
                let self_kw = self.profile.self_keyword.clone();
                let ctor_name = self.profile.constructor_name.clone();
                let is_ctor_call = method.is_none() || method.as_ref().map_or(false, |m| {
                    if self.case_sensitive {
                        m == &ctor_name || m == "new" || m == "__init__"
                    } else {
                        m.eq_ignore_ascii_case(&ctor_name) || m.eq_ignore_ascii_case("new") || m.eq_ignore_ascii_case("__init__")
                    }
                });

                if is_ctor_call {
                    // super() / MyBase.New(args) → call parent constructor
                    if let Some(ref class_name) = self.current_class.clone() {
                        if let Some(parent_name) = self.pending_classes.get(class_name.as_str()).and_then(|c| c.parent.clone()) {
                            let pname = self.canon(&parent_name);
                            let pidx = self.str_const(&pname);
                            self.emit_u16(Op::global_get, pidx);
                            for a in args { self.compile_expr(&a.value)?; }
                            self.emit_u8(Op::call_ref, args.len() as u8);
                            if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                                self.emit(Op::dup);
                                self.emit_u16(Op::local_set, slot);
                                self.emit(Op::drop);
                            }
                        } else {
                            self.emit(Op::null);
                        }
                    } else {
                        self.emit(Op::null);
                    }
                } else if let Some(ref mname) = method {
                    // MyBase.Method(args) → this.__base_method(this, args)
                    let base_name = format!("__base_{}", self.canon(mname));
                    if let Some(self_slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                        let prop = self.str_const(&base_name);
                        self.emit_u16(Op::local_get, self_slot);
                        self.emit_u16(Op::struct_get, prop);
                        self.emit_u16(Op::local_get, self_slot);
                        for a in args { self.compile_expr(&a.value)?; }
                        self.emit_u8(Op::call_ref, (args.len() + 1) as u8);
                    } else {
                        self.emit(Op::null);
                    }
                } else {
                    self.emit(Op::null);
                }
            }

            // ── Comprehension (Python) ──────────────────────────────────
            ExprKind::Comprehension { kind: _, element, generators } => {
                // Simplified: compile as loop building an array
                let line = self.line;
                self.chunks[self.current].emit_op_u16(Op::array_new, 0, line);
                let result_slot = self.scope_mut().define("__comp_result");
                self.emit_u16(Op::local_set, result_slot); self.emit(Op::drop);

                // Only handle the first generator for simplicity
                if let Some(gen) = generators.first() {
                    self.compile_expr(&gen.iter)?;
                    let arr_slot = self.scope_mut().define("__comp_iter");
                    self.emit_u16(Op::local_set, arr_slot); self.emit(Op::drop);
                    let idx_slot = self.scope_mut().define("__comp_idx");
                    let (loop_start, exit_jump) = common::loops::emit_for_in_start(
                        &mut self.chunks[self.current], arr_slot, idx_slot, line,
                    );
                    // Bind loop var
                    let var_name = match &gen.target.kind {
                        ExprKind::Ident(n) => n.clone(),
                        _ => "__comp_var".to_string(),
                    };
                    let var_slot = self.scope_mut().define(&var_name);
                    self.emit_u16(Op::local_set, var_slot); self.emit(Op::drop);

                    // Check conditions
                    let mut cond_skip = None;
                    for cond_expr in &gen.conditions {
                        self.compile_expr(cond_expr)?;
                        self.emit(Op::dyn_to_bool);
                        cond_skip = Some(self.emit_jump(Op::br_if_false));
                    }

                    // Push element
                    self.emit_u16(Op::local_get, result_slot);
                    self.compile_expr(element)?;
                    self.emit(Op::array_push);
                    self.emit(Op::drop);

                    if let Some(skip) = cond_skip { self.patch_jump(skip); }

                    common::loops::emit_for_in_end(
                        &mut self.chunks[self.current], idx_slot, loop_start, exit_jump, line,
                    );
                }

                self.emit_u16(Op::local_get, result_slot);
            }

            // ── Slice (Python) ──────────────────────────────────────────
            ExprKind::Slice { lower, upper, step } => {
                // Emit slice parts for use by Index
                if let Some(l) = lower { self.compile_expr(l)?; } else { self.emit(Op::null); }
                if let Some(u) = upper { self.compile_expr(u)?; } else { self.emit(Op::null); }
                if let Some(s) = step { self.compile_expr(s)?; } else { self.emit(Op::null); }
                let idx = self.import("vybe:array", "sliceStep");
                self.emit_host_call(idx, 4); // obj already on stack from Index parent
            }

            // ── Walrus (Python :=) ──────────────────────────────────────
            ExprKind::Walrus { target, value } => {
                self.compile_expr(value)?;
                self.emit(Op::dup);
                self.compile_assign_target(target)?;
            }

            // ── Void (JS) ───────────────────────────────────────────────
            ExprKind::Void(inner) => {
                self.compile_expr(inner)?;
                self.emit(Op::drop);
                self.emit(Op::null); // void always evaluates to undefined
            }

            // ── Delete (JS expression) ──────────────────────────────────
            ExprKind::Delete(inner) => {
                // Delete member: always returns true
                if let ExprKind::Member { object, field, .. } = &inner.kind {
                    self.compile_expr(object)?;
                    self.emit(Op::null);
                    let field_name = self.canon(field);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::struct_set, idx);
                    self.emit(Op::drop);
                } else {
                    self.compile_expr(inner)?;
                    self.emit(Op::drop);
                }
                self.emit(Op::r#true);
            }

            // ── Destructure (JS) ────────────────────────────────────────
            ExprKind::Destructure(_) => {
                // Destructure patterns are handled at assignment/declaration sites
                self.emit(Op::null);
            }

            // ── Sequence (JS comma operator) ────────────────────────────
            ExprKind::Sequence(exprs) => {
                for (i, e) in exprs.iter().enumerate() {
                    self.compile_expr(e)?;
                    if i < exprs.len() - 1 { self.emit(Op::drop); }
                }
            }

            // ── ClassExpr (JS) ──────────────────────────────────────────
            ExprKind::ClassExpr { name, parent, members } => {
                let class_name = name.clone().unwrap_or_else(|| "__anonymous_class".to_string());
                let parent_name = if let Some(p) = parent {
                    if let ExprKind::Ident(n) = &p.kind { Some(n.clone()) } else { None }
                } else { None };
                let class_name = self.canon(&class_name);
                let parent_name = parent_name.map(|p| self.canon(&p));
                self.defined_globals.insert(class_name.clone());
                self.compile_class(&class_name, &parent_name, members)?;
                self.emit_var_get(&class_name);
            }

            // ── FunctionExpr (JS) ───────────────────────────────────────
            ExprKind::FunctionExpr(stmt) => {
                if let StmtKind::FunctionDecl { name, params, return_type, body, is_sub, handles, .. } = &stmt.kind {
                    let fn_name = if name.is_empty() { "__anon_fn" } else { name };
                    self.compile_function_decl(fn_name, params, return_type, body, *is_sub, handles)?;
                    self.emit_var_get(fn_name);
                } else {
                    self.emit(Op::null);
                }
            }

            // ── Range ───────────────────────────────────────────────────
            ExprKind::Range { start, end, inclusive: _ } => {
                self.compile_expr(start)?;
                self.compile_expr(end)?;
                let line = self.line;
                common::collections::emit_range(&mut self.chunks[self.current], 2, line);
            }

            // ── StaticAccess (PHP) ──────────────────────────────────────
            ExprKind::StaticAccess { class, member } => {
                // class::member → look up class, then get static member
                self.compile_expr(class)?;
                if let ExprKind::Ident(name) = &member.kind {
                    let idx = self.str_const(name);
                    self.emit_u16(Op::struct_get, idx);
                } else {
                    self.compile_expr(member)?;
                    self.emit(Op::array_get);
                }
            }

            // ── Match expression (PHP/Rust) ─────────────────────────────
            ExprKind::Match { subject, arms } => {
                self.compile_expr(subject)?;
                let subject_slot = self.scope_mut().define("__match_subj");
                self.emit_u16(Op::local_set, subject_slot); self.emit(Op::drop);
                let mut end_patches = Vec::new();
                for arm in arms {
                    if let Some(ref conditions) = arm.conditions {
                        let mut match_patches = Vec::new();
                        for c in conditions {
                            self.emit_u16(Op::local_get, subject_slot);
                            self.compile_expr(c)?;
                            self.emit(Op::dyn_eq);
                            match_patches.push(self.emit_jump(Op::br_if_true));
                        }
                        let skip = self.emit_jump(Op::br);
                        for p in match_patches { self.patch_jump(p); }
                        self.compile_expr(&arm.body)?;
                        end_patches.push(self.emit_jump(Op::br));
                        self.patch_jump(skip);
                    } else {
                        // Default arm
                        self.compile_expr(&arm.body)?;
                        end_patches.push(self.emit_jump(Op::br));
                    }
                }
                // If no arm matched, null
                self.emit(Op::null);
                for p in end_patches { self.patch_jump(p); }
            }
        }
        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Call compilation
    // ════════════════════════════════════════════════════════════════════════

    fn compile_call(&mut self, callee: &Expression, args: &[Argument]) -> Result<(), String> {
        let arg_exprs: Vec<&Expression> = args.iter().map(|a| &a.value).collect();

        // ── super(args) → call parent constructor, store result as this ──
        if let ExprKind::Super = &callee.kind {
            if let Some(ref class_name) = self.current_class.clone() {
                if let Some(parent_name) = self.pending_classes.get(class_name.as_str()).and_then(|pc| pc.parent.clone()) {
                    let pname = self.canon(&parent_name);
                    let pidx = self.str_const(&pname);
                    self.emit_u16(Op::global_get, pidx);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::call_ref, arg_exprs.len() as u8);
                    // Store result as this
                    let self_kw = self.profile.self_keyword.clone();
                    if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                        self.emit(Op::dup);
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    }
                    return Ok(());
                }
            }
            // No parent — emit null
            self.emit(Op::null);
            return Ok(());
        }

        // ── super.method(args) → this.__base_method(args) ────────────
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if matches!(&object.kind, ExprKind::Super) {
                let base_name = format!("__base_{}", self.canon(field));
                let self_kw = self.profile.self_keyword.clone();
                if let Some(self_slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                    let prop = self.str_const(&base_name);
                    self.emit_u16(Op::local_get, self_slot);
                    self.emit_u16(Op::struct_get, prop);
                    // Call with this as first arg
                    self.emit_u16(Op::local_get, self_slot);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::call_ref, (arg_exprs.len() + 1) as u8);
                    return Ok(());
                }
            }
        }

        // ── Builtin check: Ident("print") ───────────────────────────
        if let ExprKind::Ident(name) = &callee.kind {
            if self.try_compile_builtin(name, &arg_exprs)? { return Ok(()); }
        }

        // ── Builtin check: Member("Console.WriteLine") ─────────────
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if let ExprKind::Ident(obj_name) = &object.kind {
                let compound = format!("{}.{}", obj_name, field);
                if self.try_compile_builtin(&compound, &arg_exprs)? { return Ok(()); }

                // Module alias: console.log → host call
                if let Some(module) = self.profile.lookup_module_alias(obj_name).map(|s| s.to_string()) {
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    let idx = self.import(&module, field);
                    self.emit_host_call(idx, arg_exprs.len() as u8);
                    return Ok(());
                }
            }
        }

        // ── Dotted name resolution FIRST (uses compiler_common::dotnet when use_dotnet) ──
        // Must run before value methods because value methods like "add" would
        // intercept "Controls.Add" which needs special GUI handling.
        if let ExprKind::Member { .. } = &callee.kind {
            let parts = self.flatten_member_chain(callee);
            if parts.len() >= 2 {
                let lower_parts: Vec<String> = parts.iter().map(|s| self.canon(s)).collect();

                // Use dotnet resolver when enabled
                if self.profile.namespaces.use_dotnet {
                    let imports = {
                        let mut imp = common::dotnet::default_interface_imports();
                        imp.extend(self.profile.namespaces.extra_imports.clone());
                        imp
                    };
                    let scope = self.scope().clone();
                    let defined_globals = self.defined_globals.clone();
                    let field_set: std::collections::HashSet<String> = if let Some(ref cn) = self.current_class {
                        self.pending_classes.get(cn.as_str())
                            .map(|pc| pc.fields.iter().cloned().collect())
                            .unwrap_or_default()
                    } else {
                        std::collections::HashSet::new()
                    };
                    let ctx = common::dotnet::ResolutionContext {
                        is_local: &|name: &str| {
                            scope.resolve(name).is_some()
                            || scope.resolve_ci(name).is_some()
                            || defined_globals.contains(name)
                        },
                        is_class_field: &|name: &str| field_set.contains(name),
                        is_user_type: &|name: &str| defined_globals.contains(name),
                        imports: &imports,
                    };
                    let refs: Vec<&str> = lower_parts.iter().map(|s| s.as_str()).collect();
                    let resolution = common::dotnet::resolve_dotted_name(&refs, &ctx);

                    match resolution {
                        common::dotnet::DottedResolution::HostCall { module, func } => {
                            for a in &arg_exprs { self.compile_expr(a)?; }
                            let idx = self.import(&module, &func);
                            self.emit_host_call(idx, arg_exprs.len() as u8);
                            return Ok(());
                        }
                        common::dotnet::DottedResolution::NamespaceAccess { parts: ns_parts } => {
                            // Intercept threading calls. The actual emit goes
                            // through compiler_common::dispatch so the bytecode
                            // shape is identical to what the C# profile's
                            // `common:threading.*` entries produce.
                            let dotted = ns_parts.join(".");
                            match dotted.as_str() {
                                "system.threading.task.run" | "task.run" => {
                                    if let Some(a) = arg_exprs.first() { self.compile_expr(a)?; }
                                    let line = self.line;
                                    self.emit_common("threading.task_run", line);
                                    return Ok(());
                                }
                                "system.diagnostics.process.start" | "process.start" => {
                                    // Process.Start(startInfo) → host call that runs the command
                                    for a in &arg_exprs { self.compile_expr(a)?; }
                                    let idx = self.import("vybe:types", "processStart");
                                    self.emit_host_call(idx, arg_exprs.len() as u8);
                                    return Ok(());
                                }
                                "system.threading.thread.sleep" | "thread.sleep" => {
                                    if let Some(a) = arg_exprs.first() { self.compile_expr(a)?; }
                                    let line = self.line;
                                    self.emit_common("threading.sleep", line);
                                    return Ok(());
                                }
                                _ => {}
                            }
                            let root_idx = self.str_const(&ns_parts[0]);
                            self.emit_u16(Op::global_get, root_idx);
                            for part in &ns_parts[1..] {
                                let idx = self.str_const(part);
                                self.emit_u16(Op::struct_get, idx);
                            }
                            let is_const = common::dotnet::is_known_constant(ns_parts.last().unwrap_or(&String::new()));
                            if !is_const {
                                for a in &arg_exprs { self.compile_expr(a)?; }
                                self.emit_u8(Op::call_ref, arg_exprs.len() as u8);
                            }
                            return Ok(());
                        }
                        common::dotnet::DottedResolution::InstanceMember { local, members } => {
                            // Intercept `parent.Controls.Add(child)` for GUI.
                            // The .NET WinForms surface is `Form.Controls.Add(ctrl)`,
                            // MAUI is `parent.Children.Add(ctrl)`, etc. — all
                            // resolve to the canonical gui emitter.
                            if members.len() >= 2 && members[members.len()-2] == "controls" && members[members.len()-1] == "add" {
                                let line = self.line;
                                let add_idx = self.import("vybe:gui", common::gui::HOST_FN_ADD_CHILD);
                                self.emit_var_get(&local);
                                for a in &arg_exprs { self.compile_expr(a)?; }
                                common::gui::emit_add_child(self.chunk(), add_idx, line);
                                return Ok(());
                            }
                            // Intercept Thread/Task methods → WASM stack switching opcodes
                            if members.len() == 1 {
                                let method = members[0].as_str();
                                match method {
                                    "start" => {
                                        // th.Start() — thread_spawn already started it, no-op
                                        self.emit(Op::null);
                                        return Ok(());
                                    }
                                    "join" => {
                                        // th.Join() → thread_join opcode (blocks until thread completes)
                                        self.emit_var_get(&local);
                                        let line = self.line;
                                        common::threading::emit_thread_join(self.chunk(), line);
                                        self.emit(Op::drop);
                                        return Ok(());
                                    }
                                    "waitforexit" => {
                                        // p.WaitForExit() — process ran synchronously, no-op
                                        // Must leave a value on stack (caller drops it)
                                        self.emit(Op::null);
                                        return Ok(());
                                    }
                                    _ => {}
                                }
                            }
                            // Generic instance member chain → obj.prop.method(args)
                            self.emit_var_get(&local);
                            let last_idx = members.len() - 1;
                            for (i, m) in members.iter().enumerate() {
                                let idx = self.str_const(m);
                                if i < last_idx {
                                    self.emit_u16(Op::struct_get, idx);
                                } else {
                                    // Last member is the method — struct_get then call with this
                                    self.emit(Op::dup); // keep obj for this
                                    self.emit_u16(Op::struct_get, idx);
                                    // Stack: [obj, method_fn] — swap so fn is first
                                    let fn_tmp = self.scope().resolve("__dotnet_fn")
                                        .unwrap_or_else(|| self.scope_mut().define("__dotnet_fn"));
                                    self.emit_u16(Op::local_set, fn_tmp); self.emit(Op::drop);
                                    let obj_tmp = self.scope().resolve("__dotnet_obj")
                                        .unwrap_or_else(|| self.scope_mut().define("__dotnet_obj"));
                                    self.emit_u16(Op::local_set, obj_tmp); self.emit(Op::drop);
                                    self.emit_u16(Op::local_get, fn_tmp);
                                    self.emit_u16(Op::local_get, obj_tmp);
                                    for a in &arg_exprs { self.compile_expr(a)?; }
                                    self.emit_u8(Op::call_ref, (arg_exprs.len() + 1) as u8);
                                    return Ok(());
                                }
                            }
                            // Shouldn't reach here for calls, but just in case
                            for a in &arg_exprs { self.compile_expr(a)?; }
                            self.emit_u8(Op::call_ref, arg_exprs.len() as u8);
                            return Ok(());
                        }
                        common::dotnet::DottedResolution::NoOp => {
                            self.emit(Op::null);
                            return Ok(());
                        }
                        common::dotnet::DottedResolution::Unresolved => {
                            // Fall through to value methods and other resolution
                        }
                    }
                }

                // Non-dotnet: module aliases (JS: console → wasi:cli)
                if let Some(module) = self.profile.lookup_module_alias(&lower_parts[0]).map(|s| s.to_string()) {
                    let func = if lower_parts.len() == 2 { lower_parts[1].clone() } else { lower_parts[1..].join(".") };
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    let idx = self.import(&module, &func);
                    self.emit_host_call(idx, arg_exprs.len() as u8);
                    return Ok(());
                }

                // Profile namespace roots
                if self.profile.is_namespace_root(&lower_parts[0]) {
                    let root_idx = self.str_const(&lower_parts[0]);
                    self.emit_u16(Op::global_get, root_idx);
                    for part in &lower_parts[1..] {
                        let idx = self.str_const(part);
                        self.emit_u16(Op::struct_get, idx);
                    }
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::call_ref, arg_exprs.len() as u8);
                    return Ok(());
                }
            }
        }

        // ── Static method call on user class: ClassName.Method(args) ─
        // Must run BEFORE value methods so user class names like MathUtils.Add
        // don't get hijacked by the array Add value method.
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if let ExprKind::Ident(obj_name) = &object.kind {
                let canon = self.canon(obj_name);
                let is_class = self.defined_classes.contains(&canon)
                    && self.scope().resolve(obj_name).is_none();
                if is_class {
                    // Push class, dup, struct_get(method) → [class, fn]
                    // Then swap so fn is first, class is second (as this)
                    let cls_idx = self.str_const(&canon);
                    self.emit_u16(Op::global_get, cls_idx);
                    self.emit(Op::dup);
                    let m = self.canon(field);
                    let method_idx = self.str_const(&m);
                    self.emit_u16(Op::struct_get, method_idx);
                    // Stack: [class, fn] — swap so we have [fn, class, ...args]
                    let fn_tmp = self.scope().resolve("__static_fn")
                        .unwrap_or_else(|| self.scope_mut().define("__static_fn"));
                    self.emit_u16(Op::local_set, fn_tmp); self.emit(Op::drop);
                    let cls_tmp = self.scope().resolve("__static_cls")
                        .unwrap_or_else(|| self.scope_mut().define("__static_cls"));
                    self.emit_u16(Op::local_set, cls_tmp); self.emit(Op::drop);
                    self.emit_u16(Op::local_get, fn_tmp);
                    self.emit_u16(Op::local_get, cls_tmp);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::call_ref, (arg_exprs.len() + 1) as u8);
                    return Ok(());
                }
            }
        }

        // ── Value method: obj.toUpperCase() ─────────────────────────
        // Skip if the method name is defined on a user class — that takes priority.
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            let canon_field = self.canon(field);
            if self.defined_class_methods.contains(&canon_field) {
                // Fall through — let the generic call path handle it
            } else if let Some(def) = self.profile.lookup_value_method(field, arg_exprs.len() as u8).cloned() {
                // For Stdlib calls, push func ref BEFORE args (call_ref expects [func, args...])
                if let BuiltinEmit::Stdlib(stdlib_name) = &def.emit {
                    let global_name = format!("__vybe_{}", stdlib_name);
                    let name_idx = self.str_const(&global_name);
                    self.emit_u16(Op::global_get, name_idx);
                    self.compile_expr(object)?;
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::call_ref, (arg_exprs.len() + 1) as u8);
                    return Ok(());
                }
                // Object is first arg, then explicit args
                self.compile_expr(object)?;
                for a in &arg_exprs { self.compile_expr(a)?; }
                match &def.emit {
                    BuiltinEmit::HostCall(module, func) => {
                        let idx = self.import(module, func);
                        self.emit_host_call(idx, (arg_exprs.len() + 1) as u8);
                    }
                    BuiltinEmit::Opcode(op_name) => {
                        // Object + args already on stack from above
                        self.emit_named_opcode(op_name);
                    }
                    BuiltinEmit::StrLength => {
                        let line = self.line;
                        common::strings::emit_length(self.chunk(), line);
                    }
                    BuiltinEmit::Common(name) => {
                        let line = self.line;
                        let name = name.clone();
                        self.emit_common(&name, line);
                    }
                    _ => {}
                }
                return Ok(());
            }


            // Array higher-order methods: arr.map(fn), arr.filter(fn), etc.
            // Use compiler_common::loops which emits proper loop bytecode.
            let field_lower = if self.case_sensitive { field.clone() } else { field.to_lowercase() };
            if let Some(stdlib_name) = self.profile.lookup_array_method(&field_lower).map(|s| s.to_string()) {
                // Normalize to the JS-style method name used in match below
                let field_lower = match stdlib_name.as_str() {
                    "__array_map" => "map".to_string(),
                    "__array_filter" => "filter".to_string(),
                    "__array_forEach" => "forEach".to_string(),
                    "__array_reduce" => "reduce".to_string(),
                    "__array_find" => "find".to_string(),
                    "__array_sort" => "sort".to_string(),
                    "__array_some" => "some".to_string(),
                    "__array_every" => "every".to_string(),
                    "__array_flat_map" => "flatMap".to_string(),
                    _ => field_lower,
                };
                // Compile arr and fn(s) into local slots
                self.compile_expr(object)?;
                let arr_slot = self.scope_mut().define("__hof_arr");
                self.emit_u16(Op::local_set, arr_slot); self.emit(Op::drop);

                if let Some(fn_expr) = arg_exprs.first() {
                    self.compile_expr(fn_expr)?;
                } else {
                    self.emit(Op::null);
                }
                let fn_slot = self.scope_mut().define("__hof_fn");
                self.emit_u16(Op::local_set, fn_slot); self.emit(Op::drop);

                let idx_slot = self.scope_mut().define("__hof_idx");
                let result_slot = self.scope_mut().define("__hof_result");
                let line = self.line;

                match field_lower.as_str() {
                    "map" => {
                        // emit_map leaves result on stack
                        common::loops::emit_map(self.chunk(), fn_slot, arr_slot, result_slot, idx_slot, line);
                    }
                    "filter" => {
                        let elem_slot = self.scope_mut().define("__hof_elem");
                        common::loops::emit_filter(self.chunk(), fn_slot, arr_slot, result_slot, idx_slot, elem_slot, line);
                    }
                    "reduce" => {
                        // reduce(fn, initial?) — initial is second arg
                        if let Some(init_expr) = arg_exprs.get(1) {
                            self.compile_expr(init_expr)?;
                            self.emit_u16(Op::local_set, result_slot); self.emit(Op::drop);
                        } else {
                            self.emit(Op::i32_const_0);
                            self.emit_u16(Op::local_set, result_slot); self.emit(Op::drop);
                        }
                        common::loops::emit_reduce(self.chunk(), fn_slot, arr_slot, result_slot, idx_slot, line);
                    }
                    "forEach" | "foreach" => {
                        common::loops::emit_foreach(self.chunk(), fn_slot, arr_slot, idx_slot, line);
                    }
                    "some" => {
                        common::loops::emit_any_every(self.chunk(), fn_slot, arr_slot, idx_slot, true, line);
                    }
                    "every" => {
                        common::loops::emit_any_every(self.chunk(), fn_slot, arr_slot, idx_slot, false, line);
                    }
                    "find" => {
                        // find uses includes pattern but returns element not bool
                        self.emit(Op::null);
                        self.emit_u16(Op::local_set, result_slot); self.emit(Op::drop);
                        let (loop_start, exit_jump) = common::loops::emit_for_in_start(
                            self.chunk(), arr_slot, idx_slot, line);
                        let elem_slot = self.scope_mut().define("__find_elem");
                        self.emit_u16(Op::local_set, elem_slot); self.emit(Op::drop);
                        self.emit_u16(Op::local_get, fn_slot);
                        self.emit_u16(Op::local_get, elem_slot);
                        self.emit_u8(Op::call_ref, 1);
                        self.emit(Op::dyn_to_bool);
                        let skip = self.emit_jump(Op::br_if_false);
                        self.emit_u16(Op::local_get, elem_slot);
                        self.emit_u16(Op::local_set, result_slot); self.emit(Op::drop);
                        let brk = self.emit_jump(Op::br);
                        self.patch_jump(skip);
                        common::loops::emit_for_in_end(self.chunk(), idx_slot, loop_start, exit_jump, line);
                        self.patch_jump(brk);
                        self.emit_u16(Op::local_get, result_slot);
                    }
                    "includes" => {
                        // includes uses contains from compiler_common
                        self.emit_u16(Op::local_get, arr_slot);
                        self.emit_u16(Op::local_get, fn_slot); // fn_slot holds the search value
                        common::collections::emit_contains(self.chunk(), line);
                    }
                    "sort" => {
                        // sorted from compiler_common (returns new sorted array)
                        self.emit_u16(Op::local_get, arr_slot);
                        common::collections::emit_sorted(self.chunk(), line);
                    }
                    "indexOf" | "indexof" => {
                        self.emit_u16(Op::local_get, arr_slot);
                        self.emit_u16(Op::local_get, fn_slot); // search value
                        common::collections::emit_index_of(self.chunk(), line);
                    }
                    _ => {
                        // Fallback: call as regular method
                        self.emit_u16(Op::local_get, arr_slot);
                        self.emit_u16(Op::local_get, fn_slot);
                        self.emit_u8(Op::call_ref, 2);
                    }
                }
                return Ok(());
            }
        }

        // ── Constructor call: ClassName.Create(args) ────────────────
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if let ExprKind::Ident(class_name) = &object.kind {
                let ctor_nm = &self.profile.constructor_name.clone();
                let is_ctor = if self.case_sensitive { field == ctor_nm } else { field.eq_ignore_ascii_case(ctor_nm) };
                if is_ctor && self.defined_globals.contains(class_name.as_str()) {
                    self.emit_var_get(class_name);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::call_ref, arg_exprs.len() as u8);
                    return Ok(());
                }
            }
        }

        // ── Method call: obj.method(args) ───────────────────────────
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            self.compile_expr(object)?;
            let field_name = self.canon(field);
            let prop = self.str_const(&field_name);
            self.emit(Op::dup);
            self.emit_u16(Op::struct_get, prop);
            let fn_tmp = self.scope_mut().define("__fn");
            self.emit_u16(Op::local_set, fn_tmp); self.emit(Op::drop);
            let obj_tmp = self.scope_mut().define("__obj");
            self.emit_u16(Op::local_set, obj_tmp); self.emit(Op::drop);
            self.emit_u16(Op::local_get, fn_tmp);
            self.emit_u16(Op::local_get, obj_tmp);
            for a in &arg_exprs { self.compile_expr(a)?; }
            self.emit_u8(Op::call_ref, (arg_exprs.len() + 1) as u8);
            return Ok(());
        }

        // ── Simple call: name(args) / expr(args) ────────────────────
        if let ExprKind::Ident(name) = &callee.kind {
            let is_known_func = self.defined_functions.contains(name)
                || (!self.case_sensitive && self.defined_functions.iter().any(|g| g.eq_ignore_ascii_case(name)));
            if self.try_compile_builtin(name, &arg_exprs)? { return Ok(()); }

            // VB array access: `arr(idx)` when `arr` is a known data variable
            // (local OR top-level global from `Dim arr(5)`) and is NOT a
            // declared function or class. VB syntactically overloads `()` for
            // both calls and indexing — the disambiguator is whether the head
            // is a callable function or a value. We must exclude both
            // `defined_functions` and `defined_classes` from the "looks like
            // a variable" set, otherwise `GetResult()` (function call) and
            // `New Result()` (class) would be mis-identified as indexing.
            if !is_known_func && arg_exprs.len() == 1 && self.profile.parens_for_index {
                let canon_name = self.canon(name);
                let is_local = self.scope().resolve(name).is_some()
                    || (!self.case_sensitive && self.scope().resolve_ci(name).is_some());
                let is_global_var = self.defined_globals.contains(&canon_name)
                    && !self.defined_classes.contains(&canon_name)
                    && !self.defined_functions.contains(&canon_name);
                if is_local || is_global_var {
                    self.emit_var_get(name);
                    self.compile_expr(arg_exprs[0])?;
                    self.emit(Op::array_get);
                    return Ok(());
                }
            }

            // Inside a class: bare method call → Me.method(args)
            // If name isn't a local variable and we're inside a class body,
            // resolve as Me.name() (implicit self for method calls).
            if self.current_class.is_some() && self.profile.implicit_self_fields {
                let is_local = self.scope().resolve(name).is_some()
                    || (!self.case_sensitive && self.scope().resolve_ci(name).is_some());
                if !is_local && !is_known_func {
                    let self_kw = self.profile.self_keyword.clone();
                    if let Some(self_slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                        // Me.name(args) → load Me, dup, struct_get(name), call with this
                        self.emit_u16(Op::local_get, self_slot);
                        let field_name = self.canon(name);
                        let prop = self.str_const(&field_name);
                        self.emit(Op::dup);
                        self.emit_u16(Op::struct_get, prop);
                        let fn_tmp = self.scope_mut().define("__bare_fn");
                        self.emit_u16(Op::local_set, fn_tmp); self.emit(Op::drop);
                        let obj_tmp = self.scope_mut().define("__bare_obj");
                        self.emit_u16(Op::local_set, obj_tmp); self.emit(Op::drop);
                        self.emit_u16(Op::local_get, fn_tmp);
                        self.emit_u16(Op::local_get, obj_tmp);
                        for a in &arg_exprs { self.compile_expr(a)?; }
                        self.emit_u8(Op::call_ref, (arg_exprs.len() + 1) as u8);
                        return Ok(());
                    }
                }
            }

            self.emit_var_get(name);
            for a in &arg_exprs { self.compile_expr(a)?; }
            self.emit_u8(Op::call_ref, arg_exprs.len() as u8);
            return Ok(());
        }

        // ── Fallback: general expression call ───────────────────────
        self.compile_expr(callee)?;
        for a in &arg_exprs { self.compile_expr(a)?; }
        self.emit_u8(Op::call_ref, arg_exprs.len() as u8);
        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Lambda compilation
    // ════════════════════════════════════════════════════════════════════════

    fn compile_lambda(&mut self, params: &[Param], body: &LambdaBody) -> Result<(), String> {
        let arity = params.len() as u8;
        let ci = self.chunks.len();
        let chunk = common::functions::create_function_chunk("<lambda>", arity);
        self.chunks.push(chunk);
        self.scopes.push(Scope::new_function());
        let saved = self.current;
        self.current = ci;
        for p in params { self.scope_mut().define(&p.name); }

        // Result slot for ResultSlot languages
        let result_slot = if self.profile.function_return == ReturnStyle::ResultSlot {
            let rs = self.scope_mut().define("Result");
            self.emit(Op::null); self.emit_u16(Op::local_set, rs); self.emit(Op::drop);
            let saved_fn = self.current_func_name.take();
            let saved_rs = self.current_result_slot.take();
            self.current_func_name = Some("<lambda>".into());
            self.current_result_slot = Some(rs);
            Some((rs, saved_fn, saved_rs))
        } else { None };

        match body {
            LambdaBody::Expr(expr) => {
                self.compile_expr(expr)?;
                self.emit(Op::r#return);
            }
            LambdaBody::Block(stmts) => {
                for s in stmts { self.compile_stmt(s)?; }
            }
        }

        if let Some((rs, saved_fn, saved_rs)) = result_slot {
            self.emit_u16(Op::local_get, rs);
            self.emit(Op::r#return);
            self.current_func_name = saved_fn;
            self.current_result_slot = saved_rs;
        } else if matches!(body, LambdaBody::Block(_)) {
            let line = self.line;
            common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
        }

        let locals = self.scope().next_slot;
        self.chunks[ci].local_count = locals;
        let uvs = self.scopes.last().unwrap().upvalues.clone();
        self.scopes.pop();
        self.current = saved;
        let l = self.line;
        common::functions::emit_ref_func(&mut self.chunks[self.current], ci, uvs.len() as u8, l);
        for uv in &uvs {
            self.chunks[self.current].emit(if uv.is_local { 1 } else { 0 }, l);
            self.chunks[self.current].emit(uv.index, l);
        }
        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Binary operator emission
    // ════════════════════════════════════════════════════════════════════════

    fn compile_binop(&mut self, op: &BinOp) {
        match op {
            BinOp::Add => { if self.profile.dynamic_add { self.emit(Op::dyn_add); } else { self.emit(Op::dyn_add); } }
            BinOp::Sub => self.emit(Op::f64_sub),
            BinOp::Mul => self.emit(Op::f64_mul),
            BinOp::Div => self.emit(Op::f64_div),
            BinOp::IDiv => { self.emit(Op::f64_div); let l = self.line; common::math::emit_trunc(self.chunk(), l); }
            BinOp::FloorDiv => { self.emit(Op::f64_div); let l = self.line; common::math::emit_floor(self.chunk(), l); }
            BinOp::Mod => self.emit(Op::f64_mod),
            BinOp::Pow => { let l = self.line; common::math::emit_pow(self.chunk(), l); }
            BinOp::Eq => self.emit(Op::dyn_eq),
            BinOp::NotEq => self.emit(Op::dyn_ne),
            BinOp::StrictEq => self.emit(Op::dyn_eq), // strict eq is same in our VM
            BinOp::StrictNotEq => self.emit(Op::dyn_ne),
            BinOp::Lt => self.emit(Op::dyn_lt),
            BinOp::Gt => self.emit(Op::dyn_gt),
            BinOp::LtEq => self.emit(Op::dyn_le),
            BinOp::GtEq => self.emit(Op::dyn_ge),
            BinOp::Spaceship => {
                // a <=> b: returns -1, 0, or 1
                let i = self.import("vybe:math", "spaceship");
                self.emit_host_call(i, 2);
            }
            BinOp::And | BinOp::Or => unreachable!(), // handled with short-circuit
            BinOp::Xor => self.emit(Op::i32_xor),
            BinOp::BitAnd => self.emit(Op::i32_and),
            BinOp::BitOr => self.emit(Op::i32_or),
            BinOp::BitXor => self.emit(Op::i32_xor),
            BinOp::Shl => self.emit(Op::i32_shl),
            BinOp::Shr => self.emit(Op::i32_shr_s),
            BinOp::UShr => self.emit(Op::i32_shr_u),
            BinOp::Concat => { let l = self.line; common::strings::emit_str_concat(self.chunk(), l); }
            BinOp::In => {
                // `x in arr` → contains(arr, x). Stack: [needle, haystack] (correct for VM).
                let l = self.line; common::collections::emit_contains(self.chunk(), l);
            }
            BinOp::NotIn => {
                let l = self.line; common::collections::emit_contains(self.chunk(), l);
                self.emit(Op::dyn_not);
            }
            BinOp::InstanceOf => {
                // a instanceof B → check __type chain
                self.emit(Op::dyn_eq); // simplified
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
                self.emit(Op::dyn_eq);
            }
            BinOp::IsNot => {
                self.emit(Op::dyn_eq);
                self.emit(Op::dyn_not);
            }
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Compound assignment operator emission
    // ════════════════════════════════════════════════════════════════════════

    fn compile_compound_op(&mut self, op: &CompoundOp) {
        match op {
            CompoundOp::Add => self.emit(Op::dyn_add),
            CompoundOp::Sub => self.emit(Op::f64_sub),
            CompoundOp::Mul => self.emit(Op::f64_mul),
            CompoundOp::Div => self.emit(Op::f64_div),
            CompoundOp::IDiv => { self.emit(Op::f64_div); let l = self.line; common::math::emit_trunc(self.chunk(), l); }
            CompoundOp::Mod => self.emit(Op::f64_mod),
            CompoundOp::Pow => { let l = self.line; common::math::emit_pow(self.chunk(), l); }
            CompoundOp::Concat => { let l = self.line; common::strings::emit_str_concat(self.chunk(), l); }
            CompoundOp::BitAnd => self.emit(Op::i32_and),
            CompoundOp::BitOr => self.emit(Op::i32_or),
            CompoundOp::BitXor => self.emit(Op::i32_xor),
            CompoundOp::Shl => self.emit(Op::i32_shl),
            CompoundOp::Shr => self.emit(Op::i32_shr_s),
            CompoundOp::UShr => self.emit(Op::i32_shr_u),
            CompoundOp::And => self.emit(Op::dyn_to_bool), // simplified
            CompoundOp::Or => self.emit(Op::dyn_to_bool),  // simplified
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
                    self.emit_u16(Op::global_get, tostring_global);
                    self.compile_expr(arg)?;
                    self.emit_u8(Op::call_ref, 1);
                    return Ok(true);
                }
            } else {
                // Compile args, then dispatch to canonical emitter
                for a in args { self.compile_expr(a)?; }
                common::canonical::emit_canonical(canonical_op, self.chunk(), line);
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
                        self.emit(Op::null);
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
                                "add" => self.emit(Op::dyn_add),
                                "sub" => self.emit(Op::f64_sub),
                                _ => self.emit(Op::dyn_add),
                            }
                            self.emit_var_set(&var);
                        }
                    }
                    self.emit(Op::null);
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
                    self.emit_u16(Op::global_get, name_idx);
                    for a in args { self.compile_expr(a)?; }
                    self.emit_u8(Op::call_ref, args.len() as u8);
                }
                BuiltinEmit::Noop => {
                    self.emit(Op::null);
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
        let handled = common::dispatch::emit_common(name, self.chunk(), line2);
        if !handled {
            eprintln!("Unknown common emit: {}", name);
        }
    }

    /// Emit a named opcode sequence for a builtin.
    /// Emit a single opcode by name. Used for value methods where args are already on stack.
    fn emit_named_opcode(&mut self, op_name: &str) {
        let line = self.line;
        match op_name {
            "f64_abs" => self.emit(Op::f64_abs),
            "f64_floor" => self.emit(Op::f64_floor),
            "f64_ceil" => self.emit(Op::f64_ceil),
            "f64_sqrt" => self.emit(Op::f64_sqrt),
            "f64_trunc" => self.emit(Op::f64_trunc),
            "f64_nearest" => self.emit(Op::f64_nearest),
            "f64_min" => self.emit(Op::f64_min),
            "f64_max" => self.emit(Op::f64_max),
            "i32_from_f64" => self.emit(Op::i32_from_f64),
            "f64_from_i32" => self.emit(Op::f64_from_i32),
            "dyn_to_bool" => self.emit(Op::dyn_to_bool),
            "dyn_not" => self.emit(Op::dyn_not),
            "ref_is_null" => self.emit(Op::ref_is_null),
            "ref_is_array" => self.emit(Op::ref_is_array),
            "ref_typeof" => self.emit(Op::ref_typeof),
            "str_length" => self.emit(Op::str_length),
            "str_to_upper" => self.emit(Op::str_to_upper),
            "str_to_lower" => self.emit(Op::str_to_lower),
            "str_trim" => self.emit(Op::str_trim),
            "str_trim_start" => self.emit(Op::str_trim_start),
            "str_trim_end" => self.emit(Op::str_trim_end),
            "str_reverse" => self.emit(Op::str_reverse),
            "str_from_char_code" => self.emit(Op::str_from_char_code),
            "str_char_at" => self.emit(Op::str_char_at),
            "str_char_code_at" => self.emit(Op::str_char_code_at),
            "str_starts_with" => self.emit(Op::str_starts_with),
            "str_ends_with" => self.emit(Op::str_ends_with),
            "str_index_of" => self.emit(Op::str_index_of),
            "str_last_index_of" => self.emit(Op::str_last_index_of),
            "str_includes" => self.emit(Op::str_index_of), // includes → index_of (check != -1 at runtime)
            "str_substring" => self.emit(Op::str_substring),
            "str_split" => self.emit(Op::str_split),
            "str_replace" => self.emit(Op::str_replace),
            "str_repeat" => self.emit(Op::str_repeat),
            "str_pad_start" => self.emit(Op::str_pad_start),
            "str_pad_end" => self.emit(Op::str_pad_end),
            "str_compare" => self.emit(Op::str_compare),
            "str_concat" => self.emit(Op::str_concat),
            "array_push" => self.emit(Op::array_push),
            "array_pop" => self.emit(Op::array_pop),
            "array_shift" => self.emit(Op::array_shift),
            "array_reverse" => self.emit(Op::array_reverse),
            "array_join" => self.emit(Op::array_join),
            "array_concat" => self.emit(Op::array_concat),
            "array_fill" => self.emit(Op::array_fill),
            "array_length" => self.emit(Op::array_length),
            "array_slice" => self.emit(Op::array_slice),
            "array_get" => self.emit(Op::array_get),
            "array_set" => self.emit(Op::array_set),
            _ => { let c = self.str_const(op_name); self.emit_u16(Op::global_get, c); }
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
            "min" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; common::math::emit_min(self.chunk(), line); } else { self.emit(Op::null); } }
            "max" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; common::math::emit_max(self.chunk(), line); } else { self.emit(Op::null); } }
            "sqr" => { self.compile_expr(args[0])?; self.emit(Op::dup); self.emit(Op::f64_mul); }
            "succ" => { self.compile_expr(args[0])?; self.emit_const(Value::F64(1.0)); self.emit(Op::dyn_add); }
            "pred" => { self.compile_expr(args[0])?; self.emit_const(Value::F64(1.0)); self.emit(Op::f64_sub); }
            "to_upper" => { self.compile_expr(args[0])?; common::strings::emit_to_upper(self.chunk(), line); }
            "to_lower" => { self.compile_expr(args[0])?; common::strings::emit_to_lower(self.chunk(), line); }
            "trim" => { self.compile_expr(args[0])?; common::strings::emit_trim(self.chunk(), line); }
            "concat" => { for a in args { self.compile_expr(a)?; } common::strings::emit_concat(self.chunk(), args.len(), line); }
            "replace" => { if args.len() >= 3 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; self.compile_expr(args[2])?; common::strings::emit_replace(self.chunk(), line); } }
            "repeat" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; common::strings::emit_repeat(self.chunk(), line); } }
            "leftstr" => { self.compile_expr(args[0])?; self.emit_const(Value::F64(0.0)); self.compile_expr(args[1])?; common::strings::emit_substring(self.chunk(), line); }
            "high" => { self.compile_expr(args[0])?; common::strings::emit_length(self.chunk(), line); self.emit_const(Value::F64(1.0)); self.emit(Op::f64_sub); }
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
                self.emit(Op::null);
            }
            "trim_start" => { self.compile_expr(args[0])?; common::strings::emit_trim_start(self.chunk(), line); }
            "trim_end" => { self.compile_expr(args[0])?; common::strings::emit_trim_end(self.chunk(), line); }
            "pow" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; common::math::emit_pow(self.chunk(), line); } }
            "log" => { self.compile_expr(args[0])?; common::math::emit_log(self.chunk(), line); }
            "sin" => { self.compile_expr(args[0])?; common::math::emit_sin(self.chunk(), line); }
            "cos" => { self.compile_expr(args[0])?; common::math::emit_cos(self.chunk(), line); }
            "tan" => { self.compile_expr(args[0])?; common::math::emit_tan(self.chunk(), line); }
            "exp" => { self.compile_expr(args[0])?; common::math::emit_exp(self.chunk(), line); }
            "is_null" => { self.compile_expr(args[0])?; self.emit(Op::ref_is_null); }
            "space" => { self.emit_const(Value::String(Arc::from(" "))); self.compile_expr(args[0])?; common::strings::emit_repeat(self.chunk(), line); }
            "assigned" => { self.compile_expr(args[0])?; self.emit(Op::ref_is_null); self.emit(Op::dyn_not); }
            "freeandnil" => {
                if let Some(first) = args.first() {
                    if let ExprKind::Ident(var) = &first.kind {
                        let var = var.clone();
                        self.emit(Op::null);
                        self.emit_var_set(&var);
                    }
                }
                self.emit(Op::null);
            }
            // Direct WASM opcode names
            "f64_abs" => { self.compile_expr(args[0])?; self.emit(Op::f64_abs); }
            "f64_floor" => { self.compile_expr(args[0])?; self.emit(Op::f64_floor); }
            "f64_ceil" => { self.compile_expr(args[0])?; self.emit(Op::f64_ceil); }
            "f64_sqrt" => { self.compile_expr(args[0])?; self.emit(Op::f64_sqrt); }
            "f64_trunc" => { self.compile_expr(args[0])?; self.emit(Op::f64_trunc); }
            "f64_nearest" => { self.compile_expr(args[0])?; self.emit(Op::f64_nearest); }
            "f64_min" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; self.emit(Op::f64_min); } }
            "f64_max" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; self.emit(Op::f64_max); } }
            "i32_from_f64" => { self.compile_expr(args[0])?; self.emit(Op::i32_from_f64); }
            "f64_from_i32" => { self.compile_expr(args[0])?; self.emit(Op::f64_from_i32); }
            "dyn_to_bool" => { self.compile_expr(args[0])?; self.emit(Op::dyn_to_bool); }
            "ref_is_null" => { self.compile_expr(args[0])?; self.emit(Op::ref_is_null); }
            "ref_is_array" => { self.compile_expr(args[0])?; self.emit(Op::ref_is_array); }
            "ref_typeof" => { self.compile_expr(args[0])?; self.emit(Op::ref_typeof); }
            "str_length" => { self.compile_expr(args[0])?; self.emit(Op::str_length); }
            "str_to_upper" => { self.compile_expr(args[0])?; self.emit(Op::str_to_upper); }
            "str_to_lower" => { self.compile_expr(args[0])?; self.emit(Op::str_to_lower); }
            "str_trim" => { self.compile_expr(args[0])?; self.emit(Op::str_trim); }
            "str_trim_start" => { self.compile_expr(args[0])?; self.emit(Op::str_trim_start); }
            "str_trim_end" => { self.compile_expr(args[0])?; self.emit(Op::str_trim_end); }
            "str_reverse" => { self.compile_expr(args[0])?; self.emit(Op::str_reverse); }
            "str_from_char_code" => { self.compile_expr(args[0])?; self.emit(Op::str_from_char_code); }
            "str_compare" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; self.emit(Op::str_compare); } }
            "str_split" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; self.emit(Op::str_split); } }
            "str_repeat" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; self.emit(Op::str_repeat); } }
            "array_join" => { if args.len() >= 2 { self.compile_expr(args[0])?; self.compile_expr(args[1])?; self.emit(Op::array_join); } }
            _ => { self.emit(Op::null); }
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
                self.emit(Op::i32_and);
            }
            "ubound" => {
                self.compile_expr(args[0])?;
                common::collections::emit_len(self.chunk(), line);
                self.emit_const(Value::I32(1));
                self.emit(Op::i32_sub);
            }
            "lbound" => {
                self.emit(Op::i32_const_0);
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
                    self.emit(Op::drop);
                }
                self.emit_u16(Op::array_new, 0);
            }
            "asc" => {
                self.compile_expr(args[0])?;
                self.emit(Op::i32_const_0);
                self.emit(Op::str_char_code_at);
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
                    self.emit(Op::null);
                }
            }
            "left" => {
                // Left(s, n) → substring(s, 0, n)
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.emit(Op::i32_const_0);
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::null);
                }
            }
            "string_isnullorempty" => {
                // String.IsNullOrEmpty(s) → s is null OR str_length(s) == 0.
                // Compile s, dup, ref_is_null → if true return true, else
                // str_length == 0.
                if let Some(arg) = args.first() {
                    self.compile_expr(arg)?;
                    // [s]
                    self.emit(Op::dup);
                    // [s, s]
                    self.emit(Op::ref_is_null);
                    // [s, is_null]
                    let if_null = self.emit_jump(Op::br_if_true);
                    // not null branch: [s] → str_length → cmp 0
                    common::strings::emit_length(self.chunk(), line);
                    self.emit(Op::i32_const_0);
                    self.emit(Op::dyn_eq);
                    let end = self.emit_jump(Op::br);
                    // null branch: drop [s], push true
                    self.patch_jump(if_null);
                    self.emit(Op::drop);
                    self.emit(Op::r#true);
                    self.patch_jump(end);
                } else {
                    self.emit(Op::r#true);
                }
            }
            "mid" | "mid_1based" => {
                // Mid(s, start[, len]) — 1-based
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::i32_sub); // start0
                    if args.len() >= 3 {
                        self.emit(Op::dup);
                        self.compile_expr(args[2])?;
                        common::convert::emit_to_int(self.chunk(), line);
                        self.emit(Op::i32_add); // start0 + length
                    } else {
                        self.emit_const(Value::I32(0x7FFF_FFFF));
                    }
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::null);
                }
            }
            "number_isnan" => {
                self.compile_expr(args[0])?;
                self.emit(Op::dup);
                self.emit(Op::dyn_ne);
            }
            "number_isfinite" => {
                self.compile_expr(args[0])?;
                common::math::emit_abs(self.chunk(), line);
                self.emit_const(Value::F64(f64::MAX));
                self.emit(Op::dyn_le);
            }
            "number_isinteger" => {
                self.compile_expr(args[0])?;
                self.emit(Op::dup);
                self.emit(Op::f64_trunc);
                self.emit(Op::dyn_eq);
            }
            "map_size" => {
                self.compile_expr(args[0])?;
                common::dict::emit_keys(self.chunk(), line);
                common::collections::emit_len(self.chunk(), line);
            }
            "array_at" => {
                if args.len() >= 1 {
                    // Already have object from value method dispatch
                    self.compile_expr(args[0])?;
                    common::collections::emit_get(self.chunk(), line);
                } else {
                    self.emit(Op::null);
                }
            }
            "instr" => {
                if args.len() == 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::strings::emit_index_of(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::i32_add);
                } else if args.len() == 3 {
                    self.compile_expr(args[1])?;
                    self.compile_expr(args[0])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::i32_sub);
                    self.emit_const(Value::I32(0x7FFF_FFFF));
                    common::strings::emit_substring(self.chunk(), line);
                    self.compile_expr(args[2])?;
                    common::strings::emit_index_of(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::i32_add);
                } else {
                    self.emit(Op::null);
                }
            }
            "instrrev" => {
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::strings::emit_last_index_of(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::i32_add);
                } else {
                    self.emit(Op::null);
                }
            }
            "replace" => {
                if args.len() >= 3 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    self.compile_expr(args[2])?;
                    common::strings::emit_replace(self.chunk(), line);
                } else {
                    self.emit(Op::null);
                }
            }
            "split" => {
                self.compile_expr(args[0])?;
                if args.len() >= 2 {
                    self.compile_expr(args[1])?;
                } else {
                    self.emit_const(Value::String(Arc::from(" ")));
                }
                self.emit(Op::str_split);
            }
            "join" => {
                self.compile_expr(args[0])?;
                if args.len() >= 2 {
                    self.compile_expr(args[1])?;
                } else {
                    self.emit_const(Value::String(Arc::from("")));
                }
                self.emit(Op::array_join);
            }

            // ── Pascal ordinal/array intrinsics (canonical compiler_common ops) ──

            "high" => {
                // High(arr) → __len__(arr) - 1
                self.compile_expr(args[0])?;
                common::collections::emit_len(self.chunk(), line);
                self.emit_const(Value::I32(1));
                self.emit(Op::i32_sub);
            }
            "low" => {
                // Low(arr) → 0 (always 0 for dynamic arrays in our VM)
                self.emit(Op::i32_const_0);
            }
            "succ" => {
                // Succ(x) → x + 1
                self.compile_expr(args[0])?;
                self.emit_const(Value::F64(1.0));
                self.emit(Op::dyn_add);
            }
            "pred" => {
                // Pred(x) → x - 1
                self.compile_expr(args[0])?;
                self.emit_const(Value::F64(1.0));
                self.emit(Op::f64_sub);
            }
            "sqr" => {
                // Sqr(x) → x * x (square, NOT square root)
                self.compile_expr(args[0])?;
                self.emit(Op::dup);
                self.emit(Op::f64_mul);
            }
            "assigned" => {
                // Assigned(x) → x is not null
                self.compile_expr(args[0])?;
                self.emit(Op::null);
                self.emit(Op::dyn_ne);
            }
            "sizeof" => {
                // SizeOf(x) → 4 (boxed value)
                self.compile_expr(args[0])?;
                self.emit(Op::drop);
                self.emit_const(Value::I32(4));
            }
            "classname" => {
                // ClassName(obj) → obj.__type
                self.compile_expr(args[0])?;
                let idx = self.str_const("__type");
                self.emit_u16(Op::struct_get, idx);
            }
            "pos" => {
                // Pos(substr, s) → IndexOf(s, substr) + 1 (Pascal 1-based)
                if args.len() == 2 {
                    self.compile_expr(args[1])?;
                    self.compile_expr(args[0])?;
                    common::strings::emit_index_of(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::i32_add);
                } else {
                    self.emit(Op::null);
                }
            }
            "copy" => {
                // Copy(s, start, len) → substring(s, start-1, start-1+len) — Pascal 1-based
                if args.len() >= 2 {
                    self.compile_expr(args[0])?;
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit_const(Value::I32(1));
                    self.emit(Op::i32_sub);
                    if args.len() >= 3 {
                        self.emit(Op::dup);
                        self.compile_expr(args[2])?;
                        common::convert::emit_to_int(self.chunk(), line);
                        self.emit(Op::i32_add);
                    } else {
                        self.emit_const(Value::I32(0x7FFF_FFFF));
                    }
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::null);
                }
            }
            "leftstr" => {
                // LeftStr(s, n) → substring(s, 0, n)
                if args.len() == 2 {
                    self.compile_expr(args[0])?;
                    self.emit(Op::i32_const_0);
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::null);
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
                    self.emit_u16(Op::local_set, s_slot); self.emit(Op::drop);
                    self.emit_u16(Op::local_get, s_slot);
                    self.emit_u16(Op::local_get, s_slot);
                    common::strings::emit_length(self.chunk(), line);
                    self.compile_expr(args[1])?;
                    common::convert::emit_to_int(self.chunk(), line);
                    self.emit(Op::i32_sub);
                    self.emit_u16(Op::local_get, s_slot);
                    common::strings::emit_length(self.chunk(), line);
                    common::strings::emit_substring(self.chunk(), line);
                } else {
                    self.emit(Op::null);
                }
            }

            _ => { self.emit(Op::null); }
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
