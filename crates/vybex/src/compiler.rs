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
                self.compile_expr(value)?;
                for (i, target) in targets.iter().enumerate() {
                    if i < targets.len() - 1 { self.emit(Op::DUP); }
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
            StmtKind::FunctionDecl { name, params, return_type, body, modifiers: _, handles, is_async: _, is_generator, is_sub } => {
                self.compile_function_decl(name, params, return_type, body, *is_sub, *is_generator, handles)?;
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
    // Function declaration compilation
    // ════════════════════════════════════════════════════════════════════════

    fn compile_function_decl(
        &mut self, name: &str, params: &[Param], return_type: &Option<String>,
        body: &[Statement], _is_sub: bool, _is_generator: bool, handles: &[String],
    ) -> Result<(), String> {
        let cname = self.canon(name);
        self.defined_globals.insert(cname.clone());
        self.defined_functions.insert(cname.clone());
        let name = &cname;

        let has_rest = params.last().map_or(false, |p| p.is_rest);
        let arity: u8 = if has_rest { 255 } else { params.len() as u8 };
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
                self.emit_u16(Op::LOCAL_GET, slot);
                self.emit(Op::REF_IS_NULL);
                let has_val = self.emit_jump(Op::BR_IF_FALSE);
                self.compile_expr(default)?;
                self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                self.patch_jump(has_val);
            }
        }

        // Rest param preamble: collect excess args into an array.
        // With arity=255 the VM doesn't truncate excess args. They land in
        // sequential slots after the non-rest params. We scan those slots
        // with unrolled local_get + null-check (local_get is static u16).
        // Caps at 16 rest args which covers all realistic use cases.
        if has_rest {
            let rest_name = &params.last().unwrap().name;
            let rest_slot = self.scope().resolve(rest_name).unwrap();
            // Build array from slots rest_slot..rest_slot+16, stopping at null.
            // Pattern per slot: if local[N] is null → jump to done; else arr.push(local[N])
            // Build rest array via `common::collections` so the provider
            // is swappable in one place. `wasm:js-array.push` returns
            // new_length (ECMA-262), not arr, so we stash arr in a
            // scope-local and reload each iteration.
            let line = self.line;
            // Reserve the 16 rest-arg slots before allocating `__rest_arr` so
            // the accumulator doesn't overwrite an incoming rest argument.
            // (The VM parks overflow args in slots rest_slot..rest_slot+argc-arity;
            // without this reservation, `__rest_arr` landed on the second rest arg,
            // triggering a self-referential push loop.)
            let max_rest = 16u16;
            for i in 1..max_rest {
                self.scope_mut().define(&format!("__rest_reserved_{}", i));
            }
            common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
            let rest_arr = self.scope_mut().define("__rest_arr");
            self.emit_u16(Op::LOCAL_SET, rest_arr); self.emit(Op::DROP);
            let mut done_patches: Vec<usize> = Vec::new();
            for i in 0..max_rest {
                let slot = rest_slot + i;
                self.emit_u16(Op::LOCAL_GET, slot);
                self.emit(Op::REF_IS_NULL);
                done_patches.push(self.emit_jump(Op::BR_IF_TRUE)); // null → done
                self.emit_u16(Op::LOCAL_GET, rest_arr);
                self.emit_u16(Op::LOCAL_GET, slot);
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP); // drop new_length
            }
            for p in done_patches { self.patch_jump(p); }
            // Store rest array back into the rest_slot param position.
            self.emit_u16(Op::LOCAL_GET, rest_arr);
            self.emit_u16(Op::LOCAL_SET, rest_slot);
            self.emit(Op::DROP);
        }

        // Result slot for functions with return type (Pascal/VB Function).
        // The slot name is profile-driven so VB can keep it internal
        // (`__result__`) and avoid shadowing user classes named `Result`,
        // while Pascal keeps it as `Result` (user-visible per Pascal idiom).
        let result_slot = if return_type.is_some() && self.profile.function_return == ReturnStyle::ResultSlot {
            let slot_name = self.profile.result_slot_name.clone();
            let rs = self.scope_mut().define(&slot_name);
            self.emit(Op::NULL); self.emit_u16(Op::LOCAL_SET, rs); self.emit(Op::DROP);
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
            self.emit_u16(Op::LOCAL_GET, rs);
            self.emit(Op::RETURN);
        } else {
            let line = self.line;
            common::functions::emit_function_epilogue(&mut self.chunks[func_idx], line);
        }

        let locals = self.scope().next_slot.max(self.chunks[func_idx].local_count);
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
        self.emit_u16(Op::GLOBAL_SET, idx);
        self.emit(Op::DROP);

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
                self.emit(Op::DROP); // statement: discard host call result
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
                    if let StmtKind::FunctionDecl { name: mname, params, return_type, body, modifiers, is_sub: _, .. } = &stmt.kind {
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
                                self.emit_u16(Op::LOCAL_GET, slot);
                                self.emit(Op::RETURN);
                            }
                        } else if return_type.is_some() && result_style == ReturnStyle::ResultSlot {
                            let slot_name = self.profile.result_slot_name.clone();
                            let rs = self.scope_mut().define(&slot_name);
                            self.emit(Op::NULL); self.emit_u16(Op::LOCAL_SET, rs); self.emit(Op::DROP);
                            let saved_fn = self.current_func_name.take();
                            let saved_rs = self.current_result_slot.take();
                            self.current_func_name = Some(mname.clone());
                            self.current_result_slot = Some(rs);
                            for s in body { self.compile_stmt(s)?; }
                            self.current_func_name = saved_fn;
                            self.current_result_slot = saved_rs;
                            self.emit_u16(Op::LOCAL_GET, rs);
                            self.emit(Op::RETURN);
                        } else {
                            for s in body { self.compile_stmt(s)?; }
                            let line = self.line;
                            common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
                        }

                        let locals = self.scope().next_slot.max(self.chunks[ci].local_count);
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
                                self.emit_u16(Op::LOCAL_GET, slot);
                                let backing = self.str_const(&format!("__{}", pname_canon));
                                self.emit_u16(Op::STRUCT_GET, backing);
                                self.emit(Op::RETURN);
                            }
                        } else {
                            let slot_name = self.profile.result_slot_name.clone();
                            let rs = self.scope_mut().define(&slot_name);
                            self.emit(Op::NULL); self.emit_u16(Op::LOCAL_SET, rs); self.emit(Op::DROP);
                            let saved_fn = self.current_func_name.take();
                            let saved_rs = self.current_result_slot.take();
                            self.current_func_name = Some(pname.clone());
                            self.current_result_slot = Some(rs);
                            for s in getter_body { self.compile_stmt(s)?; }
                            self.current_func_name = saved_fn;
                            self.current_result_slot = saved_rs;
                            self.emit_u16(Op::LOCAL_GET, rs);
                            self.emit(Op::RETURN);
                        }

                        let locals = self.scope().next_slot.max(self.chunks[ci].local_count);
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
                                self.emit_u16(Op::LOCAL_GET, self_slot);
                                if let Some(val_slot) = self.scope().resolve(&setter_info.param.name) {
                                    self.emit_u16(Op::LOCAL_GET, val_slot);
                                }
                                let backing = self.str_const(&format!("__{}", pname_canon));
                                self.emit_u16(Op::STRUCT_SET, backing);
                                self.emit(Op::DROP);
                            }
                        } else {
                            for s in &setter_info.body { self.compile_stmt(s)?; }
                        }

                        let line = self.line;
                        common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
                        let locals = self.scope().next_slot.max(self.chunks[ci].local_count);
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
                    self.emit_u16(Op::GLOBAL_SET, idx);
                    self.emit(Op::DROP);
                    self.defined_globals.insert(global_name);
                }
                ClassMember::Event { .. } => { /* type-level only */ }
                ClassMember::NestedType(stmt) => { self.compile_stmt(stmt)?; }
                _ => {}
            }
        }

        self.current_class = saved_class;

        // Find constructor body and its user arity
        let _ctor = method_chunks.iter().find(|(_, _, is_ctor, _)| *is_ctor);
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
        // Also handle default parameter values from the Param structs.
        let ctor_param_defaults: Vec<Option<Expression>> = ctor_body.map(|(_, params, _)| {
            let skip = if self.profile.explicit_self_param { 1 } else { 0 };
            params.iter().skip(skip).map(|p| p.default.clone()).collect()
        }).unwrap_or_default();
        for (i, p) in user_params.iter().enumerate() {
            self.scope_mut().define(p);
            if let Some(Some(ref default)) = ctor_param_defaults.get(i) {
                let slot = self.scope().resolve(p).unwrap();
                self.emit_u16(Op::LOCAL_GET, slot);
                self.emit(Op::REF_IS_NULL);
                let has_val = self.emit_jump(Op::BR_IF_FALSE);
                self.compile_expr(default)?;
                self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                self.patch_jump(has_val);
            }
        }
        self.scope_mut().define(&self_kw); // this_slot = user_arity
        let this_slot = user_arity as u16;

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
            self.emit(Op::NULL);
            self.emit_u16(Op::LOCAL_SET, this_slot);
            self.emit(Op::DROP);

            // Determine if auto_base_call should kick in:
            // ctor exists + base_args None + profile says auto + parent exists + body has no super
            let has_explicit_base = ctor_body.as_ref().map_or(false, |(_, _, ba)| ba.is_some());
            let auto_base_needed = !has_explicit_base
                && ctor_body.is_some()
                && self.profile.auto_base_call
                && parent.is_some()
                && {
                    let stmts = ctor_body.as_ref().map(|(b, _, _)| b.as_slice()).unwrap_or(&[]);
                    !body_has_super_call(stmts)
                };

            if let Some((_, _, base_args)) = &ctor_body {
                if let Some(bargs) = base_args {
                    // Explicit base_args provided (C#-style `: base(args)`)
                    if let Some(parent_name) = parent {
                        let pname = self.canon(parent_name);
                        let pidx = self.str_const(&pname);
                        self.emit_u16(Op::GLOBAL_GET, pidx);
                        for a in *bargs { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL, bargs.len() as u8);
                        self.emit_u16(Op::LOCAL_SET, this_slot);
                        self.emit(Op::DROP);
                    }
                } else if auto_base_needed {
                    // Profile-driven auto base call (VB/C#/Pascal):
                    // body has no super() → auto-call parent() with 0 args.
                    if let Some(parent_name) = parent {
                        let pname = self.canon(parent_name);
                        let pidx = self.str_const(&pname);
                        self.emit_u16(Op::GLOBAL_GET, pidx);
                        self.emit_u8(Op::CALL, 0);
                        self.emit_u16(Op::LOCAL_SET, this_slot);
                        self.emit(Op::DROP);
                    }
                }
                // else: JS pattern — body calls super() itself, sets this_slot
            } else {
                // No explicit constructor — auto-call parent with user args
                if let Some(parent_name) = parent {
                    let pname = self.canon(parent_name);
                    let pidx = self.str_const(&pname);
                    self.emit_u16(Op::GLOBAL_GET, pidx);
                    for i in 0..user_arity {
                        self.emit_u16(Op::LOCAL_GET, (i as u16) + 1);
                    }
                    self.emit_u8(Op::CALL, user_arity);
                    self.emit_u16(Op::LOCAL_SET, this_slot);
                    self.emit(Op::DROP);
                }
            }

            if has_explicit_base || auto_base_needed || ctor_body.is_none() {
                // C#-style: base call already done above, or no-ctor auto-call done above.
                // Order: re-stamp __type → fields → save base → bind methods → body
                //
                // The parent ctor stamped __type with the parent name. Re-stamp with
                // the child name so `obj is ChildType` returns true.
                self.emit_u16(Op::LOCAL_GET, this_slot);
                self.emit_const(Value::String(Arc::from(name)));
                let type_key = self.str_const("__type");
                self.emit_u16(Op::STRUCT_SET, type_key);
                self.emit(Op::DROP);

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

                // Auto-init methods from profile (e.g. InitializeComponent for .NET forms).
                // Emitted after method binding but before user body.
                {
                    let ctor_stmts: &[Statement] = ctor_body
                        .as_ref().map(|(b, _, _)| b.as_slice()).unwrap_or(&[]);
                    let auto_inits = self.profile.auto_init_methods.clone();
                    for aim in &auto_inits {
                        let has_method = instance_methods.iter()
                            .any(|(n, _, _, _)| n.eq_ignore_ascii_case(aim));
                        if has_method && !body_calls_method(ctor_stmts, aim) {
                            common::classes::emit_auto_init_call(self.chunk(), this_slot, aim, line);
                        }
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
                self.emit_u16(Op::LOCAL_GET, this_slot);
                self.emit_const(Value::String(Arc::from(name)));
                let type_key2 = self.str_const("__type");
                self.emit_u16(Op::STRUCT_SET, type_key2);
                self.emit(Op::DROP);

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

                // Auto-init methods from profile (e.g. InitializeComponent for .NET forms).
                // Emitted after method binding so struct_get finds the method,
                // but before user body so controls exist for AddHandler etc.
                let user_body = &body_stmts[preamble_end..];
                let auto_inits = self.profile.auto_init_methods.clone();
                for aim in &auto_inits {
                    let has_method = instance_methods.iter()
                        .any(|(n, _, _, _)| n.eq_ignore_ascii_case(aim));
                    if has_method && !body_calls_method(user_body, aim) {
                        common::classes::emit_auto_init_call(self.chunk(), this_slot, aim, line);
                    }
                }

                // Compile the main body (everything after the preamble).
                for s in user_body {
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

            // Auto-init methods from profile (e.g. InitializeComponent for .NET forms).
            let ctor_stmts: &[Statement] = ctor_body
                .as_ref().map(|(b, _, _)| b.as_slice()).unwrap_or(&[]);
            let auto_inits = self.profile.auto_init_methods.clone();
            for aim in &auto_inits {
                let has_method = instance_methods.iter()
                    .any(|(n, _, _, _)| n.eq_ignore_ascii_case(aim));
                if has_method && !body_calls_method(ctor_stmts, aim) {
                    common::classes::emit_auto_init_call(self.chunk(), this_slot, aim, line);
                }
            }

            // Run user constructor body
            if let Some((body, _, _)) = ctor_body {
                for s in body { self.compile_stmt(s)?; }
            }
        }

        // Finalize: instanceof chain
        common::classes::emit_instanceof_chain(&mut self.chunks, self.current, this_slot, name, line);
        common::classes::emit_constructor_return(self.chunk(), this_slot, line);

        let locals = self.scope().next_slot.max(self.chunks[ctor_idx].local_count);
        self.chunks[ctor_idx].local_count = locals;
        self.scopes.pop();
        self.current = saved_cur;
        self.current_class = saved_class2;

        // Store constructor globally and register type
        let ctor_local = self.scope_mut().define(&format!("__{}_ctor", name));
        common::classes::emit_store_constructor(self.chunk(), name, ctor_idx, ctor_local, line);

        // Initialize static fields on the constructor object
        for (fname, init) in &static_field_inits {
            self.emit_u16(Op::LOCAL_GET, ctor_local);
            if let Some(init_expr) = init {
                self.compile_expr(init_expr)?;
            } else {
                self.emit(Op::NULL);
            }
            let fk = self.str_const(fname);
            self.emit_u16(Op::STRUCT_SET, fk);
            self.emit(Op::DROP);
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
                    Literal::Bool(b) => if *b { self.emit(Op::TRUE) } else { self.emit(Op::FALSE) },
                    Literal::Null => self.emit(Op::NULL),
                    Literal::Undefined => { let l = self.line; common::expressions::emit_undefined(self.chunk(), l); }
                    Literal::Ellipsis => self.emit(Op::NULL),
                }
            }

            // ── Identifier ──────────────────────────────────────────────
            ExprKind::Ident(name) => {
                // JS global constants that aren't variables
                match name.as_str() {
                    "NaN" => { self.emit_const(Value::F64(f64::NAN)); return Ok(()); }
                    "Infinity" => { self.emit_const(Value::F64(f64::INFINITY)); return Ok(()); }
                    "undefined" if self.case_sensitive => { let l = self.line; common::expressions::emit_undefined(self.chunk(), l); return Ok(()); }
                    _ => {}
                }
                // Local variable / parameter takes priority over implicit self field
                let is_local = self.scope().resolve(name).is_some()
                    || (!self.case_sensitive && self.scope().resolve_ci(name).is_some());

                // Implicit self field access (only if NOT a local)
                if !is_local && self.is_class_field(name) {
                    let self_kw = self.profile.self_keyword.clone();
                    if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                        self.emit_u16(Op::LOCAL_GET, slot);
                        let field_name = self.canon(name);
                        let idx = self.str_const(&field_name);
                        self.emit_u16(Op::STRUCT_GET, idx);
                        return Ok(());
                    }
                }

                // Bare enum member: `Green` → `TColor.Green`
                if !is_local {
                    let canon_name = self.canon(name);
                    if let Some(enum_type) = self.enum_members.get(&canon_name).cloned() {
                        let type_idx = self.str_const(&enum_type);
                        self.emit_u16(Op::GLOBAL_GET, type_idx);
                        let mem_idx = self.str_const(&canon_name);
                        self.emit_u16(Op::STRUCT_GET, mem_idx);
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
                    self.emit_u16(Op::LOCAL_GET, slot);
                } else if self.scopes.len() > 1 {
                    // Arrow function: capture `this` from enclosing scope via upvalue
                    let kw = self.profile.self_keyword.clone();
                    if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, &kw) {
                        self.emit_u8(Op::UPVALUE_GET, uv);
                    } else {
                        self.emit(Op::NULL);
                    }
                } else {
                    self.emit(Op::NULL);
                }
            }

            ExprKind::Super => {
                // super refers to the parent class constructor.
                // Look up the parent from the current class's PendingClass info.
                if let Some(ref class_name) = self.current_class.clone() {
                    if let Some(parent_name) = self.pending_classes.get(class_name.as_str()).and_then(|pc| pc.parent.clone()) {
                        let pname = self.canon(&parent_name);
                        let idx = self.str_const(&pname);
                        self.emit_u16(Op::GLOBAL_GET, idx);
                    } else {
                        self.emit(Op::NULL);
                    }
                } else {
                    self.emit(Op::NULL);
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
                    self.emit(Op::DUP);
                    self.emit(Op::REF_IS_NULL);
                    let skip = self.emit_jump(Op::BR_IF_FALSE);
                    self.emit(Op::DROP);
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
                        if *op == UnaryOp::PostInc { self.emit(Op::DUP); }
                        self.emit_const(Value::F64(1.0));
                        self.emit(Op::DYN_ADD);
                        if *op == UnaryOp::PreInc { self.emit(Op::DUP); }
                        self.compile_assign_target(inner)?;
                    }
                    UnaryOp::PreDec | UnaryOp::PostDec => {
                        self.compile_expr(inner)?;
                        if *op == UnaryOp::PostDec { self.emit(Op::DUP); }
                        self.emit_const(Value::F64(1.0));
                        self.emit(Op::F64_SUB);
                        if *op == UnaryOp::PreDec { self.emit(Op::DUP); }
                        self.compile_assign_target(inner)?;
                    }
                    _ => {
                        self.compile_expr(inner)?;
                        match op {
                            UnaryOp::Neg => { let l = self.line; common::math::emit_neg(self.chunk(), l); }
                            UnaryOp::Pos => {
                                // JS `+v` coerces to number. Route through vybe:convert:toNumber.
                                let idx = self.import("vybe:convert", "toNumber");
                                self.emit_host_call(idx, 1);
                            }
                            UnaryOp::Not => self.emit(Op::DYN_NOT),
                            UnaryOp::BitNot => { let l = self.line; common::expressions::emit_i32_not(self.chunk(), l); }
                            UnaryOp::Typeof => self.emit(Op::REF_TYPEOF),
                            UnaryOp::Void => { self.emit(Op::DROP); self.emit(Op::NULL); }
                            UnaryOp::Delete => { self.emit(Op::DROP); self.emit(Op::TRUE); }
                            UnaryOp::Deref => { let idx = self.str_const("__value"); self.emit_u16(Op::STRUCT_GET, idx); }
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
                self.emit(Op::DYN_TO_BOOL);
                let else_j = self.emit_jump(Op::BR_IF_FALSE);
                self.compile_expr(then)?;
                let end_j = self.emit_jump(Op::BR);
                self.patch_jump(else_j);
                self.compile_expr(else_)?;
                self.patch_jump(end_j);
            }

            // ── Call ────────────────────────────────────────────────────
            ExprKind::Call { callee, args, optional } => {
                if *optional {
                    // Optional call: callee?.() — short-circuit to null if callee is null/undefined.
                    // Stack: compile callee → [func_or_null].
                    // Dup to check null while preserving the original for the call.
                    self.compile_expr(callee)?;
                    self.emit(Op::DUP);
                    self.emit(Op::REF_IS_NULL);
                    let skip = self.emit_jump(Op::BR_IF_TRUE);
                    // Not null — call it. Stack: [func]. Compile args, call.
                    for a in args { self.compile_expr(&a.value)?; }
                    self.emit_u8(Op::CALL_REF, args.len() as u8);
                    let end = self.emit_jump(Op::BR);
                    self.patch_jump(skip);
                    // Null path: the dup left [null] on stack, use it as result
                    self.patch_jump(end);
                } else {
                    self.compile_call(callee, args)?;
                }
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
                        self.emit_u8(Op::CALL_REF, 0);
                        return Ok(());
                    }
                }

                self.compile_expr(object)?;

                if *null_safe {
                    // ?. — check null before accessing
                    self.emit(Op::DUP);
                    self.emit(Op::REF_IS_NULL);
                    let skip = self.emit_jump(Op::BR_IF_FALSE);
                    // Object is null — result is null
                    let end = self.emit_jump(Op::BR);
                    self.patch_jump(skip);
                    let field_name = self.canon(field);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::STRUCT_GET, idx);
                    self.patch_jump(end);
                } else {
                    let field_name = self.canon(field);
                    let idx = self.str_const(&field_name);
                    self.emit_u16(Op::STRUCT_GET, idx);
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
                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
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
                        self.emit_u16(Op::GLOBAL_GET, idx);
                        for a in args { self.compile_expr(&a.value)?; }
                        self.emit_u8(Op::CALL_REF, args.len() as u8);
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
                        self.emit_u16(Op::STRUCT_NEW, 0);
                        self.emit(Op::DUP);
                        if let Some(msg_arg) = args.first() {
                            self.compile_expr(&msg_arg.value)?;
                        } else {
                            self.emit_const(Value::String(Arc::from("")));
                        }
                        let line = self.line;
                        common::errors::emit_exception_new_finalize(
                            self.chunk(),
                            type_name, // original casing for `name` field
                            line,
                        );
                        // Stamp `stack` = "Name: message" using locals.
                        // Stack after finalize: [obj]
                        let exc_tmp = self.scope_mut().define("__exc_tmp");
                        self.emit_u16(Op::LOCAL_SET, exc_tmp); self.emit(Op::DROP);
                        // Build "Name: " + message
                        self.emit_const(Value::String(Arc::from(format!("{}: ", type_name))));
                        self.emit_u16(Op::LOCAL_GET, exc_tmp);
                        let msg_k = self.str_const("message");
                        self.emit_u16(Op::STRUCT_GET, msg_k);
                        // Stack: ["Name: ", msg]. str_concat: a=prefix, b=msg → prefix+msg
                        self.emit(Op::STR_CONCAT);
                        // Stack: ["Name: msg"]. Save it.
                        let sv = self.scope_mut().define("__stack_val");
                        self.emit_u16(Op::LOCAL_SET, sv); self.emit(Op::DROP);
                        // Stamp: obj.stack = stack_val
                        self.emit_u16(Op::LOCAL_GET, exc_tmp);
                        self.emit_u16(Op::LOCAL_GET, sv);
                        let sk = self.str_const("stack");
                        self.emit_u16(Op::STRUCT_SET, sk);
                        self.emit(Op::DROP);
                        // Result: push obj
                        self.emit_u16(Op::LOCAL_GET, exc_tmp);
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
                self.emit_u8(Op::CALL_REF, args.len() as u8);
            }

            // ── Assignment as expression ────────────────────────────────
            ExprKind::Assign { target, value } => {
                self.compile_expr(value)?;
                self.emit(Op::DUP);
                self.compile_assign_target(target)?;
            }

            // ── Lambda ──────────────────────────────────────────────────
            ExprKind::Lambda { params, body, .. } => {
                self.compile_lambda(params, body)?;
            }

            // ── Array literal ───────────────────────────────────────────
            ExprKind::Array(elements) => {
                // Array literals funnel through `common::collections` so
                // every language and every array-literal site emits the
                // same import shape. Changing the provider (wasm:js-array
                // → vybe:array → polyfill) happens in ONE file, not here.
                // Keyed elements (PHP `['k' => v]`) still drop the key on
                // this path — full associative semantics route through
                // `wasm:js-object` and are a separate follow-up.
                let line = self.line;
                common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                for elem in elements {
                    if elem.spread {
                        // Spread: `concat(current, other)` returns a NEW
                        // array which replaces the one on TOS.
                        self.compile_expr(&elem.value)?;
                        common::collections::emit_concat(&mut self.chunks, self.current, line);
                    } else {
                        // DUP keeps the array on TOS; push returns the
                        // new length, which we drop.
                        self.emit(Op::DUP);
                        self.compile_expr(&elem.value)?;
                        common::collections::emit_push(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP);
                    }
                }
            }

            // ── Tuple (Python) ──────────────────────────────────────────
            ExprKind::Tuple(elements) => {
                let line = self.line;
                let n = elements.len();
                for elem in elements { self.compile_expr(elem)?; }
                // Allocate N consecutive slots; common::collections::emit_pack_n
                // stashes stack values and re-pushes into a fresh array —
                // same wasm:js-array.* surface as literals.
                let base = if n == 0 { 0 } else {
                    let mut first = 0u16;
                    for i in 0..n {
                        let s = self.scope_mut().define("__pack");
                        if i == 0 { first = s; }
                    }
                    first
                };
                common::collections::emit_pack_n(&mut self.chunks, self.current, n as u16, base, line);
            }

            // ── Set (Python) ────────────────────────────────────────────
            ExprKind::Set(elements) => {
                let line = self.line;
                let n = elements.len();
                for elem in elements { self.compile_expr(elem)?; }
                let base = if n == 0 { 0 } else {
                    let mut first = 0u16;
                    for i in 0..n {
                        let s = self.scope_mut().define("__pack");
                        if i == 0 { first = s; }
                    }
                    first
                };
                common::collections::emit_pack_n(&mut self.chunks, self.current, n as u16, base, line);
                // Convert to set via host call
                let idx = self.import("vybe:collections", "arrayToSet");
                self.emit_host_call(idx, 1);
            }

            // ── Object literal ──────────────────────────────────────────
            // Uses dict::emit_new to create the object WITH __keys tracking.
            // Each key is set via struct_set AND appended to __keys so that
            // Object.keys/values/entries (which read __keys) return the
            // right answer.
            ExprKind::Object(props) => {
                let line = self.line;
                common::dict::emit_new(&mut self.chunks, self.current, line);
                for prop in props {
                    match prop {
                        ObjectProperty::KeyValue { key, value } => {
                            self.emit(Op::DUP);
                            self.compile_expr(value)?;
                            if let ExprKind::Lit(Literal::Str(k)) = &key.kind {
                                let idx = self.str_const(k);
                                self.emit_u16(Op::STRUCT_SET, idx);
                                self.emit(Op::DROP);
                                // Track key in __keys
                                self.emit(Op::DUP);
                                let keys_key = self.str_const("__keys");
                                self.emit_u16(Op::STRUCT_GET, keys_key);
                                self.emit_const(Value::String(Arc::from(k.as_str())));
                                let l = self.line;
                                common::collections::emit_push(&mut self.chunks, self.current, l);
                                self.emit(Op::DROP);
                            } else {
                                // Dynamic key — save key for __keys tracking
                                self.compile_expr(key)?;
                                self.emit(Op::DUP); // [dict, val, key, key]
                                let key_tmp = self.scope_mut().define("__obj_dyn_key");
                                self.emit_u16(Op::LOCAL_SET, key_tmp); self.emit(Op::DROP);
                                let l = self.line;
                                common::collections::emit_set(&mut self.chunks, self.current, l);
                                self.emit(Op::DROP); // drop returned null
                                // Track dynamic key in __keys
                                self.emit(Op::DUP);
                                let keys_key = self.str_const("__keys");
                                self.emit_u16(Op::STRUCT_GET, keys_key);
                                self.emit_u16(Op::LOCAL_GET, key_tmp);
                                let l = self.line;
                                common::collections::emit_push(&mut self.chunks, self.current, l);
                                self.emit(Op::DROP);
                            }
                        }
                        ObjectProperty::Shorthand(name) => {
                            self.emit(Op::DUP);
                            self.emit_var_get(name);
                            let idx = self.str_const(name);
                            self.emit_u16(Op::STRUCT_SET, idx);
                            self.emit(Op::DROP);
                            // Track key in __keys
                            self.emit(Op::DUP);
                            let keys_key = self.str_const("__keys");
                            self.emit_u16(Op::STRUCT_GET, keys_key);
                            self.emit_const(Value::String(Arc::from(name.as_str())));
                            let l = self.line;
                            common::collections::emit_push(&mut self.chunks, self.current, l);
                            self.emit(Op::DROP);
                        }
                        ObjectProperty::Spread(expr) => {
                            // Object spread: merge properties from expr into current object
                            self.compile_expr(expr)?;
                            let idx = self.import("vybe:object", "assign");
                            self.emit_host_call(idx, 2);
                        }
                        ObjectProperty::Method { key, value } => {
                            self.emit(Op::DUP);
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
                                self.emit(Op::NULL);
                            }
                            let idx = self.str_const(key);
                            self.emit_u16(Op::STRUCT_SET, idx);
                            self.emit(Op::DROP);
                        }
                        ObjectProperty::Accessor { kind, key, value } => {
                            self.emit(Op::DUP);
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
                                self.emit(Op::NULL);
                            }
                            let accessor_name = match kind {
                                AccessorKind::Get => format!("__get_{}", key),
                                AccessorKind::Set => format!("__set_{}", key),
                            };
                            let idx = self.str_const(&accessor_name);
                            self.emit_u16(Op::STRUCT_SET, idx);
                            self.emit(Op::DROP);
                        }
                        ObjectProperty::Computed { key, value } => {
                            // wasm:js-array.set expects [obj, key, val] → null
                            self.emit(Op::DUP);
                            self.compile_expr(key)?;
                            self.emit(Op::DUP); // save key for __keys
                            let key_tmp = self.scope_mut().define("__obj_comp_key");
                            self.emit_u16(Op::LOCAL_SET, key_tmp); self.emit(Op::DROP);
                            self.compile_expr(value)?;
                            let l = self.line;
                            common::collections::emit_set(&mut self.chunks, self.current, l);
                            self.emit(Op::DROP); // drop returned null
                            // Track in __keys
                            self.emit(Op::DUP);
                            let keys_key = self.str_const("__keys");
                            self.emit_u16(Op::STRUCT_GET, keys_key);
                            self.emit_u16(Op::LOCAL_GET, key_tmp);
                            let l = self.line;
                            common::collections::emit_push(&mut self.chunks, self.current, l);
                            self.emit(Op::DROP);
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
                            self.emit_u16(Op::GLOBAL_GET, tostring_global);
                            self.compile_expr(e)?;
                            self.emit_u8(Op::CALL_REF, 1);
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
                self.emit_u16(Op::STRUCT_GET, key);
                self.emit_const(Value::String(Arc::from(canon_type.as_str())));
                self.emit(Op::DYN_EQ);
            }

            ExprKind::Cast { expr: inner, .. } => {
                // Cast is a no-op in our dynamic VM
                self.compile_expr(inner)?;
            }

            ExprKind::TypeOf(inner) => {
                self.compile_expr(inner)?;
                self.emit(Op::REF_TYPEOF);
            }

            // ── NullCoalesce ────────────────────────────────────────────
            ExprKind::NullCoalesce { left, right } => {
                self.compile_expr(left)?;
                self.emit(Op::DUP);
                self.emit(Op::REF_IS_NULL);
                let skip = self.emit_jump(Op::BR_IF_FALSE);
                self.emit(Op::DROP);
                self.compile_expr(right)?;
                self.patch_jump(skip);
            }

            // ── Spread ──────────────────────────────────────────────────
            ExprKind::Spread(inner) => {
                self.compile_expr(inner)?;
                self.emit(Op::SPREAD);
            }

            // ── Await ───────────────────────────────────────────────────
            ExprKind::Await(inner) => {
                // In our synchronous VM, promises are already resolved.
                // `await p` unwraps the promise's `__value` property.
                // If the inner value is not a promise, pass through.
                self.compile_expr(inner)?;
                // Save to local, try to read __value
                let await_slot = self.scope_mut().define("__await");
                self.emit_u16(Op::LOCAL_SET, await_slot); self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, await_slot);
                let vk = self.str_const("__value");
                self.emit_u16(Op::STRUCT_GET, vk);
                // If __value is null → not a promise, use original
                self.emit(Op::DUP);
                self.emit(Op::REF_IS_NULL);
                let use_original = self.emit_jump(Op::BR_IF_TRUE);
                // __value exists → use it (drop the null-check dup)
                let done = self.emit_jump(Op::BR);
                self.patch_jump(use_original);
                self.emit(Op::DROP); // drop null __value
                self.emit_u16(Op::LOCAL_GET, await_slot); // push original
                self.patch_jump(done);
            }

            // ── Yield ───────────────────────────────────────────────────
            ExprKind::Yield(val) => {
                if let Some(v) = val { self.compile_expr(v)?; } else { self.emit(Op::NULL); }
                self.emit_u16(Op::SUSPEND, 0);
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
                            self.emit_u16(Op::GLOBAL_GET, pidx);
                            for a in args { self.compile_expr(&a.value)?; }
                            self.emit_u8(Op::CALL_REF, args.len() as u8);
                            if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                                self.emit(Op::DUP);
                                self.emit_u16(Op::LOCAL_SET, slot);
                                self.emit(Op::DROP);
                            }
                        } else {
                            self.emit(Op::NULL);
                        }
                    } else {
                        self.emit(Op::NULL);
                    }
                } else if let Some(ref mname) = method {
                    // MyBase.Method(args) → this.__base_method(this, args)
                    let base_name = format!("__base_{}", self.canon(mname));
                    if let Some(self_slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                        let prop = self.str_const(&base_name);
                        self.emit_u16(Op::LOCAL_GET, self_slot);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        self.emit_u16(Op::LOCAL_GET, self_slot);
                        for a in args { self.compile_expr(&a.value)?; }
                        self.emit_u8(Op::CALL_REF, (args.len() + 1) as u8);
                    } else {
                        self.emit(Op::NULL);
                    }
                } else {
                    self.emit(Op::NULL);
                }
            }

            // ── Comprehension (Python) ──────────────────────────────────
            ExprKind::Comprehension { kind: _, element, generators } => {
                // Simplified: compile as loop building an array
                let line = self.line;
                common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                let result_slot = self.scope_mut().define("__comp_result");
                self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);

                // Only handle the first generator for simplicity
                if let Some(gen) = generators.first() {
                    self.compile_expr(&gen.iter)?;
                    let arr_slot = self.scope_mut().define("__comp_iter");
                    self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);
                    let idx_slot = self.scope_mut().define("__comp_idx");
                    let lp = common::loops::emit_for_in_start(
                        &mut self.chunks, self.current, arr_slot, idx_slot, line,
                    );
                    // Bind loop var
                    let var_name = match &gen.target.kind {
                        ExprKind::Ident(n) => n.clone(),
                        _ => "__comp_var".to_string(),
                    };
                    let var_slot = self.scope_mut().define(&var_name);
                    self.emit_u16(Op::LOCAL_SET, var_slot); self.emit(Op::DROP);

                    // Check conditions
                    let mut cond_skip = None;
                    for cond_expr in &gen.conditions {
                        self.compile_expr(cond_expr)?;
                        self.emit(Op::DYN_TO_BOOL);
                        cond_skip = Some(self.emit_jump(Op::BR_IF_FALSE));
                    }

                    // Push element via wasm:js-array.push.
                    self.emit_u16(Op::LOCAL_GET, result_slot);
                    self.compile_expr(element)?;
                    let l = self.line;
                    common::collections::emit_push(&mut self.chunks, self.current, l);
                    self.emit(Op::DROP);

                    if let Some(skip) = cond_skip { self.patch_jump(skip); }

                    common::loops::emit_for_in_end(
                        &mut self.chunks, self.current, idx_slot, lp, line,
                    );
                }

                self.emit_u16(Op::LOCAL_GET, result_slot);
            }

            // ── Slice (Python) ──────────────────────────────────────────
            ExprKind::Slice { lower, upper, step } => {
                // Emit slice parts for use by Index
                if let Some(l) = lower { self.compile_expr(l)?; } else { self.emit(Op::NULL); }
                if let Some(u) = upper { self.compile_expr(u)?; } else { self.emit(Op::NULL); }
                if let Some(s) = step { self.compile_expr(s)?; } else { self.emit(Op::NULL); }
                let idx = self.import("vybe:array", "sliceStep");
                self.emit_host_call(idx, 4); // obj already on stack from Index parent
            }

            // ── Walrus (Python :=) ──────────────────────────────────────
            ExprKind::Walrus { target, value } => {
                self.compile_expr(value)?;
                self.emit(Op::DUP);
                self.compile_assign_target(target)?;
            }

            // ── Void (JS) ───────────────────────────────────────────────
            ExprKind::Void(inner) => {
                self.compile_expr(inner)?;
                self.emit(Op::DROP);
                self.emit(Op::NULL); // void always evaluates to undefined
            }

            // ── Delete (JS expression) ──────────────────────────────────
            ExprKind::Delete(inner) => {
                // delete obj.prop → call vybe:object::deleteProperty(obj, key)
                // which removes the property and returns true.
                if let ExprKind::Member { object, field, .. } = &inner.kind {
                    self.compile_expr(object)?;
                    self.emit_const(Value::String(Arc::from(field.as_str())));
                    let idx = self.import("vybe:object", "deleteProperty");
                    self.emit_host_call(idx, 2);
                } else if let ExprKind::Index { object, index } = &inner.kind {
                    self.compile_expr(object)?;
                    self.compile_expr(index)?;
                    let idx = self.import("vybe:object", "deleteProperty");
                    self.emit_host_call(idx, 2);
                } else {
                    self.compile_expr(inner)?;
                    self.emit(Op::DROP);
                    self.emit(Op::TRUE);
                }
            }

            // ── Destructure (JS) ────────────────────────────────────────
            ExprKind::Destructure(_) => {
                // Destructure patterns are handled at assignment/declaration sites
                self.emit(Op::NULL);
            }

            // ── Sequence (JS comma operator) ────────────────────────────
            ExprKind::Sequence(exprs) => {
                for (i, e) in exprs.iter().enumerate() {
                    self.compile_expr(e)?;
                    if i < exprs.len() - 1 { self.emit(Op::DROP); }
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
                if let StmtKind::FunctionDecl { name, params, return_type, body, is_sub, is_generator, handles, .. } = &stmt.kind {
                    let fn_name = if name.is_empty() { "__anon_fn" } else { name };
                    self.compile_function_decl(fn_name, params, return_type, body, *is_sub, *is_generator, handles)?;
                    self.emit_var_get(fn_name);
                } else {
                    self.emit(Op::NULL);
                }
            }

            // ── Range ───────────────────────────────────────────────────
            ExprKind::Range { start, end, inclusive: _ } => {
                self.compile_expr(start)?;
                self.compile_expr(end)?;
                let line = self.line;
                common::collections::emit_range(&mut self.chunks, self.current, 2, line);
            }

            // ── StaticAccess (PHP) ──────────────────────────────────────
            ExprKind::StaticAccess { class, member } => {
                // class::member → look up class, then get static member
                self.compile_expr(class)?;
                if let ExprKind::Ident(name) = &member.kind {
                    let idx = self.str_const(name);
                    self.emit_u16(Op::STRUCT_GET, idx);
                } else {
                    self.compile_expr(member)?;
                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                }
            }

            // ── Match expression (PHP/Rust) ─────────────────────────────
            ExprKind::Match { subject, arms } => {
                self.compile_expr(subject)?;
                let subject_slot = self.scope_mut().define("__match_subj");
                self.emit_u16(Op::LOCAL_SET, subject_slot); self.emit(Op::DROP);
                let mut end_patches = Vec::new();
                for arm in arms {
                    if let Some(ref conditions) = arm.conditions {
                        let mut match_patches = Vec::new();
                        for c in conditions {
                            self.emit_u16(Op::LOCAL_GET, subject_slot);
                            self.compile_expr(c)?;
                            self.emit(Op::DYN_EQ);
                            match_patches.push(self.emit_jump(Op::BR_IF_TRUE));
                        }
                        let skip = self.emit_jump(Op::BR);
                        for p in match_patches { self.patch_jump(p); }
                        self.compile_expr(&arm.body)?;
                        end_patches.push(self.emit_jump(Op::BR));
                        self.patch_jump(skip);
                    } else {
                        // Default arm
                        self.compile_expr(&arm.body)?;
                        end_patches.push(self.emit_jump(Op::BR));
                    }
                }
                // If no arm matched, null
                self.emit(Op::NULL);
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
                    self.emit_u16(Op::GLOBAL_GET, pidx);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    // Store result as this
                    let self_kw = self.profile.self_keyword.clone();
                    if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                        self.emit(Op::DUP);
                        self.emit_u16(Op::LOCAL_SET, slot);
                        self.emit(Op::DROP);
                    }
                    return Ok(());
                }
            }
            // No parent — emit null
            self.emit(Op::NULL);
            return Ok(());
        }

        // ── super.method(args) → this.__base_method(args) ────────────
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if matches!(&object.kind, ExprKind::Super) {
                let base_name = format!("__base_{}", self.canon(field));
                let self_kw = self.profile.self_keyword.clone();
                if let Some(self_slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                    let prop = self.str_const(&base_name);
                    self.emit_u16(Op::LOCAL_GET, self_slot);
                    self.emit_u16(Op::STRUCT_GET, prop);
                    // Call with this as first arg
                    self.emit_u16(Op::LOCAL_GET, self_slot);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                    return Ok(());
                }
            }
        }

        // ── Debug intrinsic: __debug_dump(obj) ──────────────────────
        // Available in all languages. Prints object properties to stderr.
        if let ExprKind::Ident(name) = &callee.kind {
            if name == "__debug_dump" {
                for a in &arg_exprs { self.compile_expr(a)?; }
                let idx = self.import("vybe:debug", "dump");
                self.emit_host_call(idx, arg_exprs.len() as u8);
                return Ok(());
            }
        }

        // ── Builtin check: Ident("print") ───────────────────────────
        // Skip for user-defined functions: a VB `Function Echo(...)` must
        // dispatch to the user's chunk, not to the cross-language `echo →
        // wasi:cli.log` import shortcut.
        if let ExprKind::Ident(name) = &callee.kind {
            let shadows_builtin = self.defined_functions.contains(name)
                || (!self.case_sensitive
                    && self.defined_functions.iter().any(|g| g.eq_ignore_ascii_case(name)));
            if !shadows_builtin && self.try_compile_builtin(name, &arg_exprs)? {
                return Ok(());
            }
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
                    let scope = self.scope();
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
                            self.emit_u16(Op::GLOBAL_GET, root_idx);
                            for part in &ns_parts[1..] {
                                let idx = self.str_const(part);
                                self.emit_u16(Op::STRUCT_GET, idx);
                            }
                            let is_const = common::dotnet::is_known_constant(ns_parts.last().unwrap_or(&String::new()));
                            if !is_const {
                                for a in &arg_exprs { self.compile_expr(a)?; }
                                self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
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
                                        self.emit(Op::NULL);
                                        return Ok(());
                                    }
                                    "join" => {
                                        // th.Join() → thread_join opcode (blocks until thread
                                        // completes, pushes exit code). Leave the exit code on
                                        // stack — the statement wrapper at StmtKind::Expr adds
                                        // its own DROP.
                                        self.emit_var_get(&local);
                                        let line = self.line;
                                        common::threading::emit_thread_join(self.chunk(), line);
                                        return Ok(());
                                    }
                                    "waitforexit" => {
                                        // p.WaitForExit() — process ran synchronously, no-op
                                        // Must leave a value on stack (caller drops it)
                                        self.emit(Op::NULL);
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
                                    self.emit_u16(Op::STRUCT_GET, idx);
                                } else {
                                    // Last member is the method — struct_get then call with this
                                    self.emit(Op::DUP); // keep obj for this
                                    self.emit_u16(Op::STRUCT_GET, idx);
                                    // Stack: [obj, method_fn] — swap so fn is first
                                    let fn_tmp = self.scope().resolve("__dotnet_fn")
                                        .unwrap_or_else(|| self.scope_mut().define("__dotnet_fn"));
                                    self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                                    let obj_tmp = self.scope().resolve("__dotnet_obj")
                                        .unwrap_or_else(|| self.scope_mut().define("__dotnet_obj"));
                                    self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);
                                    self.emit_u16(Op::LOCAL_GET, fn_tmp);
                                    self.emit_u16(Op::LOCAL_GET, obj_tmp);
                                    for a in &arg_exprs { self.compile_expr(a)?; }
                                    self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                                    return Ok(());
                                }
                            }
                            // Shouldn't reach here for calls, but just in case
                            for a in &arg_exprs { self.compile_expr(a)?; }
                            self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                            return Ok(());
                        }
                        common::dotnet::DottedResolution::NoOp => {
                            self.emit(Op::NULL);
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
                    self.emit_u16(Op::GLOBAL_GET, root_idx);
                    for part in &lower_parts[1..] {
                        let idx = self.str_const(part);
                        self.emit_u16(Op::STRUCT_GET, idx);
                    }
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
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
                    self.emit_u16(Op::GLOBAL_GET, cls_idx);
                    self.emit(Op::DUP);
                    let m = self.canon(field);
                    let method_idx = self.str_const(&m);
                    self.emit_u16(Op::STRUCT_GET, method_idx);
                    // Stack: [class, fn] — swap so we have [fn, class, ...args]
                    let fn_tmp = self.scope().resolve("__static_fn")
                        .unwrap_or_else(|| self.scope_mut().define("__static_fn"));
                    self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                    let cls_tmp = self.scope().resolve("__static_cls")
                        .unwrap_or_else(|| self.scope_mut().define("__static_cls"));
                    self.emit_u16(Op::LOCAL_SET, cls_tmp); self.emit(Op::DROP);
                    self.emit_u16(Op::LOCAL_GET, fn_tmp);
                    self.emit_u16(Op::LOCAL_GET, cls_tmp);
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                    return Ok(());
                }
            }
        }

        // ── Function.prototype.call / .apply ────────────────────────
        // `fn.call(thisArg, a, b, ...)` → call `fn` with `[a, b, ...]`
        // `fn.apply(thisArg, [a, b, ...])` → same; the array form is
        // unwrapped at runtime via the spread opcode.
        //
        // We can't route this through value_methods because the standard
        // dispatch path pushes the receiver + ALL args, but here we need
        // to drop arg[0] (`thisArg`) from the middle of the stack. Skip
        // when the field is defined on a user class so user methods
        // named `call`/`apply` keep working.
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            let canon_field = self.canon(field);
            if !self.defined_class_methods.contains(&canon_field)
                && (field == "call" || field == "apply")
            {
                self.compile_expr(object)?;                       // [fn]
                if field == "call" {
                    // Skip thisArg, compile rest as positional args.
                    for a in arg_exprs.iter().skip(1) {
                        self.compile_expr(a)?;
                    }
                    let n = arg_exprs.len().saturating_sub(1);
                    self.emit_u8(Op::CALL_REF, n as u8);
                } else {
                    // apply(thisArg, argsArray) — spread the array.
                    if let Some(args_expr) = arg_exprs.get(1) {
                        self.compile_expr(args_expr)?;
                        self.emit(Op::SPREAD);
                    }
                    // Use call_ref with 0 — the spread opcode pushes
                    // each array element and bumps the call arity at
                    // runtime via Op::call_spread if available, else
                    // we fall back here. The current VM uses Op::SPREAD
                    // before call_ref to flatten the top array.
                    self.emit_u8(Op::CALL_REF, 0);
                }
                return Ok(());
            }
        }

        // ── Value method: obj.toUpperCase() ─────────────────────────
        //
        // Method name shadowing rule: a value method (e.g. `Array.push`,
        // `String.toUpperCase`) is the default for *member-access*
        // receivers like `this.items.push(x)` — the receiver is
        // structurally a property, almost certainly a built-in collection.
        //
        // For *direct* receivers (`this`, `super`, or a local variable
        // by name), if the field is a known user-class method, prefer
        // the user method via the generic call path. That preserves
        // user overrides like `class Stack { push(x) { ... } }` and
        // `class Holder { size() { ... } }` against built-in
        // `Array.push`/`map_size` shadowing.
        //
        // This is a heuristic — the cleaner fix is per-class method sets
        // plus receiver-type inference, tracked in the user's pending
        // "JS/C# compilers don't use common::classes" migration.
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            let canon_field = self.canon(field);
            let receiver_is_direct = matches!(
                object.kind,
                ExprKind::This | ExprKind::Super | ExprKind::Ident(_)
            );
            let user_method_shadow = receiver_is_direct
                && self.defined_class_methods.contains(&canon_field);
            // Also skip value_methods if the field is an array HOF method —
            // the array_methods dispatch handles it with proper HOF semantics.
            // Without this, `[1,2,3].includes(2)` routes through the string
            // `includes` value method instead of the array contains HOF.
            let field_lower_check = if self.case_sensitive { field.clone() } else { field.to_lowercase() };
            let is_array_method = self.profile.lookup_array_method(&field_lower_check).is_some();
            if user_method_shadow || is_array_method {
                // Fall through — let the HOF dispatch or generic call path handle it
            } else if let Some(def) = self.profile.lookup_value_method(field, arg_exprs.len() as u8).cloned() {
                // For Stdlib calls, push func ref BEFORE args (call_ref expects [func, args...])
                if let BuiltinEmit::Stdlib(stdlib_name) = &def.emit {
                    let global_name = format!("__vybe_{}", stdlib_name);
                    let name_idx = self.str_const(&global_name);
                    self.emit_u16(Op::GLOBAL_GET, name_idx);
                    self.compile_expr(object)?;
                    for a in &arg_exprs { self.compile_expr(a)?; }
                    self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                    return Ok(());
                }
                // Object is first arg, then explicit args
                self.compile_expr(object)?;
                for a in &arg_exprs { self.compile_expr(a)?; }
                // Some opcodes need default args when called with fewer
                // than required. Push defaults here.
                if let BuiltinEmit::Opcode(ref op) | BuiltinEmit::Common(ref op) = &def.emit {
                    match op.as_str() {
                        // array_join / collections.join needs [arr, sep]
                        "array_join" | "collections.join" if arg_exprs.is_empty() => {
                            self.emit_const(Value::String(Arc::from(",")));
                        }
                        // array_fill needs [arr, val, start, end]
                        "array_fill" if arg_exprs.len() < 2 => {
                            // Push start=0 and end=arr.length defaults
                            if arg_exprs.is_empty() {
                                self.emit(Op::NULL); // val
                            }
                            self.emit(Op::I32_CONST_0); // start
                            self.emit_const(Value::I32(i32::MAX)); // end (clamped by VM)
                        }
                        _ => {}
                    }
                }
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
                    BuiltinEmit::Invoke(method_name) => {
                        let line = self.line;
                        let name = method_name.clone();
                        common::invoke::emit_invoke_method(
                            &mut self.chunks,
                            self.current,
                            &name,
                            arg_exprs.len() as u8,
                            line,
                        );
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
                    "__array_sort_by_key" => "sort_by_key".to_string(),
                    "__array_some" => "some".to_string(),
                    "__array_every" => "every".to_string(),
                    "__array_flat_map" => "flatMap".to_string(),
                    "__array_reduce_right" => "reduceRight".to_string(),
                    _ => field_lower,
                };
                // Compile arr and fn(s) into local slots
                self.compile_expr(object)?;
                let arr_slot = self.scope_mut().define("__hof_arr");
                self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);

                if let Some(fn_expr) = arg_exprs.first() {
                    self.compile_expr(fn_expr)?;
                } else {
                    self.emit(Op::NULL);
                }
                let fn_slot = self.scope_mut().define("__hof_fn");
                self.emit_u16(Op::LOCAL_SET, fn_slot); self.emit(Op::DROP);

                let idx_slot = self.scope_mut().define("__hof_idx");
                let result_slot = self.scope_mut().define("__hof_result");
                let line = self.line;

                match field_lower.as_str() {
                    "map" => {
                        // emit_map leaves result on stack
                        common::loops::emit_map(&mut self.chunks, self.current, fn_slot, arr_slot, result_slot, idx_slot, line);
                    }
                    "filter" => {
                        let elem_slot = self.scope_mut().define("__hof_elem");
                        common::loops::emit_filter(&mut self.chunks, self.current, fn_slot, arr_slot, result_slot, idx_slot, elem_slot, line);
                    }
                    "reduce" => {
                        // reduce(fn, initial?) — initial is second arg.
                        // When initial IS provided, start from i=0 with
                        // acc=initial. emit_reduce always starts from
                        // i=1 with acc=arr[0], so we only use it for
                        // the no-initial case.
                        if let Some(init_expr) = arg_exprs.get(1) {
                            // acc = initial, i = 0
                            self.compile_expr(init_expr)?;
                            self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                            // Inline reduce loop starting from i=0
                            self.emit(Op::I32_CONST_0);
                            self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                            let loop_start = self.chunks[self.current].current_offset();
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                            self.emit(Op::DYN_LT);
                            let exit_jump = self.emit_jump(Op::BR_IF_FALSE);
                            // acc = fn(acc, arr[i])
                            self.emit_u16(Op::LOCAL_GET, fn_slot);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                            self.emit_u8(Op::CALL_REF, 2);
                            self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                            // i++
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            self.emit_const(Value::I32(1));
                            self.emit(Op::DYN_ADD);
                            self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                            self.emit_loop(loop_start);
                            self.patch_jump(exit_jump);
                            self.emit_u16(Op::LOCAL_GET, result_slot);
                        } else {
                            // No initial: emit_reduce starts from arr[0], i=1
                            common::loops::emit_reduce(&mut self.chunks, self.current, fn_slot, arr_slot, result_slot, idx_slot, line);
                        }
                    }
                    "forEach" | "foreach" => {
                        common::loops::emit_foreach(&mut self.chunks, self.current, fn_slot, arr_slot, idx_slot, line);
                    }
                    "some" => {
                        common::loops::emit_any_every(&mut self.chunks, self.current, fn_slot, arr_slot, idx_slot, true, line);
                    }
                    "every" => {
                        common::loops::emit_any_every(&mut self.chunks, self.current, fn_slot, arr_slot, idx_slot, false, line);
                    }
                    "find" => {
                        // find uses includes pattern but returns element not bool
                        self.emit(Op::NULL);
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        let lp = common::loops::emit_for_in_start(
                            &mut self.chunks, self.current, arr_slot, idx_slot, line);
                        let elem_slot = self.scope_mut().define("__find_elem");
                        self.emit_u16(Op::LOCAL_SET, elem_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.emit(Op::DYN_TO_BOOL);
                        let skip = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        let brk = self.emit_jump(Op::BR);
                        self.patch_jump(skip);
                        common::loops::emit_for_in_end(&mut self.chunks, self.current, idx_slot, lp, line);
                        self.patch_jump(brk);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    "findIndex" | "findindex" => {
                        // findIndex: like find but returns the index, not the element
                        self.emit_const(Value::I32(-1));
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        let lp = common::loops::emit_for_in_start(
                            &mut self.chunks, self.current, arr_slot, idx_slot, line);
                        let elem_slot = self.scope_mut().define("__findi_elem");
                        self.emit_u16(Op::LOCAL_SET, elem_slot); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, elem_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_u8(Op::CALL_REF, 2);
                        self.emit(Op::DYN_TO_BOOL);
                        let skip = self.emit_jump(Op::BR_IF_FALSE);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        let brk = self.emit_jump(Op::BR);
                        self.patch_jump(skip);
                        common::loops::emit_for_in_end(&mut self.chunks, self.current, idx_slot, lp, line);
                        self.patch_jump(brk);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    "includes" => {
                        // `x.includes(v)` — polymorphic: arrays do element
                        // membership, strings do substring search, user
                        // objects fall through to their own method. Route
                        // through `wasm:js-value.invokeMethod` so the
                        // emitted wasm stays spec-compliant on v8 where
                        // String.prototype.includes and Array.prototype.includes
                        // are distinct methods on distinct prototypes.
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        common::invoke::emit_invoke_method(
                            &mut self.chunks,
                            self.current,
                            "includes",
                            1,
                            line,
                        );
                    }
                    "sort" => {
                        // JS sort(comparatorFn?) — 2-arg comparator or default
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit(Op::REF_IS_NULL);
                        let no_fn = self.emit_jump(Op::BR_IF_TRUE);
                        let global = self.str_const("__vybe_sort_with_comparator");
                        self.emit_u16(Op::GLOBAL_GET, global);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u8(Op::CALL_REF, 2);
                        let done = self.emit_jump(Op::BR);
                        self.patch_jump(no_fn);
                        let sort_global = self.str_const("__vybe_sort_in_place");
                        self.emit_u16(Op::GLOBAL_GET, sort_global);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.patch_jump(done);
                    }
                    "sort_by_key" => {
                        // .NET OrderBy(keySelector) — 1-arg key extractor
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit(Op::REF_IS_NULL);
                        let no_fn = self.emit_jump(Op::BR_IF_TRUE);
                        let global = self.str_const("__vybe_sort_by_key");
                        self.emit_u16(Op::GLOBAL_GET, global);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u8(Op::CALL_REF, 2);
                        let done = self.emit_jump(Op::BR);
                        self.patch_jump(no_fn);
                        let sort_global = self.str_const("__vybe_sort_in_place");
                        self.emit_u16(Op::GLOBAL_GET, sort_global);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u8(Op::CALL_REF, 1);
                        self.patch_jump(done);
                    }
                    "indexOf" | "indexof" => {
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot); // search value
                        common::collections::emit_index_of(&mut self.chunks, self.current, line);
                    }
                    "flatMap" | "flatmap" => {
                        // arr.flatMap(fn) = arr.map(fn).flat()
                        // First emit map: result[i] = fn(arr[i])
                        let mapped_slot = self.scope_mut().define("__flatmap_mapped");
                        common::loops::emit_map(&mut self.chunks, self.current, fn_slot, arr_slot, mapped_slot, idx_slot, line);
                        // Now the mapped array is on stack. Flatten it one level.
                        let flat_idx = self.import("vybe:array", "flat");
                        self.emit_const(Value::I32(1));  // depth = 1
                        self.emit_host_call(flat_idx, 2);
                    }
                    "reduceRight" | "reduceright" => {
                        // reduceRight(fn, initial?) — iterate from end to start.
                        if let Some(init_expr) = arg_exprs.get(1) {
                            self.compile_expr(init_expr)?;
                            self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        } else {
                            // acc = arr[len-1]
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                            self.emit_const(Value::I32(1));
                            self.emit(Op::F64_SUB);
                            self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                            self.emit_u16(Op::LOCAL_GET, arr_slot);
                            self.emit_u16(Op::LOCAL_GET, idx_slot);
                            { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                            self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        }
                        // Start from len-1 (or len-2 if no initial)
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        { let l = self.line; common::collections::emit_len(&mut self.chunks, self.current, l); }
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        if arg_exprs.get(1).is_none() {
                            self.emit_const(Value::I32(1));
                            self.emit(Op::F64_SUB);
                        }
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        let loop_start = self.chunks[self.current].current_offset();
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(0));
                        self.emit(Op::DYN_GE);
                        let exit_jump = self.emit_jump(Op::BR_IF_FALSE);
                        // acc = fn(acc, arr[i])
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
                        self.emit_u8(Op::CALL_REF, 2);
                        self.emit_u16(Op::LOCAL_SET, result_slot); self.emit(Op::DROP);
                        // i--
                        self.emit_u16(Op::LOCAL_GET, idx_slot);
                        self.emit_const(Value::I32(1));
                        self.emit(Op::F64_SUB);
                        self.emit_u16(Op::LOCAL_SET, idx_slot); self.emit(Op::DROP);
                        self.emit_loop(loop_start);
                        self.patch_jump(exit_jump);
                        self.emit_u16(Op::LOCAL_GET, result_slot);
                    }
                    _ => {
                        // Fallback: call as regular method
                        self.emit_u16(Op::LOCAL_GET, arr_slot);
                        self.emit_u16(Op::LOCAL_GET, fn_slot);
                        self.emit_u8(Op::CALL_REF, 2);
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
                    self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
                    return Ok(());
                }
            }
        }

        // ── Method call: obj.method(args) ───────────────────────────
        if let ExprKind::Member { object, field, null_safe } = &callee.kind {
            self.compile_expr(object)?;

            if *null_safe {
                // obj?.method() — short-circuit to null if obj is null/undefined.
                // Stack: [obj]. Check null, if null leave null on stack and skip call.
                self.emit(Op::DUP);
                self.emit(Op::REF_IS_NULL);
                let obj_not_null = self.emit_jump(Op::BR_IF_FALSE);
                // obj IS null — leave null on stack, skip call
                let end = self.emit_jump(Op::BR);
                self.patch_jump(obj_not_null);
                // obj is not null — do the method call
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
                for a in &arg_exprs { self.compile_expr(a)?; }
                self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                self.patch_jump(end);
                return Ok(());
            }

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
            for a in &arg_exprs { self.compile_expr(a)?; }
            self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
            return Ok(());
        }

        // ── Simple call: name(args) / expr(args) ────────────────────
        if let ExprKind::Ident(name) = &callee.kind {
            let is_known_func = self.defined_functions.contains(name)
                || (!self.case_sensitive && self.defined_functions.iter().any(|g| g.eq_ignore_ascii_case(name)));
            if !is_known_func && self.try_compile_builtin(name, &arg_exprs)? {
                return Ok(());
            }

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
                    { let l = self.line; common::collections::emit_get(&mut self.chunks, self.current, l); }
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
                        self.emit_u16(Op::LOCAL_GET, self_slot);
                        let field_name = self.canon(name);
                        let prop = self.str_const(&field_name);
                        self.emit(Op::DUP);
                        self.emit_u16(Op::STRUCT_GET, prop);
                        let fn_tmp = self.scope_mut().define("__bare_fn");
                        self.emit_u16(Op::LOCAL_SET, fn_tmp); self.emit(Op::DROP);
                        let obj_tmp = self.scope_mut().define("__bare_obj");
                        self.emit_u16(Op::LOCAL_SET, obj_tmp); self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, fn_tmp);
                        self.emit_u16(Op::LOCAL_GET, obj_tmp);
                        for a in &arg_exprs { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL_REF, (arg_exprs.len() + 1) as u8);
                        return Ok(());
                    }
                }
            }

            let has_spread = args.iter().any(|a| a.spread);
            if has_spread {
                // Spread args: build a flat args array, then spread onto
                // stack and call. Stash the accumulator in a local so
                // `wasm:js-array.push` (returns new length per
                // ECMA-262) and `wasm:js-array.concat` (returns new
                // array) can both drive the same pattern.
                let line = self.line;
                let args_slot = self.scope_mut().define("__spread_args");
                common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
                self.emit_u16(Op::LOCAL_SET, args_slot); self.emit(Op::DROP);
                let mut known_len: Option<usize> = Some(0);
                for a in args {
                    if a.spread {
                        // new_arr = concat(args, spread)
                        self.emit_u16(Op::LOCAL_GET, args_slot);
                        self.compile_expr(&a.value)?;
                        common::collections::emit_concat(&mut self.chunks, self.current, line);
                        self.emit_u16(Op::LOCAL_SET, args_slot); self.emit(Op::DROP);
                        if let ExprKind::Array(elems) = &a.value.kind {
                            if let Some(ref mut k) = known_len { *k += elems.len(); }
                        } else {
                            known_len = None;
                        }
                    } else {
                        self.emit_u16(Op::LOCAL_GET, args_slot);
                        self.compile_expr(&a.value)?;
                        common::collections::emit_push(&mut self.chunks, self.current, line);
                        self.emit(Op::DROP); // drop new_length returned by push
                        if let Some(ref mut k) = known_len { *k += 1; }
                    }
                }
                self.emit_var_get(name);
                self.emit_u16(Op::LOCAL_GET, args_slot);
                self.emit(Op::SPREAD);
                let arity = known_len.unwrap_or(16) as u8;
                self.emit_u8(Op::CALL_REF, arity);
                return Ok(());
            }
            self.emit_var_get(name);
            for a in &arg_exprs { self.compile_expr(a)?; }
            self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
            return Ok(());
        }

        // ── Fallback: general expression call ───────────────────────
        self.compile_expr(callee)?;
        for a in &arg_exprs { self.compile_expr(a)?; }
        self.emit_u8(Op::CALL_REF, arg_exprs.len() as u8);
        Ok(())
    }

    // ════════════════════════════════════════════════════════════════════════
    // Lambda compilation
    // ════════════════════════════════════════════════════════════════════════

    fn compile_lambda(&mut self, params: &[Param], body: &LambdaBody) -> Result<(), String> {
        let has_rest = params.last().map_or(false, |p| p.is_rest);
        let arity = if has_rest { 255u8 } else { params.len() as u8 };
        let ci = self.chunks.len();
        let chunk = common::functions::create_function_chunk("<lambda>", arity);
        self.chunks.push(chunk);
        self.scopes.push(Scope::new_function());
        let saved = self.current;
        self.current = ci;
        for p in params {
            self.scope_mut().define(&p.name);
            if let Some(ref default) = p.default {
                let slot = self.scope().resolve(&p.name).unwrap();
                self.emit_u16(Op::LOCAL_GET, slot);
                self.emit(Op::REF_IS_NULL);
                let has_val = self.emit_jump(Op::BR_IF_FALSE);
                self.compile_expr(default)?;
                self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP);
                self.patch_jump(has_val);
            }
        }

        // Rest param preamble (same as compile_function_decl).
        // Accumulator pattern: stash arr in rest_slot and reload each
        // iteration so `wasm:js-array.push` (returns new length per
        // ECMA-262) cleanly drives the push loop.
        if has_rest {
            let rest_name = &params.last().unwrap().name;
            let rest_slot = self.scope().resolve(rest_name).unwrap();
            let line = self.line;
            let rest_arr = self.scope_mut().define("__rest_arr");
            common::collections::emit_array_new(&mut self.chunks, self.current, 0, line);
            self.emit_u16(Op::LOCAL_SET, rest_arr); self.emit(Op::DROP);
            let max_rest = 16u16;
            let mut done_patches: Vec<usize> = Vec::new();
            for i in 0..max_rest {
                let slot = rest_slot + i;
                self.emit_u16(Op::LOCAL_GET, slot);
                self.emit(Op::REF_IS_NULL);
                done_patches.push(self.emit_jump(Op::BR_IF_TRUE));
                self.emit_u16(Op::LOCAL_GET, rest_arr);
                self.emit_u16(Op::LOCAL_GET, slot);
                common::collections::emit_push(&mut self.chunks, self.current, line);
                self.emit(Op::DROP); // drop new_length
            }
            for p in done_patches { self.patch_jump(p); }
            self.emit_u16(Op::LOCAL_GET, rest_arr);
            self.emit_u16(Op::LOCAL_SET, rest_slot);
            self.emit(Op::DROP);
        }

        // Result slot for ResultSlot languages
        let result_slot = if self.profile.function_return == ReturnStyle::ResultSlot {
            let rs = self.scope_mut().define("Result");
            self.emit(Op::NULL); self.emit_u16(Op::LOCAL_SET, rs); self.emit(Op::DROP);
            let saved_fn = self.current_func_name.take();
            let saved_rs = self.current_result_slot.take();
            self.current_func_name = Some("<lambda>".into());
            self.current_result_slot = Some(rs);
            Some((rs, saved_fn, saved_rs))
        } else { None };

        match body {
            LambdaBody::Expr(expr) => {
                self.compile_expr(expr)?;
                self.emit(Op::RETURN);
            }
            LambdaBody::Block(stmts) => {
                for s in stmts { self.compile_stmt(s)?; }
            }
        }

        if let Some((rs, saved_fn, saved_rs)) = result_slot {
            self.emit_u16(Op::LOCAL_GET, rs);
            self.emit(Op::RETURN);
            self.current_func_name = saved_fn;
            self.current_result_slot = saved_rs;
        } else if matches!(body, LambdaBody::Block(_)) {
            let line = self.line;
            common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
        }

        let locals = self.scope().next_slot.max(self.chunks[ci].local_count);
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
