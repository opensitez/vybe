use std::rc::Rc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use vybe_parser_generic::*;
use vybe_parser_generic::profile::*;
use vybe_compiler_common as common;
use crate::scope::Scope;

struct LoopCtx {
    break_patches: Vec<usize>,
    continue_target: usize,
}

pub struct Compiler {
    chunks: Vec<Chunk>,
    scopes: Vec<Scope>,
    current: usize,
    loops: Vec<LoopCtx>,
    line: u32,
    defined_globals: std::collections::HashSet<String>,
    defined_functions: std::collections::HashSet<String>,
    case_sensitive: bool,
    profile: LanguageProfile,
    current_func_name: Option<String>,
    current_result_slot: Option<u16>,
    pending_classes: std::collections::HashMap<String, PendingClass>,
    current_class: Option<String>,
}

struct PendingClass {
    parent: Option<String>,
    fields: Vec<String>,
    methods: Vec<(String, usize, bool)>,  // (name, chunk_idx, is_constructor)
    ctor_arity: u8,
}

impl Compiler {
    pub fn new() -> Self {
        Self::with_profile(LanguageProfile::pascal())
    }

    pub fn with_profile(profile: LanguageProfile) -> Self {
        Self {
            chunks: vec![Chunk::new("<script>")],
            scopes: vec![Scope::new()],
            current: 0,
            loops: Vec::new(),
            line: 1,
            defined_globals: std::collections::HashSet::new(),
            defined_functions: std::collections::HashSet::new(),
            case_sensitive: true,
            profile,
            current_func_name: None,
            current_result_slot: None,
            pending_classes: std::collections::HashMap::new(),
            current_class: None,
        }
    }

    pub fn compile(mut self, module: &Module) -> Result<Vec<Chunk>, String> {
        self.case_sensitive = module.language != Lang::Pascal && module.language != Lang::VB && module.language != Lang::Cobol;
        for stmt in &module.body {
            self.compile_stmt(stmt)?;
        }
        self.emit(Op::null);
        self.emit(Op::halt);
        let locals = self.scope().next_slot;
        self.chunks[0].local_count = locals;
        common::bundle::finalize_with_stdlib(&mut self.chunks);
        Ok(self.chunks)
    }

    // ── Helpers ───────────────────────────────────────────────────────────

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
    fn str_const(&mut self, s: &str) -> u16 { self.chunks[self.current].add_constant(Value::String(Rc::from(s))) }

    fn import(&mut self, module: &str, name: &str) -> u16 { self.chunks[0].add_import(module, name) }
    fn emit_host_call(&mut self, idx: u16, argc: u8) {
        let l = self.line;
        self.chunks[self.current].emit_op_u16(Op::call_import, idx, l);
        self.chunks[self.current].emit(argc, l);
    }

    fn resolve_name(&self, name: &str) -> bool {
        if let Some(_) = self.scope().resolve(name) { return true; }
        if !self.case_sensitive {
            if let Some(_) = self.scope().resolve_ci(name) { return true; }
        }
        self.defined_globals.contains(name)
    }

    fn emit_var_get(&mut self, name: &str) {
        if let Some(slot) = self.scope().resolve(name) {
            self.emit_u16(Op::local_get, slot);
        } else if !self.case_sensitive {
            if let Some(slot) = self.scope().resolve_ci(name) {
                self.emit_u16(Op::local_get, slot);
                return;
            }
            let idx = self.str_const(name);
            self.emit_u16(Op::global_get, idx);
        } else {
            let idx = self.str_const(name);
            self.emit_u16(Op::global_get, idx);
        }
    }

    fn emit_var_set(&mut self, name: &str) {
        if let Some(slot) = self.scope().resolve(name) {
            self.emit_u16(Op::local_set, slot); self.emit(Op::drop);
        } else if !self.case_sensitive {
            if let Some(slot) = self.scope().resolve_ci(name) {
                self.emit_u16(Op::local_set, slot); self.emit(Op::drop);
                return;
            }
            let idx = self.str_const(name); self.emit_u16(Op::global_set, idx); self.emit(Op::drop);
        } else {
            let idx = self.str_const(name); self.emit_u16(Op::global_set, idx); self.emit(Op::drop);
        }
    }

    // ── Statement compilation ────────────────────────────────────────────

    fn compile_stmt(&mut self, stmt: &Statement) -> Result<(), String> {
        self.line = stmt.span.start_line;
        match &stmt.kind {
            StmtKind::Expr(expr) => {
                // For Pascal: bare identifier or field access as statement = procedure call
                match &expr.kind {
                    ExprKind::Ident(name) if self.defined_globals.contains(name.as_str()) => {
                        self.emit_var_get(name);
                        self.emit_u8(Op::call_ref, 0);
                        self.emit(Op::drop);
                    }
                    ExprKind::Member { object, field, .. } => {
                        // obj.Method → method call with 0 args
                        self.compile_expr(object)?;
                        let field_name = if self.case_sensitive { field.clone() } else { field.to_lowercase() };
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
            StmtKind::Block(stmts) => {
                // Don't create a scope for blocks that are purely declarations (const/var sections)
                let all_decls = stmts.iter().all(|s| matches!(s.kind, StmtKind::VarDecl { .. } | StmtKind::FunctionDecl { .. } | StmtKind::ClassDecl { .. } | StmtKind::EnumDecl { .. }));
                if !all_decls { self.scope_mut().begin_scope(); }
                for s in stmts { self.compile_stmt(s)?; }
                if !all_decls { self.scope_mut().end_scope(); }
            }
            StmtKind::VarDecl { name, type_hint, init, is_const, .. } => {
                if let Some(init_expr) = init {
                    self.compile_expr(init_expr)?;
                } else {
                    match type_hint.as_deref().map(|s| s.to_lowercase()).as_deref() {
                        Some("integer") | Some("int") | Some("longint") | Some("real") | Some("double") | Some("float") => self.emit(Op::f64_const_0),
                        Some("boolean") | Some("bool") => self.emit(Op::r#false),
                        Some("string") => self.emit_const(Value::String(Rc::from(""))),
                        _ => self.emit(Op::null),
                    }
                }
                // Top-level vars/consts → globals (even if inside a block from const section)
                let is_toplevel = self.scopes.len() == 1;
                if is_toplevel {
                    let idx = self.str_const(name);
                    self.emit_u16(Op::global_set, idx);
                    self.emit(Op::drop);
                    self.defined_globals.insert(name.clone());
                } else {
                    let slot = self.scope_mut().define(name);
                    self.emit_u16(Op::local_set, slot);
                    self.emit(Op::drop);
                }
            }
            StmtKind::Assign { target, value } => {
                self.compile_expr(value)?;
                self.compile_assign_target(target)?;
            }
            StmtKind::CompoundAssign { target, op, value } => {
                self.compile_expr(target)?;
                self.compile_expr(value)?;
                self.compile_binop(op);
                self.compile_assign_target(target)?;
            }
            StmtKind::If { cond, then, elifs, else_ } => {
                self.compile_expr(cond)?;
                self.emit(Op::dyn_to_bool);
                let else_j = self.emit_jump(Op::br_if_false);
                for s in then { self.compile_stmt(s)?; }
                let mut end_jumps = vec![];
                if !elifs.is_empty() || else_.is_some() {
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
                if let Some(else_body) = else_ {
                    for s in else_body { self.compile_stmt(s)?; }
                }
                for j in end_jumps { self.patch_jump(j); }
            }
            StmtKind::While { cond, body } => {
                let start = self.current_offset();
                self.loops.push(LoopCtx { break_patches: vec![], continue_target: start });
                self.compile_expr(cond)?;
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                for s in body { self.compile_stmt(s)?; }
                self.emit_loop(start);
                self.patch_jump(exit);
                let ctx = self.loops.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
            }
            StmtKind::For { init, cond, update, body } => {
                if let Some(init_stmt) = init { self.compile_stmt(init_stmt)?; }
                let start = self.current_offset();
                self.loops.push(LoopCtx { break_patches: vec![], continue_target: start });
                if let Some(c) = cond {
                    self.compile_expr(c)?;
                    self.emit(Op::dyn_to_bool);
                } else {
                    self.emit(Op::r#true);
                }
                let exit = self.emit_jump(Op::br_if_false);
                for s in body { self.compile_stmt(s)?; }
                if let Some(u) = update { self.compile_expr(u)?; self.emit(Op::drop); }
                self.emit_loop(start);
                self.patch_jump(exit);
                let ctx = self.loops.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
            }
            StmtKind::ForIn { var, iter, body } => {
                self.compile_expr(iter)?;
                let arr_slot = self.scope_mut().define("__forin_arr");
                self.emit_u16(Op::local_set, arr_slot); self.emit(Op::drop);
                let idx_slot = self.scope_mut().define("__forin_idx");
                let line = self.line;
                let (loop_start, exit_jump) = common::loops::emit_for_in_start(
                    &mut self.chunks[self.current], arr_slot, idx_slot, line,
                );
                self.emit_var_set(var);
                self.loops.push(LoopCtx { break_patches: vec![], continue_target: loop_start });
                for s in body { self.compile_stmt(s)?; }
                let ctx = self.loops.pop().unwrap();
                common::loops::emit_for_in_end(
                    &mut self.chunks[self.current], idx_slot, loop_start, exit_jump, line,
                );
                for p in ctx.break_patches { self.patch_jump(p); }
            }
            StmtKind::DoWhile { body, cond, until } => {
                let start = self.current_offset();
                self.loops.push(LoopCtx { break_patches: vec![], continue_target: start });
                for s in body { self.compile_stmt(s)?; }
                self.compile_expr(cond)?;
                self.emit(Op::dyn_to_bool);
                if *until {
                    let exit = self.emit_jump(Op::br_if_true);
                    self.emit_loop(start);
                    self.patch_jump(exit);
                } else {
                    let cont = self.emit_jump(Op::br_if_true);
                    self.patch_jump(cont); // TODO: this is wrong, need to loop back
                    // Actually: if cond is true, loop; if false, exit
                    // br_if_true jumps to... we need to loop
                }
                let ctx = self.loops.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
            }
            StmtKind::Switch { expr, cases, default } => {
                self.compile_expr(expr)?;
                let mut end_patches = Vec::new();
                for case in cases {
                    let mut match_patches = Vec::new();
                    for val in &case.values {
                        self.emit(Op::dup);
                        self.compile_expr(val)?;
                        self.emit(Op::dyn_eq);
                        match_patches.push(self.emit_jump(Op::br_if_true));
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
                self.emit(Op::drop);
            }
            StmtKind::Try { body, catches, else_: _, finally } => {
                let line = self.line;
                let catch_jump = common::errors::emit_try_start(&mut self.chunks[self.current], line);
                for s in body { self.compile_stmt(s)?; }
                common::errors::emit_try_end(&mut self.chunks[self.current], line);
                let skip = self.emit_jump(Op::br);
                common::errors::patch_catch(&mut self.chunks[self.current], catch_jump);
                if catches.is_empty() {
                    self.emit(Op::drop);
                } else {
                    for c in catches {
                        if let Some(ref var) = c.var_name {
                            let slot = self.scope_mut().define(var);
                            self.emit_u16(Op::local_set, slot); self.emit(Op::drop);
                        } else {
                            self.emit(Op::drop);
                        }
                        for s in &c.body { self.compile_stmt(s)?; }
                    }
                }
                self.patch_jump(skip);
                if let Some(fin) = finally {
                    for s in fin { self.compile_stmt(s)?; }
                }
            }
            StmtKind::Return(val) => {
                if let Some(v) = val { self.compile_expr(v)?; } else { self.emit(Op::null); }
                self.emit(Op::r#return);
            }
            StmtKind::Break(_) => {
                let p = self.emit_jump(Op::br);
                if let Some(ctx) = self.loops.last_mut() { ctx.break_patches.push(p); }
            }
            StmtKind::Continue => {
                if let Some(ctx) = self.loops.last() {
                    let target = ctx.continue_target;
                    self.emit_loop(target);
                }
            }
            StmtKind::Throw(val) => {
                if let Some(v) = val { self.compile_expr(v)?; } else { self.emit(Op::null); }
                let line = self.line;
                common::errors::emit_throw(&mut self.chunks[self.current], line);
            }
            StmtKind::Exit(val) => {
                if let Some(v) = val {
                    self.compile_expr(v)?;
                } else if let Some(rs) = self.current_result_slot {
                    self.emit_u16(Op::local_get, rs);
                } else {
                    self.emit(Op::null);
                }
                self.emit(Op::r#return);
            }
            StmtKind::Raise(val) => {
                if let Some(v) = val { self.compile_expr(v)?; } else { self.emit(Op::null); }
                let line = self.line;
                common::errors::emit_throw(&mut self.chunks[self.current], line);
            }
            StmtKind::FunctionDecl { name, params, return_type, body, modifiers } => {
                self.defined_globals.insert(name.clone());
                self.defined_functions.insert(name.clone());
                let arity: u8 = params.len() as u8;
                let func_idx = self.chunks.len();
                let chunk = common::functions::create_function_chunk(name, arity);
                self.chunks.push(chunk);
                self.scopes.push(Scope::new_function());
                let saved = self.current;
                self.current = func_idx;

                for p in params { self.scope_mut().define(&p.name); }

                // Result slot for functions with return type
                let result_slot = if return_type.is_some() {
                    let rs = self.scope_mut().define("Result");
                    self.emit(Op::null); self.emit_u16(Op::local_set, rs); self.emit(Op::drop);
                    Some(rs)
                } else { None };

                // Track function name + result slot for Exit and assign-by-name
                let saved_fn = self.current_func_name.take();
                let saved_rs = self.current_result_slot.take();
                self.current_func_name = Some(name.clone());
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
            }
            StmtKind::ClassDecl { name, parent, members, .. } => {
                self.defined_globals.insert(name.clone());
                self.compile_class(name, parent, members)?;
            }
            StmtKind::EnumDecl { name, members } => {
                // Compile enum members as global ordinal constants
                if self.profile.enum_as_ordinals {
                    for (i, m) in members.iter().enumerate() {
                        if let Some(ref val) = m.value {
                            self.compile_expr(val)?;
                        } else {
                            self.emit_const(Value::F64(i as f64));
                        }
                        let idx = self.str_const(&m.name);
                        self.emit_u16(Op::global_set, idx);
                        self.emit(Op::drop);
                        self.defined_globals.insert(m.name.clone());
                    }
                }
            }
            StmtKind::Empty => {}
            StmtKind::With { expr, body } => {
                // Simplified: just compile body
                for s in body { self.compile_stmt(s)?; }
            }
            StmtKind::ModuleDecl { name: _, body } => {
                // Compile module body — declarations become globals, Sub Main is entry point
                for s in body { self.compile_stmt(s)?; }
                // Auto-call Main if defined (VB entry point)
                if self.defined_globals.contains("main") || self.defined_globals.contains("Main") {
                    let main_name = if self.defined_globals.contains("Main") { "Main" } else { "main" };
                    self.emit_var_get(main_name);
                    self.emit_u8(Op::call_ref, 0);
                    self.emit(Op::drop);
                }
            }
            _ => {
                // InterfaceDecl, StructDecl, TypeAlias, PropertyDecl, EventDecl, Extra
                // — no-ops or TODO
            }
        }
        Ok(())
    }

    // ── Assignment target ────────────────────────────────────────────────

    // ── Class compilation (ONE path for all languages) ─────────────────

    fn compile_class(&mut self, name: &str, parent: &Option<String>, members: &[Statement]) -> Result<(), String> {
        let self_kw = self.profile.self_keyword.clone();
        let ctor_name = self.profile.constructor_name.clone();
        let implicit_fields = self.profile.implicit_self_fields;
        let result_style = self.profile.function_return.clone();

        // Collect fields and their initializers
        let mut fields = Vec::new();
        let mut field_inits: Vec<(String, Option<Expression>)> = Vec::new();
        for m in members {
            if let StmtKind::VarDecl { name: fname, init, .. } = &m.kind {
                let fname = if self.case_sensitive { fname.clone() } else { fname.to_lowercase() };
                fields.push(fname.clone());
                field_inits.push((fname, init.clone()));
            }
        }

        // Store field list for implicit self resolution during method compilation
        self.pending_classes.insert(name.to_string(), PendingClass {
            parent: parent.clone(),
            fields: fields.clone(),
            methods: Vec::new(),
            ctor_arity: 0,
        });

        // Compile methods (including constructor body)
        let mut method_chunks: Vec<(String, usize, bool)> = Vec::new(); // (name, chunk_idx, is_ctor)
        let saved_class = self.current_class.take();
        self.current_class = Some(name.to_string());

        for m in members {
            if let StmtKind::FunctionDecl { name: mname, params, return_type, body, modifiers } = &m.kind {
                if body.is_empty() { continue; } // skip empty signatures

                let is_ctor = mname.eq_ignore_ascii_case(&ctor_name)
                    || modifiers.extra.iter().any(|e| e == "constructor");

                // Method arity: self + user params (unless explicit_self_param, then self is already in params)
                let user_params: Vec<&Param> = if self.profile.explicit_self_param {
                    params.iter().skip(1).collect() // skip self param
                } else {
                    params.iter().collect()
                };
                let arity = (user_params.len() + 1) as u8; // +1 for implicit self

                let ci = self.chunks.len();
                let chunk = common::functions::create_function_chunk(mname, arity);
                self.chunks.push(chunk);
                self.scopes.push(Scope::new_function());
                let saved = self.current;
                self.current = ci;

                // Define self slot
                self.scope_mut().define(&self_kw);
                for p in &user_params { self.scope_mut().define(&p.name); }

                if is_ctor {
                    // Constructor body — sets fields on Self, returns Self
                    for s in body { self.compile_stmt(s)?; }
                    if let Some(slot) = self.scope().resolve(&self_kw) {
                        self.emit_u16(Op::local_get, slot);
                        self.emit(Op::r#return);
                    }
                } else if return_type.is_some() && result_style == ReturnStyle::ResultSlot {
                    // Result-slot function (Pascal)
                    let rs = self.scope_mut().define("Result");
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
                    // Regular method — explicit return or epilogue
                    for s in body { self.compile_stmt(s)?; }
                    let line = self.line;
                    common::functions::emit_function_epilogue(&mut self.chunks[ci], line);
                }

                let locals = self.scope().next_slot;
                self.chunks[ci].local_count = locals;
                self.scopes.pop();
                self.current = saved;

                let bound_name = if self.case_sensitive { mname.clone() } else { mname.to_lowercase() };
                if is_ctor {
                    method_chunks.push((bound_name, ci, true));
                } else {
                    method_chunks.push((bound_name, ci, false));
                }
            }
        }

        self.current_class = saved_class;

        // Find constructor
        let ctor = method_chunks.iter().find(|(_, _, is_ctor)| *is_ctor);
        let user_arity = ctor.map(|(_, ci, _)| {
            let a = self.chunks[*ci].arity;
            if a > 0 { a - 1 } else { 0 }
        }).unwrap_or(0);

        // Build constructor wrapper using common::classes
        let wrapper_idx = self.chunks.len();
        let wrapper = common::functions::create_function_chunk(&format!("{}_ctor", name), user_arity);
        self.chunks.push(wrapper);
        let line = self.line;
        let this_slot = (user_arity as u16) + 1;
        self.chunks[wrapper_idx].local_count = this_slot + 1;

        let is_child = parent.is_some();
        if is_child {
            // Child class: call PARENT constructor first to create the object
            if let Some(parent_name) = parent {
                let pidx = self.chunks[wrapper_idx].add_constant(Value::String(Rc::from(parent_name.as_str())));
                self.chunks[wrapper_idx].emit_op_u16(Op::global_get, pidx, line);
                // Pass constructor args to parent
                for i in 0..user_arity {
                    self.chunks[wrapper_idx].emit_op_u16(Op::local_get, (i as u16) + 1, line);
                }
                self.chunks[wrapper_idx].emit_op_u8(Op::call_ref, user_arity as u8, line);
                self.chunks[wrapper_idx].emit_op_u16(Op::local_set, this_slot, line);
            } else {
                self.chunks[wrapper_idx].emit_op(Op::null, line);
                self.chunks[wrapper_idx].emit_op_u16(Op::local_set, this_slot, line);
            }
            // Call child's own constructor body if present
            if let Some((_, init_ci, _)) = ctor {
                common::functions::emit_ref_func(&mut self.chunks[wrapper_idx], *init_ci, 0, line);
                self.chunks[wrapper_idx].emit_op_u16(Op::local_get, this_slot, line);
                for i in 0..user_arity {
                    self.chunks[wrapper_idx].emit_op_u16(Op::local_get, (i as u16) + 1, line);
                }
                self.chunks[wrapper_idx].emit_op_u8(Op::call_ref, (user_arity + 1) as u8, line);
                self.chunks[wrapper_idx].emit_op_u16(Op::local_set, this_slot, line);
            }
            // Bind child methods (override parent's)
            for (mname, mci, is_ctor) in &method_chunks {
                if *is_ctor { continue; }
                common::classes::emit_bind_method_with_aliases(&mut self.chunks[wrapper_idx], this_slot, mname, *mci, line);
            }
            common::classes::emit_instanceof_chain(&mut self.chunks[wrapper_idx], this_slot, name, line);
            // Re-stamp __type for child
            self.chunks[wrapper_idx].emit_op_u16(Op::local_get, this_slot, line);
            let ts = self.chunks[wrapper_idx].add_constant(Value::String(Rc::from(name)));
            let tk = self.chunks[wrapper_idx].add_constant(Value::String(Rc::from("__type")));
            self.chunks[wrapper_idx].emit_op_u16(Op::r#const, ts, line);
            self.chunks[wrapper_idx].emit_op_u16(Op::struct_set, tk, line);
            self.chunks[wrapper_idx].emit_op(Op::drop, line);
        } else {
            // Base class: create object, init fields, bind methods, call constructor
            common::classes::emit_new_typed_object(&mut self.chunks[wrapper_idx], this_slot, name, line);
            // Init fields (with initializers if present, otherwise null)
            for (fname, init) in &field_inits {
                if let Some(init_expr) = init {
                    // Compile initializer into wrapper chunk
                    self.chunks[wrapper_idx].emit_op_u16(Op::local_get, this_slot, line);
                    let saved_cur = self.current;
                    self.current = wrapper_idx;
                    self.compile_expr(init_expr)?;
                    self.current = saved_cur;
                    let fk = self.chunks[wrapper_idx].add_constant(Value::String(Rc::from(fname.as_str())));
                    self.chunks[wrapper_idx].emit_op_u16(Op::struct_set, fk, line);
                    self.chunks[wrapper_idx].emit_op(Op::drop, line);
                } else {
                    common::classes::emit_init_field_null(&mut self.chunks[wrapper_idx], this_slot, fname, line);
                }
            }
            // Bind methods
            for (mname, mci, is_ctor) in &method_chunks {
                if *is_ctor { continue; }
                common::classes::emit_bind_method_with_aliases(&mut self.chunks[wrapper_idx], this_slot, mname, *mci, line);
            }
            // Call constructor body
            if let Some((_, init_ci, _)) = ctor {
                common::functions::emit_ref_func(&mut self.chunks[wrapper_idx], *init_ci, 0, line);
                self.chunks[wrapper_idx].emit_op_u16(Op::local_get, this_slot, line);
                for i in 0..user_arity {
                    self.chunks[wrapper_idx].emit_op_u16(Op::local_get, (i as u16) + 1, line);
                }
                self.chunks[wrapper_idx].emit_op_u8(Op::call_ref, (user_arity + 1) as u8, line);
                self.chunks[wrapper_idx].emit_op(Op::drop, line);
            }
            common::classes::emit_instanceof_chain(&mut self.chunks[wrapper_idx], this_slot, name, line);
        }
        common::classes::emit_constructor_return(&mut self.chunks[wrapper_idx], this_slot, line);

        // Store constructor as global
        let ctor_local = self.scope_mut().define(&format!("__{}_ctor", name));
        common::classes::emit_store_constructor(&mut self.chunks[self.current], name, wrapper_idx, ctor_local, line);

        // Register type
        let all_methods: Vec<(String, usize)> = method_chunks.iter().map(|(n, c, _)| (n.clone(), *c)).collect();
        let parent_str = parent.clone().unwrap_or_default();
        common::classes::register_type(&mut self.chunks, name, &parent_str, fields, all_methods, false, Vec::new(), Some(wrapper_idx));

        Ok(())
    }

    /// Check if a name is a field of the current class (for implicit self resolution).
    fn is_class_field(&self, name: &str) -> bool {
        if !self.profile.implicit_self_fields { return false; }
        if let Some(ref class_name) = self.current_class {
            let mut current = Some(class_name.as_str());
            while let Some(cn) = current {
                if let Some(pc) = self.pending_classes.get(cn) {
                    if pc.fields.iter().any(|f| f.eq_ignore_ascii_case(name)) {
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
            ExprKind::Member { object, field, .. } => {
                let mut parts = self.flatten_member_chain(object);
                parts.push(field.clone());
                parts
            }
            _ => Vec::new(),
        }
    }

    fn compile_assign_target(&mut self, target: &Expression) -> Result<(), String> {
        match &target.kind {
            ExprKind::Ident(name) => {
                // FuncName := value assigns to Result slot
                if let Some(ref fn_name) = self.current_func_name {
                    let matches = if self.case_sensitive { name == fn_name } else { name.eq_ignore_ascii_case(fn_name) };
                    if matches {
                        if let Some(rs) = self.current_result_slot {
                            self.emit_u16(Op::local_set, rs);
                            self.emit(Op::drop);
                            return Ok(());
                        }
                    }
                }
                // Implicit self field write: FName := value → Self.FName := value
                if self.is_class_field(name) {
                    let self_kw = self.profile.self_keyword.clone();
                    if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                        let tmp = self.scope_mut().define("__field_tmp");
                        self.emit_u16(Op::local_set, tmp); self.emit(Op::drop);
                        self.emit_u16(Op::local_get, slot);
                        self.emit_u16(Op::local_get, tmp);
                        let field_name = if self.case_sensitive { name.clone() } else { name.to_lowercase() };
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
                let field_name = if self.case_sensitive { field.clone() } else { field.to_lowercase() };
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
            ExprKind::Call { callee, args } if args.len() == 1 => {
                let tmp = self.scope_mut().define("__tmp");
                self.emit_u16(Op::local_set, tmp); self.emit(Op::drop);
                self.compile_expr(callee)?;
                self.compile_expr(&args[0])?;
                self.emit_u16(Op::local_get, tmp);
                self.emit(Op::array_set);
                self.emit(Op::drop);
            }
            _ => {}
        }
        Ok(())
    }

    // ── Expression compilation ───────────────────────────────────────────

    fn compile_expr(&mut self, expr: &Expression) -> Result<(), String> {
        match &expr.kind {
            ExprKind::Lit(lit) => {
                match lit {
                    Literal::Int(n) => self.emit_const(Value::F64(*n as f64)),
                    Literal::Float(n) => self.emit_const(Value::F64(*n)),
                    Literal::Str(s) => self.emit_const(Value::String(Rc::from(s.as_str()))),
                    Literal::Char(c) => self.emit_const(Value::String(Rc::from(c.to_string().as_str()))),
                    Literal::Bool(b) => if *b { self.emit(Op::r#true) } else { self.emit(Op::r#false) },
                    Literal::Null => self.emit(Op::null),
                }
            }
            ExprKind::Ident(name) => {
                match name.to_lowercase().as_str() {
                    "maxint" => self.emit_const(Value::F64(2147483647.0)),
                    "pi" => self.emit_const(Value::F64(std::f64::consts::PI)),
                    _ => {
                        // Implicit self field access: FName → Self.FName
                        if self.is_class_field(name) {
                            let self_kw = self.profile.self_keyword.clone();
                            if let Some(slot) = self.scope().resolve(&self_kw).or_else(|| self.scope().resolve_ci(&self_kw)) {
                                self.emit_u16(Op::local_get, slot);
                                let field_name = if self.case_sensitive { name.clone() } else { name.to_lowercase() };
                                let idx = self.str_const(&field_name);
                                self.emit_u16(Op::struct_get, idx);
                            } else {
                                self.emit_var_get(name);
                            }
                        } else {
                            self.emit_var_get(name);
                        }
                    }
                }
            }
            ExprKind::This => {
                let self_kw = &self.profile.self_keyword;
                if let Some(slot) = self.scope().resolve(self_kw)
                    .or_else(|| self.scope().resolve_ci(self_kw))
                    .or_else(|| self.scope().resolve("Self"))
                    .or_else(|| self.scope().resolve("self"))
                    .or_else(|| self.scope().resolve("this"))
                {
                    self.emit_u16(Op::local_get, slot);
                } else { self.emit(Op::null); }
            }
            ExprKind::Super => { self.emit(Op::null); /* TODO */ }
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
                self.compile_expr(left)?;
                self.compile_expr(right)?;
                self.compile_binop(op);
            }
            ExprKind::Unary { op, expr } => {
                self.compile_expr(expr)?;
                match op {
                    UnaryOp::Neg => { let l = self.line; common::math::emit_neg(self.chunk(), l); }
                    UnaryOp::Not => self.emit(Op::dyn_not),
                    UnaryOp::BitNot => self.emit(Op::i32_not),
                    UnaryOp::Pos => {} // no-op
                    UnaryOp::Deref => { let idx = self.str_const("__value"); self.emit_u16(Op::struct_get, idx); }
                    UnaryOp::AddrOf => {} // no-op in our VM
                    UnaryOp::PreInc | UnaryOp::PostInc | UnaryOp::PreDec | UnaryOp::PostDec => {
                        // TODO: proper inc/dec
                    }
                    UnaryOp::Typeof => {} // TODO
                }
            }
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
            ExprKind::Call { callee, args } => {
                // Check for builtin calls — both Ident("print") and Member("Console.WriteLine")
                if let ExprKind::Ident(name) = &callee.kind {
                    if self.try_compile_builtin(name, args)? { return Ok(()); }
                }
                if let ExprKind::Member { object, field, .. } = &callee.kind {
                    if let ExprKind::Ident(obj_name) = &object.kind {
                        let compound = format!("{}.{}", obj_name, field);
                        if self.try_compile_builtin(&compound, args)? { return Ok(()); }
                    }
                }
                // Constructor call: ClassName.Create(args) → global_get(ClassName) + call_ref
                if let ExprKind::Member { object, field, .. } = &callee.kind {
                    if let ExprKind::Ident(class_name) = &object.kind {
                        let ctor_name = &self.profile.constructor_name;
                        let is_ctor = if self.case_sensitive { field == ctor_name } else { field.eq_ignore_ascii_case(ctor_name) };
                        if is_ctor && self.defined_globals.contains(class_name.as_str()) {
                            self.emit_var_get(class_name);
                            for a in args { self.compile_expr(a)?; }
                            self.emit_u8(Op::call_ref, args.len() as u8);
                            return Ok(());
                        }
                    }
                }
                // Namespace chain resolution: vybe.gui.setProperty(...) → host call or namespace object chain
                if let ExprKind::Member { .. } = &callee.kind {
                    let parts = self.flatten_member_chain(callee);
                    if parts.len() >= 2 {
                        let lower_parts: Vec<String> = parts.iter().map(|s| s.to_lowercase()).collect();
                        // Try dotnet interface resolution
                        let refs: Vec<&str> = lower_parts.iter().map(|s| s.as_str()).collect();
                        if let Some((module, func)) = common::dotnet::resolve_interface_call(&refs, &common::dotnet::default_interface_imports()) {
                            for a in args { self.compile_expr(a)?; }
                            let idx = self.import(&module, &func);
                            self.emit_host_call(idx, args.len() as u8);
                            return Ok(());
                        }
                        // Namespace object chain: global_get root → struct_get ... → call
                        // Skip known no-op methods (WinForms layout, etc.)
                        if let Some(last) = lower_parts.last() {
                            if common::dotnet::is_noop_method(last) {
                                self.emit(Op::null);
                                return Ok(());
                            }
                        }
                        // Only use namespace resolution if the root is NOT a local variable or known var
                        let root_is_local = self.scope().resolve(&lower_parts[0]).is_some()
                            || (!self.case_sensitive && self.scope().resolve_ci(&lower_parts[0]).is_some())
                            || self.defined_globals.contains(&lower_parts[0]);
                        if !root_is_local && common::dotnet::namespace_roots().contains(&lower_parts[0]) {
                            // vybe.gui.setProperty → call_import "vybe:gui" "setProperty"
                            if lower_parts[0] == "vybe" && parts.len() >= 3 {
                                let module = format!("vybe:{}", lower_parts[1]);
                                let func = parts[2..].join("."); // preserve original case for host fn names
                                for a in args { self.compile_expr(a)?; }
                                let idx = self.import(&module, &func);
                                self.emit_host_call(idx, args.len() as u8);
                                return Ok(());
                            }
                            let root_idx = self.str_const(&lower_parts[0]);
                            self.emit_u16(Op::global_get, root_idx);
                            for part in &lower_parts[1..] {
                                let idx = self.str_const(part);
                                self.emit_u16(Op::struct_get, idx);
                            }
                            for a in args { self.compile_expr(a)?; }
                            self.emit_u8(Op::call_ref, args.len() as u8);
                            return Ok(());
                        }
                    }
                }
                // Method call: obj.method(args)
                if let ExprKind::Member { object, field, .. } = &callee.kind {
                    self.compile_expr(object)?;
                    let field_name = if self.case_sensitive { field.clone() } else { field.to_lowercase() };
                    let prop = self.str_const(&field_name);
                    self.emit(Op::dup);
                    self.emit_u16(Op::struct_get, prop);
                    let fn_tmp = self.scope_mut().define("__fn");
                    self.emit_u16(Op::local_set, fn_tmp); self.emit(Op::drop);
                    let obj_tmp = self.scope_mut().define("__obj");
                    self.emit_u16(Op::local_set, obj_tmp); self.emit(Op::drop);
                    self.emit_u16(Op::local_get, fn_tmp);
                    self.emit_u16(Op::local_get, obj_tmp);
                    for a in args { self.compile_expr(a)?; }
                    self.emit_u8(Op::call_ref, (args.len() + 1) as u8);
                    return Ok(());
                }
                // Constructor call: ClassName(args) or ClassName.Create(args)
                if let ExprKind::Ident(name) = &callee.kind {
                    // VB: name(idx) could be array access or function call.
                    // If name is a local variable and there's 1 arg, treat as array access.
                    let is_local = self.scope().resolve(name).is_some()
                        || (!self.case_sensitive && self.scope().resolve_ci(name).is_some());
                    let is_known_func = self.defined_functions.contains(name)
                        || (!self.case_sensitive && self.defined_functions.iter().any(|g| g.eq_ignore_ascii_case(name)));
                    let is_builtin = self.try_compile_builtin(name, args)?;
                    if is_builtin { return Ok(()); }
                    if !is_known_func && args.len() == 1 && !self.case_sensitive {
                        // VB array access: arr(idx)
                        self.emit_var_get(name);
                        self.compile_expr(&args[0])?;
                        self.emit(Op::array_get);
                    } else {
                        // Function call (including forward references resolved at runtime)
                        self.emit_var_get(name);
                        for a in args { self.compile_expr(a)?; }
                        self.emit_u8(Op::call_ref, args.len() as u8);
                    }
                    return Ok(());
                }
                self.compile_expr(callee)?;
                for a in args { self.compile_expr(a)?; }
                self.emit_u8(Op::call_ref, args.len() as u8);
            }
            ExprKind::Member { object, field, .. } => {
                // ClassName.Create (no parens) → constructor call with 0 args
                if let ExprKind::Ident(class_name) = &object.kind {
                    let ctor_name = &self.profile.constructor_name;
                    let is_ctor = if self.case_sensitive { field == ctor_name } else { field.eq_ignore_ascii_case(ctor_name) };
                    if is_ctor && self.defined_globals.contains(class_name.as_str()) {
                        self.emit_var_get(class_name);
                        self.emit_u8(Op::call_ref, 0);
                        return Ok(());
                    }
                }
                self.compile_expr(object)?;
                let field_name = if self.case_sensitive { field.clone() } else { field.to_lowercase() };
                let idx = self.str_const(&field_name);
                self.emit_u16(Op::struct_get, idx);
            }
            ExprKind::Index { object, index } => {
                self.compile_expr(object)?;
                self.compile_expr(index)?;
                self.emit(Op::array_get);
            }
            ExprKind::New { class, args } => {
                // Check for known .NET types (List, Dictionary, Point, etc.)
                if let ExprKind::Ident(name) = &class.kind {
                    // Strip generic params: "list(of string)" → "list"
                    let bare = name.to_lowercase();
                    let bare = bare.split('(').next().unwrap_or(&bare).trim();
                    // Also strip namespace: "system.drawing.point" → "point"
                    let bare = bare.rsplit('.').next().unwrap_or(bare);
                    let known = common::dotnet::known_types();
                    if let Some(&(module, func)) = known.get(bare) {
                        for a in args { self.compile_expr(a)?; }
                        let idx = self.import(module, func);
                        self.emit_host_call(idx, args.len() as u8);
                        return Ok(());
                    }
                }
                // User-defined class constructor
                self.compile_expr(class)?;
                for a in args { self.compile_expr(a)?; }
                self.emit_u8(Op::call_ref, args.len() as u8);
            }
            ExprKind::Assign { target, value } => {
                self.compile_expr(value)?;
                self.emit(Op::dup);
                self.compile_assign_target(target)?;
            }
            ExprKind::Lambda { params, body, .. } => {
                let arity = params.len() as u8;
                let ci = self.chunks.len();
                let chunk = common::functions::create_function_chunk("<lambda>", arity);
                self.chunks.push(chunk);
                self.scopes.push(Scope::new_function());
                let saved = self.current;
                self.current = ci;
                for p in params { self.scope_mut().define(&p.name); }

                // For Result-slot languages, add Result slot
                let result_slot = if self.profile.function_return == ReturnStyle::ResultSlot {
                    let rs = self.scope_mut().define("Result");
                    self.emit(Op::null); self.emit_u16(Op::local_set, rs); self.emit(Op::drop);
                    let saved_fn = self.current_func_name.take();
                    let saved_rs = self.current_result_slot.take();
                    self.current_func_name = Some("<lambda>".into());
                    self.current_result_slot = Some(rs);
                    Some((rs, saved_fn, saved_rs))
                } else { None };

                for s in body { self.compile_stmt(s)?; }

                if let Some((rs, saved_fn, saved_rs)) = result_slot {
                    self.emit_u16(Op::local_get, rs);
                    self.emit(Op::r#return);
                    self.current_func_name = saved_fn;
                    self.current_result_slot = saved_rs;
                } else {
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
            }
            ExprKind::Array(items) => {
                for item in items { self.compile_expr(item)?; }
                let line = self.line;
                self.chunks[self.current].emit_op_u16(Op::array_new, items.len() as u16, line);
            }
            ExprKind::Object(pairs) => {
                let line = self.line;
                common::dict::emit_new(&mut self.chunks[self.current], line);
                for (key, val) in pairs {
                    self.emit(Op::dup);
                    self.compile_expr(val)?;
                    // Key should be a string
                    if let ExprKind::Lit(Literal::Str(k)) = &key.kind {
                        let idx = self.str_const(k);
                        self.emit_u16(Op::struct_set, idx);
                    } else {
                        self.compile_expr(key)?;
                        self.emit(Op::array_set);
                    }
                    self.emit(Op::drop);
                }
            }
            ExprKind::Inherited { method, args } => {
                // inherited Create(args) → call parent constructor, store result in Self
                if let Some(ref class_name) = self.current_class.clone() {
                    let parent = self.pending_classes.get(class_name.as_str()).and_then(|c| c.parent.clone());
                    if let Some(parent_name) = parent {
                        let parent_idx = self.str_const(&parent_name);
                        self.emit_u16(Op::global_get, parent_idx);
                        for a in args { self.compile_expr(a)?; }
                        self.emit_u8(Op::call_ref, args.len() as u8);
                        // Store result in Self slot
                        let self_kw = self.profile.self_keyword.clone();
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
            }
            ExprKind::IsType { expr, type_name } => {
                self.compile_expr(expr)?;
                let key = self.str_const("__type");
                self.emit_u16(Op::struct_get, key);
                self.emit_const(Value::String(Rc::from(type_name.as_str())));
                self.emit(Op::dyn_eq);
            }
            ExprKind::AsCast { expr, .. } => {
                self.compile_expr(expr)?;
            }
            ExprKind::Spread(inner) => {
                self.compile_expr(inner)?;
                self.emit(Op::spread);
            }
            ExprKind::Await(inner) => {
                self.compile_expr(inner)?;
                // TODO: async/await
            }
            ExprKind::Yield(val) => {
                if let Some(v) = val { self.compile_expr(v)?; } else { self.emit(Op::null); }
                self.emit_u16(Op::suspend, 0);
            }
            ExprKind::Extra { tag, exprs } => {
                match tag.as_str() {
                    "array_new" => {
                        // Dim arr(N) → create array of size N+1
                        if let Some(size_expr) = exprs.first() {
                            self.compile_expr(size_expr)?;
                            self.emit_const(Value::F64(1.0));
                            self.emit(Op::dyn_add);
                            self.emit(Op::array_new_default);
                        } else {
                            self.emit(Op::null);
                        }
                    }
                    "slice" => {
                        // [obj, start, end, step]
                        for e in exprs { self.compile_expr(e)?; }
                        let idx = self.import("vybe:array", "sliceStep");
                        self.emit_host_call(idx, exprs.len() as u8);
                    }
                    _ => self.emit(Op::null),
                }
            }
            _ => {
                self.emit(Op::null);
            }
        }
        Ok(())
    }

    fn compile_binop(&mut self, op: &BinOp) {
        match op {
            BinOp::Add => self.emit(Op::dyn_add),
            BinOp::Sub => self.emit(Op::f64_sub),
            BinOp::Mul => self.emit(Op::f64_mul),
            BinOp::Div => self.emit(Op::f64_div),
            BinOp::IDiv => { self.emit(Op::f64_div); let l = self.line; common::math::emit_trunc(self.chunk(), l); }
            BinOp::Mod => self.emit(Op::f64_mod),
            BinOp::Pow => { let i = self.import("vybe:math", "pow"); self.emit_host_call(i, 2); }
            BinOp::Eq => self.emit(Op::dyn_eq),
            BinOp::NotEq => self.emit(Op::dyn_ne),
            BinOp::Lt => self.emit(Op::dyn_lt),
            BinOp::Gt => self.emit(Op::dyn_gt),
            BinOp::Le => self.emit(Op::dyn_le),
            BinOp::Ge => self.emit(Op::dyn_ge),
            BinOp::And | BinOp::Or => unreachable!(), // handled with short-circuit
            BinOp::Xor => self.emit(Op::i32_xor),
            BinOp::BitAnd => self.emit(Op::i32_and),
            BinOp::BitOr => self.emit(Op::i32_or),
            BinOp::BitXor => self.emit(Op::i32_xor),
            BinOp::Shl => self.emit(Op::i32_shl),
            BinOp::Shr => self.emit(Op::i32_shr_s),
            BinOp::Concat => { let l = self.line; common::strings::emit_str_concat(self.chunk(), l); }
            BinOp::In | BinOp::NotIn => {
                let idx = self.import("vybe:collections", "setHas");
                self.emit_host_call(idx, 2);
                if *op == BinOp::NotIn { self.emit(Op::dyn_not); }
            }
            BinOp::NullCoalesce => { /* TODO */ }
        }
    }

    // ── Builtins ─────────────────────────────────────────────────────────

    fn try_compile_builtin(&mut self, name: &str, args: &[Expression]) -> Result<bool, String> {
        let line = self.line;

        // Check common import table first
        if let Some((module, func)) = common::imports::resolve_common_import(name) {
            for a in args { self.compile_expr(a)?; }
            let idx = self.import(module, func);
            self.emit_host_call(idx, args.len() as u8);
            return Ok(true);
        }

        // Look up in language profile
        let builtin = self.profile.lookup_builtin(name, self.case_sensitive).cloned();
        if let Some(def) = builtin {
            match &def.emit {
                BuiltinEmit::Print => {
                    for a in args { self.compile_expr(a)?; }
                    let idx = self.import("wasi:cli", "log");
                    common::io::emit_print_with_import(self.chunk(), idx, args.len() as u8, line);
                }
                BuiltinEmit::StrLength => {
                    self.compile_expr(&args[0])?;
                    common::strings::emit_length(self.chunk(), line);
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
                    if let Some(ExprKind::Ident(var)) = args.first().map(|a| &a.kind) {
                        let var = var.clone();
                        self.emit_var_get(&var);
                        if args.len() > 1 { self.compile_expr(&args[1])?; } else { self.emit_const(Value::F64(1.0)); }
                        match op.as_str() {
                            "add" => self.emit(Op::dyn_add),
                            "sub" => self.emit(Op::f64_sub),
                            _ => self.emit(Op::dyn_add),
                        }
                        self.emit_var_set(&var);
                    }
                    self.emit(Op::null);
                }
                BuiltinEmit::Noop => {
                    self.emit(Op::null);
                }
            }
            return Ok(true);
        }

        Ok(false)
    }

    /// Emit a named opcode sequence for a builtin.
    fn emit_builtin_opcode(&mut self, op_name: &str, args: &[Expression]) -> Result<(), String> {
        let line = self.line;
        match op_name {
            "abs" => { self.compile_expr(&args[0])?; common::math::emit_abs(self.chunk(), line); }
            "sqrt" => { self.compile_expr(&args[0])?; common::math::emit_sqrt(self.chunk(), line); }
            "round" => { self.compile_expr(&args[0])?; common::math::emit_round(self.chunk(), line); }
            "trunc" => { self.compile_expr(&args[0])?; common::math::emit_trunc(self.chunk(), line); }
            "floor" => { self.compile_expr(&args[0])?; common::math::emit_floor(self.chunk(), line); }
            "ceil" => { self.compile_expr(&args[0])?; common::math::emit_ceil(self.chunk(), line); }
            "min" => { self.compile_expr(&args[0])?; self.compile_expr(&args[1])?; common::math::emit_min(self.chunk(), line); }
            "max" => { self.compile_expr(&args[0])?; self.compile_expr(&args[1])?; common::math::emit_max(self.chunk(), line); }
            "sqr" => { self.compile_expr(&args[0])?; self.emit(Op::dup); self.emit(Op::f64_mul); }
            "succ" => { self.compile_expr(&args[0])?; self.emit_const(Value::F64(1.0)); self.emit(Op::dyn_add); }
            "pred" => { self.compile_expr(&args[0])?; self.emit_const(Value::F64(1.0)); self.emit(Op::f64_sub); }
            "to_upper" => { self.compile_expr(&args[0])?; common::strings::emit_to_upper(self.chunk(), line); }
            "to_lower" => { self.compile_expr(&args[0])?; common::strings::emit_to_lower(self.chunk(), line); }
            "trim" => { self.compile_expr(&args[0])?; common::strings::emit_trim(self.chunk(), line); }
            "concat" => { for a in args { self.compile_expr(a)?; } common::strings::emit_concat(self.chunk(), args.len(), line); }
            "replace" => { self.compile_expr(&args[0])?; self.compile_expr(&args[1])?; self.compile_expr(&args[2])?; common::strings::emit_replace(self.chunk(), line); }
            "repeat" => { self.compile_expr(&args[0])?; self.compile_expr(&args[1])?; common::strings::emit_repeat(self.chunk(), line); }
            "leftstr" => { self.compile_expr(&args[0])?; self.emit_const(Value::F64(0.0)); self.compile_expr(&args[1])?; common::strings::emit_substring(self.chunk(), line); }
            "high" => {
                self.compile_expr(&args[0])?;
                common::strings::emit_length(self.chunk(), line);
                self.emit_const(Value::F64(1.0));
                self.emit(Op::f64_sub);
            }
            "low" => { self.emit_const(Value::F64(0.0)); }
            "setlength" => {
                if let Some(ExprKind::Ident(var)) = args.first().map(|a| &a.kind) {
                    let var = var.clone();
                    self.compile_expr(&args[1])?;
                    let idx = self.import("vybe:array", "newWithLength");
                    self.emit_host_call(idx, 1);
                    self.emit_var_set(&var);
                }
                self.emit(Op::null);
            }
            "trim_start" => { self.compile_expr(&args[0])?; common::strings::emit_trim_start(self.chunk(), line); }
            "trim_end" => { self.compile_expr(&args[0])?; common::strings::emit_trim_end(self.chunk(), line); }
            "pow" => { self.compile_expr(&args[0])?; self.compile_expr(&args[1])?; common::math::emit_pow(self.chunk(), line); }
            "log" => { self.compile_expr(&args[0])?; common::math::emit_log(self.chunk(), line); }
            "sin" => { self.compile_expr(&args[0])?; common::math::emit_sin(self.chunk(), line); }
            "cos" => { self.compile_expr(&args[0])?; common::math::emit_cos(self.chunk(), line); }
            "tan" => { self.compile_expr(&args[0])?; common::math::emit_tan(self.chunk(), line); }
            "exp" => { self.compile_expr(&args[0])?; common::math::emit_exp(self.chunk(), line); }
            "is_null" => { self.compile_expr(&args[0])?; self.emit(Op::ref_is_null); }
            "space" => {
                // Space(n) → " " repeated n times
                self.emit_const(Value::String(Rc::from(" ")));
                self.compile_expr(&args[0])?;
                common::strings::emit_repeat(self.chunk(), line);
            }
            "assigned" => { self.compile_expr(&args[0])?; self.emit(Op::ref_is_null); self.emit(Op::dyn_not); }
            "freeandnil" => {
                if let Some(ExprKind::Ident(var)) = args.first().map(|a| &a.kind) {
                    let var = var.clone();
                    self.emit(Op::null);
                    self.emit_var_set(&var);
                }
                self.emit(Op::null);
            }
            _ => { self.emit(Op::null); }
        }
        Ok(())
    }
}
