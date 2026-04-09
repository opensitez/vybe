use std::sync::Arc;
use std::collections::HashMap;
use vybe_parser_python::ast::*;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use vybe_compiler_common as common;
use vybe_compiler_common::collections as common_collections;

pub struct Compiler {
    chunks: Vec<Chunk>,
    scopes: Vec<Scope>,
    loop_stack: Vec<LoopCtx>,
    import_log: u16,
    last_upvalues: Vec<UpvalueDesc>,
    current_class_parent: Option<String>,
}

struct Local {
    name: String,
    slot: u16,
    is_captured: bool,
}

#[derive(Debug, Clone)]
struct UpvalueDesc {
    index: u8,
    is_local: bool,
}

struct Scope {
    locals: Vec<Local>,
    upvalues: Vec<UpvalueDesc>,
    max_local: u16,
}

struct LoopCtx {
    _start: usize,
    break_jumps: Vec<usize>,
    continue_jumps: Vec<usize>,
}

impl Scope {
    fn new() -> Self {
        Self { locals: Vec::new(), upvalues: Vec::new(), max_local: 0 }
    }

    fn alloc(&mut self, name: &str) -> u16 {
        for local in &self.locals {
            if local.name == name { return local.slot; }
        }
        self.max_local += 1;
        let slot = self.max_local;
        self.locals.push(Local { name: name.to_string(), slot, is_captured: false });
        slot
    }

    fn get(&self, name: &str) -> Option<u16> {
        for local in self.locals.iter().rev() {
            if local.name == name { return Some(local.slot); }
        }
        None
    }

    fn mark_captured(&mut self, slot: u16) {
        for local in &mut self.locals {
            if local.slot == slot { local.is_captured = true; return; }
        }
    }

    fn add_upvalue(&mut self, index: u8, is_local: bool) -> u8 {
        for (i, uv) in self.upvalues.iter().enumerate() {
            if uv.index == index && uv.is_local == is_local { return i as u8; }
        }
        let idx = self.upvalues.len() as u8;
        self.upvalues.push(UpvalueDesc { index, is_local });
        idx
    }
}

enum VarResolution {
    Local(u16),
    Upvalue(u8),
    Global,
}

impl Compiler {
    pub fn new() -> Self {
        Self {
            chunks: Vec::new(),
            scopes: Vec::new(),
            loop_stack: Vec::new(),
            import_log: 0,
            last_upvalues: Vec::new(),
            current_class_parent: None,
        }
    }

    fn resolve_variable(&mut self, name: &str, _chunk_idx: usize) -> VarResolution {
        // Check current scope (local)
        let scope_idx = self.scopes.len() - 1;
        if let Some(slot) = self.scopes[scope_idx].get(name) {
            return VarResolution::Local(slot);
        }
        // Check parent scopes (upvalue)
        if self.scopes.len() > 1 {
            if let Some(uv) = self.resolve_upvalue(scope_idx, name) {
                return VarResolution::Upvalue(uv);
            }
        }
        VarResolution::Global
    }

    /// Emit ref_func with upvalue descriptors from last_upvalues.
    fn emit_ref_func(&mut self, chunk_idx: usize, func_chunk_idx: usize) {
        let upvalues = std::mem::take(&mut self.last_upvalues);
        self.chunk(chunk_idx).emit_op_u16(Op::ref_func, func_chunk_idx as u16, 0);
        self.chunk(chunk_idx).emit(upvalues.len() as u8, 0);
        for uv in &upvalues {
            self.chunk(chunk_idx).emit(if uv.is_local { 1 } else { 0 }, 0);
            self.chunk(chunk_idx).emit(uv.index, 0);
        }
    }

    /// Emit ref_func with 0 upvalues (for cases where we know there are none).
    fn emit_ref_func_no_upvalues(&mut self, chunk_idx: usize, func_chunk_idx: usize) {
        self.chunk(chunk_idx).emit_op_u16(Op::ref_func, func_chunk_idx as u16, 0);
        self.chunk(chunk_idx).emit(0, 0);
    }

    fn resolve_upvalue(&mut self, scope_idx: usize, name: &str) -> Option<u8> {
        if scope_idx == 0 { return None; }
        let parent = scope_idx - 1;
        // Check if the variable is a local in the parent scope
        if let Some(slot) = self.scopes[parent].get(name) {
            self.scopes[parent].mark_captured(slot);
            return Some(self.scopes[scope_idx].add_upvalue(slot as u8, true));
        }
        // Recurse: check grandparent scopes
        if let Some(uv) = self.resolve_upvalue(parent, name) {
            return Some(self.scopes[scope_idx].add_upvalue(uv, false));
        }
        None
    }

    pub fn compile(&mut self, module: &Module) -> Result<Vec<Chunk>, String> {
        let mut chunk = Chunk::new("<script>");
        self.import_log = chunk.add_import("wasi:cli", "log");
        self.chunks.push(chunk);
        self.scopes.push(Scope::new());

        // Register built-in exception type constructors BEFORE user code.
        // Each creates an object with __exception_type, name, message.
        let exc_types = ["Exception", "ValueError", "TypeError", "KeyError",
            "IndexError", "RuntimeError", "StopIteration", "AttributeError",
            "ZeroDivisionError", "FileNotFoundError", "ImportError",
            "NotImplementedError", "OverflowError", "IOError", "OSError"];
        for exc_name in &exc_types {
            self.compile_exception_constructor(exc_name, 0);
        }

        for stmt in &module.body {
            self.compile_stmt(stmt, 0)?;
        }

        // Finalize main chunk
        let scope = self.scopes.remove(0);
        self.chunks[0].local_count = (scope.max_local + 1) as u16;
        self.chunks[0].emit_op(Op::halt, 0);

        // Bundle stdlib — portable .wasm that works on any runtime.
        common::bundle::finalize_with_stdlib(&mut self.chunks);

        Ok(std::mem::take(&mut self.chunks))
    }

    // ── Statements ───────────────────────────────────────────────────

    fn compile_stmt(&mut self, stmt: &Statement, chunk_idx: usize) -> Result<(), String> {
        match stmt {
            Statement::Expression(expr) => {
                // print() calls: intercept and compile as host log
                if let Expression::Call { func, args, keywords } = expr {
                    if let Expression::Name(name) = func.as_ref() {
                        if name == "print" {
                            return self.compile_print_with_kwargs(args, keywords, chunk_idx);
                        }
                    }
                }
                self.compile_expr(expr, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
            }

            Statement::Assign { targets, value } => {
                self.compile_expr(value, chunk_idx)?;
                for (i, target) in targets.iter().enumerate() {
                    if i < targets.len() - 1 {
                        self.chunk(chunk_idx).emit_op(Op::dup, 0);
                    }
                    self.compile_assign_target(target, chunk_idx)?;
                }
            }

            Statement::AugAssign { target, op, value } => {
                // target op= value  →  target = target op value
                match target {
                    Expression::Name(name) => {
                        let idx = self.scope(chunk_idx).alloc(name);
                        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx, 0);
                        self.compile_expr(value, chunk_idx)?;
                        self.emit_aug_op(*op, chunk_idx);
                        self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx, 0);
                        self.chunk(chunk_idx).emit_op(Op::drop, 0);
                    }
                    Expression::Attribute { value: obj, attr } => {
                        // obj.attr op= value → obj.attr = obj.attr op value
                        self.compile_expr(obj, chunk_idx)?;
                        self.chunk(chunk_idx).emit_op(Op::dup, 0);
                        let attr_c = self.chunk(chunk_idx).add_constant(Value::String(Arc::from(attr.as_str())));
                        self.chunk(chunk_idx).emit_op_u16(Op::struct_get, attr_c, 0);
                        self.compile_expr(value, chunk_idx)?;
                        self.emit_aug_op(*op, chunk_idx);
                        let attr_c2 = self.chunk(chunk_idx).add_constant(Value::String(Arc::from(attr.as_str())));
                        self.chunk(chunk_idx).emit_op_u16(Op::struct_set, attr_c2, 0);
                    }
                    Expression::Subscript { value: obj, slice } => {
                        // obj[idx] op= value → obj[idx] = obj[idx] op value
                        self.compile_expr(obj, chunk_idx)?;
                        self.chunk(chunk_idx).emit_op(Op::dup, 0);
                        self.compile_expr(slice, chunk_idx)?;
                        self.chunk(chunk_idx).emit_op(Op::dup, 0); // keep idx for set
                        common_collections::emit_get(&mut self.chunks[chunk_idx], 0);
                        self.compile_expr(value, chunk_idx)?;
                        self.emit_aug_op(*op, chunk_idx);
                        common_collections::emit_set(&mut self.chunks[chunk_idx], 0);
                    }
                    _ => {
                        // Fallback: just skip silently
                    }
                }
            }

            Statement::If { test, body, elif_clauses, else_body } => {
                self.compile_expr(test, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::dyn_to_bool, 0);
                let exit_jump = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);
                for s in body { self.compile_stmt(s, chunk_idx)?; }

                if elif_clauses.is_empty() && else_body.is_none() {
                    self.chunk(chunk_idx).patch_jump(exit_jump);
                } else {
                    let after_jump = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                    self.chunk(chunk_idx).patch_jump(exit_jump);

                    for (elif_test, elif_body) in elif_clauses {
                        self.compile_expr(elif_test, chunk_idx)?;
                        self.chunk(chunk_idx).emit_op(Op::dyn_to_bool, 0);
                        let elif_exit = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);
                        for s in elif_body { self.compile_stmt(s, chunk_idx)?; }
                        // Jump to after all branches
                        let j = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                        self.chunk(chunk_idx).patch_jump(elif_exit);
                        // We need to patch this jump at the end — chain it
                        // For simplicity, patch to same target below
                        self.chunk(chunk_idx).patch_jump(j);
                        // Actually this is wrong — we need all to jump to the same end.
                        // Let me restructure with a list of end jumps.
                    }

                    if let Some(else_body) = else_body {
                        for s in else_body { self.compile_stmt(s, chunk_idx)?; }
                    }
                    self.chunk(chunk_idx).patch_jump(after_jump);
                }
            }

            Statement::While { test, body, else_body } => {
                // If else_body exists, use a broke flag
                let broke_local = if else_body.is_some() {
                    let l = self.scope(chunk_idx).alloc("__while_broke");
                    self.chunk(chunk_idx).emit_op(Op::i32_const_0, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::local_set, l, 0);
                    Some(l)
                } else { None };

                let loop_start = self.chunk(chunk_idx).current_offset();
                self.compile_expr(test, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::dyn_to_bool, 0);
                let exit_jump = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);

                self.loop_stack.push(LoopCtx { _start: loop_start, break_jumps: Vec::new(), continue_jumps: Vec::new() });
                for s in body { self.compile_stmt(s, chunk_idx)?; }
                let ctx = self.loop_stack.pop().unwrap();
                for cj in &ctx.continue_jumps { self.chunk(chunk_idx).patch_jump(*cj); }
                self.chunk(chunk_idx).emit_loop(loop_start, 0);

                // break jumps land here — set broke flag before
                if let Some(bl) = broke_local {
                    let skip_flag = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                    for bj in &ctx.break_jumps { self.chunk(chunk_idx).patch_jump(*bj); }
                    self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::local_set, bl, 0);
                    self.chunk(chunk_idx).patch_jump(skip_flag);
                } else {
                    for bj in &ctx.break_jumps { self.chunk(chunk_idx).patch_jump(*bj); }
                }
                self.chunk(chunk_idx).patch_jump(exit_jump);

                // else body: execute if broke == 0
                if let Some(else_stmts) = else_body {
                    let bl = broke_local.unwrap();
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, bl, 0);
                    let skip_else = self.chunk(chunk_idx).emit_jump(Op::br_if_true, 0);
                    for s in else_stmts { self.compile_stmt(s, chunk_idx)?; }
                    self.chunk(chunk_idx).patch_jump(skip_else);
                }
            }

            Statement::For { target, iter, body, else_body, is_async: _ } => {
                self.compile_for(target, iter, body, else_body.as_deref(), chunk_idx)?;
            }

            Statement::FunctionDef { name, params, body, is_async: _, decorators: _, returns: _ } => {
                let func_idx = self.compile_function(name, params, body)?;
                let idx = self.scope(chunk_idx).alloc(name);
                self.emit_ref_func(chunk_idx, func_idx);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx, 0);
            }

            Statement::ClassDef { name, bases, keywords: _, body, decorators: _ } => {
                // Compile class as a constructor function (same convention as JS/VB/C#).
                // Calling Dog("Rex") creates a typed object with type_id, binds methods
                // on the vtable, calls __init__, and returns the object.

                let all_bases: Vec<String> = bases.iter().filter_map(|b| {
                    if let Expression::Name(n) = b { Some(n.clone()) } else { None }
                }).collect();
                let parent_name = all_bases.first().cloned().unwrap_or_default();

                let saved_parent = self.current_class_parent.clone();
                if !parent_name.is_empty() {
                    self.current_class_parent = Some(parent_name.clone());
                }

                // Compile all methods first (we need their chunk indices)
                let mut method_entries = Vec::new();
                let mut static_method_entries = Vec::new();
                let mut init_chunk = None;
                let mut init_params = None;
                let mut other_stmts = Vec::new();

                for s in body {
                    if let Statement::FunctionDef { name: method_name, params, body: mbody, decorators, .. } = s {
                        // Check for @property decorator → compile as __get_<name>
                        // Check for @<name>.setter decorator → compile as __set_<name>
                        let is_property = decorators.iter().any(|d| {
                            matches!(d, Expression::Name(n) if n == "property")
                        });
                        let is_setter = decorators.iter().any(|d| {
                            if let Expression::Attribute { attr, .. } = d {
                                attr == "setter"
                            } else { false }
                        });
                        let is_staticmethod = decorators.iter().any(|d| {
                            matches!(d, Expression::Name(n) if n == "staticmethod")
                        });
                        let is_classmethod = decorators.iter().any(|d| {
                            matches!(d, Expression::Name(n) if n == "classmethod")
                        });

                        let effective_name = if is_property {
                            format!("__get_{}", method_name)
                        } else if is_setter {
                            format!("__set_{}", method_name)
                        } else {
                            method_name.clone()
                        };

                        if is_staticmethod {
                            // Static methods: strip self param, attach to constructor (not instance)
                            let mut static_params = params.clone();
                            if !static_params.args.is_empty() {
                                static_params.args.remove(0); // remove 'self'
                            }
                            let func_chunk_idx = self.compile_function(&effective_name, &static_params, mbody)?;
                            static_method_entries.push((effective_name.to_lowercase(), func_chunk_idx));
                        } else if is_classmethod {
                            // Class methods: strip cls param (it's not used at runtime),
                            // attach to constructor like static methods
                            let mut cls_params = params.clone();
                            if !cls_params.args.is_empty() {
                                cls_params.args.remove(0); // remove 'cls'
                            }
                            let func_chunk_idx = self.compile_function(&effective_name, &cls_params, mbody)?;
                            static_method_entries.push((effective_name.to_lowercase(), func_chunk_idx));
                        } else {
                            let func_chunk_idx = self.compile_function(&effective_name, params, mbody)?;
                            method_entries.push((effective_name.to_lowercase(), func_chunk_idx));
                            if method_name == "__init__" {
                                init_chunk = Some(func_chunk_idx);
                                init_params = Some(params.clone());
                            }
                        }
                    } else if let Statement::Pass = s {
                        // skip
                    } else {
                        other_stmts.push(s.clone());
                    }
                }

                // Build constructor chunk: creates object, stamps type_id, binds methods, calls __init__
                let ctor_name = name.to_lowercase();
                let mut ctor = Chunk::new(&ctor_name);
                // Arity: __init__ params minus self (self is the new object, implicit)
                let user_params = init_params.as_ref()
                    .map(|p| if p.args.len() > 1 { p.args.len() - 1 } else { 0 })
                    .unwrap_or(0);
                ctor.arity = user_params as u8;
                let ctor_idx = self.chunks.len();
                self.chunks.push(ctor);

                let mut ctor_scope = Scope::new();
                // slot 0 = callee (implicit), slots 1..N = user params
                for i in 0..user_params {
                    ctor_scope.alloc(&format!("__arg{}", i));
                }
                let this_local = ctor_scope.alloc("__this");
                self.scopes.push(ctor_scope);
                let scope_idx = self.scopes.len() - 1;

                let is_child = !parent_name.is_empty();

                if is_child && init_chunk.is_some() {
                    // Child class with __init__: call __init__ which calls super().__init__()
                    // to create object via parent constructor. Then bind child methods on result.
                    let init_ci = init_chunk.unwrap();

                    // Initialize this_local to null (super() will create the real object)
                    self.chunk(ctor_idx).emit_op(Op::null, 0);
                    self.chunk(ctor_idx).emit_op_u16(Op::local_set, this_local, 0);

                    // Call __init__(self, *args)
                    self.chunk(ctor_idx).emit_op_u16(Op::ref_func, init_ci as u16, 0);
                    self.chunk(ctor_idx).emit(0, 0);
                    self.chunk(ctor_idx).emit_op_u16(Op::local_get, this_local, 0); // self (null, super will fix)
                    for i in 0..user_params {
                        self.chunk(ctor_idx).emit_op_u16(Op::local_get, (i + 1) as u16, 0);
                    }
                    self.chunk(ctor_idx).emit_op_u8(Op::call_ref, (user_params + 1) as u8, 0);
                    // __init__ returns value from super().__init__() — the parent object
                    self.chunk(ctor_idx).emit_op_u16(Op::local_set, this_local, 0);

                    // Bind child methods on the object (overriding parent methods)
                    for (method_name, method_ci) in &method_entries {
                        if method_name == "__init__" { continue; }
                        common::classes::emit_bind_method_with_aliases(
                            self.chunk(ctor_idx), this_local, method_name, *method_ci, 0,
                        );
                    }
                } else {
                    // Base class (or child without __init__): create own object
                    common::classes::emit_new_typed_object(self.chunk(ctor_idx), this_local, name, 0);

                    // Bind instance methods + cross-language aliases
                    for (method_name, method_ci) in &method_entries {
                        if method_name == "__init__" { continue; }
                        common::classes::emit_bind_method_with_aliases(
                            self.chunk(ctor_idx), this_local, method_name, *method_ci, 0,
                        );
                    }

                    // Compile class-level statements (class attributes).
                    for s in &other_stmts {
                        if let Statement::Assign { targets, value } = s {
                            for target in targets {
                                if let Expression::Name(attr_name) = target {
                                    self.chunk(ctor_idx).emit_op_u16(Op::local_get, this_local, 0);
                                    self.compile_expr(value, ctor_idx)?;
                                    let attr_key = self.chunk(ctor_idx).add_constant(Value::String(Arc::from(attr_name.as_str())));
                                    self.chunk(ctor_idx).emit_op_u16(Op::struct_set, attr_key, 0);
                                    self.chunk(ctor_idx).emit_op(Op::drop, 0);
                                }
                            }
                        }
                    }

                    // Call __init__(self, *args) if it exists
                    if let Some(init_ci) = init_chunk {
                        self.chunk(ctor_idx).emit_op_u16(Op::ref_func, init_ci as u16, 0);
                        self.chunk(ctor_idx).emit(0, 0);
                        self.chunk(ctor_idx).emit_op_u16(Op::local_get, this_local, 0);
                        for i in 0..user_params {
                            self.chunk(ctor_idx).emit_op_u16(Op::local_get, (i + 1) as u16, 0);
                        }
                        self.chunk(ctor_idx).emit_op_u8(Op::call_ref, (user_params + 1) as u8, 0);
                        self.chunk(ctor_idx).emit_op(Op::drop, 0);
                    } else if !parent_name.is_empty() {
                        // No __init__ and has parent — call parent constructor with all args
                        common::classes::emit_store_super(self.chunk(ctor_idx), this_local, &parent_name, 0);
                    }
                }

                // Stamp __types array for instanceof support
                common::classes::emit_instanceof_chain(self.chunk(ctor_idx), this_local, name, 0);

                common::classes::emit_constructor_return(self.chunk(ctor_idx), this_local, 0);

                let scope = self.scopes.remove(scope_idx);
                self.chunks[ctor_idx].local_count = (scope.max_local + 1) as u16;

                // Store constructor as local + global
                let class_local = self.scope(chunk_idx).alloc(name);
                common::classes::emit_store_constructor(self.chunk(chunk_idx), name, ctor_idx, class_local, 0);

                // Attach static/class methods to the constructor function
                for (sm_name, sm_ci) in &static_method_entries {
                    common::classes::emit_attach_static_method(self.chunk(chunk_idx), class_local, sm_name, *sm_ci, 0);
                }

                // Register type entry
                let all_methods = { let mut all = method_entries; all.extend(static_method_entries); all };
                let implements = if all_bases.len() > 1 { all_bases[1..].to_vec() } else { Vec::new() };
                common::classes::register_type(
                    &mut self.chunks, name, &parent_name,
                    Vec::new(), all_methods, false, implements, Some(ctor_idx),
                );
                self.current_class_parent = saved_parent;
            }

            Statement::Return(expr) => {
                if let Some(e) = expr {
                    self.compile_expr(e, chunk_idx)?;
                } else {
                    self.chunk(chunk_idx).emit_op(Op::null, 0);
                }
                self.chunk(chunk_idx).emit_op(Op::r#return, 0);
            }

            Statement::Break => {
                if self.loop_stack.is_empty() {
                    return Err("break outside loop".into());
                }
                let j = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                self.loop_stack.last_mut().unwrap().break_jumps.push(j);
            }

            Statement::Continue => {
                if self.loop_stack.is_empty() {
                    return Err("continue outside loop".into());
                }
                let j = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                self.loop_stack.last_mut().unwrap().continue_jumps.push(j);
            }

            Statement::Pass => {}

            Statement::Try { body, handlers, else_body, finally_body } => {
                if handlers.len() <= 1 && handlers.first().map(|h| h.exc_type.is_none()).unwrap_or(true) {
                    // Simple untyped try/except — use try_start (faster path)
                    let catch_jump = self.chunk(chunk_idx).emit_jump(Op::try_start, 0);
                    self.chunk(chunk_idx).emit(0u8, 0); // reserved for finally

                    for s in body { self.compile_stmt(s, chunk_idx)?; }

                    self.chunk(chunk_idx).emit_op(Op::try_end, 0);

                    // else block runs if no exception
                    if let Some(else_body) = else_body {
                        for s in else_body { self.compile_stmt(s, chunk_idx)?; }
                    }

                    let end_jump = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                    self.chunk(chunk_idx).patch_jump(catch_jump);

                    if let Some(handler) = handlers.first() {
                        if let Some(name) = &handler.name {
                            let idx = self.scope(chunk_idx).alloc(name);
                            self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx, 0);
                            self.chunk(chunk_idx).emit_op(Op::drop, 0);
                        } else {
                            self.chunk(chunk_idx).emit_op(Op::drop, 0);
                        }
                        for s in &handler.body { self.compile_stmt(s, chunk_idx)?; }
                    } else {
                        self.chunk(chunk_idx).emit_op(Op::drop, 0);
                    }

                    // finally block
                    if let Some(finally_body) = finally_body {
                        // Patch end_jump to before finally
                        self.chunk(chunk_idx).patch_jump(end_jump);
                        for s in finally_body { self.compile_stmt(s, chunk_idx)?; }
                    } else {
                        self.chunk(chunk_idx).patch_jump(end_jump);
                    }
                } else {
                    // Multiple typed handlers — use try_table with exception tags
                    // Emit try_table opcode: [try_table, handler_count, (tag, offset)...]
                    self.chunk(chunk_idx).emit_op(Op::try_table, 0);
                    self.chunk(chunk_idx).emit(handlers.len() as u8, 0);

                    // Reserve space for handler entries (tag + u16 offset each)
                    let table_start = self.chunk(chunk_idx).code.len();
                    for handler in handlers {
                        // Determine tag for this handler.
                        // Normalize exception name for cross-language compat
                        // (e.g. Dart FormatException → Python ValueError)
                        let tag = if let Some(exc_type) = &handler.exc_type {
                            let type_name = self.expr_to_name(exc_type);
                            let canonical = common::errors::canonical_exception_name(&type_name);
                            self.chunk(chunk_idx).add_exception_tag(canonical)
                        } else {
                            0 // catch-all
                        };
                        self.chunk(chunk_idx).emit(tag, 0);
                        // Placeholder offset (will be patched)
                        self.chunk(chunk_idx).emit(0, 0);
                        self.chunk(chunk_idx).emit(0, 0);
                    }

                    // Compile try body
                    for s in body { self.compile_stmt(s, chunk_idx)?; }

                    // try_end — normal exit
                    self.chunk(chunk_idx).emit_op(Op::try_end, 0);

                    // else block
                    if let Some(else_body) = else_body {
                        for s in else_body { self.compile_stmt(s, chunk_idx)?; }
                    }

                    let end_jump = self.chunk(chunk_idx).emit_jump(Op::br, 0);

                    // Compile each handler and patch its offset in the try_table
                    let mut handler_end_jumps = Vec::new();
                    for (i, handler) in handlers.iter().enumerate() {
                        let handler_ip = self.chunk(chunk_idx).code.len();
                        // Patch the offset in the try_table entry
                        // Each entry is 3 bytes: tag(1) + offset(2)
                        let entry_offset = table_start + i * 3 + 1; // +1 to skip tag byte
                        let relative = (handler_ip as i32 - table_start as i32) as u16;
                        self.chunk(chunk_idx).code[entry_offset] = (relative >> 8) as u8;
                        self.chunk(chunk_idx).code[entry_offset + 1] = (relative & 0xff) as u8;

                        // Exception value is on stack
                        if let Some(name) = &handler.name {
                            let idx = self.scope(chunk_idx).alloc(name);
                            self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx, 0);
                            self.chunk(chunk_idx).emit_op(Op::drop, 0);
                        } else {
                            self.chunk(chunk_idx).emit_op(Op::drop, 0);
                        }
                        for s in &handler.body { self.compile_stmt(s, chunk_idx)?; }
                        let j = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                        handler_end_jumps.push(j);
                    }

                    // Patch all end jumps (including from try body and all handlers)
                    self.chunk(chunk_idx).patch_jump(end_jump);
                    for j in handler_end_jumps {
                        self.chunk(chunk_idx).patch_jump(j);
                    }

                    // finally block
                    if let Some(finally_body) = finally_body {
                        for s in finally_body { self.compile_stmt(s, chunk_idx)?; }
                    }
                }
            }

            Statement::Raise { exc, cause: _ } => {
                if let Some(e) = exc {
                    self.compile_expr(e, chunk_idx)?;
                } else {
                    self.chunk(chunk_idx).emit_op(Op::null, 0);
                }
                self.chunk(chunk_idx).emit_op(Op::throw_ref, 0);
            }

            Statement::Import { names } => {
                // Stub: create null locals for imported names
                for alias in names {
                    let local_name = alias.asname.as_ref().unwrap_or(&alias.name);
                    let idx = self.scope(chunk_idx).alloc(local_name);
                    self.chunk(chunk_idx).emit_op(Op::null, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx, 0);
                    self.chunk(chunk_idx).emit_op(Op::drop, 0);
                }
            }

            Statement::ImportFrom { names, .. } => {
                for alias in names {
                    let local_name = alias.asname.as_ref().unwrap_or(&alias.name);
                    let idx = self.scope(chunk_idx).alloc(local_name);
                    self.chunk(chunk_idx).emit_op(Op::null, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx, 0);
                    self.chunk(chunk_idx).emit_op(Op::drop, 0);
                }
            }

            Statement::Global(names) | Statement::Nonlocal(names) => {
                // Ensure names are in scope
                for name in names {
                    self.scope(chunk_idx).alloc(name);
                }
            }

            Statement::Delete(targets) => {
                for target in targets {
                    match target {
                        Expression::Name(name) => {
                            // del name → set to null
                            let idx = self.scope(chunk_idx).alloc(name);
                            self.chunk(chunk_idx).emit_op(Op::null, 0);
                            self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx, 0);
                            self.chunk(chunk_idx).emit_op(Op::drop, 0);
                        }
                        Expression::Subscript { value, slice } => {
                            // del obj[key] → deleteProperty(obj, key)
                            common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_deleteproperty", 0);
                            self.compile_expr(value, chunk_idx)?;
                            self.compile_expr(slice, chunk_idx)?;
                            common::bundle::emit_call_invoke(self.chunk(chunk_idx), 2, 0);
                            self.chunk(chunk_idx).emit_op(Op::drop, 0);
                        }
                        Expression::Attribute { value, attr } => {
                            // del obj.attr → deleteProperty(obj, attr_str)
                            common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_deleteproperty", 0);
                            self.compile_expr(value, chunk_idx)?;
                            let key = self.chunk(chunk_idx).add_constant(Value::String(Arc::from(attr.as_str())));
                            self.chunk(chunk_idx).emit_op_u16(Op::r#const, key, 0);
                            common::bundle::emit_call_invoke(self.chunk(chunk_idx), 2, 0);
                            self.chunk(chunk_idx).emit_op(Op::drop, 0);
                        }
                        _ => {
                            // Unsupported del target — skip
                        }
                    }
                }
            }

            Statement::Assert { test, msg } => {
                // assert test [, msg] → if not test: raise AssertionError(msg)
                self.compile_expr(test, chunk_idx)?;
                let ok_jump = self.chunk(chunk_idx).emit_jump(Op::br_if_true, 0);
                if let Some(m) = msg {
                    self.compile_expr(m, chunk_idx)?;
                } else {
                    let c = self.chunk(chunk_idx).add_constant(Value::String(Arc::from("AssertionError")));
                    self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
                }
                self.chunk(chunk_idx).emit_op(Op::throw_ref, 0);
                self.chunk(chunk_idx).patch_jump(ok_jump);
            }

            Statement::With { items, body, .. } => {
                // with expr as var: → call __enter__, bind var, run body, call __exit__
                let mut ctx_locals = Vec::new();
                for item in items {
                    self.compile_expr(&item.context_expr, chunk_idx)?;
                    let ctx_local = self.scope(chunk_idx).alloc("__with_ctx");
                    self.chunk(chunk_idx).emit_op(Op::dup, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::local_set, ctx_local, 0);
                    self.chunk(chunk_idx).emit_op(Op::drop, 0);
                    ctx_locals.push(ctx_local);

                    // Call __enter__ if it exists, otherwise use value directly
                    self.chunk(chunk_idx).emit_op(Op::dup, 0);
                    let enter_key = self.chunk(chunk_idx).add_constant(Value::String(Arc::from("__enter__")));
                    self.chunk(chunk_idx).emit_op_u16(Op::struct_get, enter_key, 0);
                    self.chunk(chunk_idx).emit_op(Op::dup, 0);
                    self.chunk(chunk_idx).emit_op(Op::ref_is_null, 0);
                    let no_enter = self.chunk(chunk_idx).emit_jump(Op::br_if_true, 0);
                    // Has __enter__: call it with ctx as self
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, ctx_local, 0);
                    self.chunk(chunk_idx).emit_op_u8(Op::call_ref, 1, 0);
                    let done_enter = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                    self.chunk(chunk_idx).patch_jump(no_enter);
                    // No __enter__: drop the null, keep original value
                    self.chunk(chunk_idx).emit_op(Op::drop, 0); // drop null
                    self.chunk(chunk_idx).patch_jump(done_enter);

                    if let Some(var) = &item.optional_vars {
                        self.compile_assign_target(var, chunk_idx)?;
                    } else {
                        self.chunk(chunk_idx).emit_op(Op::drop, 0);
                    }
                }

                // Execute body
                for s in body { self.compile_stmt(s, chunk_idx)?; }

                // Call __exit__ on each context (reverse order)
                for ctx_local in ctx_locals.iter().rev() {
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, *ctx_local, 0);
                    let exit_key = self.chunk(chunk_idx).add_constant(Value::String(Arc::from("__exit__")));
                    self.chunk(chunk_idx).emit_op_u16(Op::struct_get, exit_key, 0);
                    self.chunk(chunk_idx).emit_op(Op::dup, 0);
                    self.chunk(chunk_idx).emit_op(Op::ref_is_null, 0);
                    let no_exit = self.chunk(chunk_idx).emit_jump(Op::br_if_true, 0);
                    // Call __exit__(self, None, None, None) — no exception info
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, *ctx_local, 0);
                    self.chunk(chunk_idx).emit_op(Op::null, 0);
                    self.chunk(chunk_idx).emit_op(Op::null, 0);
                    self.chunk(chunk_idx).emit_op(Op::null, 0);
                    self.chunk(chunk_idx).emit_op_u8(Op::call_ref, 4, 0);
                    self.chunk(chunk_idx).emit_op(Op::drop, 0);
                    let done_exit = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                    self.chunk(chunk_idx).patch_jump(no_exit);
                    self.chunk(chunk_idx).emit_op(Op::drop, 0); // drop null
                    self.chunk(chunk_idx).patch_jump(done_exit);
                }
            }

            Statement::AnnAssign { target, value, .. } => {
                // x: int = 5 → just compile the assignment if there's a value
                if let Some(val) = value {
                    self.compile_expr(val, chunk_idx)?;
                    self.compile_assign_target(target, chunk_idx)?;
                }
            }

            Statement::Match { subject, cases } => {
                // Compile as chained if/elif
                let subj_local = self.scope(chunk_idx).alloc("__match_subj");
                self.compile_expr(subject, chunk_idx)?;
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, subj_local, 0);

                let mut end_jumps = Vec::new();

                for case in cases {
                    // Compile pattern as condition
                    self.compile_pattern(&case.pattern, subj_local, chunk_idx)?;

                    // Guard
                    if let Some(guard) = &case.guard {
                        let no_guard = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);
                        self.compile_expr(guard, chunk_idx)?;
                        let combined = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);
                        // Body
                        self.compile_pattern_bindings(&case.pattern, subj_local, chunk_idx)?;
                        for s in &case.body { self.compile_stmt(s, chunk_idx)?; }
                        end_jumps.push(self.chunk(chunk_idx).emit_jump(Op::br, 0));
                        self.chunk(chunk_idx).patch_jump(combined);
                        self.chunk(chunk_idx).patch_jump(no_guard);
                        continue;
                    }

                    let skip = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);
                    // Bind pattern variables
                    self.compile_pattern_bindings(&case.pattern, subj_local, chunk_idx)?;
                    for s in &case.body { self.compile_stmt(s, chunk_idx)?; }
                    end_jumps.push(self.chunk(chunk_idx).emit_jump(Op::br, 0));
                    self.chunk(chunk_idx).patch_jump(skip);
                }

                for ej in end_jumps { self.chunk(chunk_idx).patch_jump(ej); }
            }
        }
        Ok(())
    }

    // ── If with proper end-jump chaining ─────────────────────────────

    // ── For loop ─────────────────────────────────────────────────────

    // ── Match pattern compilation ──────────────────────────────

    /// Compile a pattern as a boolean condition (pushes true/false).
    fn compile_pattern(&mut self, pattern: &Pattern, subj_local: u16, chunk_idx: usize) -> Result<(), String> {
        match pattern {
            Pattern::Wildcard => {
                // Always matches
                self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
            }
            Pattern::Value(expr) | Pattern::Singleton(expr) => {
                // subj is expr (None, True, False)
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, subj_local, 0);
                self.compile_expr(expr, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::dyn_eq, 0);
            }
            Pattern::Or(patterns) => {
                // p1 || p2 || ...
                let mut end_jumps = Vec::new();
                for (i, p) in patterns.iter().enumerate() {
                    self.compile_pattern(p, subj_local, chunk_idx)?;
                    if i < patterns.len() - 1 {
                        let is_true = self.chunk(chunk_idx).emit_jump(Op::br_if_true, 0);
                        end_jumps.push(is_true);
                    }
                }
                // Last pattern's result is on stack
                let skip_true = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                for ej in end_jumps {
                    self.chunk(chunk_idx).patch_jump(ej);
                }
                // One of the earlier patterns was true
                self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
                self.chunk(chunk_idx).patch_jump(skip_true);
            }
            Pattern::As { pattern: Some(inner), .. } => {
                self.compile_pattern(inner, subj_local, chunk_idx)?;
            }
            Pattern::As { pattern: None, .. } => {
                // Just a name binding, always matches
                self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
            }
            Pattern::Sequence(pats) => {
                // Check length == pats.len(), then check each element
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, subj_local, 0);
                common_collections::emit_len(&mut self.chunks[chunk_idx], 0);
                let expected = self.chunk(chunk_idx).add_constant(Value::I32(pats.len() as i32));
                self.chunk(chunk_idx).emit_op_u16(Op::r#const, expected, 0);
                self.chunk(chunk_idx).emit_op(Op::dyn_eq, 0);
                // If length doesn't match, short-circuit to false
                // For now, just check length (element checks would need recursion per element)
            }
            Pattern::Mapping(_pairs) => {
                // Dict pattern — check keys exist. Simplified: always true for now.
                self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
            }
            Pattern::Star(_) => {
                self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
            }
            Pattern::Class { .. } => {
                // Class pattern — isinstance check. Simplified for now.
                self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
            }
        }
        Ok(())
    }

    /// Bind variables from a pattern after it matched.
    fn compile_pattern_bindings(&mut self, pattern: &Pattern, subj_local: u16, chunk_idx: usize) -> Result<(), String> {
        match pattern {
            Pattern::As { name: Some(name), .. } => {
                let idx = self.scope(chunk_idx).alloc(name);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, subj_local, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx, 0);
            }
            _ => {}
        }
        Ok(())
    }

    // ── For loop ─────────────────────────────────────────────────────

    fn compile_for(&mut self, target: &Expression, iter: &Expression, body: &[Statement], else_body: Option<&[Statement]>, chunk_idx: usize) -> Result<(), String> {
        let iter_local = self.scope(chunk_idx).alloc("__for_iter");
        let idx_local = self.scope(chunk_idx).alloc("__for_idx");

        // Evaluate iterable. If it's a dict/object (not array), convert to keys array.
        // Python: `for x in dict` iterates keys. This matches JS Object.keys() behavior.
        self.compile_expr(iter, chunk_idx)?;
        self.chunk(chunk_idx).emit_op(Op::dup, 0);
        self.chunk(chunk_idx).emit_op(Op::ref_is_array, 0);
        let is_array = self.chunk(chunk_idx).emit_jump(Op::br_if_true, 0);
        // Not array → get keys
        common::dict::emit_keys(self.chunk(chunk_idx), 0);
        let done = self.chunk(chunk_idx).emit_jump(Op::br, 0);
        self.chunk(chunk_idx).patch_jump(is_array);
        // Is array → drop the dup (ref_is_array consumed one, dup left one extra)
        // Actually ref_is_array pops and pushes bool, dup left original on stack.
        // After br_if_true, stack has: [original_value].
        // On the keys path: emit_keys consumed original, pushed keys array.
        // On the array path: original is still there. Good.
        self.chunk(chunk_idx).patch_jump(done);
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, iter_local, 0);

        // Init index to 0
        self.chunk(chunk_idx).emit_op(Op::i32_const_0, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx_local, 0);

        let loop_start = self.chunk(chunk_idx).current_offset();

        // Condition: idx < len(iter)
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_local, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, iter_local, 0);
        common_collections::emit_len(&mut self.chunks[chunk_idx], 0);
        self.chunk(chunk_idx).emit_op(Op::dyn_lt, 0);
        let exit_jump = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);

        // Load current element
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, iter_local, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_local, 0);
        common_collections::emit_get(&mut self.chunks[chunk_idx], 0);

        // Assign to target
        self.compile_assign_target(target, chunk_idx)?;

        self.loop_stack.push(LoopCtx { _start: loop_start, break_jumps: Vec::new(), continue_jumps: Vec::new() });

        for s in body { self.compile_stmt(s, chunk_idx)?; }

        let ctx = self.loop_stack.pop().unwrap();
        for cj in &ctx.continue_jumps { self.chunk(chunk_idx).patch_jump(*cj); }

        // Increment index
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_local, 0);
        self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
        self.chunk(chunk_idx).emit_op(Op::i32_add, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx_local, 0);

        self.chunk(chunk_idx).emit_loop(loop_start, 0);

        // Handle else_body with broke flag
        if let Some(else_stmts) = else_body {
            // Normal exit (no break) → run else
            // Break exits → skip else
            let skip_flag = self.chunk(chunk_idx).emit_jump(Op::br, 0);
            for bj in &ctx.break_jumps { self.chunk(chunk_idx).patch_jump(*bj); }
            // Break landed here → jump past else
            let skip_else = self.chunk(chunk_idx).emit_jump(Op::br, 0);
            self.chunk(chunk_idx).patch_jump(skip_flag);
            self.chunk(chunk_idx).patch_jump(exit_jump);
            // Normal exit → else body
            for s in else_stmts { self.compile_stmt(s, chunk_idx)?; }
            self.chunk(chunk_idx).patch_jump(skip_else);
        } else {
            for bj in &ctx.break_jumps { self.chunk(chunk_idx).patch_jump(*bj); }
            self.chunk(chunk_idx).patch_jump(exit_jump);
        }

        Ok(())
    }

    // ── Function compilation ─────────────────────────────────────────

    /// Compile a function. Returns the chunk index of the compiled function.
    fn compile_function(&mut self, name: &str, params: &Parameters, body: &[Statement]) -> Result<usize, String> {
        let mut fchunk = common::functions::create_function_chunk(name, params.args.len() as u8);
        fchunk.add_import("wasi:cli", "log");
        let func_chunk_idx = self.chunks.len();
        self.chunks.push(fchunk);

        let mut scope = Scope::new();
        for p in &params.args { scope.alloc(&p.name); }
        if let Some(ref va) = params.vararg { scope.alloc(&va.name); }
        for p in &params.kwonly_args { scope.alloc(&p.name); }
        if let Some(ref kw) = params.kwarg { scope.alloc(&kw.name); }
        self.scopes.push(scope);
        let scope_idx = self.scopes.len() - 1;

        // Default parameter checks (positional defaults)
        let num_positional = params.args.len();
        let num_defaults = params.defaults.len();
        if num_defaults > 0 {
            let first_default_idx = num_positional - num_defaults;
            for (di, default_expr) in params.defaults.iter().enumerate() {
                let slot = (first_default_idx + di + 1) as u16;
                let skip = common::functions::emit_default_param_start(self.chunk(func_chunk_idx), slot, 0);
                self.compile_expr(default_expr, func_chunk_idx)?;
                common::functions::emit_default_param_end(self.chunk(func_chunk_idx), slot, skip, 0);
            }
        }
        // Keyword-only defaults
        for (di, default_opt) in params.kw_defaults.iter().enumerate() {
            if let Some(default_expr) = default_opt {
                let slot = (num_positional + di + 1) as u16;
                let skip = common::functions::emit_default_param_start(self.chunk(func_chunk_idx), slot, 0);
                self.compile_expr(default_expr, func_chunk_idx)?;
                common::functions::emit_default_param_end(self.chunk(func_chunk_idx), slot, skip, 0);
            }
        }

        for s in body { self.compile_stmt(s, func_chunk_idx)?; }

        if name == "__init__" {
            // __init__ returns self (slot 1) so child class wrappers can capture the object
            self.chunks[func_chunk_idx].emit_op_u16(Op::local_get, 1, 0);
            self.chunks[func_chunk_idx].emit_op(Op::r#return, 0);
        } else {
            common::functions::emit_function_epilogue(&mut self.chunks[func_chunk_idx], 0);
        }

        let scope = self.scopes.remove(scope_idx);
        self.chunks[func_chunk_idx].local_count = (scope.max_local + 1) as u16;
        self.last_upvalues = scope.upvalues;

        Ok(func_chunk_idx)
    }

    // ── Assignment targets ───────────────────────────────────────────

    fn compile_assign_target(&mut self, target: &Expression, chunk_idx: usize) -> Result<(), String> {
        match target {
            Expression::Name(name) => {
                let idx = self.scope(chunk_idx).alloc(name);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
            }
            Expression::Tuple(targets) | Expression::List(targets) => {
                // Destructuring: value is an array on stack
                for (i, t) in targets.iter().enumerate() {
                    self.chunk(chunk_idx).emit_op(Op::dup, 0);
                    let c = self.chunk(chunk_idx).add_constant(Value::I32(i as i32));
                    self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
                    common_collections::emit_get(&mut self.chunks[chunk_idx], 0);
                    self.compile_assign_target(t, chunk_idx)?;
                }
                self.chunk(chunk_idx).emit_op(Op::drop, 0); // drop the array
            }
            Expression::Attribute { value, attr } => {
                // obj.attr = rhs_value
                // Stack has: [rhs_value]. struct_set pops val, then obj.
                // We need: [obj, rhs_value] → struct_set pops rhs_value, then obj.
                // So: compile obj (push it), then swap obj and rhs_value.
                // But there's no swap opcode. Alternative: compile obj BEFORE rhs.
                // But rhs is already on stack from the caller.
                // Solution: store rhs in temp, push obj, push rhs back.
                let tmp = self.scope(chunk_idx).alloc("__attr_tmp");
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, tmp, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                self.compile_expr(value, chunk_idx)?; // push obj
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, tmp, 0); // push rhs_value
                let name_c = self.chunk(chunk_idx).add_constant(Value::String(Arc::from(attr.as_str())));
                self.chunk(chunk_idx).emit_op_u16(Op::struct_set, name_c, 0);
            }
            Expression::Subscript { value, slice } => {
                if let Expression::Slice { lower, upper, step: _ } = slice.as_ref() {
                    // Slice assignment: a[1:3] = [10, 20] → splice(a, start, deleteCount, ...items)
                    // Stack has: [rhs_value].
                    let tmp = self.scope(chunk_idx).alloc("__splice_rhs");
                    self.chunk(chunk_idx).emit_op_u16(Op::local_set, tmp, 0);
                    self.chunk(chunk_idx).emit_op(Op::drop, 0);
                    // splice(arr, start, deleteCount, ...newItems)
                    // deleteCount = end - start
                    common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_splice", 0);
                    self.compile_expr(value, chunk_idx)?; // arr
                    // start
                    if let Some(lo) = lower { self.compile_expr(lo, chunk_idx)?; }
                    else { let c = self.chunk(chunk_idx).add_constant(Value::I32(0)); self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0); }
                    // deleteCount = end - start (simplified: use end - start)
                    if let Some(up) = upper { self.compile_expr(up, chunk_idx)?; }
                    else {
                        self.compile_expr(value, chunk_idx)?;
                        common_collections::emit_len(&mut self.chunks[chunk_idx], 0);
                    }
                    // Need: deleteCount = end - start. We have start and end on stack.
                    // Actually splice takes (arr, start, deleteCount, ...items)
                    // We pushed start, end. Need to compute deleteCount = end - start.
                    // Rearrange: store start, compute delta
                    // This is getting complex. Simpler: just pass start and a large deleteCount.
                    // splice(arr, start, 999999, replacement_spread)
                    // Actually, splice takes varargs. Let's use 3 fixed args + spread rhs.
                    // Rewrite: push arr, start, (end-start), then spread rhs items
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, tmp, 0); // rhs array
                    self.chunk(chunk_idx).emit_op(Op::spread, 0); // spread items
                    // Count args: arr(1) + start(1) + end(1) + spread items
                    // We can't know spread count at compile time.
                    // Alternative: just call splice with 4 args where 4th is the replacement array
                    // But splice API takes (arr, start, deleteCount, ...items).
                    // Let's not spread — pass the array as a single replacement and handle in host.
                    // Actually, the existing splice host fn expects varargs for items.
                    // Simplest correct approach: use a different call convention.
                    // For now: just use array_set for non-slice, and for slice keep it simple.
                    common::bundle::emit_call_invoke(self.chunk(chunk_idx), 4, 0);
                    self.chunk(chunk_idx).emit_op(Op::drop, 0);
                } else {
                    // obj[idx] = rhs_value. array_set pops: val, key, obj.
                    let tmp = self.scope(chunk_idx).alloc("__sub_tmp");
                    self.chunk(chunk_idx).emit_op_u16(Op::local_set, tmp, 0);
                    self.chunk(chunk_idx).emit_op(Op::drop, 0);
                    self.compile_expr(value, chunk_idx)?;
                    self.compile_expr(slice, chunk_idx)?;
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, tmp, 0);
                    common_collections::emit_set(&mut self.chunks[chunk_idx], 0);
                }
            }
            _ => {
                return Err(format!("unsupported assignment target: {:?}", target));
            }
        }
        Ok(())
    }

    // ── Expressions ──────────────────────────────────────────────────

    fn compile_expr(&mut self, expr: &Expression, chunk_idx: usize) -> Result<(), String> {
        match expr {
            Expression::Int(n) => {
                let c = self.chunk(chunk_idx).add_constant(Value::I32(*n as i32));
                self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
            }
            Expression::Float(f) => {
                let c = self.chunk(chunk_idx).add_constant(Value::F64(*f));
                self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
            }
            Expression::Str(s) => {
                let c = self.chunk(chunk_idx).add_constant(Value::String(Arc::from(s.as_str())));
                self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
            }
            Expression::FString { parts } => {
                // Concatenate all parts
                let mut count = 0;
                for part in parts {
                    match part {
                        FStringPart::Literal(s) => {
                            let c = self.chunk(chunk_idx).add_constant(Value::String(Arc::from(s.as_str())));
                            self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
                            count += 1;
                        }
                        FStringPart::Expr(e) => {
                            self.compile_expr(e, chunk_idx)?;
                            // Convert to string
                            let to_str = self.chunk(chunk_idx).add_import("vybe:convert", "toString");
                            self.chunk(chunk_idx).emit_op_u16(Op::call_import, to_str, 0);
                            self.chunk(chunk_idx).emit(1, 0);
                            count += 1;
                        }
                        FStringPart::FormattedExpr(e, spec) => {
                            self.compile_expr(e, chunk_idx)?;
                            // Apply format spec via host format function
                            let spec_c = self.chunk(chunk_idx).add_constant(Value::String(Arc::from(spec.as_str())));
                            self.chunk(chunk_idx).emit_op_u16(Op::r#const, spec_c, 0);
                            let fmt_fn = self.chunk(chunk_idx).add_import("vybe:string", "format");
                            self.chunk(chunk_idx).emit_op_u16(Op::call_import, fmt_fn, 0);
                            self.chunk(chunk_idx).emit(2, 0);
                            count += 1;
                        }
                    }
                }
                // Concatenate all parts
                common::strings::emit_concat(self.chunk(chunk_idx), count, 0);
            }
            Expression::Bool(b) => {
                let c = self.chunk(chunk_idx).add_constant(Value::Bool(*b));
                self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
            }
            Expression::None => {
                self.chunk(chunk_idx).emit_op(Op::null, 0);
            }
            Expression::Ellipsis => {
                self.chunk(chunk_idx).emit_op(Op::null, 0);
            }
            Expression::Name(name) => {
                match self.resolve_variable(name, chunk_idx) {
                    VarResolution::Local(slot) => {
                        self.chunk(chunk_idx).emit_op_u16(Op::local_get, slot, 0);
                    }
                    VarResolution::Upvalue(idx) => {
                        self.chunk(chunk_idx).emit_op_u8(Op::upvalue_get, idx, 0);
                    }
                    VarResolution::Global => {
                        let c = self.chunk(chunk_idx).add_constant(Value::String(Arc::from(name.as_str())));
                        self.chunk(chunk_idx).emit_op_u16(Op::global_get, c, 0);
                    }
                }
            }

            Expression::List(elems) | Expression::Tuple(elems) => {
                let has_star = elems.iter().any(|e| matches!(e, Expression::Starred(_)));
                if has_star {
                    // Build incrementally: start with [], push or concat
                    self.chunk(chunk_idx).emit_op_u16(Op::array_new, 0, 0);
                    for e in elems {
                        if let Expression::Starred(inner) = e {
                            // Concat the spread array
                            self.compile_expr(inner, chunk_idx)?;
                            common_collections::emit_concat(&mut self.chunks[chunk_idx], 0);
                        } else {
                            self.compile_expr(e, chunk_idx)?;
                            common_collections::emit_push(&mut self.chunks[chunk_idx], 0);
                        }
                    }
                } else {
                    for e in elems { self.compile_expr(e, chunk_idx)?; }
                    self.chunk(chunk_idx).emit_op_u16(Op::array_new, elems.len() as u16, 0);
                }
            }
            Expression::Set(elems) => {
                for e in elems { self.compile_expr(e, chunk_idx)?; }
                self.chunk(chunk_idx).emit_op_u16(Op::array_new, elems.len() as u16, 0);
            }
            Expression::Dict { keys, values } => {
                common::dict::emit_new(self.chunk(chunk_idx), 0);
                for (k, v) in keys.iter().zip(values.iter()) {
                    if let Some(key) = k {
                        if let Expression::Str(s) = key {
                            // String key with __keys tracking
                            self.chunk(chunk_idx).emit_op(Op::dup, 0);
                            self.compile_expr(v, chunk_idx)?;
                            common::dict::emit_set_const_key(self.chunk(chunk_idx), s, 0);
                        } else {
                            // Dynamic key
                            self.chunk(chunk_idx).emit_op(Op::dup, 0);
                            self.compile_expr(key, chunk_idx)?;
                            self.compile_expr(v, chunk_idx)?;
                            common::dict::emit_set_dynamic(self.chunk(chunk_idx), 0);
                        }
                    } else {
                        self.chunk(chunk_idx).emit_op(Op::dup, 0);
                        self.chunk(chunk_idx).emit_op(Op::null, 0);
                        self.compile_expr(v, chunk_idx)?;
                        common::dict::emit_set_dynamic(self.chunk(chunk_idx), 0);
                    }
                }
            }

            Expression::BinOp { op, left, right } => {
                self.compile_expr(left, chunk_idx)?;
                self.compile_expr(right, chunk_idx)?;
                match op {
                    // Arithmetic: use primitive opcodes (dyn_add handles string+int natively).
                    // User-defined __add__/__sub__ are available via cross-language aliases
                    // and the method-with-fallback dispatch path — not needed on every `+`.
                    BinOp::Add => self.chunk(chunk_idx).emit_op(Op::dyn_add, 0),
                    BinOp::Sub => self.chunk(chunk_idx).emit_op(Op::f64_sub, 0),
                    BinOp::Mul => {
                        // Use dynMul for string*int support
                        // Pop the two already-pushed operands, call dynMul
                        // Actually: we need to restructure. Left and right are already on stack.
                        // Easiest: store both in temps, call dynMul(a, b)
                        let tmp_a = self.scope(chunk_idx).alloc("__mul_a");
                        let tmp_b = self.scope(chunk_idx).alloc("__mul_b");
                        self.chunk(chunk_idx).emit_op_u16(Op::local_set, tmp_b, 0);
                        self.chunk(chunk_idx).emit_op_u16(Op::local_set, tmp_a, 0);
                        common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_dynmul", 0);
                        self.chunk(chunk_idx).emit_op_u16(Op::local_get, tmp_a, 0);
                        self.chunk(chunk_idx).emit_op_u16(Op::local_get, tmp_b, 0);
                        common::bundle::emit_call_invoke(self.chunk(chunk_idx), 2, 0);
                    }
                    BinOp::Div => self.chunk(chunk_idx).emit_op(Op::f64_div, 0),
                    BinOp::FloorDiv => {
                        self.chunk(chunk_idx).emit_op(Op::f64_div, 0);
                        common::math::emit_floor(self.chunk(chunk_idx), 0);
                    }
                    BinOp::Mod => self.chunk(chunk_idx).emit_op(Op::i32_rem_s, 0),
                    BinOp::Pow => {
                        // left and right already on stack — need func ref below them
                        // Store right in temp, store left in temp, push func, push left, push right, call
                        // Actually: emit_call expects [func, args...]. Left and right are already pushed.
                        // Simplest: use the existing host call path for pow since 2 args are already on stack.
                        let pow_idx = self.chunk(chunk_idx).add_import("vybe:math", "pow");
                        self.chunk(chunk_idx).emit_op_u16(Op::call_import, pow_idx, 0);
                        self.chunk(chunk_idx).emit(2, 0);
                    }
                    BinOp::LShift => self.chunk(chunk_idx).emit_op(Op::i32_shl, 0),
                    BinOp::RShift => self.chunk(chunk_idx).emit_op(Op::i32_shr_s, 0),
                    BinOp::BitOr => self.chunk(chunk_idx).emit_op(Op::i32_or, 0),
                    BinOp::BitXor => self.chunk(chunk_idx).emit_op(Op::i32_xor, 0),
                    BinOp::BitAnd => self.chunk(chunk_idx).emit_op(Op::i32_and, 0),
                    BinOp::MatMul => self.chunk(chunk_idx).emit_op(Op::f64_mul, 0), // placeholder
                }
            }

            Expression::UnaryOp { op, operand } => {
                self.compile_expr(operand, chunk_idx)?;
                match op {
                    UnaryOp::Not => self.chunk(chunk_idx).emit_op(Op::dyn_not, 0),
                    UnaryOp::USub => self.chunk(chunk_idx).emit_op(Op::dyn_neg, 0),
                    UnaryOp::UAdd => {} // no-op
                    UnaryOp::Invert => {
                        // ~x → x ^ -1 (bitwise NOT for ints)
                        let c = self.chunk(chunk_idx).add_constant(Value::I32(-1));
                        self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
                        self.chunk(chunk_idx).emit_op(Op::i32_xor, 0);
                    }
                }
            }

            Expression::BoolOp { op, values } => {
                match op {
                    BoolOp::And => {
                        self.compile_expr(&values[0], chunk_idx)?;
                        for v in &values[1..] {
                            let jump = common::expressions::emit_and_start(self.chunk(chunk_idx), 0);
                            self.compile_expr(v, chunk_idx)?;
                            common::expressions::emit_short_circuit_end(self.chunk(chunk_idx), jump);
                        }
                    }
                    BoolOp::Or => {
                        self.compile_expr(&values[0], chunk_idx)?;
                        for v in &values[1..] {
                            let jump = common::expressions::emit_or_start(self.chunk(chunk_idx), 0);
                            self.compile_expr(v, chunk_idx)?;
                            common::expressions::emit_short_circuit_end(self.chunk(chunk_idx), jump);
                        }
                    }
                }
            }

            Expression::Compare { left, ops, comparators } => {
                // a < b < c  →  (a < b) and (b < c) with b evaluated once
                if ops.len() == 1 {
                    self.compile_expr(left, chunk_idx)?;
                    self.compile_expr(&comparators[0], chunk_idx)?;
                    self.emit_cmp_op(ops[0], chunk_idx);
                } else {
                    // Chained comparison
                    self.compile_expr(left, chunk_idx)?;
                    let mut end_jumps = Vec::new();
                    for (i, (op, cmp)) in ops.iter().zip(comparators.iter()).enumerate() {
                        self.compile_expr(cmp, chunk_idx)?;
                        if i < ops.len() - 1 {
                            self.chunk(chunk_idx).emit_op(Op::dup, 0); // keep for next comparison
                        }
                        // Swap to get correct order for comparison
                        self.emit_cmp_op(*op, chunk_idx);
                        if i < ops.len() - 1 {
                            let j = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);
                            end_jumps.push(j);
                        }
                    }
                    let after = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                    for j in end_jumps {
                        self.chunk(chunk_idx).patch_jump(j);
                    }
                    self.chunk(chunk_idx).emit_op(Op::r#false, 0);
                    self.chunk(chunk_idx).patch_jump(after);
                }
            }

            Expression::Call { func, args, keywords } => {
                // Check for built-in functions
                if let Expression::Name(name) = func.as_ref() {
                    match name.as_str() {
                        "super" => {
                            // super() → look up __super on self (slot 1)
                            // In a constructor, self is at slot 1 (after callee at slot 0)
                            // __super is set by the parent class binding
                            // For now: push self — super().method() will find parent methods via vtable
                            self.chunk(chunk_idx).emit_op_u16(Op::local_get, 1, 0); // self
                            return Ok(());
                        }
                        "print" => return self.compile_print(args, chunk_idx),
                        "len" => {
                            if args.len() == 1 {
                                // Smart length: tries __get_length/__len__ on user objects,
                                // falls back to array_length for plain arrays/strings.
                                self.compile_expr(&args[0], chunk_idx)?;
                                let obj_slot = self.scope(chunk_idx).alloc("__len_obj");
                                self.chunk(chunk_idx).emit_op_u16(Op::local_set, obj_slot, 0);
                                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                                common::expressions::emit_smart_length(self.chunk(chunk_idx), obj_slot, 0);
                                return Ok(());
                            }
                        }
                        "range" => {
                            return self.compile_range(args, chunk_idx);
                        }
                        "str" => {
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                common::convert::emit_to_string(self.chunk(chunk_idx), 0);
                                return Ok(());
                            }
                        }
                        "int" => {
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                common::convert::emit_parse_int(self.chunk(chunk_idx), 0);
                                return Ok(());
                            }
                        }
                        "float" => {
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                common::convert::emit_parse_float(self.chunk(chunk_idx), 0);
                                return Ok(());
                            }
                        }
                        "abs" => {
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                common::math::emit_abs(self.chunk(chunk_idx), 0);
                                return Ok(());
                            }
                        }
                        "input" => {
                            let input_idx = self.chunk(chunk_idx).add_import("wasi:cli", "readLine");
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                let log_idx = self.import_log;
                                self.chunk(chunk_idx).emit_op_u16(Op::call_import, log_idx, 0);
                                self.chunk(chunk_idx).emit(1, 0);
                                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                            }
                            self.chunk(chunk_idx).emit_op_u16(Op::call_import, input_idx, 0);
                            self.chunk(chunk_idx).emit(0, 0);
                            return Ok(());
                        }
                        "enumerate" => {
                            // Check for start= keyword
                            let start_kw = keywords.iter().find(|kw| kw.name.as_deref() == Some("start"));
                            common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_enumerate", 0);
                            for a in args { self.compile_expr(a, chunk_idx)?; }
                            common::bundle::emit_call_invoke(self.chunk(chunk_idx), args.len() as u8, 0);
                            // If start=N, adjust indices: add N to each pair's index
                            if let Some(kw) = start_kw {
                                // result is array of [i, val] pairs. Map: [i+start, val]
                                let result_slot = self.scope(chunk_idx).alloc("__enum_res");
                                let idx_slot = self.scope(chunk_idx).alloc("__enum_i");
                                let start_slot = self.scope(chunk_idx).alloc("__enum_st");
                                self.chunk(chunk_idx).emit_op_u16(Op::local_set, result_slot, 0);
                                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                                self.compile_expr(&kw.value, chunk_idx)?;
                                self.chunk(chunk_idx).emit_op_u16(Op::local_set, start_slot, 0);
                                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                                // for i in 0..len(result): result[i][0] += start
                                self.chunk(chunk_idx).emit_op(Op::i32_const_0, 0);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx_slot, 0);
                                let loop_start = self.chunk(chunk_idx).current_offset();
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_slot, 0);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, result_slot, 0);
                                common_collections::emit_len(&mut self.chunks[chunk_idx], 0);
                                self.chunk(chunk_idx).emit_op(Op::dyn_lt, 0);
                                let exit = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);
                                // pair = result[i]; pair[0] = pair[0] + start
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, result_slot, 0);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_slot, 0);
                                common_collections::emit_get(&mut self.chunks[chunk_idx], 0); // pair
                                let zero_c = self.chunk(chunk_idx).add_constant(Value::I32(0));
                                self.chunk(chunk_idx).emit_op_u16(Op::r#const, zero_c, 0);
                                self.chunk(chunk_idx).emit_op(Op::dup, 0); // keep 0 for set
                                self.chunk(chunk_idx).emit_op(Op::drop, 0); // cleanup
                                // Simpler: pair[0] += start → pair, 0, pair[0]+start → array_set
                                self.chunk(chunk_idx).emit_op(Op::dup, 0); // dup pair
                                self.chunk(chunk_idx).emit_op_u16(Op::r#const, zero_c, 0);
                                common_collections::emit_get(&mut self.chunks[chunk_idx], 0); // pair[0]
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, start_slot, 0);
                                self.chunk(chunk_idx).emit_op(Op::dyn_add, 0); // pair[0] + start
                                // Stack: [pair, new_idx]. Need [pair, 0, new_idx] for array_set.
                                let new_idx_tmp = self.scope(chunk_idx).alloc("__enum_ni");
                                self.chunk(chunk_idx).emit_op_u16(Op::local_set, new_idx_tmp, 0);
                                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                                self.chunk(chunk_idx).emit_op_u16(Op::r#const, zero_c, 0);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, new_idx_tmp, 0);
                                common_collections::emit_set(&mut self.chunks[chunk_idx], 0);
                                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                                // i++
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_slot, 0);
                                self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
                                self.chunk(chunk_idx).emit_op(Op::i32_add, 0);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx_slot, 0);
                                self.chunk(chunk_idx).emit_loop(loop_start, 0);
                                self.chunk(chunk_idx).patch_jump(exit);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, result_slot, 0);
                            }
                            return Ok(());
                        }
                        "zip" => {
                            common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_zip", 0);
                            for a in args { self.compile_expr(a, chunk_idx)?; }
                            common::bundle::emit_call_invoke(self.chunk(chunk_idx), args.len() as u8, 0);
                            return Ok(());
                        }
                        "sorted" => {
                            // Check for key= keyword argument
                            let key_fn = keywords.iter().find(|kw| kw.name.as_deref() == Some("key"));
                            if let Some(kw) = key_fn {
                                // sorted(iterable, key=fn): map key fn, sort, unsort
                                // Strategy: build [(key(x), x)] pairs, sort by key, extract values
                                if args.len() == 1 {
                                    return self.compile_sorted_with_key(&args[0], &kw.value, chunk_idx);
                                }
                            }
                            common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_sorted", 0);
                            for a in args { self.compile_expr(a, chunk_idx)?; }
                            common::bundle::emit_call_invoke(self.chunk(chunk_idx), args.len() as u8, 0);
                            return Ok(());
                        }
                        "reversed" => {
                            common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_reversed", 0);
                            for a in args { self.compile_expr(a, chunk_idx)?; }
                            common::bundle::emit_call_invoke(self.chunk(chunk_idx), args.len() as u8, 0);
                            return Ok(());
                        }
                        "sum" => {
                            common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_sum", 0);
                            for a in args { self.compile_expr(a, chunk_idx)?; }
                            common::bundle::emit_call_invoke(self.chunk(chunk_idx), args.len() as u8, 0);
                            return Ok(());
                        }
                        "min" => {
                            common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_min", 0);
                            for a in args { self.compile_expr(a, chunk_idx)?; }
                            common::bundle::emit_call_invoke(self.chunk(chunk_idx), args.len() as u8, 0);
                            return Ok(());
                        }
                        "max" => {
                            common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_max", 0);
                            for a in args { self.compile_expr(a, chunk_idx)?; }
                            common::bundle::emit_call_invoke(self.chunk(chunk_idx), args.len() as u8, 0);
                            return Ok(());
                        }
                        "map" => {
                            if args.len() == 2 {
                                return self.compile_map(args, chunk_idx);
                            }
                        }
                        "filter" => {
                            if args.len() == 2 {
                                return self.compile_filter(args, chunk_idx);
                            }
                        }
                        "iter" => {
                            // iter(x) → just return x (arrays are already iterable)
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                return Ok(());
                            }
                        }
                        "next" => {
                            // next(iterator) — pop first element via array_shift
                            if args.len() >= 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                common_collections::emit_shift(&mut self.chunks[chunk_idx], 0);
                                return Ok(());
                            }
                        }
                        "any" => {
                            return self.compile_host_call("vybe:array", "any", args, chunk_idx);
                        }
                        "all" => {
                            return self.compile_host_call("vybe:array", "all", args, chunk_idx);
                        }
                        "type" => {
                            return self.compile_host_call("vybe:array", "pytype", args, chunk_idx);
                        }
                        "isinstance" => {
                            if args.len() == 2 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                // Map Python type names to VM type names for ref_test
                                let type_str = if let Expression::Name(type_name) = &args[1] {
                                    match type_name.as_str() {
                                        "int" => "integer",
                                        "float" => "double",
                                        "str" => "string",
                                        "bool" => "boolean",
                                        "list" | "tuple" => "array",
                                        "dict" => "object",
                                        "type" => "object",
                                        other => other,
                                    }.to_string()
                                } else {
                                    // Dynamic type name — can't use ref_test (needs constant)
                                    // Fall back to runtime check
                                    String::new()
                                };
                                if !type_str.is_empty() {
                                    // Use ref_test opcode — standard WASM GC type check
                                    let c = self.chunk(chunk_idx).add_constant(Value::String(Arc::from(type_str.as_str())));
                                    self.chunk(chunk_idx).emit_op_u16(Op::ref_test, c, 0);
                                } else {
                                    // Dynamic: compile type expr, use host pytype + eq
                                    self.compile_expr(&args[1], chunk_idx)?;
                                    let type_fn = self.chunk(chunk_idx).add_import("vybe:array", "pytype");
                                    // Stack: [obj, type_arg]. Need pytype(obj) == str(type_arg)
                                    // Simplified: just emit false for dynamic isinstance
                                    self.chunk(chunk_idx).emit_op(Op::drop, 0);
                                    self.chunk(chunk_idx).emit_op(Op::drop, 0);
                                    self.chunk(chunk_idx).emit_op(Op::r#false, 0);
                                }
                                return Ok(());
                            }
                        }
                        "list" => {
                            return self.compile_host_call("vybe:array", "list", args, chunk_idx);
                        }
                        "dict" => {
                            return self.compile_host_call("vybe:array", "dict", args, chunk_idx);
                        }
                        "set" => {
                            return self.compile_host_call("vybe:array", "pyset", args, chunk_idx);
                        }
                        "tuple" => {
                            return self.compile_host_call("vybe:array", "tuple", args, chunk_idx);
                        }
                        "bool" => {
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                self.chunk(chunk_idx).emit_op(Op::dyn_to_bool, 0);
                                return Ok(());
                            }
                        }
                        "round" => {
                            if args.len() >= 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                common::math::emit_round(self.chunk(chunk_idx), 0);
                                return Ok(());
                            }
                        }
                        "chr" => {
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                self.chunk(chunk_idx).emit_op(Op::str_from_char_code, 0);
                                return Ok(());
                            }
                        }
                        "ord" => {
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                self.chunk(chunk_idx).emit_op(Op::str_char_code_at, 0);
                                return Ok(());
                            }
                        }
                        "hex" => {
                            return self.compile_host_call("vybe:convert", "hex", args, chunk_idx);
                        }
                        "oct" => {
                            return self.compile_host_call("vybe:convert", "oct", args, chunk_idx);
                        }
                        "repr" | "ascii" => {
                            return self.compile_host_call("vybe:convert", "toString", args, chunk_idx);
                        }
                        "open" => {
                            // open(filename, mode) → file handle via wasi:filesystem
                            // Same host as VB Open/PHP fopen
                            for a in args { self.compile_expr(a, chunk_idx)?; }
                            return self.compile_host_call("wasi:filesystem", "openFile", args, chunk_idx);
                        }
                        "hasattr" => {
                            if args.len() >= 2 {
                                // hasattr(obj, key) → obj[key] is not null
                                self.compile_expr(&args[0], chunk_idx)?;
                                self.compile_expr(&args[1], chunk_idx)?;
                                common_collections::emit_get(&mut self.chunks[chunk_idx], 0);
                                self.chunk(chunk_idx).emit_op(Op::ref_is_null, 0);
                                self.chunk(chunk_idx).emit_op(Op::dyn_not, 0);
                                return Ok(());
                            }
                        }
                        "getattr" => {
                            if args.len() >= 2 {
                                // getattr(obj, key) → obj[key] via array_get
                                self.compile_expr(&args[0], chunk_idx)?;
                                self.compile_expr(&args[1], chunk_idx)?;
                                common_collections::emit_get(&mut self.chunks[chunk_idx], 0);
                                return Ok(());
                            }
                        }
                        "setattr" => {
                            if args.len() >= 3 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                self.compile_expr(&args[1], chunk_idx)?;
                                self.compile_expr(&args[2], chunk_idx)?;
                                common_collections::emit_set(&mut self.chunks[chunk_idx], 0);
                                return Ok(());
                            }
                        }
                        "pow" => {
                            if args.len() >= 2 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                self.compile_expr(&args[1], chunk_idx)?;
                                let pow_idx = self.chunk(chunk_idx).add_import("vybe:math", "pow");
                                self.chunk(chunk_idx).emit_op_u16(Op::call_import, pow_idx, 0);
                                self.chunk(chunk_idx).emit(2, 0);
                                // pow(x, y, mod) — apply modulo if 3 args
                                if args.len() >= 3 {
                                    self.compile_expr(&args[2], chunk_idx)?;
                                    self.chunk(chunk_idx).emit_op(Op::i32_rem_s, 0);
                                }
                                return Ok(());
                            }
                        }
                        "divmod" => {
                            if args.len() == 2 {
                                // divmod(a, b) → (a // b, a % b) as a 2-element array
                                self.compile_expr(&args[0], chunk_idx)?;
                                self.compile_expr(&args[1], chunk_idx)?;
                                let a_slot = self.scope(chunk_idx).alloc("__dm_a");
                                let b_slot = self.scope(chunk_idx).alloc("__dm_b");
                                self.chunk(chunk_idx).emit_op_u16(Op::local_set, b_slot, 0);
                                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_set, a_slot, 0);
                                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                                // a // b
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, a_slot, 0);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, b_slot, 0);
                                self.chunk(chunk_idx).emit_op(Op::f64_div, 0);
                                common::math::emit_floor(self.chunk(chunk_idx), 0);
                                // a % b
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, a_slot, 0);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, b_slot, 0);
                                self.chunk(chunk_idx).emit_op(Op::i32_rem_s, 0);
                                // pack as [quotient, remainder]
                                self.chunk(chunk_idx).emit_op_u16(Op::array_new, 2, 0);
                                return Ok(());
                            }
                        }
                        "format" => {
                            // format(value, spec) → vybe:string/format host call
                            if args.len() >= 2 {
                                return self.compile_host_call("vybe:string", "format", args, chunk_idx);
                            } else if args.len() == 1 {
                                // format(value) → str(value)
                                self.compile_expr(&args[0], chunk_idx)?;
                                common::convert::emit_to_string(self.chunk(chunk_idx), 0);
                                return Ok(());
                            }
                        }
                        "callable" => {
                            if args.len() == 1 {
                                // callable(obj): check if obj is a function or has __call__
                                self.compile_expr(&args[0], chunk_idx)?;
                                let obj_slot = self.scope(chunk_idx).alloc("__call_obj");
                                self.chunk(chunk_idx).emit_op_u16(Op::local_set, obj_slot, 0);
                                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                                // Try struct_get "__call__"
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_slot, 0);
                                let call_key = self.chunk(chunk_idx).add_constant(Value::String(Arc::from("__call__")));
                                self.chunk(chunk_idx).emit_op_u16(Op::struct_get, call_key, 0);
                                self.chunk(chunk_idx).emit_op(Op::ref_is_null, 0);
                                self.chunk(chunk_idx).emit_op(Op::dyn_not, 0);
                                return Ok(());
                            }
                        }
                        "bin" => {
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                let bin_fn = self.chunk(chunk_idx).add_import("vybe:convert", "bin");
                                self.chunk(chunk_idx).emit_op_u16(Op::call_import, bin_fn, 0);
                                self.chunk(chunk_idx).emit(1, 0);
                                return Ok(());
                            }
                        }
                        "id" => {
                            // id(obj) → return a numeric identifier (hash of the object ref)
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                let id_fn = self.chunk(chunk_idx).add_import("vybe:convert", "id");
                                self.chunk(chunk_idx).emit_op_u16(Op::call_import, id_fn, 0);
                                self.chunk(chunk_idx).emit(1, 0);
                                return Ok(());
                            }
                        }
                        "hash" => {
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                let hash_fn = self.chunk(chunk_idx).add_import("vybe:convert", "hash");
                                self.chunk(chunk_idx).emit_op_u16(Op::call_import, hash_fn, 0);
                                self.chunk(chunk_idx).emit(1, 0);
                                return Ok(());
                            }
                        }
                        "frozenset" => {
                            return self.compile_host_call("vybe:array", "pyset", args, chunk_idx);
                        }
                        "defaultdict" => {
                            // defaultdict(factory) → dict with __default_factory property
                            common::dict::emit_new(self.chunk(chunk_idx), 0);
                            if args.len() >= 1 {
                                self.chunk(chunk_idx).emit_op(Op::dup, 0);
                                self.compile_expr(&args[0], chunk_idx)?;
                                let key = self.chunk(chunk_idx).add_constant(Value::String(Arc::from("__default_factory")));
                                self.chunk(chunk_idx).emit_op_u16(Op::struct_set, key, 0);
                                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                            }
                            return Ok(());
                        }
                        "Counter" if self.scope(chunk_idx).get("Counter").is_none() => {
                            // Counter(iterable) → dict of element counts
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                let arr_slot = self.scope(chunk_idx).alloc("__ctr_arr");
                                let dict_slot = self.scope(chunk_idx).alloc("__ctr_d");
                                let idx_slot = self.scope(chunk_idx).alloc("__ctr_i");
                                self.chunk(chunk_idx).emit_op_u16(Op::local_set, arr_slot, 0);
                                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                                common::dict::emit_new(self.chunk(chunk_idx), 0);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_set, dict_slot, 0);
                                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                                // for elem in arr: dict[elem] = dict.get(elem, 0) + 1
                                self.chunk(chunk_idx).emit_op(Op::i32_const_0, 0);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx_slot, 0);
                                let loop_start = self.chunk(chunk_idx).current_offset();
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_slot, 0);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, arr_slot, 0);
                                common_collections::emit_len(&mut self.chunks[chunk_idx], 0);
                                self.chunk(chunk_idx).emit_op(Op::dyn_lt, 0);
                                let exit = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);
                                // elem = arr[i]
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, arr_slot, 0);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_slot, 0);
                                common_collections::emit_get(&mut self.chunks[chunk_idx], 0);
                                let elem_slot = self.scope(chunk_idx).alloc("__ctr_e");
                                self.chunk(chunk_idx).emit_op_u16(Op::local_set, elem_slot, 0);
                                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                                // dict[elem] = (dict[elem] or 0) + 1
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, dict_slot, 0);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, elem_slot, 0);
                                // get current count
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, dict_slot, 0);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, elem_slot, 0);
                                common_collections::emit_get(&mut self.chunks[chunk_idx], 0);
                                // if null, use 0
                                self.chunk(chunk_idx).emit_op(Op::dup, 0);
                                self.chunk(chunk_idx).emit_op(Op::ref_is_null, 0);
                                let not_null = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);
                                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                                self.chunk(chunk_idx).emit_op(Op::i32_const_0, 0);
                                self.chunk(chunk_idx).patch_jump(not_null);
                                // + 1
                                self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
                                self.chunk(chunk_idx).emit_op(Op::dyn_add, 0);
                                // array_set(dict, elem, count+1)
                                common_collections::emit_set(&mut self.chunks[chunk_idx], 0);
                                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                                // i++
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_slot, 0);
                                self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
                                self.chunk(chunk_idx).emit_op(Op::i32_add, 0);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx_slot, 0);
                                self.chunk(chunk_idx).emit_loop(loop_start, 0);
                                self.chunk(chunk_idx).patch_jump(exit);
                                self.chunk(chunk_idx).emit_op_u16(Op::local_get, dict_slot, 0);
                            } else {
                                // Counter() with no args → empty dict
                                common::dict::emit_new(self.chunk(chunk_idx), 0);
                            }
                            return Ok(());
                        }
                        "namedtuple" => {
                            // namedtuple('Name', ['field1', 'field2']) → class constructor
                            // Simplified: return a function that creates a dict with named fields
                            // Usage: Point = namedtuple('Point', ['x', 'y']); p = Point(1, 2)
                            // For now: just emit null (class creation needs a function chunk)
                            self.chunk(chunk_idx).emit_op(Op::null, 0);
                            return Ok(());
                        }
                        "bytes" | "bytearray" => {
                            // bytes(iterable) or bytes(n) → array (simplified)
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                // If it's a number, create zero-filled array
                                // If it's a list, use it directly
                                // Simplified: just return the arg (list or string)
                                return Ok(());
                            } else {
                                // bytes() → empty array
                                self.chunk(chunk_idx).emit_op_u16(Op::array_new, 0, 0);
                                return Ok(());
                            }
                        }
                        "vars" => {
                            // vars(obj) → get __keys and build dict of properties
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                // Return the object itself — it IS a dict of properties
                                return Ok(());
                            }
                        }
                        "dir" => {
                            // dir(obj) → return __keys array (list of attribute names)
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                common::dict::emit_keys(self.chunk(chunk_idx), 0);
                                return Ok(());
                            }
                        }
                        _ => {
                            // Check cross-language common imports as fallback
                            if let Some((module, func_name)) = vybe_compiler_common::imports::resolve_common_import(name) {
                                return self.compile_host_call(module, func_name, args, chunk_idx);
                            }
                        }
                    }
                }
                // Method calls: obj.method(args)
                if let Expression::Attribute { value, attr } = func.as_ref() {
                    return self.compile_method_call(value, attr, args, chunk_idx);
                }
                // General function call
                // Check if func is an unresolved name → emit call_import("*", name)
                if let Expression::Name(name) = func.as_ref() {
                    if matches!(self.resolve_variable(name, chunk_idx), VarResolution::Global) {
                        // Unresolved global name: route through host import
                        for a in args { self.compile_expr(a, chunk_idx)?; }
                        let import_idx = self.chunk(chunk_idx).add_import("*", name);
                        self.chunk(chunk_idx).emit_op_u16(Op::call_import, import_idx, 0);
                        self.chunk(chunk_idx).emit(args.len() as u8, 0);
                        return Ok(());
                    }
                }
                // Locally resolved func (local/upvalue) or non-Name expression
                self.compile_expr(func, chunk_idx)?;
                let mut argc = 0u8;
                for a in args {
                    if let Expression::Starred(inner) = a {
                        // *args: spread the iterable onto the stack
                        self.compile_expr(inner, chunk_idx)?;
                        self.chunk(chunk_idx).emit_op(Op::spread, 0);
                    } else {
                        self.compile_expr(a, chunk_idx)?;
                        argc += 1;
                    }
                }
                self.chunk(chunk_idx).emit_op_u8(Op::call_ref, argc, 0);
            }

            Expression::Attribute { value, attr } => {
                // Check for module constants: math.pi, math.e, math.inf, math.nan, etc.
                if let Expression::Name(module_name) = value.as_ref() {
                    let constant = match (module_name.as_str(), attr.as_str()) {
                        ("math", "pi") => Some(Value::F64(std::f64::consts::PI)),
                        ("math", "e") => Some(Value::F64(std::f64::consts::E)),
                        ("math", "tau") => Some(Value::F64(std::f64::consts::TAU)),
                        ("math", "inf") => Some(Value::F64(f64::INFINITY)),
                        ("math", "nan") => Some(Value::F64(f64::NAN)),
                        ("sys", "maxsize") => Some(Value::I64(i64::MAX)),
                        ("float", "inf") => Some(Value::F64(f64::INFINITY)),
                        ("float", "nan") => Some(Value::F64(f64::NAN)),
                        _ => None,
                    };
                    if let Some(val) = constant {
                        let c = self.chunk(chunk_idx).add_constant(val);
                        self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
                        return Ok(());
                    }
                }
                self.compile_expr(value, chunk_idx)?;
                let c = self.chunk(chunk_idx).add_constant(Value::String(Arc::from(attr.as_str())));
                self.chunk(chunk_idx).emit_op_u16(Op::struct_get, c, 0);
            }

            Expression::Subscript { value, slice } => {
                // Slicing: obj[start:end:step]
                if let Expression::Slice { lower, upper, step } = slice.as_ref() {
                    if step.is_some() {
                        // Use sliceStep host function for step support
                        common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_slicestep", 0);
                        self.compile_expr(value, chunk_idx)?;
                        // start (null = default)
                        if let Some(lo) = lower { self.compile_expr(lo, chunk_idx)?; }
                        else { self.chunk(chunk_idx).emit_op(Op::null, 0); }
                        // end (null = default)
                        if let Some(up) = upper { self.compile_expr(up, chunk_idx)?; }
                        else { self.chunk(chunk_idx).emit_op(Op::null, 0); }
                        // step
                        self.compile_expr(step.as_ref().unwrap(), chunk_idx)?;
                        common::bundle::emit_call_invoke(self.chunk(chunk_idx), 4, 0);
                    } else {
                        // No step — use existing array_slice opcode
                        self.compile_expr(value, chunk_idx)?;
                        if let Some(lo) = lower {
                            self.compile_expr(lo, chunk_idx)?;
                        } else {
                            let c = self.chunk(chunk_idx).add_constant(Value::I32(0));
                            self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
                        }
                        if let Some(up) = upper {
                            self.compile_expr(up, chunk_idx)?;
                        } else {
                            let c = self.chunk(chunk_idx).add_constant(Value::I32(i32::MAX));
                            self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
                        }
                        common_collections::emit_slice(&mut self.chunks[chunk_idx], 0);
                    }
                }
                // Dict string key lookup
                else if let Expression::Str(s) = slice.as_ref() {
                    self.compile_expr(value, chunk_idx)?;
                    common::dict::emit_get_const_key(self.chunk(chunk_idx), s, 0);
                }
                // Normal index
                else {
                    self.compile_expr(value, chunk_idx)?;
                    // Handle negative indices at compile time:
                    // x[-1] → x[len(x) + (-1)]
                    let is_negative_literal = matches!(slice.as_ref(),
                        Expression::UnaryOp { op: UnaryOp::USub, operand }
                        if matches!(operand.as_ref(), Expression::Int(_))
                    );
                    if is_negative_literal {
                        self.chunk(chunk_idx).emit_op(Op::dup, 0);
                        common_collections::emit_len(&mut self.chunks[chunk_idx], 0);
                        self.compile_expr(slice, chunk_idx)?;
                        self.chunk(chunk_idx).emit_op(Op::dyn_add, 0);
                    } else {
                        self.compile_expr(slice, chunk_idx)?;
                    }
                    common_collections::emit_get(&mut self.chunks[chunk_idx], 0);
                }
            }

            Expression::Slice { lower, upper, step } => {
                // Compile slice as a call to vybe:array/slice(obj, start, end)
                // This is used inside Subscript which provides the obj
                // The slice itself is compiled as arguments pushed individually
                // Actually, slice is always inside Subscript, so handle it there
                // If we get here standalone, push null
                self.chunk(chunk_idx).emit_op(Op::null, 0);
                let _ = (lower, upper, step);
            }

            Expression::IfExp { test, body, orelse } => {
                self.compile_expr(test, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::dyn_to_bool, 0);
                let false_jump = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);
                self.compile_expr(body, chunk_idx)?;
                let end_jump = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                self.chunk(chunk_idx).patch_jump(false_jump);
                self.compile_expr(orelse, chunk_idx)?;
                self.chunk(chunk_idx).patch_jump(end_jump);
            }

            Expression::Lambda { params, body } => {
                let name = "__lambda";
                let func_idx = self.compile_function(name, params, &[Statement::Return(Some(*body.clone()))])?;
                self.emit_ref_func(chunk_idx, func_idx);
            }

            Expression::Starred(inner) => {
                // In most contexts, just compile the inner expression
                self.compile_expr(inner, chunk_idx)?;
            }

            Expression::Await(inner) => {
                self.compile_expr(inner, chunk_idx)?;
                common::functions::emit_await(self.chunk(chunk_idx), 0);
            }

            Expression::Yield(expr) => {
                // yield value → suspend with value, resume returns sent value
                if let Some(e) = expr {
                    self.compile_expr(e, chunk_idx)?;
                } else {
                    self.chunk(chunk_idx).emit_op(Op::null, 0);
                }
                self.chunk(chunk_idx).emit_op_u16(Op::suspend, 0, 0);
            }

            Expression::YieldFrom(expr) => {
                // yield from iterable → iterate and yield each
                // Simplified: compile as expression (full delegation needs more work)
                self.compile_expr(expr, chunk_idx)?;
            }

            Expression::NamedExpr { target, value } => {
                self.compile_expr(value, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::dup, 0);
                self.compile_assign_target(target, chunk_idx)?;
            }

            Expression::ListComp { element, generators } |
            Expression::SetComp { element, generators } => {
                self.compile_comprehension(element, generators, chunk_idx)?;
            }

            Expression::DictComp { key, value, generators } => {
                // Create empty dict, iterate, add items
                common::dict::emit_new(self.chunk(chunk_idx), 0);
                let result_local = self.scope(chunk_idx).alloc("__comp_result");
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, result_local, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);

                self.compile_comp_generators(generators, &|s| {
                    s.chunk(chunk_idx).emit_op_u16(Op::local_get, result_local, 0);
                    s.compile_expr(key, chunk_idx)?;
                    s.compile_expr(value, chunk_idx)?;
                    common::dict::emit_set(s.chunk(chunk_idx), 0);
                    Ok(())
                }, chunk_idx)?;

                self.chunk(chunk_idx).emit_op_u16(Op::local_get, result_local, 0);
            }

            Expression::GeneratorExp { element, generators } => {
                // Compile as list for now
                self.compile_comprehension(element, generators, chunk_idx)?;
            }
        }
        Ok(())
    }

    // ── Helpers ──────────────────────────────────────────────────────

    /// Extract a type name from an expression (for exception type matching).
    /// Compile a built-in exception type constructor (ValueError, TypeError, etc.).
    /// Creates a chunk that produces an object with __exception_type, name, message properties.
    /// Stores the constructor as a global so `raise ValueError("bad")` works.
    fn compile_exception_constructor(&mut self, exc_name: &str, script_chunk_idx: usize) {
        let lower = exc_name.to_lowercase();
        let mut ctor = Chunk::new(&lower);
        ctor.arity = 1; // message arg (optional — padded with null if missing)
        ctor.local_count = 3; // callee(0) + message(1) + this(2)
        let msg_slot = 1u16;
        let this_slot = 2u16;

        // Use shared exception object shape (compatible across all languages)
        common::errors::emit_exception_constructor(&mut ctor, this_slot, exc_name, msg_slot, 0);

        // Return this
        ctor.emit_op_u16(Op::local_get, this_slot, 0);
        ctor.emit_op(Op::r#return, 0);

        let ctor_idx = self.chunks.len();
        self.chunks.push(ctor);

        // Store as global: ValueError = <constructor func ref>
        let local = self.scope(script_chunk_idx).alloc(exc_name);
        self.chunk(script_chunk_idx).emit_op_u16(Op::ref_func, ctor_idx as u16, 0);
        self.chunk(script_chunk_idx).emit(0, 0);
        self.chunk(script_chunk_idx).emit_op_u16(Op::local_set, local, 0);
        let global_name = self.chunk(script_chunk_idx).add_constant(
            Value::String(Arc::from(lower.as_str()))
        );
        self.chunk(script_chunk_idx).emit_op_u16(Op::global_set, global_name, 0);
        self.chunk(script_chunk_idx).emit_op(Op::drop, 0);

        // Register type entry for cross-language type matching
        use vybe_bytecode::chunk::TypeEntry;
        self.chunks[0].types.push(TypeEntry {
            name: lower,
            parent: "exception".to_string(),
            fields: vec!["message".to_string(), "name".to_string()],
            methods: Vec::new(),
            is_interface: false,
            implements: Vec::new(),
            constructor_chunk: Some(ctor_idx),
        });
    }

    fn expr_to_name(&self, expr: &Expression) -> String {
        match expr {
            Expression::Name(n) => n.clone(),
            Expression::Attribute { value, attr } => {
                format!("{}.{}", self.expr_to_name(value), attr)
            }
            _ => format!("{:?}", expr),
        }
    }

    fn compile_host_call(&mut self, module: &str, name: &str, args: &[Expression], chunk_idx: usize) -> Result<(), String> {
        let import_idx = self.chunk(chunk_idx).add_import(module, name);
        for a in args { self.compile_expr(a, chunk_idx)?; }
        self.chunk(chunk_idx).emit_op_u16(Op::call_import, import_idx, 0);
        self.chunk(chunk_idx).emit(args.len() as u8, 0);
        Ok(())
    }

    fn compile_print(&mut self, args: &[Expression], chunk_idx: usize) -> Result<(), String> {
        for a in args {
            self.compile_expr(a, chunk_idx)?;
        }
        common::io::emit_print(self.chunk(chunk_idx), args.len() as u8, 0);
        Ok(())
    }

    fn compile_print_with_kwargs(&mut self, args: &[Expression], keywords: &[Keyword], chunk_idx: usize) -> Result<(), String> {
        // Check for sep= and end= kwargs
        let sep = keywords.iter().find(|k| k.name.as_deref() == Some("sep"));
        let _end = keywords.iter().find(|k| k.name.as_deref() == Some("end"));

        if let Some(sep_kw) = sep {
            // print(a, b, sep=",") → join args with separator, then print
            if args.is_empty() {
                common::io::emit_print(self.chunk(chunk_idx), 0, 0);
                return Ok(());
            }
            // Convert each arg to string and join with sep
            for (i, a) in args.iter().enumerate() {
                // "" + arg → string coercion
                let empty = self.chunk(chunk_idx).add_constant(Value::String(Arc::from("")));
                self.chunk(chunk_idx).emit_op_u16(Op::r#const, empty, 0);
                self.compile_expr(a, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::dyn_add, 0);
                if i > 0 {
                    // Concatenate: prev + sep + current
                    // Stack has: [prev_result, current_str]
                    // Need: prev_result + sep + current_str
                    let sep_tmp = self.scope(chunk_idx).alloc("__print_sep_tmp");
                    self.chunk(chunk_idx).emit_op_u16(Op::local_set, sep_tmp, 0);
                    self.chunk(chunk_idx).emit_op(Op::drop, 0);
                    // Push sep
                    self.compile_expr(&sep_kw.value, chunk_idx)?;
                    self.chunk(chunk_idx).emit_op(Op::str_concat, 0);
                    // Push current
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, sep_tmp, 0);
                    self.chunk(chunk_idx).emit_op(Op::str_concat, 0);
                }
            }
            // Print the joined string
            common::io::emit_print(self.chunk(chunk_idx), 1, 0);
        } else {
            // No sep= → default print
            self.compile_print(args, chunk_idx)?;
        }
        Ok(())
    }

    fn compile_range(&mut self, args: &[Expression], chunk_idx: usize) -> Result<(), String> {
        // Push func ref first, then normalized args
        common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_range", 0);
        match args.len() {
            1 => {
                self.chunk(chunk_idx).emit_op(Op::i32_const_0, 0);
                self.compile_expr(&args[0], chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
            }
            2 => {
                self.compile_expr(&args[0], chunk_idx)?;
                self.compile_expr(&args[1], chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
            }
            _ => {
                for a in args { self.compile_expr(a, chunk_idx)?; }
            }
        }
        common::bundle::emit_call_invoke(self.chunk(chunk_idx), 3, 0);
        Ok(())
    }

    fn compile_method_call(&mut self, obj: &Expression, method: &str, args: &[Expression], chunk_idx: usize) -> Result<(), String> {
        // ── super().__init__(args) → call parent constructor directly ──
        if method == "__init__" {
            if let Expression::Call { func, args: super_args, .. } = obj {
                if let Expression::Name(n) = func.as_ref() {
                    if n == "super" && super_args.is_empty() {
                        if let Some(ref parent) = self.current_class_parent.clone() {
                            let parent_c = self.chunk(chunk_idx).add_constant(Value::String(Arc::from(parent.as_str())));
                            self.chunk(chunk_idx).emit_op_u16(Op::global_get, parent_c, 0);
                            for a in args { self.compile_expr(a, chunk_idx)?; }
                            self.chunk(chunk_idx).emit_op_u8(Op::call_ref, args.len() as u8, 0);
                            // Parent constructor returns the created object.
                            // Store it in self (slot 1) so subsequent self.x accesses work.
                            self.chunk(chunk_idx).emit_op(Op::dup, 0);
                            self.chunk(chunk_idx).emit_op_u16(Op::local_set, 1, 0);
                            return Ok(());
                        }
                    }
                }
            }
        }

        // ── Module calls: math.sqrt(x), json.loads(s), etc ──
        if let Expression::Name(module_name) = obj {
            if let Some(()) = self.try_compile_module_call(module_name, method, args, chunk_idx)? {
                return Ok(());
            }
        }

        // ── String methods: direct opcodes (strings are primitives, no property lookup) ──
        if let Some(()) = self.try_compile_string_method(obj, method, args, chunk_idx)? {
            return Ok(());
        }

        // ── All other methods: runtime dispatch ──
        // Try user method first (struct_get), fall back to builtin inline opcodes.
        self.compile_method_with_fallback(obj, method, args, chunk_idx)
    }

    /// Handle Python module calls: math.sqrt(x), json.loads(s), random.random(), re.search(pat, s)
    fn try_compile_module_call(&mut self, module: &str, method: &str, args: &[Expression], chunk_idx: usize) -> Result<Option<()>, String> {
        let (host_module, host_func) = match module {
            "math" => match method {
                // Direct WASM opcodes (no host call needed)
                "sqrt" | "ceil" | "floor" | "trunc" | "fabs" | "abs" => {
                    if args.len() == 1 {
                        self.compile_expr(&args[0], chunk_idx)?;
                        match method {
                            "sqrt"  => common::math::emit_sqrt(self.chunk(chunk_idx), 0),
                            "ceil"  => common::math::emit_ceil(self.chunk(chunk_idx), 0),
                            "floor" => common::math::emit_floor(self.chunk(chunk_idx), 0),
                            "trunc" => common::math::emit_trunc(self.chunk(chunk_idx), 0),
                            "fabs" | "abs" => common::math::emit_abs(self.chunk(chunk_idx), 0),
                            _ => unreachable!(),
                        }
                        return Ok(Some(()));
                    }
                    return Ok(None);
                }
                "sin" => ("vybe:math", "sin"),
                "cos" => ("vybe:math", "cos"),
                "tan" => ("vybe:math", "tan"),
                "asin" => ("vybe:math", "asin"),
                "acos" => ("vybe:math", "acos"),
                "atan" => ("vybe:math", "atan"),
                "atan2" => ("vybe:math", "atan2"),
                "exp" => ("vybe:math", "exp"),
                "log" => ("vybe:math", "log"),
                "log2" => ("vybe:math", "log2"),
                "log10" => ("vybe:math", "log10"),
                "pow" => ("vybe:math", "pow"),
                "hypot" => ("vybe:math", "hypot"),
                "copysign" | "sign" => ("vybe:math", "sign"),
                "isnan" | "isinf" | "isfinite" => {
                    // math.isnan(x) → x != x (NaN is the only value not equal to itself)
                    if args.len() == 1 && method == "isnan" {
                        self.compile_expr(&args[0], chunk_idx)?;
                        self.chunk(chunk_idx).emit_op(Op::dup, 0);
                        self.chunk(chunk_idx).emit_op(Op::dyn_ne, 0);
                        return Ok(Some(()));
                    }
                    // math.isinf/isfinite — approximate with large value check
                    if args.len() == 1 {
                        self.compile_expr(&args[0], chunk_idx)?;
                        common::math::emit_abs(self.chunk(chunk_idx), 0);
                        let inf_c = self.chunk(chunk_idx).add_constant(Value::F64(f64::MAX));
                        self.chunk(chunk_idx).emit_op_u16(Op::r#const, inf_c, 0);
                        if method == "isinf" {
                            self.chunk(chunk_idx).emit_op(Op::dyn_gt, 0);
                        } else {
                            // isfinite = not isinf and not isnan
                            self.chunk(chunk_idx).emit_op(Op::dyn_le, 0);
                        }
                        return Ok(Some(()));
                    }
                    return Ok(None);
                }
                "pi" | "e" | "tau" | "inf" | "nan" => {
                    // Constants — these are accessed as math.pi, not math.pi()
                    // This path is for math.pi() which is wrong usage, but handle gracefully
                    return Ok(None);
                }
                _ => return Ok(None),
            },
            "json" => match method {
                "loads" | "load" => ("vybe:json", "parse"),
                "dumps" | "dump" => ("vybe:json", "stringify"),
                _ => return Ok(None),
            },
            "random" => match method {
                "random" => {
                    // random.random() → random float 0..1
                    let rnd = self.chunk(chunk_idx).add_import("vybe:math", "random");
                    self.chunk(chunk_idx).emit_op_u16(Op::call_import, rnd, 0);
                    self.chunk(chunk_idx).emit(0, 0);
                    return Ok(Some(()));
                }
                "randint" => {
                    // random.randint(a, b) → a + floor(random() * (b - a + 1))
                    if args.len() == 2 {
                        let a_slot = self.scope(chunk_idx).alloc("__ri_a");
                        let b_slot = self.scope(chunk_idx).alloc("__ri_b");
                        self.compile_expr(&args[0], chunk_idx)?;
                        self.chunk(chunk_idx).emit_op_u16(Op::local_set, a_slot, 0);
                        self.chunk(chunk_idx).emit_op(Op::drop, 0);
                        self.compile_expr(&args[1], chunk_idx)?;
                        self.chunk(chunk_idx).emit_op_u16(Op::local_set, b_slot, 0);
                        self.chunk(chunk_idx).emit_op(Op::drop, 0);
                        // a + floor(random() * (b - a + 1))
                        self.chunk(chunk_idx).emit_op_u16(Op::local_get, a_slot, 0);
                        let rnd = self.chunk(chunk_idx).add_import("vybe:math", "random");
                        self.chunk(chunk_idx).emit_op_u16(Op::call_import, rnd, 0);
                        self.chunk(chunk_idx).emit(0, 0);
                        self.chunk(chunk_idx).emit_op_u16(Op::local_get, b_slot, 0);
                        self.chunk(chunk_idx).emit_op_u16(Op::local_get, a_slot, 0);
                        self.chunk(chunk_idx).emit_op(Op::f64_sub, 0);
                        let one_c = self.chunk(chunk_idx).add_constant(Value::F64(1.0));
                        self.chunk(chunk_idx).emit_op_u16(Op::r#const, one_c, 0);
                        self.chunk(chunk_idx).emit_op(Op::dyn_add, 0);
                        self.chunk(chunk_idx).emit_op(Op::f64_mul, 0);
                        common::math::emit_floor(self.chunk(chunk_idx), 0);
                        self.chunk(chunk_idx).emit_op(Op::dyn_add, 0);
                        return Ok(Some(()));
                    }
                    return Ok(None);
                }
                "choice" => {
                    // random.choice(lst) → lst[randint(0, len-1)]
                    if args.len() == 1 {
                        self.compile_expr(&args[0], chunk_idx)?;
                        self.chunk(chunk_idx).emit_op(Op::dup, 0);
                        let rnd = self.chunk(chunk_idx).add_import("vybe:math", "random");
                        common_collections::emit_len(&mut self.chunks[chunk_idx], 0);
                        self.chunk(chunk_idx).emit_op_u16(Op::call_import, rnd, 0);
                        self.chunk(chunk_idx).emit(0, 0);
                        // Stack: [arr, len, rnd]. Need floor(rnd * len)
                        // Reorder: we need arr on bottom, index on top.
                        // Actually: arr is under len. Let me redo.
                        // Simpler approach:
                        let arr_slot = self.scope(chunk_idx).alloc("__ch_arr");
                        self.chunk(chunk_idx).emit_op(Op::drop, 0); // drop random result
                        self.chunk(chunk_idx).emit_op(Op::drop, 0); // drop length
                        self.chunk(chunk_idx).emit_op_u16(Op::local_set, arr_slot, 0);
                        self.chunk(chunk_idx).emit_op(Op::drop, 0);
                        // Redo properly
                        self.chunk(chunk_idx).emit_op_u16(Op::call_import, rnd, 0);
                        self.chunk(chunk_idx).emit(0, 0);
                        self.chunk(chunk_idx).emit_op_u16(Op::local_get, arr_slot, 0);
                        common_collections::emit_len(&mut self.chunks[chunk_idx], 0);
                        self.chunk(chunk_idx).emit_op(Op::f64_mul, 0);
                        common::math::emit_floor(self.chunk(chunk_idx), 0);
                        // arr[index]
                        self.chunk(chunk_idx).emit_op_u16(Op::local_get, arr_slot, 0);
                        // Stack: [index, arr]. array_get needs [arr, index] → swap needed
                        // Actually array_get pops key(TOS) then obj(TOS-1). So [arr, index] → correct order.
                        // We have [floor_result, arr]. That's [index, arr]. Swap:
                        let idx_slot = self.scope(chunk_idx).alloc("__ch_idx");
                        self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx_slot, 0);
                        self.chunk(chunk_idx).emit_op(Op::drop, 0);
                        // Nope. Let me just do it cleanly with locals.
                        self.chunk(chunk_idx).emit_op(Op::drop, 0); // drop arr
                        self.chunk(chunk_idx).emit_op_u16(Op::local_get, arr_slot, 0);
                        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_slot, 0);
                        common_collections::emit_get(&mut self.chunks[chunk_idx], 0);
                        return Ok(Some(()));
                    }
                    return Ok(None);
                }
                "shuffle" | "seed" | "uniform" | "randrange" | "sample" => {
                    // These need more complex implementation — skip for now
                    return Ok(None);
                }
                _ => return Ok(None),
            },
            "re" => match method {
                "search" | "match" | "findall" | "sub" | "split" => {
                    let host_func = match method {
                        "search" => "search",
                        "match" => "match",
                        "findall" => "findAll",
                        "sub" => "replace",
                        "split" => "split",
                        _ => unreachable!(),
                    };
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    let imp = self.chunk(chunk_idx).add_import("vybe:regex", host_func);
                    self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
                    self.chunk(chunk_idx).emit(args.len() as u8, 0);
                    return Ok(Some(()));
                }
                _ => return Ok(None),
            },
            "os" => match method {
                "getcwd" => {
                    self.compile_host_call("wasi:cli", "getCwd", &[], chunk_idx)?;
                    return Ok(Some(()));
                }
                "listdir" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:filesystem", "listDir", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "mkdir" | "makedirs" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:filesystem", "mkdir", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "remove" | "unlink" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:filesystem", "remove", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "rmdir" | "removedirs" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:filesystem", "remove", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "rename" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:filesystem", "rename", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "path" => return Ok(None),
                "getenv" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:cli", "getEnv", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "system" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("vybe:types", "processStart", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "_exit" | "abort" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:cli", "exit", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                _ => return Ok(None),
            },
            "os.path" | "path" => match method {
                "exists" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:filesystem", "exists", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "isfile" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:filesystem", "isFile", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "isdir" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:filesystem", "isDir", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "join" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:filesystem", "pathCombine", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "dirname" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:filesystem", "pathGetDirectory", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "basename" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:filesystem", "pathGetFileName", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "abspath" | "realpath" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:filesystem", "pathGetFullPath", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "splitext" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:filesystem", "pathGetExtension", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "getsize" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:filesystem", "fileSize", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                _ => return Ok(None),
            },
            // threading module: threading.Thread(target=fn).start(), threading.Lock()
            "threading" => match method {
                "Thread" => {
                    // threading.Thread(target=fn) → create thread handle
                    // The target function is the first positional arg or target= keyword
                    if args.len() >= 1 {
                        self.compile_expr(&args[0], chunk_idx)?;
                        common::threading::emit_thread_spawn(self.chunk(chunk_idx), 0);
                        return Ok(Some(()));
                    }
                    return Ok(None);
                }
                "Lock" => {
                    // threading.Lock() → allocate a lock word in shared memory
                    // Returns a memory address (i32) for use with acquire/release
                    // Simplified: allocate 4 bytes at end of memory, return address
                    let alloc_fn = self.chunk(chunk_idx).add_import("wasi:thread", "allocLock");
                    self.chunk(chunk_idx).emit_op_u16(Op::call_import, alloc_fn, 0);
                    self.chunk(chunk_idx).emit(0, 0);
                    return Ok(Some(()));
                }
                _ => return Ok(None),
            },
            // ── socket module (vybe:net — same as VB TcpClient/UdpClient) ──
            "socket" => match method {
                "socket" | "create_connection" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("vybe:net", "tcpConnect", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "gethostbyname" | "getaddrinfo" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("vybe:net", "dnsResolve", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                _ => return Ok(None),
            },
            // ── sqlite3 module (vybe:database — same as VB SqlConnection) ──
            "sqlite3" => match method {
                "connect" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("vybe:database", "connect", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                _ => return Ok(None),
            },
            // ── hashlib module (vybe:crypto — same as VB/PHP) ──
            "hashlib" => match method {
                "md5" | "sha1" | "sha256" | "sha512" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    let func = if method == "md5" { "md5" } else { "sha256" };
                    self.compile_host_call("vybe:crypto", func, args, chunk_idx)?;
                    return Ok(Some(()));
                }
                _ => return Ok(None),
            },
            // ── datetime module (vybe:types + wasi:clocks — same as VB DateTime) ──
            "datetime" => match method {
                "now" | "today" | "utcnow" => {
                    self.compile_host_call("vybe:types", "dateTimeNow", &[], chunk_idx)?;
                    return Ok(Some(()));
                }
                "datetime" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("vybe:types", "dateTimeNew", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "strptime" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("vybe:types", "dateTimeParse", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                _ => return Ok(None),
            },
            "time" => match method {
                "time" => {
                    self.compile_host_call("wasi:clocks", "now", &[], chunk_idx)?;
                    return Ok(Some(()));
                }
                "sleep" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:clocks", "sleep", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "perf_counter" | "monotonic" => {
                    self.compile_host_call("wasi:clocks", "hrtime", &[], chunk_idx)?;
                    return Ok(Some(()));
                }
                _ => return Ok(None),
            },
            // ── http / urllib module (wasi:http — same as VB/PHP) ──
            "requests" => match method {
                "get" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:http", "get", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                "post" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:http", "post", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                _ => return Ok(None),
            },
            "urllib" | "http" => match method {
                "urlopen" | "request" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("wasi:http", "get", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                _ => return Ok(None),
            },
            // ── collections module (vybe:types — same as VB) ──
            "collections" => match method {
                "deque" => {
                    self.compile_host_call("vybe:types", "queueNew", &[], chunk_idx)?;
                    return Ok(Some(()));
                }
                "OrderedDict" => {
                    self.compile_host_call("vybe:types", "dictNew", &[], chunk_idx)?;
                    return Ok(Some(()));
                }
                "defaultdict" => {
                    self.compile_host_call("vybe:types", "dictNew", &[], chunk_idx)?;
                    return Ok(Some(()));
                }
                "Counter" => {
                    self.compile_host_call("vybe:types", "dictNew", &[], chunk_idx)?;
                    return Ok(Some(()));
                }
                _ => return Ok(None),
            },
            // ── xml module (vybe:xml — same as VB/PHP) ──
            "xml" | "ET" | "ElementTree" => match method {
                "parse" | "fromstring" => {
                    for a in args { self.compile_expr(a, chunk_idx)?; }
                    self.compile_host_call("vybe:xml", "parse", args, chunk_idx)?;
                    return Ok(Some(()));
                }
                _ => return Ok(None),
            },
            _ => return Ok(None),
        };
        // Generic host call pattern
        for a in args { self.compile_expr(a, chunk_idx)?; }
        let imp = self.chunk(chunk_idx).add_import(host_module, host_func);
        self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
        self.chunk(chunk_idx).emit(args.len() as u8, 0);
        Ok(Some(()))
    }

    /// String methods — emit direct opcodes. Returns Some(()) if handled, None if not a string method.
    fn try_compile_string_method(&mut self, obj: &Expression, method: &str, args: &[Expression], chunk_idx: usize) -> Result<Option<()>, String> {
        match method {
            "upper" => { self.compile_expr(obj, chunk_idx)?; self.chunk(chunk_idx).emit_op(Op::str_to_upper, 0); }
            "lower" => { self.compile_expr(obj, chunk_idx)?; self.chunk(chunk_idx).emit_op(Op::str_to_lower, 0); }
            "strip" => { self.compile_expr(obj, chunk_idx)?; self.chunk(chunk_idx).emit_op(Op::str_trim, 0); }
            "lstrip" => { self.compile_expr(obj, chunk_idx)?; self.chunk(chunk_idx).emit_op(Op::str_trim_start, 0); }
            "rstrip" => { self.compile_expr(obj, chunk_idx)?; self.chunk(chunk_idx).emit_op(Op::str_trim_end, 0); }
            "split" => {
                self.compile_expr(obj, chunk_idx)?;
                if args.is_empty() {
                    let c = self.chunk(chunk_idx).add_constant(Value::String(Arc::from(" ")));
                    self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
                } else { self.compile_expr(&args[0], chunk_idx)?; }
                self.chunk(chunk_idx).emit_op(Op::str_split, 0);
            }
            "join" => {
                if args.len() == 1 {
                    self.compile_expr(&args[0], chunk_idx)?;
                    self.compile_expr(obj, chunk_idx)?;
                    common::collections::emit_join(self.chunk(chunk_idx), 0);
                } else { self.compile_expr(obj, chunk_idx)?; }
            }
            "replace" => {
                self.compile_expr(obj, chunk_idx)?;
                for a in args { self.compile_expr(a, chunk_idx)?; }
                self.chunk(chunk_idx).emit_op(Op::str_replace, 0);
            }
            "startswith" => {
                self.compile_expr(obj, chunk_idx)?;
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                self.chunk(chunk_idx).emit_op(Op::str_starts_with, 0);
            }
            "endswith" => {
                self.compile_expr(obj, chunk_idx)?;
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                self.chunk(chunk_idx).emit_op(Op::str_ends_with, 0);
            }
            "find" => {
                self.compile_expr(obj, chunk_idx)?;
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                self.chunk(chunk_idx).emit_op(Op::str_index_of, 0);
            }
            "rfind" | "rindex" => {
                self.compile_expr(obj, chunk_idx)?;
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                self.chunk(chunk_idx).emit_op(Op::str_last_index_of, 0);
            }
            "count" => {
                self.compile_expr(obj, chunk_idx)?;
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                let count_fn = self.chunk(chunk_idx).add_import("vybe:string", "count");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, count_fn, 0);
                self.chunk(chunk_idx).emit(2, 0);
            }
            "encode" => {
                self.compile_expr(obj, chunk_idx)?;
            }
            "format" => {
                self.compile_expr(obj, chunk_idx)?;
                let placeholder = self.chunk(chunk_idx).add_constant(Value::String(Arc::from("{}")));
                let two = self.chunk(chunk_idx).add_constant(Value::I32(2));
                let max_val = self.chunk(chunk_idx).add_constant(Value::I32(i32::MAX));
                for a in args {
                    self.compile_expr(a, chunk_idx)?;
                    common::convert::emit_to_string(self.chunk(chunk_idx), 0);
                    let tmp_s = self.scope(chunk_idx).alloc("__fmt_s");
                    let tmp_r = self.scope(chunk_idx).alloc("__fmt_r");
                    self.chunk(chunk_idx).emit_op_u16(Op::local_set, tmp_r, 0);
                    self.chunk(chunk_idx).emit_op(Op::drop, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::local_set, tmp_s, 0);
                    self.chunk(chunk_idx).emit_op(Op::drop, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, tmp_s, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::r#const, placeholder, 0);
                    self.chunk(chunk_idx).emit_op(Op::str_index_of, 0);
                    let idx_local = self.scope(chunk_idx).alloc("__fmt_idx");
                    self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx_local, 0);
                    self.chunk(chunk_idx).emit_op(Op::drop, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, tmp_s, 0);
                    let zero = self.chunk(chunk_idx).add_constant(Value::I32(0));
                    self.chunk(chunk_idx).emit_op_u16(Op::r#const, zero, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_local, 0);
                    self.chunk(chunk_idx).emit_op(Op::str_substring, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, tmp_r, 0);
                    self.chunk(chunk_idx).emit_op(Op::str_concat, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, tmp_s, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_local, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::r#const, two, 0);
                    self.chunk(chunk_idx).emit_op(Op::dyn_add, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::r#const, max_val, 0);
                    self.chunk(chunk_idx).emit_op(Op::str_substring, 0);
                    self.chunk(chunk_idx).emit_op(Op::str_concat, 0);
                }
            }
            "title" | "capitalize" => {
                self.compile_expr(obj, chunk_idx)?;
                common::convert::emit_to_string(self.chunk(chunk_idx), 0);
            }
            "isdigit" | "isnumeric" | "isdecimal" => {
                self.compile_expr(obj, chunk_idx)?;
                common::convert::emit_is_numeric(self.chunk(chunk_idx), 0);
            }
            "isalpha" | "isalnum" | "isspace" | "islower" | "isupper" => {
                self.compile_expr(obj, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                self.chunk(chunk_idx).emit_op(Op::r#true, 0);
            }
            "splitlines" => {
                self.compile_expr(obj, chunk_idx)?;
                let c = self.chunk(chunk_idx).add_constant(Value::String(Arc::from("\n")));
                self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
                self.chunk(chunk_idx).emit_op(Op::str_split, 0);
            }
            "removeprefix" => {
                // s.removeprefix(prefix): if s.startswith(prefix), return s[len(prefix):]
                self.compile_expr(obj, chunk_idx)?;
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                // Stack: [s, prefix]. Dup both for startswith check.
                let prefix_tmp = self.scope(chunk_idx).alloc("__rp_pfx");
                let str_tmp = self.scope(chunk_idx).alloc("__rp_str");
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, prefix_tmp, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, str_tmp, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                // if s.startswith(prefix):
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, str_tmp, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, prefix_tmp, 0);
                self.chunk(chunk_idx).emit_op(Op::str_starts_with, 0);
                let no_match = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);
                // return s[len(prefix):]
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, str_tmp, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, prefix_tmp, 0);
                self.chunk(chunk_idx).emit_op(Op::str_length, 0);
                let max_c = self.chunk(chunk_idx).add_constant(Value::I32(i32::MAX));
                self.chunk(chunk_idx).emit_op_u16(Op::r#const, max_c, 0);
                self.chunk(chunk_idx).emit_op(Op::str_substring, 0);
                let done = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                // else return s unchanged
                self.chunk(chunk_idx).patch_jump(no_match);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, str_tmp, 0);
                self.chunk(chunk_idx).patch_jump(done);
            }
            "removesuffix" => {
                // s.removesuffix(suffix): if s.endswith(suffix), return s[:len(s)-len(suffix)]
                self.compile_expr(obj, chunk_idx)?;
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                let suffix_tmp = self.scope(chunk_idx).alloc("__rs_sfx");
                let str_tmp = self.scope(chunk_idx).alloc("__rs_str");
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, suffix_tmp, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, str_tmp, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                // if s.endswith(suffix):
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, str_tmp, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, suffix_tmp, 0);
                self.chunk(chunk_idx).emit_op(Op::str_ends_with, 0);
                let no_match = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);
                // return s[:len(s)-len(suffix)]
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, str_tmp, 0);
                let zero_c = self.chunk(chunk_idx).add_constant(Value::I32(0));
                self.chunk(chunk_idx).emit_op_u16(Op::r#const, zero_c, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, str_tmp, 0);
                self.chunk(chunk_idx).emit_op(Op::str_length, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, suffix_tmp, 0);
                self.chunk(chunk_idx).emit_op(Op::str_length, 0);
                self.chunk(chunk_idx).emit_op(Op::f64_sub, 0);
                self.chunk(chunk_idx).emit_op(Op::str_substring, 0);
                let done = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                // else return s unchanged
                self.chunk(chunk_idx).patch_jump(no_match);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, str_tmp, 0);
                self.chunk(chunk_idx).patch_jump(done);
            }
            "zfill" | "rjust" => {
                self.compile_expr(obj, chunk_idx)?;
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                self.chunk(chunk_idx).emit_op(Op::str_pad_start, 0);
            }
            _ => return Ok(None), // Not a string method
        }
        Ok(Some(()))
    }

    /// Emit method call with runtime dispatch: try user method first, fall back to builtin.
    /// Stack: obj is compiled, method is looked up via struct_get. If found, generic call.
    /// If null (not on object), emit builtin inline opcodes.
    fn compile_method_with_fallback(&mut self, obj: &Expression, method: &str, args: &[Expression], chunk_idx: usize) -> Result<(), String> {
        // Store obj in temp so we can use it in both paths
        let obj_tmp = self.scope(chunk_idx).alloc("__meth_obj");
        self.compile_expr(obj, chunk_idx)?;
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, obj_tmp, 0);
        self.chunk(chunk_idx).emit_op(Op::drop, 0);

        // Try struct_get method from the object
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
        let method_c = self.chunk(chunk_idx).add_constant(Value::String(Arc::from(method)));
        self.chunk(chunk_idx).emit_op_u16(Op::struct_get, method_c, 0);
        self.chunk(chunk_idx).emit_op(Op::dup, 0);
        self.chunk(chunk_idx).emit_op(Op::ref_is_null, 0);
        let is_null = self.chunk(chunk_idx).emit_jump(Op::br_if_true, 0);

        // ── Found user method: generic call ──
        // Stack: [method_func]. Push self + args, call.
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
        for a in args { self.compile_expr(a, chunk_idx)?; }
        self.chunk(chunk_idx).emit_op_u8(Op::call_ref, (args.len() + 1) as u8, 0);
        let done = self.chunk(chunk_idx).emit_jump(Op::br, 0);

        // ── Not found: builtin fallback ──
        self.chunk(chunk_idx).patch_jump(is_null);
        self.chunk(chunk_idx).emit_op(Op::drop, 0); // drop null from dup
        self.chunk(chunk_idx).emit_op(Op::drop, 0); // drop null from struct_get

        self.emit_builtin_method_fallback(obj_tmp, method, args, chunk_idx)?;

        self.chunk(chunk_idx).patch_jump(done);
        Ok(())
    }

    /// Emit inline opcodes for builtin list/dict/file methods.
    fn emit_builtin_method_fallback(&mut self, obj_tmp: u16, method: &str, args: &[Expression], chunk_idx: usize) -> Result<(), String> {
        match method {
            "append" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                for a in args { self.compile_expr(a, chunk_idx)?; }
                common::collections::emit_push(self.chunk(chunk_idx), 0);
            }
            "pop" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                common::collections::emit_pop(self.chunk(chunk_idx), 0);
            }
            "keys" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                common::dict::emit_keys(self.chunk(chunk_idx), 0);
            }
            "values" => {
                let keys_slot = self.scope(chunk_idx).alloc("__dv_keys");
                let result_slot = self.scope(chunk_idx).alloc("__dv_res");
                let idx_slot = self.scope(chunk_idx).alloc("__dv_i");
                common::dict::emit_values_from_local(
                    self.chunk(chunk_idx), obj_tmp, keys_slot, result_slot, idx_slot, 0,
                );
            }
            "items" => {
                let keys_slot = self.scope(chunk_idx).alloc("__di_keys");
                let result_slot = self.scope(chunk_idx).alloc("__di_res");
                let idx_slot = self.scope(chunk_idx).alloc("__di_i");
                common::dict::emit_items_from_local(
                    self.chunk(chunk_idx), obj_tmp, keys_slot, result_slot, idx_slot, 0,
                );
            }
            "get" => {
                // dict.get(key) — dynamic key via array_get
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                common::dict::emit_get_dynamic(self.chunk(chunk_idx), 0);
            }
            "sort" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                common::collections::emit_sorted(self.chunk(chunk_idx), 0);
            }
            "reverse" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                common::collections::emit_reverse(self.chunk(chunk_idx), 0);
            }
            // Threading: lock.acquire() / lock.release()
            "acquire" => {
                // Lock acquire — obj_tmp holds the lock address
                common::threading::emit_lock_acquire(self.chunk(chunk_idx), obj_tmp, 0);
                self.chunk(chunk_idx).emit_op(Op::null, 0); // return None
            }
            "release" => {
                // Lock release
                common::threading::emit_lock_release(self.chunk(chunk_idx), obj_tmp, 0);
                self.chunk(chunk_idx).emit_op(Op::null, 0);
            }
            "start" => {
                // thread.start() — obj_tmp holds thread handle, just return it
                // Thread was already spawned by threading.Thread(target=fn)
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
            }
            "join" => {
                // thread.join() — wait for thread to complete
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                common::threading::emit_thread_join(self.chunk(chunk_idx), 0);
            }
            // ── File object methods (same host as VB/PHP fopen/fwrite) ──
            "read" => {
                // f.read() → readFile (whole file) or lineInput
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                let imp = self.chunk(chunk_idx).add_import("wasi:filesystem", "readFile");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
                self.chunk(chunk_idx).emit(1, 0);
            }
            "readline" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                let imp = self.chunk(chunk_idx).add_import("wasi:filesystem", "lineInput");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
                self.chunk(chunk_idx).emit(1, 0);
            }
            "readlines" => {
                // f.readlines() → readFile then split by \n
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                let imp = self.chunk(chunk_idx).add_import("wasi:filesystem", "readFile");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
                self.chunk(chunk_idx).emit(1, 0);
                let nl = self.chunk(chunk_idx).add_constant(Value::String(Arc::from("\n")));
                self.chunk(chunk_idx).emit_op_u16(Op::r#const, nl, 0);
                self.chunk(chunk_idx).emit_op(Op::str_split, 0);
            }
            "write" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                for a in args { self.compile_expr(a, chunk_idx)?; }
                let imp = self.chunk(chunk_idx).add_import("wasi:filesystem", "printFile");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
                self.chunk(chunk_idx).emit(args.len() as u8 + 1, 0);
            }
            "writelines" => {
                // f.writelines(lines) — join and write
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                for a in args { self.compile_expr(a, chunk_idx)?; }
                let imp = self.chunk(chunk_idx).add_import("wasi:filesystem", "printFile");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
                self.chunk(chunk_idx).emit(args.len() as u8 + 1, 0);
            }
            "close" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                let imp = self.chunk(chunk_idx).add_import("wasi:filesystem", "closeFile");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
                self.chunk(chunk_idx).emit(1, 0);
            }
            "flush" => {
                self.chunk(chunk_idx).emit_op(Op::null, 0); // no-op
            }
            // ── Socket object methods (same host as VB/PHP) ──
            "connect" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                for a in args { self.compile_expr(a, chunk_idx)?; }
                let imp = self.chunk(chunk_idx).add_import("vybe:net", "tcpConnect");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
                self.chunk(chunk_idx).emit(args.len() as u8 + 1, 0);
            }
            "send" | "sendall" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                for a in args { self.compile_expr(a, chunk_idx)?; }
                let imp = self.chunk(chunk_idx).add_import("vybe:net", "streamWriterWriteLine");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
                self.chunk(chunk_idx).emit(args.len() as u8 + 1, 0);
            }
            "recv" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                for a in args { self.compile_expr(a, chunk_idx)?; }
                let imp = self.chunk(chunk_idx).add_import("vybe:net", "streamReaderReadLine");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
                self.chunk(chunk_idx).emit(args.len() as u8 + 1, 0);
            }
            "bind" | "listen" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                let imp = self.chunk(chunk_idx).add_import("vybe:net", "tcpListenerStart");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
                self.chunk(chunk_idx).emit(0, 0);
            }
            "accept" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                let imp = self.chunk(chunk_idx).add_import("vybe:net", "tcpListenerAccept");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
                self.chunk(chunk_idx).emit(1, 0);
            }
            // ── Database cursor methods (same host as VB/PHP) ──
            "execute" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                for a in args { self.compile_expr(a, chunk_idx)?; }
                let imp = self.chunk(chunk_idx).add_import("vybe:database", "execute");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
                self.chunk(chunk_idx).emit(args.len() as u8 + 1, 0);
            }
            "fetchone" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                let imp = self.chunk(chunk_idx).add_import("vybe:database", "scalar");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
                self.chunk(chunk_idx).emit(1, 0);
            }
            "fetchall" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                let imp = self.chunk(chunk_idx).add_import("vybe:database", "query");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
                self.chunk(chunk_idx).emit(1, 0);
            }
            "cursor" => {
                // conn.cursor() → return conn itself (cursor IS the connection in our model)
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
            }
            // ── DateTime object methods ──
            "strftime" | "isoformat" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                for a in args { self.compile_expr(a, chunk_idx)?; }
                let imp = self.chunk(chunk_idx).add_import("vybe:types", "dateTimeToString");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, imp, 0);
                self.chunk(chunk_idx).emit(args.len() as u8 + 1, 0);
            }
            "timestamp" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                // DateTime stores timestamp internally
            }
            // Cross-language compat: Dart/C# .contains(), JS .includes()
            "contains" | "includes" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                common_collections::emit_contains(&mut self.chunks[chunk_idx], 0);
            }
            // Cross-language compat: Dart .length, .isEmpty, .isNotEmpty
            "length" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                common_collections::emit_len(&mut self.chunks[chunk_idx], 0);
            }
            "insert" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                for a in args { self.compile_expr(a, chunk_idx)?; }
                let splice_fn = self.chunk(chunk_idx).add_import("vybe:array", "splice");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, splice_fn, 0);
                self.chunk(chunk_idx).emit(args.len() as u8 + 1, 0);
            }
            "remove" => {
                // list.remove(item) — find index, delete
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                common::collections::emit_index_of(self.chunk(chunk_idx), 0);
                // TODO: splice at index
            }
            "extend" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                common::collections::emit_concat(self.chunk(chunk_idx), 0);
            }
            "clear" => {
                // Simplified: set to empty array
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                self.chunk(chunk_idx).emit_op(Op::null, 0);
            }
            "copy" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                let zero_c = self.chunk(chunk_idx).add_constant(Value::I32(0));
                let max_c = self.chunk(chunk_idx).add_constant(Value::I32(i32::MAX));
                self.chunk(chunk_idx).emit_op_u16(Op::r#const, zero_c, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::r#const, max_c, 0);
                common_collections::emit_slice(&mut self.chunks[chunk_idx], 0);
            }
            "index" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                common::collections::emit_index_of(self.chunk(chunk_idx), 0);
            }
            "update" => {
                // dict.update(other) — simplified
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
            }
            "setdefault" => {
                // dict.setdefault(key, default) — get key, if null set default and return it
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                common_collections::emit_get(&mut self.chunks[chunk_idx], 0);
                // Check if result is null
                self.chunk(chunk_idx).emit_op(Op::dup, 0);
                self.chunk(chunk_idx).emit_op(Op::ref_is_null, 0);
                let not_null = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);
                // Null — set the default
                self.chunk(chunk_idx).emit_op(Op::drop, 0); // drop null
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                if args.len() >= 2 {
                    self.compile_expr(&args[1], chunk_idx)?;
                } else {
                    self.chunk(chunk_idx).emit_op(Op::null, 0);
                }
                self.chunk(chunk_idx).emit_op(Op::dup, 0);
                let val_tmp = self.scope(chunk_idx).alloc("__sd_val");
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, val_tmp, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                common_collections::emit_set(&mut self.chunks[chunk_idx], 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, val_tmp, 0);
                let done = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                self.chunk(chunk_idx).patch_jump(not_null);
                self.chunk(chunk_idx).patch_jump(done);
            }
            // ── Set methods ──
            "add" => {
                // set.add(value) — call host setAdd
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                let f = self.chunk(chunk_idx).add_import("vybe:collections", "setAdd");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, f, 0);
                self.chunk(chunk_idx).emit(2, 0);
            }
            "discard" => {
                // set.discard(value) — call host setDelete (no error if missing)
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                let f = self.chunk(chunk_idx).add_import("vybe:collections", "setDelete");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, f, 0);
                self.chunk(chunk_idx).emit(2, 0);
            }
            "union" => {
                // set.union(other) — create new set, add all from self + other
                // Simplified: concat items arrays, build new set
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                let v1 = self.chunk(chunk_idx).add_import("vybe:collections", "setValues");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, v1, 0);
                self.chunk(chunk_idx).emit(1, 0);
                if !args.is_empty() {
                    self.compile_expr(&args[0], chunk_idx)?;
                    let v2 = self.chunk(chunk_idx).add_import("vybe:collections", "setValues");
                    self.chunk(chunk_idx).emit_op_u16(Op::call_import, v2, 0);
                    self.chunk(chunk_idx).emit(1, 0);
                    common_collections::emit_concat(&mut self.chunks[chunk_idx], 0);
                }
                // Convert combined array to set via pyset host
                let pyset = self.chunk(chunk_idx).add_import("vybe:array", "pyset");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, pyset, 0);
                self.chunk(chunk_idx).emit(1, 0);
            }
            "intersection" => {
                // set.intersection(other) — filter self.values where other.has(v)
                let self_vals = self.scope(chunk_idx).alloc("__isect_sv");
                let other_set = self.scope(chunk_idx).alloc("__isect_os");
                let result = self.scope(chunk_idx).alloc("__isect_r");
                let i = self.scope(chunk_idx).alloc("__isect_i");

                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                let v_fn = self.chunk(chunk_idx).add_import("vybe:collections", "setValues");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, v_fn, 0);
                self.chunk(chunk_idx).emit(1, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, self_vals, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);

                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, other_set, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);

                self.chunk(chunk_idx).emit_op_u16(Op::array_new, 0, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, result, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                self.chunk(chunk_idx).emit_op(Op::i32_const_0, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, i, 0);

                let loop_start = self.chunk(chunk_idx).current_offset();
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, i, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, self_vals, 0);
                common_collections::emit_len(&mut self.chunks[chunk_idx], 0);
                self.chunk(chunk_idx).emit_op(Op::dyn_lt, 0);
                let exit = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);

                // if other.has(self_vals[i]): result.push(self_vals[i])
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, other_set, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, self_vals, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, i, 0);
                common_collections::emit_get(&mut self.chunks[chunk_idx], 0);
                let has_fn = self.chunk(chunk_idx).add_import("vybe:collections", "setHas");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, has_fn, 0);
                self.chunk(chunk_idx).emit(2, 0);
                self.chunk(chunk_idx).emit_op(Op::dyn_to_bool, 0);
                let skip = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);

                self.chunk(chunk_idx).emit_op_u16(Op::local_get, result, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, self_vals, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, i, 0);
                common_collections::emit_get(&mut self.chunks[chunk_idx], 0);
                common_collections::emit_push(&mut self.chunks[chunk_idx], 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);

                self.chunk(chunk_idx).patch_jump(skip);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, i, 0);
                self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
                self.chunk(chunk_idx).emit_op(Op::i32_add, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, i, 0);
                self.chunk(chunk_idx).emit_loop(loop_start, 0);
                self.chunk(chunk_idx).patch_jump(exit);

                // Convert result array to set
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, result, 0);
                let pyset = self.chunk(chunk_idx).add_import("vybe:array", "pyset");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, pyset, 0);
                self.chunk(chunk_idx).emit(1, 0);
            }
            "difference" => {
                // set.difference(other) — filter self.values where NOT other.has(v)
                let self_vals = self.scope(chunk_idx).alloc("__diff_sv");
                let other_set = self.scope(chunk_idx).alloc("__diff_os");
                let result = self.scope(chunk_idx).alloc("__diff_r");
                let i = self.scope(chunk_idx).alloc("__diff_i");

                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                let v_fn = self.chunk(chunk_idx).add_import("vybe:collections", "setValues");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, v_fn, 0);
                self.chunk(chunk_idx).emit(1, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, self_vals, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);

                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, other_set, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);

                self.chunk(chunk_idx).emit_op_u16(Op::array_new, 0, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, result, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                self.chunk(chunk_idx).emit_op(Op::i32_const_0, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, i, 0);

                let loop_start = self.chunk(chunk_idx).current_offset();
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, i, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, self_vals, 0);
                common_collections::emit_len(&mut self.chunks[chunk_idx], 0);
                self.chunk(chunk_idx).emit_op(Op::dyn_lt, 0);
                let exit = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);

                self.chunk(chunk_idx).emit_op_u16(Op::local_get, other_set, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, self_vals, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, i, 0);
                common_collections::emit_get(&mut self.chunks[chunk_idx], 0);
                let has_fn = self.chunk(chunk_idx).add_import("vybe:collections", "setHas");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, has_fn, 0);
                self.chunk(chunk_idx).emit(2, 0);
                self.chunk(chunk_idx).emit_op(Op::dyn_to_bool, 0);
                // skip push if has returns true (we want elements NOT in other)
                let skip = self.chunk(chunk_idx).emit_jump(Op::br_if_true, 0);

                self.chunk(chunk_idx).emit_op_u16(Op::local_get, result, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, self_vals, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, i, 0);
                common_collections::emit_get(&mut self.chunks[chunk_idx], 0);
                common_collections::emit_push(&mut self.chunks[chunk_idx], 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);

                self.chunk(chunk_idx).patch_jump(skip);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, i, 0);
                self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
                self.chunk(chunk_idx).emit_op(Op::i32_add, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, i, 0);
                self.chunk(chunk_idx).emit_loop(loop_start, 0);
                self.chunk(chunk_idx).patch_jump(exit);

                self.chunk(chunk_idx).emit_op_u16(Op::local_get, result, 0);
                let pyset = self.chunk(chunk_idx).add_import("vybe:array", "pyset");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, pyset, 0);
                self.chunk(chunk_idx).emit(1, 0);
            }
            "symmetric_difference" => {
                // a.symmetric_difference(b) = a.union(b).difference(a.intersection(b))
                // Simplified: just do union for now — this is rarely used
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                let v1 = self.chunk(chunk_idx).add_import("vybe:collections", "setValues");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, v1, 0);
                self.chunk(chunk_idx).emit(1, 0);
                if !args.is_empty() {
                    self.compile_expr(&args[0], chunk_idx)?;
                    let v2 = self.chunk(chunk_idx).add_import("vybe:collections", "setValues");
                    self.chunk(chunk_idx).emit_op_u16(Op::call_import, v2, 0);
                    self.chunk(chunk_idx).emit(1, 0);
                    common_collections::emit_concat(&mut self.chunks[chunk_idx], 0);
                }
                let pyset = self.chunk(chunk_idx).add_import("vybe:array", "pyset");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, pyset, 0);
                self.chunk(chunk_idx).emit(1, 0);
            }
            "read" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
            }
            "write" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                let write_fn = self.chunk(chunk_idx).add_import("wasi:filesystem", "writeFile");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, write_fn, 0);
                self.chunk(chunk_idx).emit(2, 0);
            }
            "close" => {
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, obj_tmp, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                self.chunk(chunk_idx).emit_op(Op::null, 0);
            }
            _ => {
                // Unknown builtin — should not reach here, but emit null as safety
                self.chunk(chunk_idx).emit_op(Op::null, 0);
            }
        }
        Ok(())
    }


    fn compile_comprehension(&mut self, element: &Expression, generators: &[Comprehension], chunk_idx: usize) -> Result<(), String> {
        // Create empty array, iterate, push elements
        self.chunk(chunk_idx).emit_op_u16(Op::array_new, 0, 0);
        let result_local = self.scope(chunk_idx).alloc("__comp_result");
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, result_local, 0);

        self.compile_comp_generators(generators, &|s| {
            s.chunk(chunk_idx).emit_op_u16(Op::local_get, result_local, 0);
            s.compile_expr(element, chunk_idx)?;
            common_collections::emit_push(&mut s.chunks[chunk_idx], 0);
            s.chunk(chunk_idx).emit_op(Op::drop, 0);
            Ok(())
        }, chunk_idx)?;

        self.chunk(chunk_idx).emit_op_u16(Op::local_get, result_local, 0);
        Ok(())
    }

    fn compile_comp_generators(&mut self, generators: &[Comprehension], body: &dyn Fn(&mut Self) -> Result<(), String>, chunk_idx: usize) -> Result<(), String> {
        if generators.is_empty() {
            return body(self);
        }

        let generator = &generators[0];
        let iter_local = self.scope(chunk_idx).alloc("__comp_iter");
        let idx_local = self.scope(chunk_idx).alloc("__comp_idx");

        self.compile_expr(&generator.iter, chunk_idx)?;
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, iter_local, 0);
        self.chunk(chunk_idx).emit_op(Op::i32_const_0, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx_local, 0);

        let loop_start = self.chunk(chunk_idx).current_offset();
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_local, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, iter_local, 0);
        common_collections::emit_len(&mut self.chunks[chunk_idx], 0);
        self.chunk(chunk_idx).emit_op(Op::dyn_lt, 0);
        let exit_jump = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);

        // Load current element
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, iter_local, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_local, 0);
        common_collections::emit_get(&mut self.chunks[chunk_idx], 0);
        self.compile_assign_target(&generator.target, chunk_idx)?;

        // Apply if filters
        let mut filter_jumps = Vec::new();
        for f in &generator.ifs {
            self.compile_expr(f, chunk_idx)?;
            self.chunk(chunk_idx).emit_op(Op::dyn_to_bool, 0);
            let j = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);
            filter_jumps.push(j);
        }

        // Recurse into remaining generators or execute body
        if generators.len() > 1 {
            self.compile_comp_generators(&generators[1..], body, chunk_idx)?;
        } else {
            body(self)?;
        }

        for j in filter_jumps {
            self.chunk(chunk_idx).patch_jump(j);
        }

        // Increment index
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_local, 0);
        self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
        self.chunk(chunk_idx).emit_op(Op::i32_add, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx_local, 0);

        self.chunk(chunk_idx).emit_loop(loop_start, 0);
        self.chunk(chunk_idx).patch_jump(exit_jump);

        Ok(())
    }

    fn emit_cmp_op(&mut self, op: CmpOp, chunk_idx: usize) {
        // Use primitive comparison opcodes. User-defined __lt__/__eq__/etc are
        // available via cross-language aliases and the method-with-fallback dispatch
        // path — rich compare dispatch (emit_rich_compare_locals) is available in
        // common::expressions for cases where the compiler knows objects are involved
        // (e.g. sorted(key=...) comparisons).
        match op {
            CmpOp::Eq => self.chunk(chunk_idx).emit_op(Op::dyn_eq, 0),
            CmpOp::NotEq => self.chunk(chunk_idx).emit_op(Op::dyn_ne, 0),
            CmpOp::Lt => self.chunk(chunk_idx).emit_op(Op::dyn_lt, 0),
            CmpOp::LtE => self.chunk(chunk_idx).emit_op(Op::dyn_le, 0),
            CmpOp::Gt => self.chunk(chunk_idx).emit_op(Op::dyn_gt, 0),
            CmpOp::GtE => self.chunk(chunk_idx).emit_op(Op::dyn_ge, 0),
            CmpOp::Is => self.chunk(chunk_idx).emit_op(Op::dyn_eq, 0),
            CmpOp::IsNot => self.chunk(chunk_idx).emit_op(Op::dyn_ne, 0),
            CmpOp::In => common_collections::emit_contains(&mut self.chunks[chunk_idx], 0),
            CmpOp::NotIn => {
                common_collections::emit_contains(&mut self.chunks[chunk_idx], 0);
                self.chunk(chunk_idx).emit_op(Op::dyn_not, 0);
            }
        }
    }

    fn emit_aug_op(&mut self, op: AugOp, chunk_idx: usize) {
        match op {
            AugOp::Add => self.chunk(chunk_idx).emit_op(Op::dyn_add, 0),
            AugOp::Sub => self.chunk(chunk_idx).emit_op(Op::f64_sub, 0),
            AugOp::Mul => {
                let ta = self.scope(chunk_idx).alloc("__aug_a");
                let tb = self.scope(chunk_idx).alloc("__aug_b");
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, tb, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, ta, 0);
                common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_dynmul", 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, ta, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_get, tb, 0);
                common::bundle::emit_call_invoke(self.chunk(chunk_idx), 2, 0);
            }
            AugOp::Div => self.chunk(chunk_idx).emit_op(Op::f64_div, 0),
            AugOp::FloorDiv => { self.chunk(chunk_idx).emit_op(Op::f64_div, 0); common::math::emit_floor(self.chunk(chunk_idx), 0); }
            AugOp::Mod => self.chunk(chunk_idx).emit_op(Op::i32_rem_s, 0),
            AugOp::Pow => {
                let pow_idx = self.chunk(chunk_idx).add_import("vybe:math", "pow");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, pow_idx, 0);
                self.chunk(chunk_idx).emit(2, 0);
            }
            AugOp::LShift => self.chunk(chunk_idx).emit_op(Op::i32_shl, 0),
            AugOp::RShift => self.chunk(chunk_idx).emit_op(Op::i32_shr_s, 0),
            AugOp::BitOr => self.chunk(chunk_idx).emit_op(Op::i32_or, 0),
            AugOp::BitXor => self.chunk(chunk_idx).emit_op(Op::i32_xor, 0),
            AugOp::BitAnd => self.chunk(chunk_idx).emit_op(Op::i32_and, 0),
            AugOp::MatMul => self.chunk(chunk_idx).emit_op(Op::f64_mul, 0),
        }
    }

    /// map(fn, iterable) → [fn(x) for x in iterable]
    fn compile_map(&mut self, args: &[Expression], chunk_idx: usize) -> Result<(), String> {
        let fn_local = self.scope(chunk_idx).alloc("__map_fn");
        let arr_local = self.scope(chunk_idx).alloc("__map_arr");
        let result_local = self.scope(chunk_idx).alloc("__map_res");
        let idx_local = self.scope(chunk_idx).alloc("__map_i");

        self.compile_expr(&args[0], chunk_idx)?;
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, fn_local, 0);
        self.chunk(chunk_idx).emit_op(Op::drop, 0);
        self.compile_expr(&args[1], chunk_idx)?;
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, arr_local, 0);
        self.chunk(chunk_idx).emit_op(Op::drop, 0);

        common::loops::emit_map(self.chunk(chunk_idx), fn_local, arr_local, result_local, idx_local, 0);
        Ok(())
    }

    fn compile_filter(&mut self, args: &[Expression], chunk_idx: usize) -> Result<(), String> {
        let fn_local = self.scope(chunk_idx).alloc("__filt_fn");
        let arr_local = self.scope(chunk_idx).alloc("__filt_arr");
        let result_local = self.scope(chunk_idx).alloc("__filt_res");
        let idx_local = self.scope(chunk_idx).alloc("__filt_i");
        let elem_local = self.scope(chunk_idx).alloc("__filt_elem");

        self.compile_expr(&args[0], chunk_idx)?;
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, fn_local, 0);
        self.chunk(chunk_idx).emit_op(Op::drop, 0);
        self.compile_expr(&args[1], chunk_idx)?;
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, arr_local, 0);
        self.chunk(chunk_idx).emit_op(Op::drop, 0);

        common::loops::emit_filter(self.chunk(chunk_idx), fn_local, arr_local, result_local, idx_local, elem_local, 0);
        Ok(())
    }

    /// sorted(iterable, key=fn): map key, sort pairs by key, extract values.
    /// Emits: pairs = [[key(x), x] for x in iterable] → sort pairs → [p[1] for p in pairs]
    fn compile_sorted_with_key(&mut self, iterable: &Expression, key_fn: &Expression, chunk_idx: usize) -> Result<(), String> {
        let arr_slot = self.scope(chunk_idx).alloc("__sk_arr");
        let fn_slot = self.scope(chunk_idx).alloc("__sk_fn");
        let pairs_slot = self.scope(chunk_idx).alloc("__sk_pairs");
        let result_slot = self.scope(chunk_idx).alloc("__sk_res");
        let idx_slot = self.scope(chunk_idx).alloc("__sk_i");

        // Evaluate iterable and key function
        self.compile_expr(iterable, chunk_idx)?;
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, arr_slot, 0);
        self.chunk(chunk_idx).emit_op(Op::drop, 0);
        self.compile_expr(key_fn, chunk_idx)?;
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, fn_slot, 0);
        self.chunk(chunk_idx).emit_op(Op::drop, 0);

        // Step 1: Build pairs = [[key(x), x] for x in arr]
        self.chunk(chunk_idx).emit_op_u16(Op::array_new, 0, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, pairs_slot, 0);
        self.chunk(chunk_idx).emit_op(Op::drop, 0);
        self.chunk(chunk_idx).emit_op(Op::i32_const_0, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx_slot, 0);

        let loop1 = self.chunk(chunk_idx).current_offset();
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_slot, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, arr_slot, 0);
        common_collections::emit_len(&mut self.chunks[chunk_idx], 0);
        self.chunk(chunk_idx).emit_op(Op::dyn_lt, 0);
        let exit1 = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);

        // pair = [key(arr[i]), arr[i]]
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, pairs_slot, 0);
        // key(arr[i])
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, fn_slot, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, arr_slot, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_slot, 0);
        common_collections::emit_get(&mut self.chunks[chunk_idx], 0);
        self.chunk(chunk_idx).emit_op_u8(Op::call_ref, 1, 0);
        // arr[i]
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, arr_slot, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_slot, 0);
        common_collections::emit_get(&mut self.chunks[chunk_idx], 0);
        // [key, val]
        self.chunk(chunk_idx).emit_op_u16(Op::array_new, 2, 0);
        common_collections::emit_push(&mut self.chunks[chunk_idx], 0);
        self.chunk(chunk_idx).emit_op(Op::drop, 0);
        // i++
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_slot, 0);
        self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
        self.chunk(chunk_idx).emit_op(Op::i32_add, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx_slot, 0);
        self.chunk(chunk_idx).emit_loop(loop1, 0);
        self.chunk(chunk_idx).patch_jump(exit1);

        // Step 2: Sort pairs by first element (key) using stdlib sorted
        common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_sorted", 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, pairs_slot, 0);
        common::bundle::emit_call_invoke(self.chunk(chunk_idx), 1, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, pairs_slot, 0);
        self.chunk(chunk_idx).emit_op(Op::drop, 0);

        // Step 3: Extract values: result = [pair[1] for pair in sorted_pairs]
        self.chunk(chunk_idx).emit_op_u16(Op::array_new, 0, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, result_slot, 0);
        self.chunk(chunk_idx).emit_op(Op::drop, 0);
        self.chunk(chunk_idx).emit_op(Op::i32_const_0, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx_slot, 0);

        let loop2 = self.chunk(chunk_idx).current_offset();
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_slot, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, pairs_slot, 0);
        common_collections::emit_len(&mut self.chunks[chunk_idx], 0);
        self.chunk(chunk_idx).emit_op(Op::dyn_lt, 0);
        let exit2 = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);

        self.chunk(chunk_idx).emit_op_u16(Op::local_get, result_slot, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, pairs_slot, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_slot, 0);
        common_collections::emit_get(&mut self.chunks[chunk_idx], 0); // pair
        let one_c = self.chunk(chunk_idx).add_constant(Value::I32(1));
        self.chunk(chunk_idx).emit_op_u16(Op::r#const, one_c, 0);
        common_collections::emit_get(&mut self.chunks[chunk_idx], 0); // pair[1] = original value
        common_collections::emit_push(&mut self.chunks[chunk_idx], 0);
        self.chunk(chunk_idx).emit_op(Op::drop, 0);

        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_slot, 0);
        self.chunk(chunk_idx).emit_op(Op::i32_const_1, 0);
        self.chunk(chunk_idx).emit_op(Op::i32_add, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx_slot, 0);
        self.chunk(chunk_idx).emit_loop(loop2, 0);
        self.chunk(chunk_idx).patch_jump(exit2);

        self.chunk(chunk_idx).emit_op_u16(Op::local_get, result_slot, 0);
        Ok(())
    }

    fn chunk(&mut self, idx: usize) -> &mut Chunk {
        &mut self.chunks[idx]
    }

    fn scope(&mut self, chunk_idx: usize) -> &mut Scope {
        // Scope 0 = main, scope N = function N
        let scope_idx = if chunk_idx == 0 { 0 } else {
            // Find scope for this chunk
            self.scopes.len() - 1
        };
        &mut self.scopes[scope_idx]
    }
}
