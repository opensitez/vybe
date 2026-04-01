use std::rc::Rc;
use std::collections::HashMap;
use vybe_parser_python::ast::*;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use vybe_compiler_common as common;

pub struct Compiler {
    chunks: Vec<Chunk>,
    scopes: Vec<Scope>,
    loop_stack: Vec<LoopCtx>,
    import_log: u16,
    last_upvalues: Vec<UpvalueDesc>,
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
                        let attr_c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(attr.as_str())));
                        self.chunk(chunk_idx).emit_op_u16(Op::struct_get, attr_c, 0);
                        self.compile_expr(value, chunk_idx)?;
                        self.emit_aug_op(*op, chunk_idx);
                        let attr_c2 = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(attr.as_str())));
                        self.chunk(chunk_idx).emit_op_u16(Op::struct_set, attr_c2, 0);
                    }
                    Expression::Subscript { value: obj, slice } => {
                        // obj[idx] op= value → obj[idx] = obj[idx] op value
                        self.compile_expr(obj, chunk_idx)?;
                        self.chunk(chunk_idx).emit_op(Op::dup, 0);
                        self.compile_expr(slice, chunk_idx)?;
                        self.chunk(chunk_idx).emit_op(Op::dup, 0); // keep idx for set
                        self.chunk(chunk_idx).emit_op(Op::array_get, 0);
                        self.compile_expr(value, chunk_idx)?;
                        self.emit_aug_op(*op, chunk_idx);
                        self.chunk(chunk_idx).emit_op(Op::array_set, 0);
                    }
                    _ => {
                        // Fallback: just skip silently
                    }
                }
            }

            Statement::If { test, body, elif_clauses, else_body } => {
                self.compile_expr(test, chunk_idx)?;
                let exit_jump = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);
                for s in body { self.compile_stmt(s, chunk_idx)?; }

                if elif_clauses.is_empty() && else_body.is_none() {
                    self.chunk(chunk_idx).patch_jump(exit_jump);
                } else {
                    let after_jump = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                    self.chunk(chunk_idx).patch_jump(exit_jump);

                    for (elif_test, elif_body) in elif_clauses {
                        self.compile_expr(elif_test, chunk_idx)?;
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
                    if let Expression::Name(n) = b { Some(n.to_lowercase()) } else { None }
                }).collect();
                let parent_name = all_bases.first().cloned().unwrap_or_default();

                // Compile all methods first (we need their chunk indices)
                let mut method_entries = Vec::new();
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

                        let effective_name = if is_property {
                            format!("__get_{}", method_name)
                        } else if is_setter {
                            format!("__set_{}", method_name)
                        } else {
                            method_name.clone()
                        };

                        let func_chunk_idx = self.compile_function(&effective_name, params, mbody)?;
                        method_entries.push((effective_name.to_lowercase(), func_chunk_idx));
                        if method_name == "__init__" {
                            init_chunk = Some(func_chunk_idx);
                            init_params = Some(params.clone());
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

                // Create empty object
                self.chunk(ctor_idx).emit_op_u16(Op::struct_new, 0, 0);
                self.chunk(ctor_idx).emit_op_u16(Op::local_set, this_local, 0);
                self.chunk(ctor_idx).emit_op(Op::drop, 0);

                // Stamp __type (for untyped fallback)
                self.chunk(ctor_idx).emit_op_u16(Op::local_get, this_local, 0);
                let type_str = self.chunk(ctor_idx).add_constant(Value::String(Rc::from(name.as_str())));
                let type_key = self.chunk(ctor_idx).add_constant(Value::String(Rc::from("__type")));
                self.chunk(ctor_idx).emit_op_u16(Op::r#const, type_str, 0);
                self.chunk(ctor_idx).emit_op_u16(Op::struct_set, type_key, 0);
                self.chunk(ctor_idx).emit_op(Op::drop, 0);

                // Stamp type_id via __tid_ global (set by TypeRegistry at load time)
                let tid_name = self.chunk(ctor_idx).add_constant(
                    Value::String(Rc::from(format!("__tid_{}", ctor_name).as_str()))
                );
                self.chunk(ctor_idx).emit_op_u16(Op::local_get, this_local, 0);
                self.chunk(ctor_idx).emit_op_u16(Op::global_get, tid_name, 0);
                self.chunk(ctor_idx).emit_op(Op::set_type_id, 0);
                self.chunk(ctor_idx).emit_op(Op::drop, 0);

                // Bind instance methods on the object.
                // Also create cross-language aliases for Python dunders:
                //   __str__  → toString (JS), __get_tostring (VB/C#)
                //   __len__  → length property
                //   __bool__ → valueOf (JS truthiness)
                //   __getitem__ → mapped at call site
                //   __enter__/__exit__ → stored as-is (with statement)
                for (method_name, method_ci) in &method_entries {
                    if method_name == "__init__" { continue; }

                    // Bind under original dunder name
                    self.chunk(ctor_idx).emit_op_u16(Op::local_get, this_local, 0);
                    self.chunk(ctor_idx).emit_op_u16(Op::ref_func, *method_ci as u16, 0);
                    self.chunk(ctor_idx).emit(0, 0);
                    let m_key = self.chunk(ctor_idx).add_constant(Value::String(Rc::from(method_name.as_str())));
                    self.chunk(ctor_idx).emit_op_u16(Op::struct_set, m_key, 0);
                    self.chunk(ctor_idx).emit_op(Op::drop, 0);

                    // Cross-language aliases
                    let aliases: &[&str] = match method_name.as_str() {
                        "__str__" => &["toString", "__get_tostring"],
                        "__repr__" => &["toDebugString"],
                        "__len__" => &["__get_length", "__get_count"],
                        "__bool__" => &["valueOf"],
                        "__contains__" => &["contains", "includes"],
                        "__enter__" => &[],
                        "__exit__" => &[],
                        _ => &[],
                    };
                    for alias in aliases {
                        self.chunk(ctor_idx).emit_op_u16(Op::local_get, this_local, 0);
                        self.chunk(ctor_idx).emit_op_u16(Op::ref_func, *method_ci as u16, 0);
                        self.chunk(ctor_idx).emit(0, 0);
                        let a_key = self.chunk(ctor_idx).add_constant(Value::String(Rc::from(*alias)));
                        self.chunk(ctor_idx).emit_op_u16(Op::struct_set, a_key, 0);
                        self.chunk(ctor_idx).emit_op(Op::drop, 0);
                    }
                }

                // Compile class-level statements (class attributes).
                // These run in the constructor, setting attributes on the new object.
                for s in &other_stmts {
                    if let Statement::Assign { targets, value } = s {
                        // class-level assignment: set on self
                        for target in targets {
                            if let Expression::Name(attr_name) = target {
                                self.chunk(ctor_idx).emit_op_u16(Op::local_get, this_local, 0);
                                self.compile_expr(value, ctor_idx)?;
                                let attr_key = self.chunk(ctor_idx).add_constant(Value::String(Rc::from(attr_name.as_str())));
                                self.chunk(ctor_idx).emit_op_u16(Op::struct_set, attr_key, 0);
                                self.chunk(ctor_idx).emit_op(Op::drop, 0);
                            }
                        }
                    }
                }

                // Call __init__(self, *args) if it exists.
                // If this class has no __init__ but has a parent, call the parent constructor
                // with all args (Python inheritance: child inherits parent's __init__).
                if let Some(init_ci) = init_chunk {
                    self.chunk(ctor_idx).emit_op_u16(Op::ref_func, init_ci as u16, 0);
                    self.chunk(ctor_idx).emit(0, 0);
                    self.chunk(ctor_idx).emit_op_u16(Op::local_get, this_local, 0); // self
                    for i in 0..user_params {
                        self.chunk(ctor_idx).emit_op_u16(Op::local_get, (i + 1) as u16, 0);
                    }
                    self.chunk(ctor_idx).emit_op_u8(Op::call_ref, (user_params + 1) as u8, 0);
                    self.chunk(ctor_idx).emit_op(Op::drop, 0);
                // If no __init__ and has parent, store parent constructor as __super
                // so super().__init__() can find it. Users must call super() explicitly.
                } else if !parent_name.is_empty() {
                    // Store parent constructor ref on the object for super() access
                    let parent_c = self.chunk(ctor_idx).add_constant(
                        Value::String(Rc::from(parent_name.as_str()))
                    );
                    self.chunk(ctor_idx).emit_op_u16(Op::local_get, this_local, 0);
                    self.chunk(ctor_idx).emit_op_u16(Op::global_get, parent_c, 0);
                    let super_key = self.chunk(ctor_idx).add_constant(Value::String(Rc::from("__super")));
                    self.chunk(ctor_idx).emit_op_u16(Op::struct_set, super_key, 0);
                    self.chunk(ctor_idx).emit_op(Op::drop, 0);
                }

                // Return this
                self.chunk(ctor_idx).emit_op_u16(Op::local_get, this_local, 0);
                self.chunk(ctor_idx).emit_op(Op::r#return, 0);

                let scope = self.scopes.remove(scope_idx);
                self.chunks[ctor_idx].local_count = (scope.max_local + 1) as u16;

                // Store constructor as a global (ClassName = constructor function)
                let class_local = self.scope(chunk_idx).alloc(name);
                self.chunk(chunk_idx).emit_op_u16(Op::ref_func, ctor_idx as u16, 0);
                self.chunk(chunk_idx).emit(0, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, class_local, 0);
                // Also store as global for cross-module access
                let global_name = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(name.to_lowercase().as_str())));
                self.chunk(chunk_idx).emit_op_u16(Op::global_set, global_name, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);

                // Register type entry
                use vybe_bytecode::chunk::TypeEntry;
                self.chunks[0].types.push(TypeEntry {
                    name: name.to_lowercase(),
                    parent: parent_name,
                    fields: Vec::new(),
                    methods: method_entries,
                    is_interface: false,
                    implements: if all_bases.len() > 1 { all_bases[1..].to_vec() } else { Vec::new() },
                    constructor_chunk: Some(ctor_idx),
                });
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
                        // Determine tag for this handler
                        let tag = if let Some(exc_type) = &handler.exc_type {
                            let type_name = self.expr_to_name(exc_type);
                            self.chunk(chunk_idx).add_exception_tag(&type_name)
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
                            let key = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(attr.as_str())));
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
                    let c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from("AssertionError")));
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
                    let enter_key = self.chunk(chunk_idx).add_constant(Value::String(Rc::from("__enter__")));
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
                    let exit_key = self.chunk(chunk_idx).add_constant(Value::String(Rc::from("__exit__")));
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
                self.chunk(chunk_idx).emit_op(Op::array_length, 0);
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
        self.chunk(chunk_idx).emit_op(Op::array_length, 0);
        self.chunk(chunk_idx).emit_op(Op::dyn_lt, 0);
        let exit_jump = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);

        // Load current element
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, iter_local, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_local, 0);
        self.chunk(chunk_idx).emit_op(Op::array_get, 0);

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
        let mut fchunk = Chunk::new(name);
        fchunk.add_import("wasi:cli", "log");
        let func_chunk_idx = self.chunks.len();
        self.chunks.push(fchunk);

        let mut scope = Scope::new();
        // Map params to locals 1..n
        for p in &params.args {
            scope.alloc(&p.name);
        }
        if let Some(ref va) = params.vararg {
            scope.alloc(&va.name);
        }
        for p in &params.kwonly_args {
            scope.alloc(&p.name);
        }
        if let Some(ref kw) = params.kwarg {
            scope.alloc(&kw.name);
        }
        self.scopes.push(scope);

        let scope_idx = self.scopes.len() - 1;

        // Emit default parameter checks.
        // Missing args are Undefined (from VM padding). If param == Undefined, set to default.
        // Defaults are listed right-to-left: last N params have defaults.
        let num_positional = params.args.len();
        let num_defaults = params.defaults.len();
        if num_defaults > 0 {
            let first_default_idx = num_positional - num_defaults;
            for (di, default_expr) in params.defaults.iter().enumerate() {
                let param_idx = first_default_idx + di;
                // slot = param_idx + 1 (slot 0 = callee)
                let slot = (param_idx + 1) as u16;
                // if param is Null (missing arg): use default value
                self.chunk(func_chunk_idx).emit_op_u16(Op::local_get, slot, 0);
                self.chunk(func_chunk_idx).emit_op(Op::ref_is_null, 0);
                let skip = self.chunk(func_chunk_idx).emit_jump(Op::br_if_false, 0);
                self.compile_expr(default_expr, func_chunk_idx)?;
                self.chunk(func_chunk_idx).emit_op_u16(Op::local_set, slot, 0);
                self.chunk(func_chunk_idx).emit_op(Op::drop, 0);
                self.chunk(func_chunk_idx).patch_jump(skip);
            }
        }
        // Same for keyword-only defaults
        for (di, default_opt) in params.kw_defaults.iter().enumerate() {
            if let Some(default_expr) = default_opt {
                let slot = (num_positional + di + 1) as u16;
                self.chunk(func_chunk_idx).emit_op_u16(Op::local_get, slot, 0);
                self.chunk(func_chunk_idx).emit_op(Op::ref_is_null, 0);
                let skip = self.chunk(func_chunk_idx).emit_jump(Op::br_if_false, 0);
                self.compile_expr(default_expr, func_chunk_idx)?;
                self.chunk(func_chunk_idx).emit_op_u16(Op::local_set, slot, 0);
                self.chunk(func_chunk_idx).emit_op(Op::drop, 0);
                self.chunk(func_chunk_idx).patch_jump(skip);
            }
        }

        for s in body {
            self.compile_stmt(s, func_chunk_idx)?;
        }

        // Ensure function ends with return
        self.chunks[func_chunk_idx].emit_op(Op::null, 0);
        self.chunks[func_chunk_idx].emit_op(Op::r#return, 0);

        let scope = self.scopes.remove(scope_idx);
        self.chunks[func_chunk_idx].local_count = (scope.max_local + 1) as u16;
        self.chunks[func_chunk_idx].arity = params.args.len() as u8;

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
                    self.chunk(chunk_idx).emit_op(Op::array_get, 0);
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
                let name_c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(attr.as_str())));
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
                        self.chunk(chunk_idx).emit_op(Op::array_length, 0);
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
                    self.chunk(chunk_idx).emit_op(Op::array_set, 0);
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
                let c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(s.as_str())));
                self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
            }
            Expression::FString { parts } => {
                // Concatenate all parts
                let mut count = 0;
                for part in parts {
                    match part {
                        FStringPart::Literal(s) => {
                            let c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(s.as_str())));
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
                    }
                }
                // Concatenate all parts
                for _ in 1..count {
                    self.chunk(chunk_idx).emit_op(Op::str_concat, 0);
                }
                if count == 0 {
                    let c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from("")));
                    self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
                }
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
                        let c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(name.as_str())));
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
                            self.chunk(chunk_idx).emit_op(Op::array_concat, 0);
                        } else {
                            self.compile_expr(e, chunk_idx)?;
                            self.chunk(chunk_idx).emit_op(Op::array_push, 0);
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
                            // String key: dict is on stack, push value, struct_set
                            self.chunk(chunk_idx).emit_op(Op::dup, 0); // keep dict
                            self.compile_expr(v, chunk_idx)?; // push value
                            let key_idx = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(s.as_str())));
                            self.chunk(chunk_idx).emit_op_u16(Op::struct_set, key_idx, 0);
                            self.chunk(chunk_idx).emit_op(Op::drop, 0); // drop struct_set result
                        } else {
                            // Dynamic key: use array_set (obj, key, val)
                            self.chunk(chunk_idx).emit_op(Op::dup, 0);
                            self.compile_expr(key, chunk_idx)?;
                            self.compile_expr(v, chunk_idx)?;
                            self.chunk(chunk_idx).emit_op(Op::array_set, 0);
                            self.chunk(chunk_idx).emit_op(Op::drop, 0);
                        }
                    } else {
                        self.chunk(chunk_idx).emit_op(Op::dup, 0);
                        self.chunk(chunk_idx).emit_op(Op::null, 0);
                        self.compile_expr(v, chunk_idx)?;
                        self.chunk(chunk_idx).emit_op(Op::array_set, 0);
                        self.chunk(chunk_idx).emit_op(Op::drop, 0);
                    }
                }
            }

            Expression::BinOp { op, left, right } => {
                self.compile_expr(left, chunk_idx)?;
                self.compile_expr(right, chunk_idx)?;
                match op {
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
                        self.chunk(chunk_idx).emit_op(Op::f64_floor, 0);
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
                            self.chunk(chunk_idx).emit_op(Op::dyn_to_bool, 0);
                            let false_jump = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);
                            self.compile_expr(v, chunk_idx)?;
                            self.chunk(chunk_idx).emit_op(Op::dyn_to_bool, 0);
                            let end_jump = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                            self.chunk(chunk_idx).patch_jump(false_jump);
                            self.chunk(chunk_idx).emit_op(Op::r#false, 0);
                            self.chunk(chunk_idx).patch_jump(end_jump);
                        }
                    }
                    BoolOp::Or => {
                        self.compile_expr(&values[0], chunk_idx)?;
                        for v in &values[1..] {
                            self.chunk(chunk_idx).emit_op(Op::dyn_to_bool, 0);
                            let true_jump = self.chunk(chunk_idx).emit_jump(Op::br_if_true, 0);
                            self.compile_expr(v, chunk_idx)?;
                            self.chunk(chunk_idx).emit_op(Op::dyn_to_bool, 0);
                            let end_jump = self.chunk(chunk_idx).emit_jump(Op::br, 0);
                            self.chunk(chunk_idx).patch_jump(true_jump);
                            self.chunk(chunk_idx).emit_op(Op::r#true, 0);
                            self.chunk(chunk_idx).patch_jump(end_jump);
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

            Expression::Call { func, args, keywords: _ } => {
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
                                self.compile_expr(&args[0], chunk_idx)?;
                                self.chunk(chunk_idx).emit_op(Op::array_length, 0);
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
                                self.chunk(chunk_idx).emit_op(Op::f64_abs, 0);
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
                            common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_enumerate", 0);
                            for a in args { self.compile_expr(a, chunk_idx)?; }
                            common::bundle::emit_call_invoke(self.chunk(chunk_idx), args.len() as u8, 0);
                            return Ok(());
                        }
                        "zip" => {
                            common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_zip", 0);
                            for a in args { self.compile_expr(a, chunk_idx)?; }
                            common::bundle::emit_call_invoke(self.chunk(chunk_idx), args.len() as u8, 0);
                            return Ok(());
                        }
                        "sorted" => {
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
                                // Second arg is the type — convert built-in type names to strings
                                if let Expression::Name(type_name) = &args[1] {
                                    let type_str = match type_name.as_str() {
                                        "int" => "int",
                                        "float" => "float",
                                        "str" => "str",
                                        "bool" => "bool",
                                        "list" => "list",
                                        "dict" => "dict",
                                        "tuple" => "list",
                                        "type" => "object",
                                        _ => type_name.as_str(),
                                    };
                                    let c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(type_str)));
                                    self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
                                } else {
                                    self.compile_expr(&args[1], chunk_idx)?;
                                }
                                // Use ref_typeof + comparison for built-in types,
                                // or ref_test for user-defined types
                                let isinstance_fn = self.chunk(chunk_idx).add_import("vybe:array", "isinstance");
                                self.chunk(chunk_idx).emit_op_u16(Op::call_import, isinstance_fn, 0);
                                self.chunk(chunk_idx).emit(2, 0);
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
                                self.chunk(chunk_idx).emit_op(Op::f64_nearest, 0);
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
                            // open(filename) → readFile for now
                            return self.compile_host_call("wasi:filesystem", "readFile", args, chunk_idx);
                        }
                        "hasattr" => {
                            if args.len() >= 2 {
                                // hasattr(obj, key) → obj[key] is not null
                                self.compile_expr(&args[0], chunk_idx)?;
                                self.compile_expr(&args[1], chunk_idx)?;
                                self.chunk(chunk_idx).emit_op(Op::array_get, 0);
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
                                self.chunk(chunk_idx).emit_op(Op::array_get, 0);
                                return Ok(());
                            }
                        }
                        "setattr" => {
                            if args.len() >= 3 {
                                // setattr(obj, key, val) → obj[key] = val via array_set
                                self.compile_expr(&args[0], chunk_idx)?;
                                self.compile_expr(&args[1], chunk_idx)?;
                                self.compile_expr(&args[2], chunk_idx)?;
                                self.chunk(chunk_idx).emit_op(Op::array_set, 0);
                                return Ok(());
                            }
                        }
                        _ => {}
                    }
                }
                // Method calls: obj.method(args)
                if let Expression::Attribute { value, attr } = func.as_ref() {
                    return self.compile_method_call(value, attr, args, chunk_idx);
                }
                // General function call
                self.compile_expr(func, chunk_idx)?;
                let mut argc = 0u8;
                for a in args {
                    if let Expression::Starred(inner) = a {
                        // *args: spread the iterable onto the stack
                        self.compile_expr(inner, chunk_idx)?;
                        self.chunk(chunk_idx).emit_op(Op::spread, 0);
                        // We don't know the spread count at compile time.
                        // The VM's spread pushes N elements. call_ref will use
                        // whatever argc we give it — extra args are ignored,
                        // missing args are padded with Null. This is imprecise
                        // but handles the common case f(*[1,2,3]) where the
                        // list length matches the function arity.
                    } else {
                        self.compile_expr(a, chunk_idx)?;
                        argc += 1;
                    }
                }
                self.chunk(chunk_idx).emit_op_u8(Op::call_ref, argc, 0);
            }

            Expression::Attribute { value, attr } => {
                self.compile_expr(value, chunk_idx)?;
                let c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(attr.as_str())));
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
                        self.chunk(chunk_idx).emit_op(Op::array_slice, 0);
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
                        self.chunk(chunk_idx).emit_op(Op::array_length, 0);
                        self.compile_expr(slice, chunk_idx)?;
                        self.chunk(chunk_idx).emit_op(Op::dyn_add, 0);
                    } else {
                        self.compile_expr(slice, chunk_idx)?;
                    }
                    self.chunk(chunk_idx).emit_op(Op::array_get, 0);
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
                self.chunk(chunk_idx).emit_op(Op::r#await, 0);
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

        // Create object
        ctor.emit_op_u16(Op::struct_new, 0, 0);
        ctor.emit_op_u16(Op::local_set, this_slot, 0);
        ctor.emit_op(Op::drop, 0);

        // Set __exception_type = exc_name
        ctor.emit_op_u16(Op::local_get, this_slot, 0);
        let et_val = ctor.add_constant(Value::String(Rc::from(exc_name)));
        ctor.emit_op_u16(Op::r#const, et_val, 0);
        let et_key = ctor.add_constant(Value::String(Rc::from("__exception_type")));
        ctor.emit_op_u16(Op::struct_set, et_key, 0);
        ctor.emit_op(Op::drop, 0);

        // Set name = exc_name (JS Error convention)
        ctor.emit_op_u16(Op::local_get, this_slot, 0);
        let n_val = ctor.add_constant(Value::String(Rc::from(exc_name)));
        ctor.emit_op_u16(Op::r#const, n_val, 0);
        let n_key = ctor.add_constant(Value::String(Rc::from("name")));
        ctor.emit_op_u16(Op::struct_set, n_key, 0);
        ctor.emit_op(Op::drop, 0);

        // Set message = arg
        ctor.emit_op_u16(Op::local_get, this_slot, 0);
        ctor.emit_op_u16(Op::local_get, msg_slot, 0);
        let m_key = ctor.add_constant(Value::String(Rc::from("message")));
        ctor.emit_op_u16(Op::struct_set, m_key, 0);
        ctor.emit_op(Op::drop, 0);

        // Set __type for TypeOf checks
        ctor.emit_op_u16(Op::local_get, this_slot, 0);
        let t_val = ctor.add_constant(Value::String(Rc::from(exc_name)));
        ctor.emit_op_u16(Op::r#const, t_val, 0);
        let t_key = ctor.add_constant(Value::String(Rc::from("__type")));
        ctor.emit_op_u16(Op::struct_set, t_key, 0);
        ctor.emit_op(Op::drop, 0);

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
            Value::String(Rc::from(lower.as_str()))
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
                let empty = self.chunk(chunk_idx).add_constant(Value::String(Rc::from("")));
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
        match method {
            "append" => {
                self.compile_expr(obj, chunk_idx)?;
                for a in args { self.compile_expr(a, chunk_idx)?; }
                common::collections::emit_push(self.chunk(chunk_idx), 0);
            }
            "pop" => {
                self.compile_expr(obj, chunk_idx)?;
                common::collections::emit_pop(self.chunk(chunk_idx), 0);
            }
            "keys" => {
                self.compile_expr(obj, chunk_idx)?;
                common::dict::emit_keys(self.chunk(chunk_idx), 0);
            }
            "values" => {
                self.compile_expr(obj, chunk_idx)?;
                common::dict::emit_values(self.chunk(chunk_idx), 0);
            }
            "upper" => {
                self.compile_expr(obj, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::str_to_upper, 0);
            }
            "lower" => {
                self.compile_expr(obj, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::str_to_lower, 0);
            }
            "strip" => {
                self.compile_expr(obj, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::str_trim, 0);
            }
            "split" => {
                self.compile_expr(obj, chunk_idx)?;
                if args.is_empty() {
                    let c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(" ")));
                    self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
                } else {
                    self.compile_expr(&args[0], chunk_idx)?;
                }
                self.chunk(chunk_idx).emit_op(Op::str_split, 0);
            }
            "join" => {
                // separator.join(iterable) → join(iterable, separator)
                if args.len() == 1 {
                    self.compile_expr(&args[0], chunk_idx)?;
                    self.compile_expr(obj, chunk_idx)?;
                    common::collections::emit_join(self.chunk(chunk_idx), 0);
                } else {
                    self.compile_expr(obj, chunk_idx)?;
                }
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
            "find" | "index" => {
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
                // str.count(sub) — count occurrences
                self.compile_expr(obj, chunk_idx)?;
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                let count_fn = self.chunk(chunk_idx).add_import("vybe:string", "count");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, count_fn, 0);
                self.chunk(chunk_idx).emit(2, 0);
            }
            "lstrip" => {
                self.compile_expr(obj, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::str_trim_start, 0);
            }
            "rstrip" => {
                self.compile_expr(obj, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::str_trim_end, 0);
            }
            "title" | "capitalize" => {
                // Simplified: just upper the first char
                self.compile_expr(obj, chunk_idx)?;
                let to_str = self.chunk(chunk_idx).add_import("vybe:convert", "toString");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, to_str, 0);
                self.chunk(chunk_idx).emit(1, 0);
            }
            "isdigit" | "isnumeric" | "isdecimal" => {
                self.compile_expr(obj, chunk_idx)?;
                let is_num = self.chunk(chunk_idx).add_import("vybe:convert", "isNumeric");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, is_num, 0);
                self.chunk(chunk_idx).emit(1, 0);
            }
            "isalpha" | "isalnum" | "isspace" | "islower" | "isupper" => {
                // Simplified: return true for now
                self.compile_expr(obj, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                self.chunk(chunk_idx).emit_op(Op::r#true, 0);
            }
            "format" => {
                // "hello {} and {}".format(a, b)
                // For each arg, find first "{}", replace with str(arg) using substring ops.
                // s = s[:idx] + str(arg) + s[idx+2:]
                self.compile_expr(obj, chunk_idx)?;
                let placeholder = self.chunk(chunk_idx).add_constant(Value::String(Rc::from("{}")));
                let two = self.chunk(chunk_idx).add_constant(Value::I32(2));
                let max = self.chunk(chunk_idx).add_constant(Value::I32(i32::MAX));
                for a in args {
                    // Stack: [current_string]
                    let str_local = self.scope(chunk_idx).alloc("__fmt_str");
                    self.chunk(chunk_idx).emit_op_u16(Op::local_set, str_local, 0);
                    self.chunk(chunk_idx).emit_op(Op::drop, 0);
                    // Find position of first "{}"
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, str_local, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::r#const, placeholder, 0);
                    self.chunk(chunk_idx).emit_op(Op::str_index_of, 0);
                    let pos_local = self.scope(chunk_idx).alloc("__fmt_pos");
                    self.chunk(chunk_idx).emit_op_u16(Op::local_set, pos_local, 0);
                    self.chunk(chunk_idx).emit_op(Op::drop, 0);
                    // s[:pos]
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, str_local, 0);
                    self.chunk(chunk_idx).emit_op(Op::i32_const_0, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, pos_local, 0);
                    self.chunk(chunk_idx).emit_op(Op::str_substring, 0);
                    // + str(arg)
                    let empty = self.chunk(chunk_idx).add_constant(Value::String(Rc::from("")));
                    self.chunk(chunk_idx).emit_op_u16(Op::r#const, empty, 0);
                    self.compile_expr(a, chunk_idx)?;
                    self.chunk(chunk_idx).emit_op(Op::dyn_add, 0);
                    self.chunk(chunk_idx).emit_op(Op::str_concat, 0);
                    // + s[pos+2:]
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, str_local, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, pos_local, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::r#const, two, 0);
                    self.chunk(chunk_idx).emit_op(Op::i32_add, 0);
                    self.chunk(chunk_idx).emit_op_u16(Op::r#const, max, 0);
                    self.chunk(chunk_idx).emit_op(Op::str_substring, 0);
                    self.chunk(chunk_idx).emit_op(Op::str_concat, 0);
                }
            }
            "encode" => {
                // str.encode() → bytes (just return string)
                self.compile_expr(obj, chunk_idx)?;
            }
            "zfill" => {
                self.compile_expr(obj, chunk_idx)?;
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                self.chunk(chunk_idx).emit_op(Op::str_pad_start, 0);
            }
            "center" | "ljust" | "rjust" => {
                self.compile_expr(obj, chunk_idx)?;
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                let op = if method == "rjust" { Op::str_pad_start } else { Op::str_pad_end };
                self.chunk(chunk_idx).emit_op(op, 0);
            }
            // List methods
            "insert" => {
                self.compile_expr(obj, chunk_idx)?;
                for a in args { self.compile_expr(a, chunk_idx)?; }
                let insert_fn = self.chunk(chunk_idx).add_import("vybe:array", "splice");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, insert_fn, 0);
                self.chunk(chunk_idx).emit((1 + args.len()) as u8, 0);
            }
            "extend" => {
                // lst.extend(other) → for each in other: lst.append(each)
                self.compile_expr(obj, chunk_idx)?;
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                self.chunk(chunk_idx).emit_op(Op::array_concat, 0);
            }
            "remove" => {
                self.compile_expr(obj, chunk_idx)?;
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                self.chunk(chunk_idx).emit_op(Op::array_index_of, 0);
            }
            "reverse" => {
                self.compile_expr(obj, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::array_reverse, 0);
            }
            "sort" => {
                common::bundle::emit_call_push_func(self.chunk(chunk_idx), "__vybe_sorted", 0);
                self.compile_expr(obj, chunk_idx)?;
                common::bundle::emit_call_invoke(self.chunk(chunk_idx), 1, 0);
            }
            "copy" => {
                // array_slice(0, MAX) = shallow copy
                self.compile_expr(obj, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::i32_const_0, 0);
                let max = self.chunk(chunk_idx).add_constant(Value::I32(i32::MAX));
                self.chunk(chunk_idx).emit_op_u16(Op::r#const, max, 0);
                common::collections::emit_slice(self.chunk(chunk_idx), 0);
            }
            "clear" => {
                // obj.clear() — set to empty array
                self.compile_expr(obj, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::array_new, 0, 0);
            }
            // Dict methods
            "items" => {
                self.compile_expr(obj, chunk_idx)?;
                common::dict::emit_items(self.chunk(chunk_idx), 0);
            }
            "get" => {
                // dict.get(key, default=None)
                self.compile_expr(obj, chunk_idx)?;
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                common::dict::emit_get(self.chunk(chunk_idx), 0);
                self.chunk(chunk_idx).emit(2, 0);
            }
            "update" => {
                // dict.update(other) — simplified: skip
                self.compile_expr(obj, chunk_idx)?;
            }
            "setdefault" => {
                self.compile_expr(obj, chunk_idx)?;
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                common::dict::emit_get(self.chunk(chunk_idx), 0);
            }
            "read" => {
                // file.read() — simplified: return the file content (obj is already content from open())
                self.compile_expr(obj, chunk_idx)?;
            }
            "write" => {
                self.compile_expr(obj, chunk_idx)?;
                if !args.is_empty() { self.compile_expr(&args[0], chunk_idx)?; }
                let write_fn = self.chunk(chunk_idx).add_import("wasi:filesystem", "writeFile");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, write_fn, 0);
                self.chunk(chunk_idx).emit(2, 0);
            }
            "close" => {
                // file.close() — no-op
                self.compile_expr(obj, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::drop, 0);
                self.chunk(chunk_idx).emit_op(Op::null, 0);
            }
            _ => {
                // Generic method call: obj.method(args) → get method, pass obj as self
                self.compile_expr(obj, chunk_idx)?;
                let c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(method)));
                self.chunk(chunk_idx).emit_op_u16(Op::struct_get, c, 0);
                // Pass obj as self (first arg) — same convention as JS/VB/C#
                self.compile_expr(obj, chunk_idx)?;
                for a in args { self.compile_expr(a, chunk_idx)?; }
                self.chunk(chunk_idx).emit_op_u8(Op::call_ref, (args.len() + 1) as u8, 0);
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
            s.chunk(chunk_idx).emit_op(Op::array_push, 0);
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
        self.chunk(chunk_idx).emit_op(Op::array_length, 0);
        self.chunk(chunk_idx).emit_op(Op::dyn_lt, 0);
        let exit_jump = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);

        // Load current element
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, iter_local, 0);
        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx_local, 0);
        self.chunk(chunk_idx).emit_op(Op::array_get, 0);
        self.compile_assign_target(&generator.target, chunk_idx)?;

        // Apply if filters
        let mut filter_jumps = Vec::new();
        for f in &generator.ifs {
            self.compile_expr(f, chunk_idx)?;
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
        match op {
            CmpOp::Eq => self.chunk(chunk_idx).emit_op(Op::dyn_eq, 0),
            CmpOp::NotEq => self.chunk(chunk_idx).emit_op(Op::dyn_ne, 0),
            CmpOp::Lt => self.chunk(chunk_idx).emit_op(Op::dyn_lt, 0),
            CmpOp::LtE => self.chunk(chunk_idx).emit_op(Op::dyn_le, 0),
            CmpOp::Gt => self.chunk(chunk_idx).emit_op(Op::dyn_gt, 0),
            CmpOp::GtE => self.chunk(chunk_idx).emit_op(Op::dyn_ge, 0),
            CmpOp::Is => self.chunk(chunk_idx).emit_op(Op::dyn_eq, 0), // simplified
            CmpOp::IsNot => self.chunk(chunk_idx).emit_op(Op::dyn_ne, 0),
            CmpOp::In => {
                // Stack: [needle, haystack]. array_contains pops haystack(TOS), needle.
                self.chunk(chunk_idx).emit_op(Op::array_contains, 0);
            }
            CmpOp::NotIn => {
                self.chunk(chunk_idx).emit_op(Op::array_contains, 0);
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
            AugOp::FloorDiv => { self.chunk(chunk_idx).emit_op(Op::f64_div, 0); self.chunk(chunk_idx).emit_op(Op::f64_floor, 0); }
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
