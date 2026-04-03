use std::rc::Rc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use vybe_parser_pascal::ast::*;
use vybe_compiler_common as common;
use crate::scope::Scope;

struct LoopCtx {
    break_patches: Vec<usize>,
    continue_target: usize,
}

/// Compiled method for a class (stored until constructor is built).
struct CompiledMethod {
    name: String,
    chunk_idx: usize,
    is_static: bool,
}

/// Registered class info (from type section, before method impls).
struct ClassInfo {
    parent: Option<String>,
    fields: Vec<String>,
    method_sigs: Vec<MethodSig>,
    compiled_methods: Vec<CompiledMethod>,
    properties: Vec<PropertyDef>,
}

pub struct Compiler {
    chunks: Vec<Chunk>,
    scopes: Vec<Scope>,
    current: usize,
    loops: Vec<LoopCtx>,
    line: u32,
    defined_globals: std::collections::HashSet<String>,
    current_func_name: Option<String>,
    current_result_slot: Option<u16>,
    classes: std::collections::HashMap<String, ClassInfo>,
    current_class: Option<String>,
}

use vybe_parser_pascal::ast::{ClassDef, ClassMember, MethodSig, MethodKind, MethodImpl, PropertyDef};

impl Compiler {
    pub fn new() -> Self {
        Self {
            chunks: vec![Chunk::new("<program>")],
            scopes: vec![Scope::new()],
            current: 0,
            loops: Vec::new(),
            line: 1,
            defined_globals: std::collections::HashSet::new(),
            current_func_name: None,
            current_result_slot: None,
            classes: std::collections::HashMap::new(),
            current_class: None,
        }
    }

    pub fn compile(mut self, program: &Program) -> Result<Vec<Chunk>, String> {
        for decl in &program.decls { self.compile_decl(decl)?; }
        // Build class constructors AFTER all method implementations are compiled
        self.finalize_classes()?;
        for stmt in &program.body { self.compile_stmt(stmt)?; }
        self.emit(Op::null);
        self.emit(Op::halt);
        let locals = self.scope().next_slot;
        self.chunks[0].local_count = locals;
        common::bundle::finalize_with_stdlib(&mut self.chunks);
        Ok(self.chunks)
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    fn scope(&self) -> &Scope { self.scopes.last().unwrap() }
    fn scope_mut(&mut self) -> &mut Scope { self.scopes.last_mut().unwrap() }

    fn emit(&mut self, op: Op) {
        let line = self.line;
        self.chunks[self.current].emit_op(op, line);
    }
    fn emit_u16(&mut self, op: Op, v: u16) {
        let line = self.line;
        self.chunks[self.current].emit_op_u16(op, v, line);
    }
    fn emit_u8(&mut self, op: Op, v: u8) {
        let line = self.line;
        self.chunks[self.current].emit_op_u8(op, v, line);
    }
    fn emit_const(&mut self, val: Value) {
        let idx = self.chunks[self.current].add_constant(val);
        self.emit_u16(Op::r#const, idx);
    }
    fn emit_jump(&mut self, op: Op) -> usize {
        let line = self.line;
        self.chunks[self.current].emit_jump(op, line)
    }
    fn patch_jump(&mut self, offset: usize) {
        self.chunks[self.current].patch_jump(offset);
    }
    fn emit_loop(&mut self, target: usize) {
        let line = self.line;
        self.chunks[self.current].emit_loop(target, line);
    }
    fn current_offset(&self) -> usize { self.chunks[self.current].current_offset() }
    fn str_const(&mut self, s: &str) -> u16 {
        self.chunks[self.current].add_constant(Value::String(Rc::from(s)))
    }
    fn chunk(&mut self) -> &mut Chunk { &mut self.chunks[self.current] }

    /// Add import to chunk 0 (VM resolves imports from chunk 0 only).
    fn import(&mut self, module: &str, name: &str) -> u16 {
        self.chunks[0].add_import(module, name)
    }
    fn emit_host_call(&mut self, import_idx: u16, argc: u8) {
        let line = self.line;
        self.chunks[self.current].emit_op_u16(Op::call_import, import_idx, line);
        self.chunks[self.current].emit(argc, line);
    }

    // ── Variable resolution ───────────────────────────────────────────────────

    fn resolve_var(&mut self, name: &str) -> VarRes {
        if let Some(slot) = self.scope().resolve(name) { return VarRes::Local(slot); }
        if self.scopes.len() > 1 {
            if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, name) {
                return VarRes::Upvalue(uv);
            }
        }
        VarRes::Global
    }

    fn resolve_upvalue(&mut self, scope_idx: usize, name: &str) -> Option<u8> {
        if scope_idx == 0 { return None; }
        let parent = scope_idx - 1;
        if let Some(slot) = self.scopes[parent].resolve(name) {
            self.scopes[parent].mark_captured(slot);
            return Some(self.scopes[scope_idx].add_upvalue(slot as u8, true));
        }
        if let Some(uv) = self.resolve_upvalue(parent, name) {
            return Some(self.scopes[scope_idx].add_upvalue(uv, false));
        }
        None
    }

    fn emit_var_get(&mut self, name: &str) {
        match self.resolve_var(name) {
            VarRes::Local(slot) => self.emit_u16(Op::local_get, slot),
            VarRes::Upvalue(idx) => self.emit_u8(Op::upvalue_get, idx),
            VarRes::Global => { let idx = self.str_const(name); self.emit_u16(Op::global_get, idx); }
        }
    }

    fn emit_var_set(&mut self, name: &str) {
        match self.resolve_var(name) {
            VarRes::Local(slot) => { self.emit_u16(Op::local_set, slot); self.emit(Op::drop); }
            VarRes::Upvalue(idx) => { self.emit_u8(Op::upvalue_set, idx); self.emit(Op::drop); }
            VarRes::Global => { let idx = self.str_const(name); self.emit_u16(Op::global_set, idx); self.emit(Op::drop); }
        }
    }

    // ── Declarations ──────────────────────────────────────────────────────────

    fn compile_decl(&mut self, decl: &Decl) -> Result<(), String> {
        match decl {
            Decl::Var(vars) => {
                for v in vars {
                    for name in &v.names {
                        if let Some(init) = &v.init {
                            self.compile_expr(init)?;
                        } else {
                            match v.type_name.name.to_lowercase().as_str() {
                                "integer" | "longint" | "int64" | "cardinal" | "byte" | "word" | "shortint"
                                | "real" | "double" | "single" | "extended" => self.emit(Op::f64_const_0),
                                "boolean" => self.emit(Op::r#false),
                                "string" => self.emit_const(Value::String(Rc::from(""))),
                                _ => self.emit(Op::null),
                            }
                        }
                        if self.scopes.len() == 1 && self.scope().depth == 0 {
                            let idx = self.str_const(name);
                            self.emit_u16(Op::global_set, idx);
                            self.emit(Op::drop);
                        } else {
                            let slot = self.scope_mut().define(name);
                            self.emit_u16(Op::local_set, slot);
                            self.emit(Op::drop);
                        }
                    }
                }
            }
            Decl::Const(consts) => {
                for c in consts {
                    self.compile_expr(&c.value)?;
                    if self.scopes.len() == 1 {
                        let idx = self.str_const(&c.name);
                        self.emit_u16(Op::global_set, idx);
                        self.emit(Op::drop);
                    } else {
                        let slot = self.scope_mut().define(&c.name);
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    }
                }
            }
            Decl::Type(types) => {
                for t in types {
                    if let TypeDef::Class(ref class_def) = t.def {
                        self.register_class(&t.name, class_def)?;
                    }
                }
            }
            Decl::Procedure(proc) => self.compile_proc(proc)?,
            Decl::Function(func) => self.compile_func(func)?,
            Decl::Method(method) => self.compile_method_impl(method)?,
        }
        Ok(())
    }

    fn compile_proc(&mut self, proc: &ProcDecl) -> Result<(), String> {
        if proc.is_forward { return Ok(()); }
        self.defined_globals.insert(proc.name.clone());
        let arity: u8 = proc.params.iter().map(|p| p.names.len() as u8).sum();
        let func_idx = self.chunks.len();
        let chunk = common::functions::create_function_chunk(&proc.name, arity);
        self.chunks.push(chunk);
        self.scopes.push(Scope::new_function());
        let saved = self.current;
        self.current = func_idx;

        for param in &proc.params {
            for name in &param.names { self.scope_mut().define(name); }
        }
        for decl in &proc.decls { self.compile_decl(decl)?; }
        for stmt in &proc.body { self.compile_stmt(stmt)?; }

        let line = self.line;
        common::functions::emit_function_epilogue(&mut self.chunks[func_idx], line);

        let locals = self.scope().next_slot;
        self.chunks[func_idx].local_count = locals;
        let uvs = self.scopes.last().unwrap().upvalues.clone();
        self.scopes.pop();
        self.current = saved;

        let line = self.line;
        common::functions::emit_ref_func(&mut self.chunks[self.current], func_idx, uvs.len() as u8, line);
        // Emit upvalue descriptors
        for uv in &uvs {
            self.chunks[self.current].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current].emit(uv.index, line);
        }
        let idx = self.str_const(&proc.name);
        self.emit_u16(Op::global_set, idx);
        self.emit(Op::drop);
        self.defined_globals.insert(proc.name.clone());
        Ok(())
    }

    fn compile_func(&mut self, func: &FuncDecl) -> Result<(), String> {
        if func.is_forward { return Ok(()); }
        // Pre-register so recursive calls resolve as global_get + call_ref
        self.defined_globals.insert(func.name.clone());
        let arity: u8 = func.params.iter().map(|p| p.names.len() as u8).sum();
        let func_idx = self.chunks.len();
        let chunk = common::functions::create_function_chunk(&func.name, arity);
        self.chunks.push(chunk);
        self.scopes.push(Scope::new_function());
        let saved = self.current;
        self.current = func_idx;

        for param in &func.params {
            for name in &param.names { self.scope_mut().define(name); }
        }
        // Result slot — Pascal functions assign to Result or function name
        let result_slot = self.scope_mut().define("Result");
        self.emit(Op::null);
        self.emit_u16(Op::local_set, result_slot);
        self.emit(Op::drop);
        // Track function name → Result slot so `FuncName := value` works
        let saved_func_name = self.current_func_name.take();
        let saved_result_slot = self.current_result_slot.take();
        self.current_func_name = Some(func.name.clone());
        self.current_result_slot = Some(result_slot);


        for decl in &func.decls { self.compile_decl(decl)?; }
        for stmt in &func.body { self.compile_stmt(stmt)?; }

        // Return Result
        self.emit_u16(Op::local_get, result_slot);
        self.emit(Op::r#return);

        self.current_func_name = saved_func_name;
        self.current_result_slot = saved_result_slot;

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
        let idx = self.str_const(&func.name);
        self.emit_u16(Op::global_set, idx);
        self.emit(Op::drop);
        Ok(())
    }

    // ── Statements ────────────────────────────────────────────────────────────

    fn compile_stmt(&mut self, stmt: &Statement) -> Result<(), String> {
        match stmt {
            Statement::Empty => {}
            Statement::Block(stmts) => {
                self.scope_mut().begin_scope();
                for s in stmts { self.compile_stmt(s)?; }
                self.scope_mut().end_scope();
            }
            Statement::Assign { target, value } => {
                self.compile_expr(value)?;
                self.compile_assign_target(target)?;
            }
            Statement::Call { name, args } => {
                if args.is_empty() {
                    // Parser wrapped the expression in name.
                    // If it's a bare Field access (e.g. c.Inc), treat as method call with no args.
                    if let Expression::Field { record, field } = name {
                        let call = Expression::Call {
                            callee: Box::new(Expression::Field { record: record.clone(), field: field.clone() }),
                            args: Vec::new(),
                        };
                        self.compile_expr(&call)?;
                    } else if let Expression::Identifier(ident) = name {
                        // Bare identifier call (e.g. SomeProcedure)
                        let call = Expression::Call {
                            callee: Box::new(Expression::Identifier(ident.clone())),
                            args: Vec::new(),
                        };
                        self.compile_expr(&call)?;
                    } else {
                        self.compile_expr(name)?;
                    }
                } else {
                    let call = Expression::Call { callee: Box::new(name.clone()), args: args.clone() };
                    self.compile_expr(&call)?;
                }
                self.emit(Op::drop);
            }
            Statement::If { cond, then, else_ } => {
                self.compile_expr(cond)?;
                self.emit(Op::dyn_to_bool);
                let else_j = self.emit_jump(Op::br_if_false);
                self.compile_stmt(then)?;
                if let Some(alt) = else_ {
                    let end_j = self.emit_jump(Op::br);
                    self.patch_jump(else_j);
                    self.compile_stmt(alt)?;
                    self.patch_jump(end_j);
                } else {
                    self.patch_jump(else_j);
                }
            }
            Statement::For { var, from, to, downto, body } => {
                // var := from
                self.compile_expr(from)?;
                self.emit_var_set(var);

                let loop_start = self.current_offset();
                self.loops.push(LoopCtx { break_patches: vec![], continue_target: loop_start });

                // condition: var <= to  (or >= for downto)
                self.emit_var_get(var);
                self.compile_expr(to)?;
                if *downto { self.emit(Op::dyn_ge); } else { self.emit(Op::dyn_le); }
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);

                self.compile_stmt(body)?;

                // inc/dec
                self.emit_var_get(var);
                self.emit_const(Value::F64(1.0));
                if *downto {
                    self.emit(Op::f64_sub);
                } else {
                    self.emit(Op::dyn_add);
                }
                self.emit_var_set(var);

                self.emit_loop(loop_start);
                self.patch_jump(exit);
                let ctx = self.loops.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
            }
            Statement::While { cond, body } => {
                let start = self.current_offset();
                self.loops.push(LoopCtx { break_patches: vec![], continue_target: start });
                self.compile_expr(cond)?;
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                self.compile_stmt(body)?;
                self.emit_loop(start);
                self.patch_jump(exit);
                let ctx = self.loops.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
            }
            Statement::Repeat { body, until } => {
                let start = self.current_offset();
                self.loops.push(LoopCtx { break_patches: vec![], continue_target: start });
                for s in body { self.compile_stmt(s)?; }
                self.compile_expr(until)?;
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_true);
                self.emit_loop(start);
                self.patch_jump(exit);
                let ctx = self.loops.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
            }
            Statement::Case { expr, arms, else_ } => {
                self.compile_expr(expr)?;
                let mut end_patches = Vec::new();
                for arm in arms {
                    let mut match_patches = Vec::new();
                    for val in &arm.values {
                        match val {
                            CaseValue::Single(v) => {
                                self.emit(Op::dup);
                                self.compile_expr(v)?;
                                self.emit(Op::dyn_eq);
                                match_patches.push(self.emit_jump(Op::br_if_true));
                            }
                            CaseValue::Range(lo, hi) => {
                                self.emit(Op::dup);
                                self.compile_expr(lo)?;
                                self.emit(Op::dyn_ge);
                                let left_ok = self.emit_jump(Op::br_if_false);
                                self.emit(Op::dup);
                                self.compile_expr(hi)?;
                                self.emit(Op::dyn_le);
                                match_patches.push(self.emit_jump(Op::br_if_true));
                                self.patch_jump(left_ok);
                            }
                        }
                    }
                    let skip = self.emit_jump(Op::br);
                    for p in match_patches { self.patch_jump(p); }
                    for s in &arm.body { self.compile_stmt(s)?; }
                    end_patches.push(self.emit_jump(Op::br));
                    self.patch_jump(skip);
                }
                if let Some(else_stmts) = else_ {
                    for s in else_stmts { self.compile_stmt(s)?; }
                }
                for p in end_patches { self.patch_jump(p); }
                self.emit(Op::drop);
            }
            Statement::With { vars: _, body } => {
                self.compile_stmt(body)?;
            }
            Statement::Try { body, handler } => {
                let line = self.line;
                let catch_jump = common::errors::emit_try_start(&mut self.chunks[self.current], line);

                for s in body { self.compile_stmt(s)?; }
                common::errors::emit_try_end(&mut self.chunks[self.current], line);
                let skip = self.emit_jump(Op::br);

                common::errors::patch_catch(&mut self.chunks[self.current], catch_jump);

                match handler {
                    TryHandler::Except(clauses, else_stmts) => {
                        if clauses.is_empty() {
                            self.emit(Op::drop);
                            if let Some(stmts) = else_stmts {
                                for s in stmts { self.compile_stmt(s)?; }
                            }
                        } else {
                            for clause in clauses {
                                if let Some(var) = &clause.var_name {
                                    let slot = self.scope_mut().define(var);
                                    self.emit_u16(Op::local_set, slot);
                                    self.emit(Op::drop);
                                } else {
                                    self.emit(Op::drop);
                                }
                                for s in &clause.body { self.compile_stmt(s)?; }
                            }
                        }
                    }
                    TryHandler::Finally(stmts) => {
                        self.emit(Op::drop);
                        for s in stmts { self.compile_stmt(s)?; }
                    }
                }
                self.patch_jump(skip);
            }
            Statement::Raise(expr) => {
                if let Some(e) = expr { self.compile_expr(e)?; } else { self.emit(Op::null); }
                let line = self.line;
                common::errors::emit_throw(&mut self.chunks[self.current], line);
            }
            Statement::Exit(val) => {
                if let Some(v) = val {
                    self.compile_expr(v)?;
                } else if let Some(slot) = self.current_result_slot {
                    // In a function, Exit without value returns Result
                    self.emit_u16(Op::local_get, slot);
                } else {
                    self.emit(Op::null);
                }
                self.emit(Op::r#return);
            }
            Statement::Break => {
                let p = self.emit_jump(Op::br);
                if let Some(ctx) = self.loops.last_mut() { ctx.break_patches.push(p); }
            }
            Statement::Continue => {
                if let Some(ctx) = self.loops.last() {
                    let target = ctx.continue_target;
                    self.emit_loop(target);
                }
            }
        }
        Ok(())
    }

    fn compile_assign_target(&mut self, target: &Expression) -> Result<(), String> {
        match target {
            Expression::Identifier(name) => {
                // `FuncName := value` in Pascal → assign to Result slot
                if let Some(ref func_name) = self.current_func_name {
                    if name.eq_ignore_ascii_case(func_name) {
                        if let Some(slot) = self.current_result_slot {
                            self.emit_u16(Op::local_set, slot);
                            self.emit(Op::drop);
                            return Ok(());
                        }
                    }
                }
                // Inside class method, field assignment: FName := value → Self.FName := value
                if self.current_class.is_some() && self.is_class_field(name) {
                    if let Some(slot) = self.scope().resolve("Self") {
                        let tmp = self.scope_mut().define("__field_tmp");
                        self.emit_u16(Op::local_set, tmp);
                        self.emit(Op::drop);
                        self.emit_u16(Op::local_get, slot); // Self
                        self.emit_u16(Op::local_get, tmp);  // value
                        let idx = self.str_const(name);
                        self.emit_u16(Op::struct_set, idx);
                        self.emit(Op::drop);
                        return Ok(());
                    }
                }
                self.emit_var_set(name);
            }
            Expression::Field { record, field } => {
                // value on stack; need [obj, val] for struct_set
                let tmp = self.scope_mut().define("__tmp");
                self.emit_u16(Op::local_set, tmp);
                self.emit(Op::drop);
                self.compile_expr(record)?;
                self.emit_u16(Op::local_get, tmp);
                let idx = self.str_const(field);
                self.emit_u16(Op::struct_set, idx);
                self.emit(Op::drop);
            }
            Expression::Index { array, index } => {
                let tmp = self.scope_mut().define("__tmp");
                self.emit_u16(Op::local_set, tmp);
                self.emit(Op::drop);
                self.compile_expr(array)?;
                self.compile_expr(index)?;
                self.emit_u16(Op::local_get, tmp);
                self.emit(Op::array_set);
                self.emit(Op::drop);
            }
            _ => return Err(format!("Invalid assignment target: {:?}", target)),
        }
        Ok(())
    }

    // ── Expressions ───────────────────────────────────────────────────────────

    fn compile_expr(&mut self, expr: &Expression) -> Result<(), String> {
        let line = self.line;
        match expr {
            Expression::Int(n) => {
                let idx = self.chunks[self.current].add_constant(Value::F64(*n as f64));
                self.chunks[self.current].emit_op_u16(Op::r#const, idx, line);
            }
            Expression::Real(n) => {
                let idx = self.chunks[self.current].add_constant(Value::F64(*n));
                self.chunks[self.current].emit_op_u16(Op::r#const, idx, line);
            }
            Expression::Bool(b) => {
                if *b { self.emit(Op::r#true); } else { self.emit(Op::r#false); }
            }
            Expression::Nil => self.emit(Op::null),
            Expression::Str(s) => {
                let idx = self.chunks[self.current].add_constant(Value::String(Rc::from(s.as_str())));
                self.chunks[self.current].emit_op_u16(Op::r#const, idx, line);
            }
            Expression::Char(c) => {
                let idx = self.chunks[self.current].add_constant(Value::String(Rc::from(c.to_string().as_str())));
                self.chunks[self.current].emit_op_u16(Op::r#const, idx, line);
            }
            Expression::Identifier(name) => {
                match name.to_lowercase().as_str() {
                    "maxint" => self.emit_const(Value::F64(2147483647.0)),
                    "pi" => self.emit_const(Value::F64(std::f64::consts::PI)),
                    "self" => {
                        if let Some(slot) = self.scope().resolve("Self") {
                            self.emit_u16(Op::local_get, slot);
                        } else {
                            self.emit(Op::null);
                        }
                    }
                    _ => {
                        // Inside a class method, check if name is a field
                        if self.current_class.is_some() && self.is_class_field(name) {
                            // FName → Self.FName
                            if let Some(slot) = self.scope().resolve("Self") {
                                self.emit_u16(Op::local_get, slot);
                                let idx = self.str_const(name);
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
            Expression::Binary { op, left, right } => {
                // Short-circuit for And/Or
                if *op == BinOp::And {
                    self.compile_expr(left)?;
                    let skip = common::expressions::emit_and_start(&mut self.chunks[self.current], line);
                    self.compile_expr(right)?;
                    common::expressions::emit_short_circuit_end(&mut self.chunks[self.current], skip);
                    return Ok(());
                }
                if *op == BinOp::Or {
                    self.compile_expr(left)?;
                    let skip = common::expressions::emit_or_start(&mut self.chunks[self.current], line);
                    self.compile_expr(right)?;
                    common::expressions::emit_short_circuit_end(&mut self.chunks[self.current], skip);
                    return Ok(());
                }

                self.compile_expr(left)?;
                self.compile_expr(right)?;
                match op {
                    BinOp::Add => self.emit(Op::dyn_add),
                    BinOp::Sub => self.emit(Op::f64_sub),
                    BinOp::Mul => self.emit(Op::f64_mul),
                    BinOp::Div => self.emit(Op::f64_div),
                    BinOp::IDiv => {
                        self.emit(Op::f64_div);
                        common::math::emit_trunc(self.chunk(), line);
                    }
                    BinOp::Mod => self.emit(Op::f64_mod),
                    BinOp::Eq => self.emit(Op::dyn_eq),
                    BinOp::NotEq => self.emit(Op::dyn_ne),
                    BinOp::Lt => self.emit(Op::dyn_lt),
                    BinOp::Gt => self.emit(Op::dyn_gt),
                    BinOp::Le => self.emit(Op::dyn_le),
                    BinOp::Ge => self.emit(Op::dyn_ge),
                    BinOp::And | BinOp::Or => unreachable!(), // handled above
                    BinOp::Xor => self.emit(Op::i32_xor),
                    BinOp::Shl => self.emit(Op::i32_shl),
                    BinOp::Shr => self.emit(Op::i32_shr_s),
                    BinOp::In => {
                        let idx = self.import("vybe:collections", "setHas");
                        self.emit_host_call(idx, 2);
                    }
                }
            }
            Expression::Unary { op, expr } => {
                self.compile_expr(expr)?;
                match op {
                    UnaryOp::Neg => common::math::emit_neg(self.chunk(), line),
                    UnaryOp::Not => self.emit(Op::dyn_not),
                    UnaryOp::Deref => {
                        let idx = self.str_const("__value");
                        self.emit_u16(Op::struct_get, idx);
                    }
                    UnaryOp::AddrOf => {} // no-op in WASM
                }
            }
            Expression::AddrOf(e) => self.compile_expr(e)?,
            Expression::Deref(e) => {
                self.compile_expr(e)?;
                let idx = self.str_const("__value");
                self.emit_u16(Op::struct_get, idx);
            }
            Expression::Field { record, field } => {
                // TClassName.Create → constructor call (Pascal allows no parens)
                if let Expression::Identifier(class_name) = record.as_ref() {
                    if self.classes.contains_key(class_name.as_str()) && field.eq_ignore_ascii_case("Create") {
                        let idx = self.str_const(class_name);
                        self.emit_u16(Op::global_get, idx);
                        self.emit_u8(Op::call_ref, 0);
                        return Ok(());
                    }
                }
                self.compile_expr(record)?;
                let idx = self.str_const(field);
                self.emit_u16(Op::struct_get, idx);
            }
            Expression::Index { array, index } => {
                self.compile_expr(array)?;
                self.compile_expr(index)?;
                self.emit(Op::array_get);
            }
            Expression::Cast { type_name, expr } => {
                self.compile_expr(expr)?;
                match type_name.to_lowercase().as_str() {
                    "integer" | "longint" | "int64" | "cardinal" | "byte" | "word" | "shortint" => {
                        let i = self.import("vybe:convert", "cint");
                        self.emit_host_call(i, 1);
                    }
                    "real" | "double" | "single" | "extended" => {
                        let i = self.import("vybe:convert", "cdbl");
                        self.emit_host_call(i, 1);
                    }
                    "string" => {
                        let i = self.import("vybe:convert", "toString");
                        self.emit_host_call(i, 1);
                    }
                    "boolean" => common::convert::emit_to_bool(self.chunk(), line),
                    _ => {}
                }
            }
            Expression::SetLiteral(elems) | Expression::ArrayLiteral(elems) => {
                for e in elems { self.compile_expr(e)?; }
                self.chunks[self.current].emit_op_u16(Op::array_new, elems.len() as u16, line);
            }
            Expression::Call { callee, args } => {
                self.compile_call(callee, args)?;
            }
            Expression::Inherited { method, args } => {
                // inherited Create(args) → call parent constructor
                if let Some(method_name) = method {
                    if let Some(ref class_name) = self.current_class.clone() {
                        let parent = self.classes.get(class_name).and_then(|c| c.parent.clone());
                        if let Some(parent_name) = parent {
                            let parent_idx = self.str_const(&parent_name);
                            self.emit_u16(Op::global_get, parent_idx);
                            for a in args { self.compile_expr(a)?; }
                            self.emit_u8(Op::call_ref, args.len() as u8);
                            if method_name.to_lowercase() == "create" {
                                // Store result in Self slot
                                self.emit(Op::dup);
                                if let Some(slot) = self.scope().resolve("Self") {
                                    self.emit_u16(Op::local_set, slot);
                                    self.emit(Op::drop);
                                }
                            }
                        }
                    }
                }
            }
        }
        Ok(())
    }

    fn compile_call(&mut self, callee: &Expression, args: &[Expression]) -> Result<(), String> {
        let line = self.line;
        // Builtin function check
        if let Expression::Identifier(name) = callee {
            if self.try_compile_builtin(name, args)? { return Ok(()); }
        }
        // Class constructor/static call: TClassName.Create(args) or TClassName.StaticMethod(args)
        if let Expression::Field { record, field } = callee {
            if let Expression::Identifier(class_name) = record.as_ref() {
                if self.classes.contains_key(class_name.as_str()) {
                    let lower = field.to_lowercase();
                    if lower == "create" {
                        // Constructor: global_get(class) + args + call_ref(argc)
                        let idx = self.str_const(class_name);
                        self.emit_u16(Op::global_get, idx);
                        for a in args { self.compile_expr(a)?; }
                        self.emit_u8(Op::call_ref, args.len() as u8);
                    } else {
                        // Static method: global_get(class) + struct_get(method) + args + call_ref
                        let idx = self.str_const(class_name);
                        self.emit_u16(Op::global_get, idx);
                        let prop_idx = self.str_const(field);
                        self.emit_u16(Op::struct_get, prop_idx);
                        for a in args { self.compile_expr(a)?; }
                        self.emit_u8(Op::call_ref, args.len() as u8);
                    }
                    return Ok(());
                }
            }
        }

        // Method call: obj.method(args) — get method from obj, pass obj as self
        if let Expression::Field { record, field } = callee {
            self.compile_expr(record)?;        // push obj
            let prop_idx = self.str_const(field);
            self.emit(Op::dup);                // [obj, obj]
            self.emit_u16(Op::struct_get, prop_idx); // [obj, method_ref]
            // Swap so method_ref is below obj (call_ref expects [fn, self, args...])
            // Actually call_ref expects fn on stack first, then args.
            // But our stack is [obj, method_ref]. We need [method_ref, obj, args].
            // Use a temp local to reorder.
            let fn_tmp = self.scope_mut().define("__fn_tmp");
            self.emit_u16(Op::local_set, fn_tmp);
            self.emit(Op::drop);
            // Stack: [obj]. Push fn, then obj, then args.
            let obj_tmp = self.scope_mut().define("__obj_tmp");
            self.emit_u16(Op::local_set, obj_tmp);
            self.emit(Op::drop);
            self.emit_u16(Op::local_get, fn_tmp);   // [fn]
            self.emit_u16(Op::local_get, obj_tmp);   // [fn, obj]
            for a in args { self.compile_expr(a)?; }  // [fn, obj, args...]
            self.emit_u8(Op::call_ref, (args.len() + 1) as u8);
            return Ok(());
        }
        // Regular function call
        if let Expression::Identifier(name) = callee {
            // Check if resolved locally/globally or needs import
            let is_resolved = self.resolve_var(name) != VarRes::Global
                || self.defined_globals.contains(name.as_str());
            if is_resolved {
                self.emit_var_get(name);
                for a in args { self.compile_expr(a)?; }
                self.emit_u8(Op::call_ref, args.len() as u8);
            } else {
                // Unresolved — emit call_import for cross-language interop
                for a in args { self.compile_expr(a)?; }
                let import_idx = self.import("*", name);
                self.emit_host_call(import_idx, args.len() as u8);
            }
        } else {
            self.compile_expr(callee)?;
            for a in args { self.compile_expr(a)?; }
            self.emit_u8(Op::call_ref, args.len() as u8);
        }
        Ok(())
    }

    fn try_compile_builtin(&mut self, name: &str, args: &[Expression]) -> Result<bool, String> {
        let line = self.line;

        // Check common cross-language import table
        if let Some((module, func)) = common::imports::resolve_common_import(name) {
            for a in args { self.compile_expr(a)?; }
            let idx = self.import(module, func);
            self.emit_host_call(idx, args.len() as u8);
            return Ok(true);
        }

        match name.to_lowercase().as_str() {
            // I/O — use chunk 0 imports
            "writeln" | "write" => {
                for a in args { self.compile_expr(a)?; }
                let idx = self.import("wasi:cli", "log");
                common::io::emit_print_with_import(self.chunk(), idx, args.len() as u8, line);
                Ok(true)
            }
            "readln" | "read" => {
                if args.is_empty() {
                    let idx = self.import("wasi:cli", "readLine");
                    common::io::emit_input_with_import(self.chunk(), idx, line);
                    self.emit(Op::drop);
                } else if let Some(Expression::Identifier(var)) = args.first() {
                    let idx = self.import("wasi:cli", "readLine");
                    common::io::emit_input_with_import(self.chunk(), idx, line);
                    let var = var.clone();
                    self.emit_var_set(&var);
                }
                Ok(true)
            }
            // Math — common::math (direct WASM opcodes)
            "abs" => { self.compile_expr(&args[0])?; common::math::emit_abs(self.chunk(), line); Ok(true) }
            "sqrt" => { self.compile_expr(&args[0])?; common::math::emit_sqrt(self.chunk(), line); Ok(true) }
            "round" => { self.compile_expr(&args[0])?; common::math::emit_round(self.chunk(), line); Ok(true) }
            "trunc" | "int" => { self.compile_expr(&args[0])?; common::math::emit_trunc(self.chunk(), line); Ok(true) }
            "sin" => { self.compile_expr(&args[0])?; let i = self.import("vybe:math","sin"); self.emit_host_call(i, 1); Ok(true) }
            "cos" => { self.compile_expr(&args[0])?; let i = self.import("vybe:math","cos"); self.emit_host_call(i, 1); Ok(true) }
            "exp" => { self.compile_expr(&args[0])?; let i = self.import("vybe:math","exp"); self.emit_host_call(i, 1); Ok(true) }
            "ln" => { self.compile_expr(&args[0])?; let i = self.import("vybe:math","log"); self.emit_host_call(i, 1); Ok(true) }
            "sqr" => {
                self.compile_expr(&args[0])?;
                self.emit(Op::dup);
                self.emit(Op::f64_mul);
                Ok(true)
            }
            "power" => { self.compile_expr(&args[0])?; self.compile_expr(&args[1])?; let i = self.import("vybe:math","pow"); self.emit_host_call(i, 2); Ok(true) }
            "random" => { let i = self.import("vybe:math","random"); self.emit_host_call(i, 0); Ok(true) }
            "randomize" => { self.emit(Op::null); Ok(true) }
            "min" => { self.compile_expr(&args[0])?; self.compile_expr(&args[1])?; common::math::emit_min(self.chunk(), line); Ok(true) }
            "max" => { self.compile_expr(&args[0])?; self.compile_expr(&args[1])?; common::math::emit_max(self.chunk(), line); Ok(true) }
            "floor" => { self.compile_expr(&args[0])?; common::math::emit_floor(self.chunk(), line); Ok(true) }
            "ceil" => { self.compile_expr(&args[0])?; common::math::emit_ceil(self.chunk(), line); Ok(true) }
            // String — common::strings
            "length" => { self.compile_expr(&args[0])?; common::strings::emit_length(self.chunk(), line); Ok(true) }
            "upcase" | "uppercase" => { self.compile_expr(&args[0])?; common::strings::emit_to_upper(self.chunk(), line); Ok(true) }
            "lowercase" => { self.compile_expr(&args[0])?; common::strings::emit_to_lower(self.chunk(), line); Ok(true) }
            "trim" => { self.compile_expr(&args[0])?; common::strings::emit_trim(self.chunk(), line); Ok(true) }
            "pos" => { self.compile_expr(&args[0])?; self.compile_expr(&args[1])?; common::strings::emit_index_of(self.chunk(), line); Ok(true) }
            "copy" => {
                self.compile_expr(&args[0])?;
                self.compile_expr(&args[1])?;
                if args.len() > 2 { self.compile_expr(&args[2])?; }
                common::strings::emit_substring(self.chunk(), line);
                Ok(true)
            }
            "concat" => {
                for a in args { self.compile_expr(a)?; }
                common::strings::emit_concat(self.chunk(), args.len(), line);
                Ok(true)
            }
            // Type conversions — route imports through chunk 0
            "inttostr" | "floattostr" => {
                self.compile_expr(&args[0])?;
                let idx = self.import("vybe:convert", "toString");
                self.emit_host_call(idx, 1);
                Ok(true)
            }
            "strtoint" | "strtointdef" => {
                self.compile_expr(&args[0])?;
                let idx = self.import("vybe:convert", "parseInt");
                self.emit_host_call(idx, 1);
                Ok(true)
            }
            "strtofloat" => {
                self.compile_expr(&args[0])?;
                let idx = self.import("vybe:convert", "parseFloat");
                self.emit_host_call(idx, 1);
                Ok(true)
            }
            "chr" => {
                self.compile_expr(&args[0])?;
                let idx = self.import("vybe:string", "chr");
                self.emit_host_call(idx, 1);
                Ok(true)
            }
            "ord" => {
                self.compile_expr(&args[0])?;
                let idx = self.import("vybe:string", "charCodeAt");
                self.emit_host_call(idx, 1);
                Ok(true)
            }
            "succ" => { self.compile_expr(&args[0])?; self.emit_const(Value::F64(1.0)); self.emit(Op::dyn_add); Ok(true) }
            "pred" => { self.compile_expr(&args[0])?; self.emit_const(Value::F64(1.0)); self.emit(Op::f64_sub); Ok(true) }
            "inc" => {
                if let Some(Expression::Identifier(var)) = args.first() {
                    let var = var.clone();
                    self.emit_var_get(&var);
                    if args.len() > 1 { self.compile_expr(&args[1])?; } else { self.emit_const(Value::F64(1.0)); }
                    self.emit(Op::dyn_add);
                    self.emit_var_set(&var);
                }
                self.emit(Op::null);
                Ok(true)
            }
            "dec" => {
                if let Some(Expression::Identifier(var)) = args.first() {
                    let var = var.clone();
                    self.emit_var_get(&var);
                    if args.len() > 1 { self.compile_expr(&args[1])?; } else { self.emit_const(Value::F64(1.0)); }
                    self.emit(Op::f64_sub);
                    self.emit_var_set(&var);
                }
                self.emit(Op::null);
                Ok(true)
            }
            "assigned" => { self.compile_expr(&args[0])?; self.emit(Op::ref_is_null); self.emit(Op::dyn_not); Ok(true) }
            "high" => {
                self.compile_expr(&args[0])?;
                common::strings::emit_length(self.chunk(), line);
                self.emit_const(Value::F64(1.0));
                self.emit(Op::f64_sub);
                Ok(true)
            }
            "low" => { self.emit_const(Value::F64(0.0)); Ok(true) }
            "setlength" => {
                if let Some(Expression::Identifier(var)) = args.first() {
                    let var = var.clone();
                    self.compile_expr(&args[1])?;
                    let idx = self.import("vybe:array", "newWithLength");
                    self.emit_host_call(idx, 1);
                    self.emit_var_set(&var);
                }
                self.emit(Op::null);
                Ok(true)
            }
            _ => Ok(false),
        }
    }

    fn is_class_field(&self, name: &str) -> bool {
        if let Some(ref class_name) = self.current_class {
            let mut current = Some(class_name.as_str());
            while let Some(cn) = current {
                if let Some(info) = self.classes.get(cn) {
                    if info.fields.iter().any(|f| f.eq_ignore_ascii_case(name)) {
                        return true;
                    }
                    current = info.parent.as_deref();
                } else {
                    break;
                }
            }
        }
        false
    }

    // ── Class compilation ──────────────────────────────────────────────────

    /// Register a class from its type declaration. Methods are compiled later
    /// when their implementations are encountered.
    fn register_class(&mut self, name: &str, class_def: &ClassDef) -> Result<(), String> {
        let mut fields = Vec::new();
        let mut method_sigs = Vec::new();
        let mut properties = Vec::new();

        for member in &class_def.members {
            match member {
                ClassMember::Field(v) => {
                    for fname in &v.names { fields.push(fname.clone()); }
                }
                ClassMember::MethodDecl(sig) => {
                    method_sigs.push(sig.clone());
                }
                ClassMember::PropertyDecl(prop) => {
                    properties.push(prop.clone());
                }
            }
        }

        self.classes.insert(name.to_string(), ClassInfo {
            parent: class_def.parent.clone(),
            fields,
            method_sigs,
            compiled_methods: Vec::new(),
            properties,
        });
        self.defined_globals.insert(name.to_string());
        Ok(())
    }

    /// Compile a method implementation (constructor/destructor/procedure/function ClassName.MethodName).
    fn compile_method_impl(&mut self, method: &MethodImpl) -> Result<(), String> {
        let class_name = &method.class_name;
        let method_name = &method.method_name;

        // Determine arity: self + user params
        let user_arity: u8 = method.params.iter().map(|p| p.names.len() as u8).sum();
        let is_constructor = method.kind == MethodKind::Constructor;
        let is_function = method.kind == MethodKind::Function;

        let func_idx = self.chunks.len();
        // All methods (including constructors) take Self as first param
        let full_arity = user_arity + 1;
        let chunk = common::functions::create_function_chunk(
            &format!("{}.{}", class_name, method_name), full_arity,
        );
        self.chunks.push(chunk);
        self.scopes.push(Scope::new_function());
        let saved = self.current;
        self.current = func_idx;

        let saved_class = self.current_class.take();
        self.current_class = Some(class_name.clone());

        // All methods get Self as first param
        self.scope_mut().define("Self");
        for param in &method.params {
            for pname in &param.names { self.scope_mut().define(pname); }
        }

        if is_constructor {
            // Constructor body: just compile the body statements.
            // The wrapper (built in finalize_classes) handles object creation + method binding.
            // Return Self at end (for child classes, `inherited Create` sets Self).
            for decl in &method.decls { self.compile_decl(decl)?; }
            for stmt in &method.body { self.compile_stmt(stmt)?; }
            // Return Self so the wrapper can capture it
            if let Some(slot) = self.scope().resolve("Self") {
                self.emit_u16(Op::local_get, slot);
                self.emit(Op::r#return);
            } else {
                let line = self.line;
                common::functions::emit_function_epilogue(&mut self.chunks[func_idx], line);
            }
        } else if is_function {
            // Function method: has Result slot
            let result_slot = self.scope_mut().define("Result");
            self.emit(Op::null);
            self.emit_u16(Op::local_set, result_slot);
            self.emit(Op::drop);

            let saved_fn = self.current_func_name.take();
            let saved_rs = self.current_result_slot.take();
            self.current_func_name = Some(method_name.clone());
            self.current_result_slot = Some(result_slot);

            for decl in &method.decls { self.compile_decl(decl)?; }
            for stmt in &method.body { self.compile_stmt(stmt)?; }

            self.emit_u16(Op::local_get, result_slot);
            self.emit(Op::r#return);

            self.current_func_name = saved_fn;
            self.current_result_slot = saved_rs;
        } else {
            // Procedure method or destructor
            for decl in &method.decls { self.compile_decl(decl)?; }
            for stmt in &method.body { self.compile_stmt(stmt)?; }
            let line = self.line;
            common::functions::emit_function_epilogue(&mut self.chunks[func_idx], line);
        }

        self.current_class = saved_class;

        let locals = self.scope().next_slot;
        self.chunks[func_idx].local_count = locals;
        self.scopes.pop();
        self.current = saved;

        // Store compiled method in class info
        let is_static = self.classes.get(class_name)
            .and_then(|c| c.method_sigs.iter().find(|s| s.name == *method_name))
            .map(|s| s.is_static)
            .unwrap_or(false);

        if let Some(info) = self.classes.get_mut(class_name) {
            info.compiled_methods.push(CompiledMethod {
                name: method_name.clone(),
                chunk_idx: func_idx,
                is_static,
            });
        }

        Ok(())
    }

    /// Build constructor wrappers for all registered classes.
    /// Called after all method implementations have been compiled.
    fn finalize_classes(&mut self) -> Result<(), String> {
        let class_names: Vec<String> = self.classes.keys().cloned().collect();
        for class_name in &class_names {
            let info = self.classes.get(class_name).unwrap();
            let parent = info.parent.clone();
            let is_child = parent.is_some();
            let fields = info.fields.clone();
            let methods: Vec<_> = info.compiled_methods.iter()
                .filter(|m| !m.is_static)
                .map(|m| (m.name.clone(), m.chunk_idx))
                .collect();
            let statics: Vec<_> = info.compiled_methods.iter()
                .filter(|m| m.is_static)
                .map(|m| (m.name.clone(), m.chunk_idx))
                .collect();
            let all_methods: Vec<_> = info.compiled_methods.iter()
                .map(|m| (m.name.clone(), m.chunk_idx))
                .collect();

            // Find constructor impl chunk
            let ctor_impl = info.compiled_methods.iter()
                .find(|m| m.name.eq_ignore_ascii_case("Create"))
                .map(|m| m.chunk_idx);

            // Determine user arity from the constructor chunk
            let user_arity = ctor_impl
                .map(|ci| {
                    let a = self.chunks[ci].arity;
                    if a > 0 { a - 1 } else { 0 } // subtract 1 for Self
                })
                .unwrap_or(0);

            // Build wrapper chunk
            let wrapper_idx = self.chunks.len();
            let wrapper = common::functions::create_function_chunk(
                &format!("{}_ctor", class_name), user_arity,
            );
            self.chunks.push(wrapper);
            let line = self.line;

            // Params at slots 1..N, this at slot N+1
            let this_slot = (user_arity as u16) + 1;
            self.chunks[wrapper_idx].local_count = this_slot + 1;

            if is_child {
                // Child class: same pattern as Ruby — init body calls inherited Create,
                // which creates the object and returns it.
                self.chunks[wrapper_idx].local_count = this_slot + 1;

                if let Some(init_ci) = ctor_impl {
                    self.chunks[wrapper_idx].emit_op(Op::null, line);
                    self.chunks[wrapper_idx].emit_op_u16(Op::local_set, this_slot, line);
                    // ref_func for init method
                    common::functions::emit_ref_func(&mut self.chunks[wrapper_idx], init_ci, 0, line);
                    self.chunks[wrapper_idx].emit_op_u16(Op::local_get, this_slot, line);
                    for i in 0..user_arity {
                        self.chunks[wrapper_idx].emit_op_u16(Op::local_get, (i as u16) + 1, line);
                    }
                    self.chunks[wrapper_idx].emit_op_u8(Op::call_ref, (user_arity + 1) as u8, line);
                    // init returns the object — store as this
                    self.chunks[wrapper_idx].emit_op_u16(Op::local_set, this_slot, line);
                }

                // Bind child methods (override parent's)
                for (mname, mci) in &methods {
                    common::classes::emit_bind_method_with_aliases(
                        &mut self.chunks[wrapper_idx], this_slot, mname, *mci, line,
                    );
                }

                common::classes::emit_instanceof_chain(&mut self.chunks[wrapper_idx], this_slot, class_name, line);
            } else {
                // Base class: same pattern as Ruby — create object, bind methods, call init.
                common::classes::emit_new_typed_object(&mut self.chunks[wrapper_idx], this_slot, class_name, line);
                self.chunks[wrapper_idx].local_count = this_slot + 1;

                for fname in &fields {
                    common::classes::emit_init_field_null(&mut self.chunks[wrapper_idx], this_slot, fname, line);
                }

                for (mname, mci) in &methods {
                    common::classes::emit_bind_method_with_aliases(
                        &mut self.chunks[wrapper_idx], this_slot, mname, *mci, line,
                    );
                }

                // Call init body: ref_func + self + args + call_ref + drop
                if let Some(init_ci) = ctor_impl {
                    common::functions::emit_ref_func(&mut self.chunks[wrapper_idx], init_ci, 0, line);
                    self.chunks[wrapper_idx].emit_op_u16(Op::local_get, this_slot, line);
                    for i in 0..user_arity {
                        self.chunks[wrapper_idx].emit_op_u16(Op::local_get, (i as u16) + 1, line);
                    }
                    self.chunks[wrapper_idx].emit_op_u8(Op::call_ref, (user_arity + 1) as u8, line);
                    self.chunks[wrapper_idx].emit_op(Op::drop, line);
                }

                common::classes::emit_instanceof_chain(&mut self.chunks[wrapper_idx], this_slot, class_name, line);
            }

            common::classes::emit_instanceof_chain(&mut self.chunks[wrapper_idx], this_slot, class_name, line);
            common::classes::emit_constructor_return(&mut self.chunks[wrapper_idx], this_slot, line);

            // Store wrapper as global constructor
            let ctor_local = self.scope_mut().define(&format!("__{}_ctor", class_name));
            common::classes::emit_store_constructor(
                &mut self.chunks[self.current], class_name, wrapper_idx, ctor_local, line,
            );

            // Attach static methods
            for (sname, sci) in &statics {
                common::classes::emit_attach_static_method(
                    &mut self.chunks[self.current], ctor_local, sname, *sci, line,
                );
            }

            // Register type
            let parent_str = parent.unwrap_or_default();
            common::classes::register_type(
                &mut self.chunks, class_name, &parent_str,
                fields.clone(), all_methods, false, Vec::new(), Some(wrapper_idx),
            );
        }
        Ok(())
    }
}

#[derive(PartialEq)]
enum VarRes { Local(u16), Upvalue(u8), Global }
