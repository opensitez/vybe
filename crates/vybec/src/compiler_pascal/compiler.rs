use std::sync::Arc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use crate::parser_pascal::ast::*;
use vybe_compiler_common as common;
use super::scope::Scope;

struct LoopCtx {
    break_label_depth: u32,
    continue_label_depth: u32,
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
    #[allow(dead_code)]
    properties: Vec<PropertyDef>,
}

pub struct Compiler {
    chunks: Vec<Chunk>,
    scopes: Vec<Scope>,
    current: usize,
    loops: Vec<LoopCtx>,
    label_depth: u32,
    line: u32,
    defined_globals: std::collections::HashSet<String>,
    current_func_name: Option<String>,
    current_result_slot: Option<u16>,
    classes: std::collections::HashMap<String, ClassInfo>,
    current_class: Option<String>,
}

use crate::parser_pascal::ast::{ClassDef, ClassMember, MethodSig, MethodKind, MethodImpl, PropertyDef};

impl Compiler {
    pub fn new() -> Self {
        Self {
            chunks: vec![Chunk::new("<program>")],
            scopes: vec![Scope::new()],
            current: 0,
            loops: Vec::new(),
            label_depth: 0,
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
        self.emit(Op::NULL);
        self.emit(Op::HALT);
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
        self.emit_u16(Op::CONST, idx);
    }
    fn chunk(&mut self) -> &mut Chunk { &mut self.chunks[self.current] }
    fn str_const(&mut self, s: &str) -> u16 {
        self.chunks[self.current].add_constant(Value::String(Arc::from(s)))
    }

    /// Add import to chunk 0 (VM resolves imports from chunk 0 only).
    fn import(&mut self, module: &str, name: &str) -> u16 {
        self.chunks[0].add_import(module, name)
    }
    fn emit_host_call(&mut self, import_idx: u16, argc: u8) {
        let line = self.line;
        self.chunks[self.current].emit_op_u16(Op::CALL_IMPORT, import_idx, line);
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
            VarRes::Local(slot) => self.emit_u16(Op::LOCAL_GET, slot),
            VarRes::Upvalue(idx) => self.emit_u8(Op::UPVALUE_GET, idx),
            VarRes::Global => { let idx = self.str_const(name); self.emit_u16(Op::GLOBAL_GET, idx); }
        }
    }

    fn emit_var_set(&mut self, name: &str) {
        match self.resolve_var(name) {
            VarRes::Local(slot) => { self.emit_u16(Op::LOCAL_SET, slot); self.emit(Op::DROP); }
            VarRes::Upvalue(idx) => { self.emit_u8(Op::UPVALUE_SET, idx); self.emit(Op::DROP); }
            VarRes::Global => { let idx = self.str_const(name); self.emit_u16(Op::GLOBAL_SET, idx); self.emit(Op::DROP); }
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
                                | "real" | "double" | "single" | "extended" => self.emit(Op::F64_CONST_0),
                                "boolean" => self.emit(Op::FALSE),
                                "string" => self.emit_const(Value::String(Arc::from(""))),
                                _ => self.emit(Op::NULL),
                            }
                        }
                        if self.scopes.len() == 1 && self.scope().depth == 0 {
                            let idx = self.str_const(name);
                            self.emit_u16(Op::GLOBAL_SET, idx);
                            self.emit(Op::DROP);
                            self.defined_globals.insert(name.clone());
                        } else {
                            let slot = self.scope_mut().define(name);
                            self.emit_u16(Op::LOCAL_SET, slot);
                            self.emit(Op::DROP);
                        }
                    }
                }
            }
            Decl::Const(consts) => {
                for c in consts {
                    self.compile_expr(&c.value)?;
                    if self.scopes.len() == 1 {
                        let idx = self.str_const(&c.name);
                        self.emit_u16(Op::GLOBAL_SET, idx);
                        self.emit(Op::DROP);
                    } else {
                        let slot = self.scope_mut().define(&c.name);
                        self.emit_u16(Op::LOCAL_SET, slot);
                        self.emit(Op::DROP);
                    }
                }
            }
            Decl::Type(types) => {
                for t in types {
                    match &t.def {
                        TypeDef::Class(class_def) => {
                            self.register_class(&t.name, class_def)?;
                        }
                        TypeDef::Enum(values) => {
                            // Compile enum: each value becomes a global constant (ordinal)
                            for (i, ev) in values.iter().enumerate() {
                                let val = if let Some(ref expr) = ev.value {
                                    // Will be evaluated at runtime, but for simple ints:
                                    self.compile_expr(expr)?;
                                } else {
                                    self.emit_const(Value::F64(i as f64));
                                };
                                let _ = val;
                                let idx = self.str_const(&ev.name);
                                self.emit_u16(Op::GLOBAL_SET, idx);
                                self.emit(Op::DROP);
                                self.defined_globals.insert(ev.name.clone());
                            }
                        }
                        TypeDef::InterfaceDef(_) => {
                            // Interfaces are compile-time only (method contracts)
                            self.defined_globals.insert(t.name.clone());
                        }
                        TypeDef::Record(_) => {
                            // Record types: compile-time only for now
                            // (record instances are created as objects)
                        }
                        _ => {} // Alias, Array, Pointer — type-only
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
        self.emit_u16(Op::GLOBAL_SET, idx);
        self.emit(Op::DROP);
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
        self.emit(Op::NULL);
        self.emit_u16(Op::LOCAL_SET, result_slot);
        self.emit(Op::DROP);
        // Track function name → Result slot so `FuncName := value` works
        let saved_func_name = self.current_func_name.take();
        let saved_result_slot = self.current_result_slot.take();
        self.current_func_name = Some(func.name.clone());
        self.current_result_slot = Some(result_slot);


        for decl in &func.decls { self.compile_decl(decl)?; }
        for stmt in &func.body { self.compile_stmt(stmt)?; }

        // Return Result
        self.emit_u16(Op::LOCAL_GET, result_slot);
        self.emit(Op::RETURN);

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
        self.emit_u16(Op::GLOBAL_SET, idx);
        self.emit(Op::DROP);
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
                self.emit(Op::DROP);
            }
            Statement::If { cond, then, else_ } => {
                let line = self.line;
                let outer = self.chunk().emit_block(line); self.label_depth += 1;
                let then_block = self.chunk().emit_block(line); self.label_depth += 1;
                self.compile_expr(cond)?;
                self.emit(Op::DYN_TO_BOOL);
                self.emit(Op::DYN_NOT);
                self.chunk().emit_br_if(0, line); // skip then if false
                self.compile_stmt(then)?;
                self.chunk().emit_br(1, line); // jump to outer end
                self.chunk().emit_end(line); self.chunk().patch_block(then_block); self.label_depth -= 1;
                if let Some(alt) = else_ {
                    self.compile_stmt(alt)?;
                }
                self.chunk().emit_end(line); self.chunk().patch_block(outer); self.label_depth -= 1;
            }
            Statement::For { var, from, to, downto, body } => {
                let line = self.line;
                // var := from
                self.compile_expr(from)?;
                self.emit_var_set(var);

                let block_p = self.chunk().emit_block(line);
                let (loop_p, _) = self.chunk().emit_loop_s(line);
                self.label_depth += 2;

                // For loop has update after body, use body block for continue
                let body_block_p = self.chunk().emit_block(line);
                self.label_depth += 1;

                let break_depth = self.label_depth - 2; // the outer block
                let continue_depth = self.label_depth; // the body block (continue skips to increment)
                self.loops.push(LoopCtx { break_label_depth: break_depth, continue_label_depth: continue_depth });

                // condition: var <= to  (or >= for downto)
                self.emit_var_get(var);
                self.compile_expr(to)?;
                if *downto { self.emit(Op::DYN_GE); } else { self.emit(Op::DYN_LE); }
                common::loops::emit_loop_cond(self.chunk(), line);

                self.compile_stmt(body)?;

                // end body block (continue lands here)
                self.chunk().emit_end(line); self.chunk().patch_block(body_block_p); self.label_depth -= 1;

                // inc/dec
                self.emit_var_get(var);
                self.emit_const(Value::F64(1.0));
                if *downto {
                    self.emit(Op::F64_SUB);
                } else {
                    self.emit(Op::DYN_ADD);
                }
                self.emit_var_set(var);

                self.chunk().emit_br(0, line); // continue loop
                self.chunk().emit_end(line); self.chunk().patch_loop(loop_p); self.label_depth -= 1;
                self.chunk().emit_end(line); self.chunk().patch_block(block_p); self.label_depth -= 1;
                self.loops.pop();
            }
            Statement::While { cond, body } => {
                let line = self.line;
                let lp = common::loops::emit_loop_start(self.chunk(), line);
                self.label_depth += 2;
                self.loops.push(LoopCtx { break_label_depth: self.label_depth - 1, continue_label_depth: self.label_depth });
                self.compile_expr(cond)?;
                common::loops::emit_loop_cond(self.chunk(), line);
                self.compile_stmt(body)?;
                common::loops::emit_loop_end(self.chunk(), lp, line);
                self.label_depth -= 2;
                self.loops.pop();
            }
            Statement::Repeat { body, until } => {
                let line = self.line;
                let lp = common::loops::emit_do_loop_start(self.chunk(), line);
                self.label_depth += 2;
                self.loops.push(LoopCtx { break_label_depth: self.label_depth - 1, continue_label_depth: self.label_depth });
                for s in body { self.compile_stmt(s)?; }
                self.compile_expr(until)?;
                // repeat..until: negate=true because we loop while condition is FALSE
                common::loops::emit_do_loop_end(self.chunk(), lp, true, line);
                self.label_depth -= 2;
                self.loops.pop();
            }
            Statement::Case { expr, arms, else_ } => {
                let line = self.line;
                self.compile_expr(expr)?;

                // Outer block for the whole case statement
                let outer = self.chunk().emit_block(line); self.label_depth += 1;

                for arm in arms {
                    // Each arm: block { check conditions, skip if no match; body; br outer }
                    let arm_block = self.chunk().emit_block(line); self.label_depth += 1;

                    // Match block: conditions jump INTO body if matched
                    let match_block = self.chunk().emit_block(line); self.label_depth += 1;

                    for val in &arm.values {
                        match val {
                            CaseValue::Single(v) => {
                                self.emit(Op::DUP);
                                self.compile_expr(v)?;
                                self.emit(Op::DYN_EQ);
                                self.emit(Op::DYN_TO_BOOL);
                                // If match, skip to end of match_block (fall through to body)
                                self.chunk().emit_br_if(0, line);
                            }
                            CaseValue::Range(lo, hi) => {
                                // Check lo <= val <= hi
                                self.emit(Op::DUP);
                                self.compile_expr(lo)?;
                                self.emit(Op::DYN_GE);
                                self.emit(Op::DYN_TO_BOOL);
                                self.emit(Op::DYN_NOT);
                                // If not >= lo, skip this range check
                                let range_skip = self.chunk().emit_block(line); self.label_depth += 1;
                                self.chunk().emit_br_if(0, line); // skip if not >= lo
                                self.emit(Op::DUP);
                                self.compile_expr(hi)?;
                                self.emit(Op::DYN_LE);
                                self.emit(Op::DYN_TO_BOOL);
                                // If match, skip to end of match_block
                                self.chunk().emit_br_if(1, line);
                                self.chunk().emit_end(line); self.chunk().patch_block(range_skip); self.label_depth -= 1;
                            }
                        }
                    }
                    // No condition matched — skip body
                    self.chunk().emit_br(1, line); // skip arm_block to next arm
                    self.chunk().emit_end(line); self.chunk().patch_block(match_block); self.label_depth -= 1;

                    // Body (reached when a condition matched)
                    for s in &arm.body { self.compile_stmt(s)?; }
                    // Jump to outer end (skip remaining arms)
                    // depth to outer = self.label_depth - (outer_label_depth)
                    // outer is at label_depth - 1 (arm_block) - 1 (outer) from current
                    self.chunk().emit_br(1, line);
                    self.chunk().emit_end(line); self.chunk().patch_block(arm_block); self.label_depth -= 1;
                }

                if let Some(else_stmts) = else_ {
                    for s in else_stmts { self.compile_stmt(s)?; }
                }
                self.chunk().emit_end(line); self.chunk().patch_block(outer); self.label_depth -= 1;
                self.emit(Op::DROP);
            }
            Statement::With { vars, body } => {
                // `with obj do` — compile body. In a proper implementation
                // we'd add obj's fields to scope. For now, store obj in a temp
                // and make it accessible as the implicit self.
                if let Some(first) = vars.first() {
                    self.compile_expr(first)?;
                    let with_slot = self.scope_mut().define("__with_obj");
                    self.emit_u16(Op::LOCAL_SET, with_slot);
                    self.emit(Op::DROP);
                }
                self.compile_stmt(body)?;
            }
            Statement::Try { body, handler } => {
                let line = self.line;
                let outer = self.chunk().emit_block(line); self.label_depth += 1;
                let try_block = self.chunk().emit_block(line); self.label_depth += 1;

                let catch_jump = common::errors::emit_try_start(&mut self.chunks[self.current], line);

                for s in body { self.compile_stmt(s)?; }
                common::errors::emit_try_end(&mut self.chunks[self.current], line);
                self.chunk().emit_br(1, line); // skip to outer end

                self.chunk().emit_end(line); self.chunk().patch_block(try_block); self.label_depth -= 1;

                common::errors::patch_catch(&mut self.chunks[self.current], catch_jump);

                match handler {
                    TryHandler::Except(clauses, else_stmts) => {
                        if clauses.is_empty() {
                            self.emit(Op::DROP);
                            if let Some(stmts) = else_stmts {
                                for s in stmts { self.compile_stmt(s)?; }
                            }
                        } else {
                            for clause in clauses {
                                if let Some(var) = &clause.var_name {
                                    let slot = self.scope_mut().define(var);
                                    self.emit_u16(Op::LOCAL_SET, slot);
                                    self.emit(Op::DROP);
                                } else {
                                    self.emit(Op::DROP);
                                }
                                for s in &clause.body { self.compile_stmt(s)?; }
                            }
                        }
                    }
                    TryHandler::Finally(stmts) => {
                        self.emit(Op::DROP);
                        for s in stmts { self.compile_stmt(s)?; }
                    }
                }
                self.chunk().emit_end(line); self.chunk().patch_block(outer); self.label_depth -= 1;
            }
            Statement::Raise(expr) => {
                if let Some(e) = expr { self.compile_expr(e)?; } else { self.emit(Op::NULL); }
                let line = self.line;
                common::errors::emit_throw(&mut self.chunks[self.current], line);
            }
            Statement::Exit(val) => {
                if let Some(v) = val {
                    self.compile_expr(v)?;
                } else if let Some(slot) = self.current_result_slot {
                    // In a function, Exit without value returns Result
                    self.emit_u16(Op::LOCAL_GET, slot);
                } else {
                    self.emit(Op::NULL);
                }
                self.emit(Op::RETURN);
            }
            Statement::Break => {
                let line = self.line;
                if let Some(ctx) = self.loops.last() {
                    let depth = (self.label_depth - ctx.break_label_depth) as u8;
                    self.chunk().emit_br(depth, line);
                }
            }
            Statement::Continue => {
                let line = self.line;
                if let Some(ctx) = self.loops.last() {
                    let depth = (self.label_depth - ctx.continue_label_depth) as u8;
                    self.chunk().emit_br(depth, line);
                }
            }
            Statement::CompoundAssign { target, op, value } => {
                // target += value → target := target + value
                self.compile_expr(target)?;
                self.compile_expr(value)?;
                match op {
                    CompoundOp::Add => self.emit(Op::DYN_ADD),
                    CompoundOp::Sub => self.emit(Op::F64_SUB),
                    CompoundOp::Mul => self.emit(Op::F64_MUL),
                    CompoundOp::Div => self.emit(Op::F64_DIV),
                }
                self.compile_assign_target(target)?;
            }
            Statement::ForIn { var, collection, body } => {
                // for item in collection do body
                // Uses common::loops::emit_for_in_start/end (already structured CF)
                self.compile_expr(collection)?;
                let arr_slot = self.scope_mut().define("__forin_arr");
                self.emit_u16(Op::LOCAL_SET, arr_slot);
                self.emit(Op::DROP);
                let idx_slot = self.scope_mut().define("__forin_idx");
                let line = self.line;
                let lp = common::loops::emit_for_in_start(
                    &mut self.chunks[self.current], arr_slot, idx_slot, line,
                );
                // emit_for_in_start emits: block + loop + cond + block $body = 3 labels
                let break_depth = self.label_depth + 1; // outer block
                let continue_depth = self.label_depth + 3; // body block (innermost)
                self.label_depth += 3;
                // Element is on stack from emit_for_in_start
                self.emit_var_set(var);
                self.loops.push(LoopCtx { break_label_depth: break_depth, continue_label_depth: continue_depth });
                self.compile_stmt(body)?;
                self.loops.pop();
                common::loops::emit_for_in_end(
                    &mut self.chunks[self.current], idx_slot, lp, line,
                );
                self.label_depth -= 3;
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
                            self.emit_u16(Op::LOCAL_SET, slot);
                            self.emit(Op::DROP);
                            return Ok(());
                        }
                    }
                }
                // Inside class method, field assignment: FName := value → Self.FName := value
                if self.current_class.is_some() && self.is_class_field(name) {
                    if let Some(slot) = self.scope().resolve("Self") {
                        let tmp = self.scope_mut().define("__field_tmp");
                        self.emit_u16(Op::LOCAL_SET, tmp);
                        self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, slot); // Self
                        self.emit_u16(Op::LOCAL_GET, tmp);  // value
                        let idx = self.str_const(name);
                        self.emit_u16(Op::STRUCT_SET, idx);
                        self.emit(Op::DROP);
                        return Ok(());
                    }
                }
                self.emit_var_set(name);
            }
            Expression::Field { record, field } => {
                // value on stack; need [obj, val] for struct_set
                let tmp = self.scope_mut().define("__tmp");
                self.emit_u16(Op::LOCAL_SET, tmp);
                self.emit(Op::DROP);
                self.compile_expr(record)?;
                self.emit_u16(Op::LOCAL_GET, tmp);
                let idx = self.str_const(field);
                self.emit_u16(Op::STRUCT_SET, idx);
                self.emit(Op::DROP);
            }
            Expression::Index { array, index } => {
                let tmp = self.scope_mut().define("__tmp");
                self.emit_u16(Op::LOCAL_SET, tmp);
                self.emit(Op::DROP);
                self.compile_expr(array)?;
                self.compile_expr(index)?;
                self.emit_u16(Op::LOCAL_GET, tmp);
                self.emit(Op::ARRAY_SET);
                self.emit(Op::DROP);
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
                self.chunks[self.current].emit_op_u16(Op::CONST, idx, line);
            }
            Expression::Real(n) => {
                let idx = self.chunks[self.current].add_constant(Value::F64(*n));
                self.chunks[self.current].emit_op_u16(Op::CONST, idx, line);
            }
            Expression::Bool(b) => {
                if *b { self.emit(Op::TRUE); } else { self.emit(Op::FALSE); }
            }
            Expression::Nil => self.emit(Op::NULL),
            Expression::Str(s) => {
                let idx = self.chunks[self.current].add_constant(Value::String(Arc::from(s.as_str())));
                self.chunks[self.current].emit_op_u16(Op::CONST, idx, line);
            }
            Expression::Char(c) => {
                let idx = self.chunks[self.current].add_constant(Value::String(Arc::from(c.to_string().as_str())));
                self.chunks[self.current].emit_op_u16(Op::CONST, idx, line);
            }
            Expression::Identifier(name) => {
                match name.to_lowercase().as_str() {
                    "maxint" => self.emit_const(Value::F64(2147483647.0)),
                    "pi" => self.emit_const(Value::F64(std::f64::consts::PI)),
                    "self" => {
                        if let Some(slot) = self.scope().resolve("Self") {
                            self.emit_u16(Op::LOCAL_GET, slot);
                        } else {
                            self.emit(Op::NULL);
                        }
                    }
                    _ => {
                        // Inside a class method, check if name is a field
                        if self.current_class.is_some() && self.is_class_field(name) {
                            // FName → Self.FName
                            if let Some(slot) = self.scope().resolve("Self") {
                                self.emit_u16(Op::LOCAL_GET, slot);
                                let idx = self.str_const(name);
                                self.emit_u16(Op::STRUCT_GET, idx);
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
                    BinOp::Add => self.emit(Op::DYN_ADD),
                    BinOp::Sub => self.emit(Op::F64_SUB),
                    BinOp::Mul => self.emit(Op::F64_MUL),
                    BinOp::Div => self.emit(Op::F64_DIV),
                    BinOp::IDiv => {
                        self.emit(Op::F64_DIV);
                        common::math::emit_trunc(self.chunk(), line);
                    }
                    BinOp::Mod => { let idx = self.import("vybe:math", "fmod"); let l = self.line; vybe_compiler_common::expressions::emit_f64_mod_with_import(self.chunk(), idx, l); }
                    BinOp::Eq => self.emit(Op::DYN_EQ),
                    BinOp::NotEq => self.emit(Op::DYN_NE),
                    BinOp::Lt => self.emit(Op::DYN_LT),
                    BinOp::Gt => self.emit(Op::DYN_GT),
                    BinOp::Le => self.emit(Op::DYN_LE),
                    BinOp::Ge => self.emit(Op::DYN_GE),
                    BinOp::And | BinOp::Or => unreachable!(), // handled above
                    BinOp::Xor => self.emit(Op::I32_XOR),
                    BinOp::Shl => self.emit(Op::I32_SHL),
                    BinOp::Shr => self.emit(Op::I32_SHR_S),
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
                    UnaryOp::Not => self.emit(Op::DYN_NOT),
                    UnaryOp::Deref => {
                        let idx = self.str_const("__value");
                        self.emit_u16(Op::STRUCT_GET, idx);
                    }
                    UnaryOp::AddrOf => {} // no-op in WASM
                }
            }
            Expression::AddrOf(e) => self.compile_expr(e)?,
            Expression::Deref(e) => {
                self.compile_expr(e)?;
                let idx = self.str_const("__value");
                self.emit_u16(Op::STRUCT_GET, idx);
            }
            Expression::Field { record, field } => {
                // TClassName.Create → constructor call (Pascal allows no parens)
                if let Expression::Identifier(class_name) = record.as_ref() {
                    if self.classes.contains_key(class_name.as_str()) && field.eq_ignore_ascii_case("Create") {
                        let idx = self.str_const(class_name);
                        self.emit_u16(Op::GLOBAL_GET, idx);
                        self.emit_u8(Op::CALL_REF, 0);
                        return Ok(());
                    }
                    // ClassName.ClassVar → global lookup
                    if self.classes.contains_key(class_name.as_str()) {
                        let global_name = format!("{}.{}", class_name, field);
                        let idx = self.str_const(&global_name);
                        self.emit_u16(Op::GLOBAL_GET, idx);
                        return Ok(());
                    }
                }
                // Self.FName → explicit self field access
                if let Expression::Identifier(name) = record.as_ref() {
                    if name.eq_ignore_ascii_case("Self") {
                        if let Some(slot) = self.scope().resolve("Self") {
                            self.emit_u16(Op::LOCAL_GET, slot);
                            let idx = self.str_const(field);
                            self.emit_u16(Op::STRUCT_GET, idx);
                            return Ok(());
                        }
                    }
                }
                self.compile_expr(record)?;
                let idx = self.str_const(field);
                self.emit_u16(Op::STRUCT_GET, idx);
            }
            Expression::Index { array, index } => {
                self.compile_expr(array)?;
                self.compile_expr(index)?;
                self.emit(Op::ARRAY_GET);
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
                self.chunks[self.current].emit_op_u16(Op::ARRAY_NEW, elems.len() as u16, line);
            }
            Expression::Call { callee, args } => {
                self.compile_call(callee, args)?;
            }
            Expression::Lambda { params, return_type, body } => {
                // Anonymous procedure/function → compile as closure
                let arity: u8 = params.iter().map(|p| p.names.len() as u8).sum();
                let has_result = return_type.is_some();
                let func_idx = self.chunks.len();
                let chunk = common::functions::create_function_chunk("<lambda>", arity);
                self.chunks.push(chunk);
                self.scopes.push(Scope::new_function());
                let saved = self.current;
                self.current = func_idx;
                for param in params {
                    for pname in &param.names { self.scope_mut().define(pname); }
                }
                let result_slot = if has_result {
                    let rs = self.scope_mut().define("Result");
                    self.emit(Op::NULL);
                    self.emit_u16(Op::LOCAL_SET, rs);
                    self.emit(Op::DROP);
                    Some(rs)
                } else { None };

                let saved_fn = self.current_func_name.take();
                let saved_rs = self.current_result_slot.take();
                if has_result {
                    self.current_func_name = Some("<lambda>".to_string());
                    self.current_result_slot = result_slot;
                }

                for stmt in body { self.compile_stmt(stmt)?; }

                if let Some(rs) = result_slot {
                    self.emit_u16(Op::LOCAL_GET, rs);
                    self.emit(Op::RETURN);
                } else {
                    let line = self.line;
                    common::functions::emit_function_epilogue(&mut self.chunks[func_idx], line);
                }

                self.current_func_name = saved_fn;
                self.current_result_slot = saved_rs;

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
            }
            Expression::IsCheck { expr, type_name } => {
                // obj is TClassName → check __type field
                // For child classes, emit_instanceof_chain re-stamps __type
                self.compile_expr(expr)?;
                let type_key = self.str_const("__type");
                self.emit_u16(Op::STRUCT_GET, type_key);
                self.emit_const(Value::String(Arc::from(type_name.as_str())));
                self.emit(Op::DYN_EQ);
            }
            Expression::AsCast { expr, type_name: _ } => {
                // obj as TClassName → runtime: just return the object (no type erasure in our VM)
                self.compile_expr(expr)?;
            }
            Expression::Inherited { method, args } => {
                // inherited Create(args) → call parent constructor
                if let Some(method_name) = method {
                    if let Some(ref class_name) = self.current_class.clone() {
                        let parent = self.classes.get(class_name).and_then(|c| c.parent.clone());
                        if let Some(parent_name) = parent {
                            let parent_idx = self.str_const(&parent_name);
                            self.emit_u16(Op::GLOBAL_GET, parent_idx);
                            for a in args { self.compile_expr(a)?; }
                            self.emit_u8(Op::CALL_REF, args.len() as u8);
                            if method_name.to_lowercase() == "create" {
                                // Store result in Self slot
                                self.emit(Op::DUP);
                                if let Some(slot) = self.scope().resolve("Self") {
                                    self.emit_u16(Op::LOCAL_SET, slot);
                                    self.emit(Op::DROP);
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
        let _line = self.line;
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
                        self.emit_u16(Op::GLOBAL_GET, idx);
                        for a in args { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL_REF, args.len() as u8);
                    } else {
                        // Static method: global_get(class) + struct_get(method) + args + call_ref
                        let idx = self.str_const(class_name);
                        self.emit_u16(Op::GLOBAL_GET, idx);
                        let prop_idx = self.str_const(field);
                        self.emit_u16(Op::STRUCT_GET, prop_idx);
                        for a in args { self.compile_expr(a)?; }
                        self.emit_u8(Op::CALL_REF, args.len() as u8);
                    }
                    return Ok(());
                }
            }
        }

        // Method call: obj.method(args) — get method from obj, pass obj as self
        if let Expression::Field { record, field } = callee {
            self.compile_expr(record)?;        // push obj
            let prop_idx = self.str_const(field);
            self.emit(Op::DUP);                // [obj, obj]
            self.emit_u16(Op::STRUCT_GET, prop_idx); // [obj, method_ref]
            // Swap so method_ref is below obj (call_ref expects [fn, self, args...])
            // Actually call_ref expects fn on stack first, then args.
            // But our stack is [obj, method_ref]. We need [method_ref, obj, args].
            // Use a temp local to reorder.
            let fn_tmp = self.scope_mut().define("__fn_tmp");
            self.emit_u16(Op::LOCAL_SET, fn_tmp);
            self.emit(Op::DROP);
            // Stack: [obj]. Push fn, then obj, then args.
            let obj_tmp = self.scope_mut().define("__obj_tmp");
            self.emit_u16(Op::LOCAL_SET, obj_tmp);
            self.emit(Op::DROP);
            self.emit_u16(Op::LOCAL_GET, fn_tmp);   // [fn]
            self.emit_u16(Op::LOCAL_GET, obj_tmp);   // [fn, obj]
            for a in args { self.compile_expr(a)?; }  // [fn, obj, args...]
            self.emit_u8(Op::CALL_REF, (args.len() + 1) as u8);
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
                self.emit_u8(Op::CALL_REF, args.len() as u8);
            } else {
                // Unresolved — emit call_import for cross-language interop
                for a in args { self.compile_expr(a)?; }
                let import_idx = self.import("*", name);
                self.emit_host_call(import_idx, args.len() as u8);
            }
        } else {
            self.compile_expr(callee)?;
            for a in args { self.compile_expr(a)?; }
            self.emit_u8(Op::CALL_REF, args.len() as u8);
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
                    self.emit(Op::DROP);
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
                self.emit(Op::DUP);
                self.emit(Op::F64_MUL);
                Ok(true)
            }
            "power" => { self.compile_expr(&args[0])?; self.compile_expr(&args[1])?; let i = self.import("vybe:math","pow"); self.emit_host_call(i, 2); Ok(true) }
            "random" => { let i = self.import("vybe:math","random"); self.emit_host_call(i, 0); Ok(true) }
            "randomize" => { self.emit(Op::NULL); Ok(true) }
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
            "succ" => { self.compile_expr(&args[0])?; self.emit_const(Value::F64(1.0)); self.emit(Op::DYN_ADD); Ok(true) }
            "pred" => { self.compile_expr(&args[0])?; self.emit_const(Value::F64(1.0)); self.emit(Op::F64_SUB); Ok(true) }
            "inc" => {
                if let Some(Expression::Identifier(var)) = args.first() {
                    let var = var.clone();
                    self.emit_var_get(&var);
                    if args.len() > 1 { self.compile_expr(&args[1])?; } else { self.emit_const(Value::F64(1.0)); }
                    self.emit(Op::DYN_ADD);
                    self.emit_var_set(&var);
                }
                self.emit(Op::NULL);
                Ok(true)
            }
            "dec" => {
                if let Some(Expression::Identifier(var)) = args.first() {
                    let var = var.clone();
                    self.emit_var_get(&var);
                    if args.len() > 1 { self.compile_expr(&args[1])?; } else { self.emit_const(Value::F64(1.0)); }
                    self.emit(Op::F64_SUB);
                    self.emit_var_set(&var);
                }
                self.emit(Op::NULL);
                Ok(true)
            }
            "assigned" => { self.compile_expr(&args[0])?; self.emit(Op::REF_IS_NULL); self.emit(Op::DYN_NOT); Ok(true) }
            "high" => {
                self.compile_expr(&args[0])?;
                common::strings::emit_length(self.chunk(), line);
                self.emit_const(Value::F64(1.0));
                self.emit(Op::F64_SUB);
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
                self.emit(Op::NULL);
                Ok(true)
            }
            // ── String manipulation ────────────────────────────────────
            "format" => {
                // Format('Hello %s, you are %d', [name, age])
                // Simplified: just concat the format string and args via host
                for a in args { self.compile_expr(a)?; }
                let idx = self.import("vybe:string", "format");
                self.emit_host_call(idx, args.len() as u8);
                Ok(true)
            }
            "delete" => {
                // Delete(s, index, count) — mutates string variable
                if args.len() >= 3 {
                    if let Expression::Identifier(var) = &args[0] {
                        let var = var.clone();
                        self.emit_var_get(&var);
                        self.compile_expr(&args[1])?;
                        self.compile_expr(&args[2])?;
                        let idx = self.import("vybe:string", "deleteStr");
                        self.emit_host_call(idx, 3);
                        self.emit_var_set(&var);
                    }
                }
                self.emit(Op::NULL);
                Ok(true)
            }
            "insert" => {
                // Insert(source, s, index) — mutates s
                if args.len() >= 3 {
                    if let Expression::Identifier(var) = &args[1] {
                        let var = var.clone();
                        self.compile_expr(&args[0])?; // source
                        self.emit_var_get(&var);       // s
                        self.compile_expr(&args[2])?; // index
                        let idx = self.import("vybe:string", "insertStr");
                        self.emit_host_call(idx, 3);
                        self.emit_var_set(&var);
                    }
                }
                self.emit(Op::NULL);
                Ok(true)
            }
            "stringreplace" => {
                // StringReplace(s, old, new, [flags]) → new string
                if args.len() >= 3 {
                    self.compile_expr(&args[0])?;
                    self.compile_expr(&args[1])?;
                    self.compile_expr(&args[2])?;
                    common::strings::emit_replace(self.chunk(), line);
                } else { self.emit(Op::NULL); }
                Ok(true)
            }
            "stringofchar" => {
                // StringOfChar(ch, count) → string
                self.compile_expr(&args[0])?;
                self.compile_expr(&args[1])?;
                common::strings::emit_repeat(self.chunk(), line);
                Ok(true)
            }
            "comparestr" | "comparetext" => {
                self.compile_expr(&args[0])?;
                self.compile_expr(&args[1])?;
                // Simple: return 0 if equal, -1 if a < b, 1 if a > b
                let idx = self.import("vybe:string", "compare");
                self.emit_host_call(idx, 2);
                Ok(true)
            }
            "leftstr" => {
                // LeftStr(s, count) → s[0..count]
                self.compile_expr(&args[0])?;
                self.emit_const(Value::F64(0.0)); // start = 0
                self.compile_expr(&args[1])?;     // length
                common::strings::emit_substring(self.chunk(), line);
                Ok(true)
            }
            "rightstr" => {
                // RightStr(s, count) → s[len-count..]
                self.compile_expr(&args[0])?;
                self.emit(Op::DUP);
                common::strings::emit_length(self.chunk(), line);
                self.compile_expr(&args[1])?;
                self.emit(Op::F64_SUB); // start = len - count
                self.compile_expr(&args[1])?; // length = count
                common::strings::emit_substring(self.chunk(), line);
                Ok(true)
            }
            // ── Object lifecycle ──────────────────────────────────────
            "free" | "freeandnil" => {
                // In our GC'd VM, Free is a no-op
                if let Some(Expression::Identifier(var)) = args.first() {
                    if name.to_lowercase() == "freeandnil" {
                        self.emit(Op::NULL);
                        let var = var.clone();
                        self.emit_var_set(&var);
                    }
                }
                self.emit(Op::NULL);
                Ok(true)
            }
            // ── Type info ─────────────────────────────────────────────
            "classname" | "typename" => {
                if !args.is_empty() {
                    self.compile_expr(&args[0])?;
                    let type_key = self.str_const("__type");
                    self.emit_u16(Op::STRUCT_GET, type_key);
                } else { self.emit(Op::NULL); }
                Ok(true)
            }
            // ── Array operations ──────────────────────────────────────
            "append" => {
                // Append to dynamic array: SetLength + assign last
                if args.len() >= 2 {
                    self.compile_expr(&args[0])?;
                    self.compile_expr(&args[1])?;
                    let idx = self.import("vybe:array", "push");
                    self.emit_host_call(idx, 2);
                } else { self.emit(Op::NULL); }
                Ok(true)
            }
            "sort" => {
                if !args.is_empty() {
                    self.compile_expr(&args[0])?;
                    let idx = self.import("vybe:array", "sort");
                    self.emit_host_call(idx, 1);
                } else { self.emit(Op::NULL); }
                Ok(true)
            }
            "reverse" => {
                if !args.is_empty() {
                    self.compile_expr(&args[0])?;
                    let idx = self.import("vybe:array", "reverse");
                    self.emit_host_call(idx, 1);
                } else { self.emit(Op::NULL); }
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
        let mut class_vars = Vec::new();

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
                ClassMember::ClassVar(v) => {
                    for cname in &v.names { class_vars.push(cname.clone()); }
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

        // Initialize class variables as globals
        let _line = self.line;
        for cv in &class_vars {
            let global_name = format!("{}.{}", name, cv);
            self.emit(Op::NULL);
            let idx = self.str_const(&global_name);
            self.emit_u16(Op::GLOBAL_SET, idx);
            self.emit(Op::DROP);
        }

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
                self.emit_u16(Op::LOCAL_GET, slot);
                self.emit(Op::RETURN);
            } else {
                let line = self.line;
                common::functions::emit_function_epilogue(&mut self.chunks[func_idx], line);
            }
        } else if is_function {
            // Function method: has Result slot
            let result_slot = self.scope_mut().define("Result");
            self.emit(Op::NULL);
            self.emit_u16(Op::LOCAL_SET, result_slot);
            self.emit(Op::DROP);

            let saved_fn = self.current_func_name.take();
            let saved_rs = self.current_result_slot.take();
            self.current_func_name = Some(method_name.clone());
            self.current_result_slot = Some(result_slot);

            for decl in &method.decls { self.compile_decl(decl)?; }
            for stmt in &method.body { self.compile_stmt(stmt)?; }

            self.emit_u16(Op::LOCAL_GET, result_slot);
            self.emit(Op::RETURN);

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
                    self.chunks[wrapper_idx].emit_op(Op::NULL, line);
                    self.chunks[wrapper_idx].emit_op_u16(Op::LOCAL_SET, this_slot, line);
                    // ref_func for init method
                    common::functions::emit_ref_func(&mut self.chunks[wrapper_idx], init_ci, 0, line);
                    self.chunks[wrapper_idx].emit_op_u16(Op::LOCAL_GET, this_slot, line);
                    for i in 0..user_arity {
                        self.chunks[wrapper_idx].emit_op_u16(Op::LOCAL_GET, (i as u16) + 1, line);
                    }
                    self.chunks[wrapper_idx].emit_op_u8(Op::CALL_REF, (user_arity + 1) as u8, line);
                    // init returns the object — store as this
                    self.chunks[wrapper_idx].emit_op_u16(Op::LOCAL_SET, this_slot, line);
                }

                // Bind child methods (override parent's)
                for (mname, mci) in &methods {
                    common::classes::emit_bind_method_with_aliases(
                        &mut self.chunks[wrapper_idx], this_slot, mname, *mci, line,
                    );
                }

                common::classes::emit_instanceof_chain(&mut self.chunks[wrapper_idx], this_slot, class_name, line);
                // Re-stamp __type for child class (parent constructor set it to parent name)
                self.chunks[wrapper_idx].emit_op_u16(Op::LOCAL_GET, this_slot, line);
                let type_str = self.chunks[wrapper_idx].add_constant(Value::String(Arc::from(class_name.as_str())));
                let type_key = self.chunks[wrapper_idx].add_constant(Value::String(Arc::from("__type")));
                self.chunks[wrapper_idx].emit_op_u16(Op::CONST, type_str, line);
                self.chunks[wrapper_idx].emit_op_u16(Op::STRUCT_SET, type_key, line);
                self.chunks[wrapper_idx].emit_op(Op::DROP, line);
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
                    self.chunks[wrapper_idx].emit_op_u16(Op::LOCAL_GET, this_slot, line);
                    for i in 0..user_arity {
                        self.chunks[wrapper_idx].emit_op_u16(Op::LOCAL_GET, (i as u16) + 1, line);
                    }
                    self.chunks[wrapper_idx].emit_op_u8(Op::CALL_REF, (user_arity + 1) as u8, line);
                    self.chunks[wrapper_idx].emit_op(Op::DROP, line);
                }

                common::classes::emit_instanceof_chain(&mut self.chunks[wrapper_idx], this_slot, class_name, line);
            }

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
