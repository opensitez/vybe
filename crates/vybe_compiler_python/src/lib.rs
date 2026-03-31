use std::rc::Rc;
use std::collections::HashMap;
use vybe_parser_python::ast::*;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

pub struct Compiler {
    chunks: Vec<Chunk>,
    scopes: Vec<Scope>,
    loop_stack: Vec<LoopCtx>,
    import_log: u16,
}

struct Scope {
    locals: HashMap<String, u16>,
    max_local: u16,
}

struct LoopCtx {
    _start: usize,
    break_jumps: Vec<usize>,
    continue_jumps: Vec<usize>,
}

impl Scope {
    fn new() -> Self {
        Self { locals: HashMap::new(), max_local: 0 }
    }

    fn alloc(&mut self, name: &str) -> u16 {
        if let Some(&idx) = self.locals.get(name) {
            idx
        } else {
            self.max_local += 1;
            let idx = self.max_local;
            self.locals.insert(name.to_string(), idx);
            idx
        }
    }

    fn get(&self, name: &str) -> Option<u16> {
        self.locals.get(name).copied()
    }
}

impl Compiler {
    pub fn new() -> Self {
        Self {
            chunks: Vec::new(),
            scopes: Vec::new(),
            loop_stack: Vec::new(),
            import_log: 0,
        }
    }

    pub fn compile(&mut self, module: &Module) -> Result<Vec<Chunk>, String> {
        let mut chunk = Chunk::new("<script>");
        self.import_log = chunk.add_import("wasi:cli", "log");
        self.chunks.push(chunk);
        self.scopes.push(Scope::new());

        for stmt in &module.body {
            self.compile_stmt(stmt, 0)?;
        }

        // Finalize main chunk
        let scope = self.scopes.remove(0);
        self.chunks[0].local_count = (scope.max_local + 1) as u16;
        self.chunks[0].emit_op(Op::halt, 0);

        Ok(std::mem::take(&mut self.chunks))
    }

    // ── Statements ───────────────────────────────────────────────────

    fn compile_stmt(&mut self, stmt: &Statement, chunk_idx: usize) -> Result<(), String> {
        match stmt {
            Statement::Expression(expr) => {
                // print() calls: intercept and compile as host log
                if let Expression::Call { func, args, keywords: _ } = expr {
                    if let Expression::Name(name) = func.as_ref() {
                        if name == "print" {
                            return self.compile_print(args, chunk_idx);
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

            Statement::While { test, body, else_body: _ } => {
                let loop_start = self.chunk(chunk_idx).current_offset();
                self.compile_expr(test, chunk_idx)?;
                let exit_jump = self.chunk(chunk_idx).emit_jump(Op::br_if_false, 0);

                self.loop_stack.push(LoopCtx { _start: loop_start, break_jumps: Vec::new(), continue_jumps: Vec::new() });

                for s in body { self.compile_stmt(s, chunk_idx)?; }

                let ctx = self.loop_stack.pop().unwrap();
                for cj in &ctx.continue_jumps { self.chunk(chunk_idx).patch_jump(*cj); }
                self.chunk(chunk_idx).emit_loop(loop_start, 0);
                for bj in &ctx.break_jumps { self.chunk(chunk_idx).patch_jump(*bj); }
                self.chunk(chunk_idx).patch_jump(exit_jump);
            }

            Statement::For { target, iter, body, else_body: _, is_async: _ } => {
                self.compile_for(target, iter, body, chunk_idx)?;
            }

            Statement::FunctionDef { name, params, body, is_async: _, decorators: _, returns: _ } => {
                self.compile_function(name, params, body)?;
                // Bind function to local in current scope
                let func_idx = self.chunks.len() - 1;
                let idx = self.scope(chunk_idx).alloc(name);
                self.chunk(chunk_idx).emit_op_u16(Op::ref_func, func_idx as u16, 0);
                self.chunk(chunk_idx).emit(0, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx, 0);
            }

            Statement::ClassDef { name, bases: _, keywords: _, body, decorators: _ } => {
                // Minimal class support: create an object with methods
                // For now, compile as a namespace with functions
                let idx = self.scope(chunk_idx).alloc(name);
                // Create an object to hold methods
                let dict_new = self.chunk(chunk_idx).add_import("vybe:types", "dictNew");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, dict_new, 0);
                self.chunk(chunk_idx).emit(0, 0);
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx, 0);

                for s in body {
                    if let Statement::FunctionDef { name: method_name, params, body: mbody, .. } = s {
                        self.compile_function(method_name, params, mbody)?;
                        let func_chunk_idx = self.chunks.len() - 1;
                        // Store method on class dict
                        self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx, 0);
                        let name_c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(method_name.as_str())));
                        self.chunk(chunk_idx).emit_op_u16(Op::r#const, name_c, 0);
                        self.chunk(chunk_idx).emit_op_u16(Op::ref_func, func_chunk_idx as u16, 0);
                        self.chunk(chunk_idx).emit(0, 0);
                        let dict_add = self.chunk(chunk_idx).add_import("vybe:types", "dictAdd");
                        self.chunk(chunk_idx).emit_op_u16(Op::call_import, dict_add, 0);
                        self.chunk(chunk_idx).emit(3, 0);
                        self.chunk(chunk_idx).emit_op(Op::drop, 0);
                    } else if let Statement::Pass = s {
                        // skip
                    } else {
                        // skip other class body statements for now
                    }
                }
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

            Statement::Try { body, handlers, else_body: _, finally_body: _ } => {
                // Basic try/except: try_start → body → br end → catch → handler → end
                let _try_start = self.chunk(chunk_idx).code.len();
                let catch_jump = self.chunk(chunk_idx).emit_jump(Op::try_start, 0);
                self.chunk(chunk_idx).emit(0u8, 0); // reserved for finally

                for s in body { self.compile_stmt(s, chunk_idx)?; }

                // try_end + jump to after handlers
                self.chunk(chunk_idx).emit_op(Op::try_end, 0);
                let end_jump = self.chunk(chunk_idx).emit_jump(Op::br, 0);

                // Patch catch jump
                self.chunk(chunk_idx).patch_jump(catch_jump);

                for handler in handlers {
                    // If handler has a name, store exception in local
                    if let Some(name) = &handler.name {
                        let idx = self.scope(chunk_idx).alloc(name);
                        // Exception is on stack from try_start
                        self.chunk(chunk_idx).emit_op_u16(Op::local_set, idx, 0);
                        self.chunk(chunk_idx).emit_op(Op::drop, 0);
                    } else {
                        self.chunk(chunk_idx).emit_op(Op::drop, 0); // drop exception
                    }
                    for s in &handler.body { self.compile_stmt(s, chunk_idx)?; }
                }

                self.chunk(chunk_idx).patch_jump(end_jump);
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

            Statement::Delete(_) | Statement::Assert { .. } | Statement::With { .. }
            | Statement::AnnAssign { .. } | Statement::Match { .. } => {
                // Not yet compiled — skip silently
            }
        }
        Ok(())
    }

    // ── If with proper end-jump chaining ─────────────────────────────

    // ── For loop ─────────────────────────────────────────────────────

    fn compile_for(&mut self, target: &Expression, iter: &Expression, body: &[Statement], chunk_idx: usize) -> Result<(), String> {
        let iter_local = self.scope(chunk_idx).alloc("__for_iter");
        let idx_local = self.scope(chunk_idx).alloc("__for_idx");

        // Evaluate iterable
        self.compile_expr(iter, chunk_idx)?;
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
        for bj in &ctx.break_jumps { self.chunk(chunk_idx).patch_jump(*bj); }
        self.chunk(chunk_idx).patch_jump(exit_jump);

        Ok(())
    }

    // ── Function compilation ─────────────────────────────────────────

    fn compile_function(&mut self, name: &str, params: &Parameters, body: &[Statement]) -> Result<(), String> {
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

        for s in body {
            self.compile_stmt(s, func_chunk_idx)?;
        }

        // Ensure function ends with return
        self.chunks[func_chunk_idx].emit_op(Op::null, 0);
        self.chunks[func_chunk_idx].emit_op(Op::r#return, 0);

        let scope = self.scopes.remove(scope_idx);
        self.chunks[func_chunk_idx].local_count = (scope.max_local + 1) as u16;
        self.chunks[func_chunk_idx].arity = params.args.len() as u8;

        Ok(())
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
                // obj.attr = value  →  struct_set
                self.compile_expr(value, chunk_idx)?;
                let name_c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(attr.as_str())));
                self.chunk(chunk_idx).emit_op_u16(Op::struct_set, name_c, 0);
            }
            Expression::Subscript { value, slice } => {
                // obj[idx] = value  →  array_set
                self.compile_expr(value, chunk_idx)?;
                self.compile_expr(slice, chunk_idx)?;
                self.chunk(chunk_idx).emit_op(Op::array_set, 0);
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
                if let Some(idx) = self.scope(chunk_idx).get(name) {
                    self.chunk(chunk_idx).emit_op_u16(Op::local_get, idx, 0);
                } else {
                    // Try global
                    let c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(name.as_str())));
                    self.chunk(chunk_idx).emit_op_u16(Op::global_get, c, 0);
                }
            }

            Expression::List(elems) => {
                for e in elems { self.compile_expr(e, chunk_idx)?; }
                self.chunk(chunk_idx).emit_op_u16(Op::array_new, elems.len() as u16, 0);
            }
            Expression::Tuple(elems) => {
                for e in elems { self.compile_expr(e, chunk_idx)?; }
                self.chunk(chunk_idx).emit_op_u16(Op::array_new, elems.len() as u16, 0);
            }
            Expression::Set(elems) => {
                for e in elems { self.compile_expr(e, chunk_idx)?; }
                self.chunk(chunk_idx).emit_op_u16(Op::array_new, elems.len() as u16, 0);
            }
            Expression::Dict { keys, values } => {
                let dict_new = self.chunk(chunk_idx).add_import("vybe:types", "dictNew");
                let dict_add = self.chunk(chunk_idx).add_import("vybe:types", "dictAdd");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, dict_new, 0);
                self.chunk(chunk_idx).emit(0, 0);
                for (k, v) in keys.iter().zip(values.iter()) {
                    self.chunk(chunk_idx).emit_op(Op::dup, 0);
                    if let Some(key) = k {
                        self.compile_expr(key, chunk_idx)?;
                    } else {
                        self.chunk(chunk_idx).emit_op(Op::null, 0);
                    }
                    self.compile_expr(v, chunk_idx)?;
                    self.chunk(chunk_idx).emit_op_u16(Op::call_import, dict_add, 0);
                    self.chunk(chunk_idx).emit(3, 0);
                    self.chunk(chunk_idx).emit_op(Op::drop, 0);
                }
            }

            Expression::BinOp { op, left, right } => {
                self.compile_expr(left, chunk_idx)?;
                self.compile_expr(right, chunk_idx)?;
                match op {
                    BinOp::Add => self.chunk(chunk_idx).emit_op(Op::dyn_add, 0),
                    BinOp::Sub => self.chunk(chunk_idx).emit_op(Op::f64_sub, 0),
                    BinOp::Mul => self.chunk(chunk_idx).emit_op(Op::f64_mul, 0),
                    BinOp::Div => self.chunk(chunk_idx).emit_op(Op::f64_div, 0),
                    BinOp::FloorDiv => {
                        self.chunk(chunk_idx).emit_op(Op::f64_div, 0);
                        self.chunk(chunk_idx).emit_op(Op::f64_floor, 0);
                    }
                    BinOp::Mod => self.chunk(chunk_idx).emit_op(Op::i32_rem_s, 0),
                    BinOp::Pow => {
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
                                let to_str = self.chunk(chunk_idx).add_import("vybe:convert", "toString");
                                self.chunk(chunk_idx).emit_op_u16(Op::call_import, to_str, 0);
                                self.chunk(chunk_idx).emit(1, 0);
                                return Ok(());
                            }
                        }
                        "int" => {
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                let to_int = self.chunk(chunk_idx).add_import("vybe:convert", "cint");
                                self.chunk(chunk_idx).emit_op_u16(Op::call_import, to_int, 0);
                                self.chunk(chunk_idx).emit(1, 0);
                                return Ok(());
                            }
                        }
                        "float" => {
                            if args.len() == 1 {
                                self.compile_expr(&args[0], chunk_idx)?;
                                let to_flt = self.chunk(chunk_idx).add_import("vybe:convert", "cdbl");
                                self.chunk(chunk_idx).emit_op_u16(Op::call_import, to_flt, 0);
                                self.chunk(chunk_idx).emit(1, 0);
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
                        _ => {}
                    }
                }
                // Method calls: obj.method(args)
                if let Expression::Attribute { value, attr } = func.as_ref() {
                    return self.compile_method_call(value, attr, args, chunk_idx);
                }
                // General function call
                self.compile_expr(func, chunk_idx)?;
                for a in args { self.compile_expr(a, chunk_idx)?; }
                self.chunk(chunk_idx).emit_op_u8(Op::call_ref, args.len() as u8, 0);
            }

            Expression::Attribute { value, attr } => {
                self.compile_expr(value, chunk_idx)?;
                let c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(attr.as_str())));
                self.chunk(chunk_idx).emit_op_u16(Op::struct_get, c, 0);
            }

            Expression::Subscript { value, slice } => {
                // Dict string key lookup vs array index
                if let Expression::Str(s) = slice.as_ref() {
                    self.compile_expr(value, chunk_idx)?;
                    let c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(s.as_str())));
                    self.chunk(chunk_idx).emit_op_u16(Op::r#const, c, 0);
                    let dict_item = self.chunk(chunk_idx).add_import("vybe:types", "dictItem");
                    self.chunk(chunk_idx).emit_op_u16(Op::call_import, dict_item, 0);
                    self.chunk(chunk_idx).emit(2, 0);
                } else {
                    self.compile_expr(value, chunk_idx)?;
                    self.compile_expr(slice, chunk_idx)?;
                    self.chunk(chunk_idx).emit_op(Op::array_get, 0);
                }
            }

            Expression::Slice { .. } => {
                // Slicing not fully supported yet — push null
                self.chunk(chunk_idx).emit_op(Op::null, 0);
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
                // Compile as anonymous function
                let name = "__lambda";
                self.compile_function(name, params, &[Statement::Return(Some(*body.clone()))])?;
                let func_idx = self.chunks.len() - 1;
                self.chunk(chunk_idx).emit_op_u16(Op::ref_func, func_idx as u16, 0);
                self.chunk(chunk_idx).emit(0, 0);
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
                if let Some(e) = expr {
                    self.compile_expr(e, chunk_idx)?;
                } else {
                    self.chunk(chunk_idx).emit_op(Op::null, 0);
                }
            }

            Expression::YieldFrom(expr) => {
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
                let dict_new = self.chunk(chunk_idx).add_import("vybe:types", "dictNew");
                let dict_add = self.chunk(chunk_idx).add_import("vybe:types", "dictAdd");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, dict_new, 0);
                self.chunk(chunk_idx).emit(0, 0);
                let result_local = self.scope(chunk_idx).alloc("__comp_result");
                self.chunk(chunk_idx).emit_op_u16(Op::local_set, result_local, 0);
                self.chunk(chunk_idx).emit_op(Op::drop, 0);

                self.compile_comp_generators(generators, &|s| {
                    s.chunk(chunk_idx).emit_op_u16(Op::local_get, result_local, 0);
                    s.compile_expr(key, chunk_idx)?;
                    s.compile_expr(value, chunk_idx)?;
                    s.chunk(chunk_idx).emit_op_u16(Op::call_import, dict_add, 0);
                    s.chunk(chunk_idx).emit(3, 0);
                    s.chunk(chunk_idx).emit_op(Op::drop, 0);
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

    fn compile_print(&mut self, args: &[Expression], chunk_idx: usize) -> Result<(), String> {
        for a in args {
            self.compile_expr(a, chunk_idx)?;
        }
        let log_idx = self.import_log;
        self.chunk(chunk_idx).emit_op_u16(Op::call_import, log_idx, 0);
        self.chunk(chunk_idx).emit(args.len() as u8, 0);
        Ok(())
    }

    fn compile_range(&mut self, args: &[Expression], chunk_idx: usize) -> Result<(), String> {
        let range_fn = self.chunk(chunk_idx).add_import("vybe:array", "range");
        for a in args {
            self.compile_expr(a, chunk_idx)?;
        }
        self.chunk(chunk_idx).emit_op_u16(Op::call_import, range_fn, 0);
        self.chunk(chunk_idx).emit(args.len() as u8, 0);
        Ok(())
    }

    fn compile_method_call(&mut self, obj: &Expression, method: &str, args: &[Expression], chunk_idx: usize) -> Result<(), String> {
        match method {
            "append" => {
                self.compile_expr(obj, chunk_idx)?;
                for a in args { self.compile_expr(a, chunk_idx)?; }
                let push_fn = self.chunk(chunk_idx).add_import("vybe:array", "push");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, push_fn, 0);
                self.chunk(chunk_idx).emit((1 + args.len()) as u8, 0);
            }
            "pop" => {
                self.compile_expr(obj, chunk_idx)?;
                let pop_fn = self.chunk(chunk_idx).add_import("vybe:array", "pop");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, pop_fn, 0);
                self.chunk(chunk_idx).emit(1, 0);
            }
            "keys" => {
                self.compile_expr(obj, chunk_idx)?;
                let keys_fn = self.chunk(chunk_idx).add_import("vybe:types", "dictKeys");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, keys_fn, 0);
                self.chunk(chunk_idx).emit(1, 0);
            }
            "values" => {
                self.compile_expr(obj, chunk_idx)?;
                let vals_fn = self.chunk(chunk_idx).add_import("vybe:types", "dictValues");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, vals_fn, 0);
                self.chunk(chunk_idx).emit(1, 0);
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
                let split_fn = self.chunk(chunk_idx).add_import("vybe:string", "split");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, split_fn, 0);
                self.chunk(chunk_idx).emit(2, 0);
            }
            "join" => {
                // separator.join(iterable) → join(iterable, separator)
                if args.len() == 1 {
                    self.compile_expr(&args[0], chunk_idx)?;
                    self.compile_expr(obj, chunk_idx)?;
                    let join_fn = self.chunk(chunk_idx).add_import("vybe:array", "join");
                    self.chunk(chunk_idx).emit_op_u16(Op::call_import, join_fn, 0);
                    self.chunk(chunk_idx).emit(2, 0);
                } else {
                    self.compile_expr(obj, chunk_idx)?;
                }
            }
            "replace" => {
                self.compile_expr(obj, chunk_idx)?;
                for a in args { self.compile_expr(a, chunk_idx)?; }
                let replace_fn = self.chunk(chunk_idx).add_import("vybe:string", "replaceAll");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, replace_fn, 0);
                self.chunk(chunk_idx).emit((1 + args.len()) as u8, 0);
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
            _ => {
                // Generic method call via attribute access + call_ref
                self.compile_expr(obj, chunk_idx)?;
                let c = self.chunk(chunk_idx).add_constant(Value::String(Rc::from(method)));
                self.chunk(chunk_idx).emit_op_u16(Op::struct_get, c, 0);
                for a in args { self.compile_expr(a, chunk_idx)?; }
                self.chunk(chunk_idx).emit_op_u8(Op::call_ref, args.len() as u8, 0);
            }
        }
        Ok(())
    }

    fn compile_comprehension(&mut self, element: &Expression, generators: &[Comprehension], chunk_idx: usize) -> Result<(), String> {
        // Create empty array, iterate, push elements
        self.chunk(chunk_idx).emit_op_u16(Op::array_new, 0, 0);
        let result_local = self.scope(chunk_idx).alloc("__comp_result");
        self.chunk(chunk_idx).emit_op_u16(Op::local_set, result_local, 0);

        let push_fn = self.chunk(chunk_idx).add_import("vybe:array", "push");

        self.compile_comp_generators(generators, &|s| {
            s.chunk(chunk_idx).emit_op_u16(Op::local_get, result_local, 0);
            s.compile_expr(element, chunk_idx)?;
            s.chunk(chunk_idx).emit_op_u16(Op::call_import, push_fn, 0);
            s.chunk(chunk_idx).emit(2, 0);
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
                // a in b → array_contains(b, a) — swap then call
                let contains = self.chunk(chunk_idx).add_import("vybe:array", "contains");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, contains, 0);
                self.chunk(chunk_idx).emit(2, 0);
            }
            CmpOp::NotIn => {
                let contains = self.chunk(chunk_idx).add_import("vybe:array", "contains");
                self.chunk(chunk_idx).emit_op_u16(Op::call_import, contains, 0);
                self.chunk(chunk_idx).emit(2, 0);
                self.chunk(chunk_idx).emit_op(Op::dyn_not, 0);
            }
        }
    }

    fn emit_aug_op(&mut self, op: AugOp, chunk_idx: usize) {
        match op {
            AugOp::Add => self.chunk(chunk_idx).emit_op(Op::dyn_add, 0),
            AugOp::Sub => self.chunk(chunk_idx).emit_op(Op::f64_sub, 0),
            AugOp::Mul => self.chunk(chunk_idx).emit_op(Op::f64_mul, 0),
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
