use std::rc::Rc;

use vybe_bytecode::{Chunk, Value, Op};
use vybe_parser_js::ast::*;

use crate::scope::Scope;

struct LoopContext {
    start_offset: usize,
    break_patches: Vec<usize>,
    continue_patches: Vec<usize>,
}

pub struct Compiler {
    chunks: Vec<Chunk>,
    scopes: Vec<Scope>,
    current_chunk_idx: usize,
    loop_stack: Vec<LoopContext>,
    line: u32,
    in_method: bool,
}

impl Compiler {
    pub fn new() -> Self {
        Compiler {
            chunks: vec![Chunk::new("<script>")],
            scopes: vec![Scope::new()],
            current_chunk_idx: 0,
            loop_stack: Vec::new(),
            line: 1,
            in_method: false,
        }
    }

    pub fn compile(mut self, program: &Program) -> Result<Vec<Chunk>, String> {
        for stmt in &program.body {
            self.compile_statement(stmt)?;
        }
        self.emit(Op::null);
        self.emit(Op::halt);
        let local_count = self.current_scope().next_slot;
        self.chunks[0].local_count = local_count;
        Ok(self.chunks)
    }

    // -- Import helper: adds to chunk 0's import table, returns import index --

    fn import(&mut self, module: &str, name: &str) -> u16 {
        self.chunks[0].add_import(module, name)
    }

    // -- Emit helpers --

    fn emit(&mut self, op: Op) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op(op, line);
    }
    fn emit_u16(&mut self, op: Op, operand: u16) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u16(op, operand, line);
    }
    fn emit_u8(&mut self, op: Op, operand: u8) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u8(op, operand, line);
    }
    fn emit_constant(&mut self, value: Value) {
        let idx = self.chunks[self.current_chunk_idx].add_constant(value);
        self.emit_u16(Op::r#const, idx);
    }
    fn emit_jump(&mut self, op: Op) -> usize {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_jump(op, line)
    }
    fn patch_jump(&mut self, offset: usize) {
        self.chunks[self.current_chunk_idx].patch_jump(offset);
    }
    fn current_offset(&self) -> usize {
        self.chunks[self.current_chunk_idx].current_offset()
    }
    fn emit_loop(&mut self, target: usize) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_loop(target, line);
    }
    fn add_string_constant(&mut self, s: &str) -> u16 {
        self.chunks[self.current_chunk_idx].add_constant(Value::String(Rc::from(s)))
    }

    /// Emit CallHost: [CallHost, import_hi, import_lo, argc]
    fn emit_host_call(&mut self, import_idx: u16, argc: u8) {
        let line = self.line;
        let c = &mut self.chunks[self.current_chunk_idx];
        c.emit_op(Op::call_import, line);
        c.emit((import_idx >> 8) as u8, line);
        c.emit((import_idx & 0xff) as u8, line);
        c.emit(argc, line);
    }

    /// Emit truthy check: value → Bool (inline VM op, no host call)
    fn emit_to_bool(&mut self) {
        self.emit(Op::dyn_to_bool);
    }

    /// Emit JS add: string concat or numeric add (inline VM op, no host call)
    fn emit_js_add(&mut self) {
        self.emit(Op::dyn_add);
    }

    // -- Scope helpers --

    fn current_scope(&self) -> &Scope { self.scopes.last().unwrap() }
    fn current_scope_mut(&mut self) -> &mut Scope { self.scopes.last_mut().unwrap() }
    fn define_local(&mut self, name: &str) -> u16 { self.current_scope_mut().define_local(name) }

    fn resolve_variable(&mut self, name: &str) -> VarResolution {
        if let Some(slot) = self.current_scope().resolve_local(name) {
            return VarResolution::Local(slot);
        }
        if self.scopes.len() > 1 {
            if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, name) {
                return VarResolution::Upvalue(uv);
            }
        }
        VarResolution::Global
    }

    fn resolve_upvalue(&mut self, scope_idx: usize, name: &str) -> Option<u8> {
        if scope_idx == 0 { return None; }
        let parent = scope_idx - 1;
        if let Some(slot) = self.scopes[parent].resolve_local(name) {
            self.scopes[parent].mark_captured(slot);
            return Some(self.scopes[scope_idx].add_upvalue(slot as u8, true));
        }
        if let Some(uv) = self.resolve_upvalue(parent, name) {
            return Some(self.scopes[scope_idx].add_upvalue(uv, false));
        }
        None
    }

    // -- Resolve host calls: maps JS syntax → VSI (module, name) --

    /// Check if an identifier is a local, upvalue, or declared global variable.
    fn is_known_variable(&self, name: &str) -> bool {
        if self.current_scope().resolve_local(name).is_some() {
            return true;
        }
        // Check upvalue scopes
        for scope in self.scopes.iter().rev().skip(1) {
            if scope.resolve_local(name).is_some() {
                return true;
            }
        }
        false
    }

    /// JS-to-module alias table. Maps JS namespace objects to VSI module names.
    /// If a JS identifier isn't a variable and isn't in this table, it's passed
    /// through as-is (e.g. "MyModule" → ("MyModule", "method")).
    fn js_module_alias(obj_name: &str) -> &str {
        match obj_name {
            // JS standard objects → VSI modules
            "console" => "vybe:console",
            "Math"    => "vybe:math",
            "JSON"    => "vybe:json",
            "Date"    => "vybe:clock",
            // Vybe platform modules — map to vybe: prefix
            "fs"      => "vybe:fs",
            "clock"   => "vybe:clock",
            "env"     => "vybe:env",
            "random"  => "vybe:random",
            "http"    => "vybe:http",
            "gui"     => "vybe:gui",
            // Unknown → pass through as-is (user module or future VB namespace)
            _ => obj_name,
        }
    }

    /// JS-specific name remapping. Most methods pass through as-is.
    /// Only remap when JS convention differs from VSI naming.
    fn js_remap<'a>(module: &'a str, method: &'a str) -> (&'a str, &'a str) {
        match (module, method) {
            ("vybe:math", "random") => ("vybe:random", "random"),
            _ => (module, method),
        }
    }

    /// Resolve bare global calls that are host imports, not user functions.
    /// Each language compiler defines its own set of bare imports.
    fn resolve_bare_import(&mut self, name: &str) -> Option<u16> {
        match name {
            // JS standard globals that are host functions
            "parseInt" | "parseFloat" | "isNaN" | "isFinite" => Some(self.import("vybe:convert", name)),
            _ => None,
        }
    }

    /// Resolve method calls on values: str.toUpperCase(), arr.push(), etc.
    /// These are instance methods — the object is passed as the first arg.
    fn resolve_value_method(&mut self, method: &str) -> Option<u16> {
        // Only methods that are called ON a value (not a namespace).
        // The host function receives the object as first arg.
        let (module, name) = match method {
            // String methods
            "toUpperCase" | "toLowerCase" | "trim" | "startsWith" | "endsWith" |
            "charAt" | "substring" | "split" | "replace" => ("vybe:string", method),
            // Array methods
            "push" | "pop" | "shift" | "join" | "reverse" | "concat" => ("vybe:array", method),
            // Shared — host dispatches by type at runtime
            "slice" => ("vybe:array", "slice"),
            "indexOf" => ("vybe:string", "indexOf"),
            "includes" => ("vybe:string", "includes"),
            _ => return None,
        };
        Some(self.import(module, name))
    }

    // -- Statements --

    fn compile_statement(&mut self, stmt: &Statement) -> Result<(), String> {
        match stmt {
            Statement::Expression(expr) => {
                self.compile_expression(expr)?;
                self.emit(Op::drop);
            }
            Statement::Block(stmts) => {
                self.current_scope_mut().begin_scope();
                for s in stmts { self.compile_statement(s)?; }
                self.current_scope_mut().end_scope();
            }
            Statement::VariableDeclaration { kind, declarations } => {
                for decl in declarations {
                    if let Some(init) = &decl.init {
                        self.compile_expression(init)?;
                    } else {
                        self.emit(Op::null);
                    }
                    match self.resolve_variable(&decl.name) {
                        VarResolution::Local(_) => {
                            let slot = self.current_scope().resolve_local(&decl.name).unwrap();
                            self.emit_u16(Op::local_set, slot);
                            self.emit(Op::drop);
                        }
                        _ => {
                            if self.scopes.len() == 1 && self.current_scope().depth == 0 && *kind == VarKind::Var {
                                let idx = self.add_string_constant(&decl.name);
                                self.emit_u16(Op::global_set, idx);
                                self.emit(Op::drop);
                            } else {
                                let slot = self.define_local(&decl.name);
                                self.emit_u16(Op::local_set, slot);
                                self.emit(Op::drop);
                            }
                        }
                    }
                }
            }
            Statement::FunctionDeclaration(func) => {
                self.compile_function(func)?;
                if let Some(name) = &func.name {
                    if self.scopes.len() == 1 && self.current_scope().depth == 0 {
                        let idx = self.add_string_constant(name);
                        self.emit_u16(Op::global_set, idx);
                        self.emit(Op::drop);
                    } else {
                        let slot = self.define_local(name);
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    }
                }
            }
            Statement::If { test, consequent, alternate } => {
                self.compile_expression(test)?;
                self.emit_to_bool();
                let else_j = self.emit_jump(Op::br_if_false);
                self.compile_statement(consequent)?;
                if let Some(alt) = alternate {
                    let end_j = self.emit_jump(Op::br);
                    self.patch_jump(else_j);
                    self.compile_statement(alt)?;
                    self.patch_jump(end_j);
                } else { self.patch_jump(else_j); }
            }
            Statement::While { test, body } => {
                let start = self.current_offset();
                self.loop_stack.push(LoopContext { start_offset: start, break_patches: vec![], continue_patches: vec![] });
                self.compile_expression(test)?;
                self.emit_to_bool();
                let exit = self.emit_jump(Op::br_if_false);
                self.compile_statement(body)?;
                for p in self.loop_stack.last().unwrap().continue_patches.clone() { self.patch_jump(p); }
                self.emit_loop(start);
                self.patch_jump(exit);
                let ctx = self.loop_stack.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
            }
            Statement::DoWhile { body, test } => {
                let start = self.current_offset();
                self.loop_stack.push(LoopContext { start_offset: start, break_patches: vec![], continue_patches: vec![] });
                self.compile_statement(body)?;
                for p in self.loop_stack.last().unwrap().continue_patches.clone() { self.patch_jump(p); }
                self.compile_expression(test)?;
                self.emit_to_bool();
                let exit = self.emit_jump(Op::br_if_false);
                self.emit_loop(start);
                self.patch_jump(exit);
                let ctx = self.loop_stack.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
            }
            Statement::For { init, test, update, body } => {
                self.current_scope_mut().begin_scope();
                if let Some(init) = init {
                    match init {
                        ForInit::VarDecl(kind, decls) => {
                            self.compile_statement(&Statement::VariableDeclaration { kind: *kind, declarations: decls.clone() })?;
                        }
                        ForInit::Expression(expr) => { self.compile_expression(expr)?; self.emit(Op::drop); }
                    }
                }
                let start = self.current_offset();
                self.loop_stack.push(LoopContext { start_offset: start, break_patches: vec![], continue_patches: vec![] });
                let exit = if let Some(test) = test {
                    self.compile_expression(test)?;
                    self.emit_to_bool();
                    Some(self.emit_jump(Op::br_if_false))
                } else { None };
                self.compile_statement(body)?;
                for p in self.loop_stack.last().unwrap().continue_patches.clone() { self.patch_jump(p); }
                if let Some(update) = update { self.compile_expression(update)?; self.emit(Op::drop); }
                self.emit_loop(start);
                if let Some(e) = exit { self.patch_jump(e); }
                let ctx = self.loop_stack.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
                self.current_scope_mut().end_scope();
            }
            Statement::Return(value) => {
                if let Some(expr) = value { self.compile_expression(expr)?; }
                else { self.emit(Op::null); }
                self.emit(Op::r#return);
            }
            Statement::Break(_) => {
                let p = self.emit_jump(Op::br);
                if let Some(ctx) = self.loop_stack.last_mut() { ctx.break_patches.push(p); }
            }
            Statement::Continue(_) => {
                let p = self.emit_jump(Op::br);
                if let Some(ctx) = self.loop_stack.last_mut() { ctx.continue_patches.push(p); }
            }
            Statement::Throw(expr) => { self.compile_expression(expr)?; self.emit(Op::throw); }
            Statement::Try { block, handler, finalizer } => {
                let line = self.line;
                let c = &mut self.chunks[self.current_chunk_idx];
                c.emit_op(Op::try_start, line); c.emit(0, line); c.emit(0, line); c.emit(0, line); c.emit(0, line);
                for s in block { self.compile_statement(s)?; }
                self.emit(Op::try_end);
                if let Some(h) = handler { for s in &h.body { self.compile_statement(s)?; } }
                if let Some(f) = finalizer { for s in f { self.compile_statement(s)?; } }
            }
            Statement::Switch { discriminant, cases } => {
                self.compile_expression(discriminant)?;
                self.loop_stack.push(LoopContext { start_offset: 0, break_patches: vec![], continue_patches: vec![] });
                let mut next: Option<usize> = None;
                for case in cases {
                    if let Some(p) = next.take() { self.patch_jump(p); }
                    if let Some(test) = &case.test {
                        self.emit(Op::dup);
                        self.compile_expression(test)?;
                        self.emit(Op::eq);
                        next = Some(self.emit_jump(Op::br_if_false));
                    }
                    for s in &case.consequent { self.compile_statement(s)?; }
                }
                if let Some(p) = next { self.patch_jump(p); }
                let ctx = self.loop_stack.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
                self.emit(Op::drop);
            }
            Statement::ClassDeclaration(class) => {
                self.compile_class(class)?;
                if let Some(name) = &class.name {
                    if self.scopes.len() == 1 && self.current_scope().depth == 0 {
                        let idx = self.add_string_constant(name);
                        self.emit_u16(Op::global_set, idx);
                        self.emit(Op::drop);
                    } else {
                        let slot = self.define_local(name);
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    }
                }
            }
            Statement::ForIn { .. } | Statement::ForOf { .. } => {
                return Err("for-in/for-of not yet implemented".into());
            }
            Statement::Labeled { body, .. } => { self.compile_statement(body)?; }
            Statement::Empty => {}
        }
        Ok(())
    }

    // -- Expressions --

    fn compile_expression(&mut self, expr: &Expression) -> Result<(), String> {
        match expr {
            Expression::Number(n) => { self.emit_constant(Value::F64(*n)); }
            Expression::String(s) => { self.emit_constant(Value::String(Rc::from(s.as_str()))); }
            Expression::Boolean(true) => self.emit(Op::r#true),
            Expression::Boolean(false) => self.emit(Op::r#false),
            Expression::Null | Expression::Undefined => self.emit(Op::null),
            Expression::This => {
                if self.in_method { self.emit_u16(Op::local_get, 1); }
                else { self.emit(Op::null); }
            }
            Expression::Identifier(name) => {
                match self.resolve_variable(name) {
                    VarResolution::Local(slot) => self.emit_u16(Op::local_get, slot),
                    VarResolution::Upvalue(idx) => self.emit_u8(Op::upvalue_get, idx),
                    VarResolution::Global => {
                        let idx = self.add_string_constant(name);
                        self.emit_u16(Op::global_get, idx);
                    }
                }
            }
            Expression::Binary { op, left, right } => {
                if *op == BinaryOp::NullishCoalescing {
                    self.compile_expression(left)?;
                    self.emit(Op::dup);
                    let not_null = self.emit_jump(Op::br_if_null);
                    let end = self.emit_jump(Op::br);
                    self.patch_jump(not_null);
                    self.emit(Op::drop);
                    self.compile_expression(right)?;
                    self.patch_jump(end);
                    return Ok(());
                }
                self.compile_expression(left)?;
                self.compile_expression(right)?;
                match op {
                    BinaryOp::Add => self.emit_js_add(),
                    BinaryOp::Sub => self.emit(Op::f64_sub),
                    BinaryOp::Mul => self.emit(Op::f64_mul),
                    BinaryOp::Div => self.emit(Op::f64_div),
                    BinaryOp::Mod => self.emit(Op::f64_mod),
                    BinaryOp::Exp => self.emit(Op::f64_mul), // TODO
                    BinaryOp::BitAnd => self.emit(Op::i32_and),
                    BinaryOp::BitOr => self.emit(Op::i32_or),
                    BinaryOp::BitXor => self.emit(Op::i32_xor),
                    BinaryOp::Shl => self.emit(Op::i32_shl),
                    BinaryOp::Shr => self.emit(Op::i32_shr_s),
                    BinaryOp::UShr => self.emit(Op::i32_shr_u),
                    BinaryOp::Eq => self.emit(Op::dyn_eq),
                    BinaryOp::Neq => self.emit(Op::dyn_ne),
                    BinaryOp::SEq => self.emit(Op::dyn_eq),
                    BinaryOp::SNeq => self.emit(Op::dyn_ne),
                    BinaryOp::Lt => self.emit(Op::dyn_lt),
                    BinaryOp::Gt => self.emit(Op::dyn_gt),
                    BinaryOp::Le => self.emit(Op::dyn_le),
                    BinaryOp::Ge => self.emit(Op::dyn_ge),
                    BinaryOp::InstanceOf | BinaryOp::In => {
                        self.emit(Op::drop); self.emit(Op::drop); self.emit(Op::r#false);
                    }
                    BinaryOp::NullishCoalescing => unreachable!(),
                }
            }
            Expression::Logical { op, left, right } => {
                self.compile_expression(left)?;
                match op {
                    LogicalOp::And => {
                        self.emit(Op::dup); self.emit_to_bool();
                        let end = self.emit_jump(Op::br_if_false);
                        self.emit(Op::drop);
                        self.compile_expression(right)?;
                        self.patch_jump(end);
                    }
                    LogicalOp::Or => {
                        self.emit(Op::dup); self.emit_to_bool();
                        let end = self.emit_jump(Op::br_if_true);
                        self.emit(Op::drop);
                        self.compile_expression(right)?;
                        self.patch_jump(end);
                    }
                }
            }
            Expression::Unary { op, argument } => {
                self.compile_expression(argument)?;
                match op {
                    UnaryOp::Neg => self.emit(Op::dyn_neg),
                    UnaryOp::Pos => {}
                    UnaryOp::Not => self.emit(Op::dyn_not),
                    UnaryOp::BitNot => self.emit(Op::i32_not),
                }
            }
            Expression::Update { op, prefix, argument } => {
                if *prefix {
                    self.compile_expression(argument)?;
                    self.emit_constant(Value::F64(1.0));
                    match op { UpdateOp::Increment => self.emit_js_add(), UpdateOp::Decrement => self.emit(Op::f64_sub) }
                    self.emit(Op::dup);
                    self.compile_store(argument)?;
                } else {
                    self.compile_expression(argument)?;
                    self.emit(Op::dup);
                    self.emit_constant(Value::F64(1.0));
                    match op { UpdateOp::Increment => self.emit_js_add(), UpdateOp::Decrement => self.emit(Op::f64_sub) }
                    self.compile_store(argument)?;
                }
            }
            Expression::Assignment { op, left, right } => {
                match left.as_ref() {
                    Expression::Member { object, property } => {
                        self.compile_expression(object)?;
                        if *op == AssignOp::Assign {
                            self.compile_expression(right)?;
                        } else {
                            self.emit(Op::dup);
                            let idx = self.add_string_constant(property);
                            self.emit_u16(Op::struct_get, idx);
                            self.compile_expression(right)?;
                            self.emit_compound_op(op);
                        }
                        let idx = self.add_string_constant(property);
                        self.emit_u16(Op::struct_set, idx);
                    }
                    Expression::ComputedMember { object, property } => {
                        self.compile_expression(object)?;
                        self.compile_expression(property)?;
                        if *op == AssignOp::Assign { self.compile_expression(right)?; }
                        else { self.compile_expression(right)?; }
                        self.emit(Op::array_set);
                    }
                    _ => {
                        if *op == AssignOp::Assign { self.compile_expression(right)?; }
                        else {
                            self.compile_expression(left)?;
                            self.compile_expression(right)?;
                            self.emit_compound_op(op);
                        }
                        self.emit(Op::dup);
                        self.compile_store(left)?;
                    }
                }
            }
            Expression::Conditional { test, consequent, alternate } => {
                self.compile_expression(test)?; self.emit_to_bool();
                let else_j = self.emit_jump(Op::br_if_false);
                self.compile_expression(consequent)?;
                let end_j = self.emit_jump(Op::br);
                self.patch_jump(else_j);
                self.compile_expression(alternate)?;
                self.patch_jump(end_j);
            }
            Expression::Member { object, property } => {
                self.compile_expression(object)?;
                let idx = self.add_string_constant(property);
                self.emit_u16(Op::struct_get, idx);
            }
            Expression::ComputedMember { object, property } => {
                self.compile_expression(object)?;
                self.compile_expression(property)?;
                self.emit(Op::array_get);
            }
            Expression::Call { callee, arguments } => {
                self.compile_call(callee, arguments)?;
            }
            Expression::New { callee, arguments } => {
                self.compile_expression(callee)?;
                self.emit_u16(Op::struct_new, 0);
                for arg in arguments { self.compile_expression(arg)?; }
                self.emit_u8(Op::call, (arguments.len() + 1) as u8);
            }
            Expression::Array(elements) => {
                for e in elements { self.compile_expression(e)?; }
                self.emit_u16(Op::array_new, elements.len() as u16);
            }
            Expression::Object(properties) => {
                let mut count = 0u16;
                for prop in properties {
                    match prop {
                        PropertyDef::KeyValue { key, value } => {
                            self.emit_constant(Value::String(Rc::from(key.as_str())));
                            self.compile_expression(value)?;
                            count += 1;
                        }
                        PropertyDef::Shorthand(name) => {
                            self.emit_constant(Value::String(Rc::from(name.as_str())));
                            match self.resolve_variable(name) {
                                VarResolution::Local(slot) => self.emit_u16(Op::local_get, slot),
                                VarResolution::Upvalue(idx) => self.emit_u8(Op::upvalue_get, idx),
                                VarResolution::Global => {
                                    let idx = self.add_string_constant(name);
                                    self.emit_u16(Op::global_get, idx);
                                }
                            }
                            count += 1;
                        }
                        PropertyDef::Method { key, value } => {
                            self.emit_constant(Value::String(Rc::from(key.as_str())));
                            self.compile_function(value)?;
                            count += 1;
                        }
                        PropertyDef::Spread(_) => {}
                    }
                }
                self.emit_u16(Op::struct_new, count);
            }
            Expression::Function(func) => { self.compile_function(func)?; }
            Expression::ArrowFunction { params, body } => {
                let func = match body {
                    ArrowBody::Block(stmts) => FunctionDecl { name: None, params: params.clone(), body: stmts.clone() },
                    ArrowBody::Expression(expr) => FunctionDecl { name: None, params: params.clone(), body: vec![Statement::Return(Some(*expr.clone()))] },
                };
                self.compile_function(&func)?;
            }
            Expression::TemplateLiteral { quasis, expressions } => {
                let to_str = self.import("js:coerce", "toString");
                let mut count = 0u8;
                for (i, quasi) in quasis.iter().enumerate() {
                    if !quasi.is_empty() || expressions.is_empty() {
                        self.emit_constant(Value::String(Rc::from(quasi.as_str())));
                        count += 1;
                    }
                    if i < expressions.len() {
                        self.compile_expression(&expressions[i])?;
                        self.emit_host_call(to_str, 1);
                        count += 1;
                    }
                }
                if count == 0 { self.emit_constant(Value::String(Rc::from(""))); }
                else if count > 1 { self.emit_u8(Op::str_concat_n, count); }
            }
            Expression::Typeof(arg) => {
                self.compile_expression(arg)?;
                let idx = self.import("js:coerce", "typeof");
                self.emit_host_call(idx, 1);
            }
            Expression::Void(arg) => {
                self.compile_expression(arg)?;
                self.emit(Op::drop);
                self.emit(Op::null);
            }
            Expression::Delete(_) => { self.emit(Op::r#true); }
            Expression::Spread(inner) => { self.compile_expression(inner)?; }
            Expression::Sequence(exprs) => {
                for (i, e) in exprs.iter().enumerate() {
                    self.compile_expression(e)?;
                    if i < exprs.len() - 1 { self.emit(Op::drop); }
                }
            }
        }
        Ok(())
    }

    /// Compile a function/method call.
    ///
    /// Resolution order for `obj.method(args)`:
    ///   1. If `obj` is NOT a known variable → module import: (module, method)
    ///   2. If `method` is a known value method (push, slice, etc.) → host call with obj as first arg
    ///   3. Otherwise → regular method call on object (obj.method with this binding)
    ///
    /// Resolution for `func(args)`:
    ///   1. If `func` is NOT a known variable → bare import: ("vybe:convert", func) for known globals
    ///   2. Otherwise → regular function call
    fn compile_call(&mut self, callee: &Expression, arguments: &[Expression]) -> Result<(), String> {
        if let Expression::Member { object, property } = callee {
            // obj.method() pattern
            if let Expression::Identifier(obj_name) = object.as_ref() {
                if !self.is_known_variable(obj_name) {
                    // obj is not a variable → it's a module namespace.
                    // Emit as import call: (module_alias, method)
                    let module = Self::js_module_alias(obj_name);
                    let (module, method) = Self::js_remap(module, property);
                    let idx = self.import(module, method);
                    for arg in arguments { self.compile_expression(arg)?; }
                    self.emit_host_call(idx, arguments.len() as u8);
                    return Ok(());
                }
            }

            // obj IS a variable — check for value methods (push, slice, etc.)
            if let Some(idx) = self.resolve_value_method(property) {
                self.compile_expression(object)?;
                for arg in arguments { self.compile_expression(arg)?; }
                self.emit_host_call(idx, (arguments.len() + 1) as u8);
                return Ok(());
            }

            // Regular method call on object: obj.method() with this binding
            self.compile_expression(object)?;
            let prop_idx = self.add_string_constant(property);
            self.emit_u16(Op::struct_get, prop_idx);
            self.compile_expression(object)?;
            for arg in arguments { self.compile_expression(arg)?; }
            self.emit_u8(Op::call, (arguments.len() + 1) as u8);
            return Ok(());
        }

        // Bare function call: func(args)
        // Only specific JS globals are bare imports. User-defined functions
        // are resolved via GlobalGet at runtime.
        if let Expression::Identifier(name) = callee {
            if let Some(idx) = self.resolve_bare_import(name) {
                for arg in arguments { self.compile_expression(arg)?; }
                self.emit_host_call(idx, arguments.len() as u8);
                return Ok(());
            }
        }

        // Regular function call
        self.compile_expression(callee)?;
        for arg in arguments { self.compile_expression(arg)?; }
        self.emit_u8(Op::call, arguments.len() as u8);
        Ok(())
    }

    fn emit_compound_op(&mut self, op: &AssignOp) {
        match op {
            AssignOp::AddAssign => self.emit_js_add(),
            AssignOp::SubAssign => self.emit(Op::f64_sub),
            AssignOp::MulAssign => self.emit(Op::f64_mul),
            AssignOp::DivAssign => self.emit(Op::f64_div),
            AssignOp::ModAssign => self.emit(Op::f64_mod),
            AssignOp::BitAndAssign => self.emit(Op::i32_and),
            AssignOp::BitOrAssign => self.emit(Op::i32_or),
            AssignOp::BitXorAssign => self.emit(Op::i32_xor),
            AssignOp::ShlAssign => self.emit(Op::i32_shl),
            AssignOp::ShrAssign => self.emit(Op::i32_shr_s),
            AssignOp::UShrAssign => self.emit(Op::i32_shr_u),
            _ => {}
        }
    }

    fn compile_store(&mut self, target: &Expression) -> Result<(), String> {
        match target {
            Expression::Identifier(name) => {
                match self.resolve_variable(name) {
                    VarResolution::Local(slot) => self.emit_u16(Op::local_set, slot),
                    VarResolution::Upvalue(idx) => self.emit_u8(Op::upvalue_set, idx),
                    VarResolution::Global => {
                        let idx = self.add_string_constant(name);
                        self.emit_u16(Op::global_set, idx);
                    }
                }
                self.emit(Op::drop);
            }
            Expression::Member { object, property } => {
                self.compile_expression(object)?;
                let idx = self.add_string_constant(property);
                self.emit_u16(Op::struct_set, idx);
            }
            Expression::ComputedMember { object, property } => {
                self.compile_expression(object)?;
                self.compile_expression(property)?;
                self.emit(Op::array_set);
            }
            _ => { self.emit(Op::drop); }
        }
        Ok(())
    }

    fn compile_function(&mut self, func: &FunctionDecl) -> Result<(), String> {
        let name = func.name.clone().unwrap_or_else(|| "<anonymous>".into());
        let mut chunk = Chunk::new(&name);
        chunk.arity = func.params.len() as u8;
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        for param in &func.params { scope.define_local(param); }

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        for stmt in &func.body { self.compile_statement(stmt)?; }
        self.emit(Op::null);
        self.emit(Op::r#return);

        let lc = self.current_scope().next_slot;
        self.chunks[idx].local_count = lc;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved;

        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u16(Op::ref_func, idx as u16, line);
        self.chunks[self.current_chunk_idx].emit(upvalues.len() as u8, line);
        for uv in &upvalues {
            self.chunks[self.current_chunk_idx].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current_chunk_idx].emit(uv.index, line);
        }
        Ok(())
    }

    fn compile_method(&mut self, func: &FunctionDecl) -> Result<(), String> {
        let saved_method = self.in_method;
        self.in_method = true;

        let name = func.name.clone().unwrap_or_else(|| "<method>".into());
        let mut chunk = Chunk::new(&name);
        chunk.arity = (func.params.len() + 1) as u8; // +1 for this

        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        scope.define_local("this");
        for param in &func.params { scope.define_local(param); }

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        for stmt in &func.body { self.compile_statement(stmt)?; }
        self.emit_u16(Op::local_get, 1); // return this
        self.emit(Op::r#return);

        let lc = self.current_scope().next_slot;
        self.chunks[idx].local_count = lc;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved;

        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u16(Op::ref_func, idx as u16, line);
        self.chunks[self.current_chunk_idx].emit(upvalues.len() as u8, line);
        for uv in &upvalues {
            self.chunks[self.current_chunk_idx].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current_chunk_idx].emit(uv.index, line);
        }

        self.in_method = saved_method;
        Ok(())
    }

    fn compile_class(&mut self, class: &ClassDecl) -> Result<(), String> {
        let name = class.name.as_deref().unwrap_or("<class>");
        let mut constructor = None;
        let mut methods: Vec<(String, FunctionDecl)> = Vec::new();
        for member in &class.body {
            if let ClassMember::Method { key, value, kind, .. } = member {
                if *kind == MethodKind::Constructor { constructor = Some(value.clone()); }
                else { methods.push((key.clone(), value.clone())); }
            }
        }
        let ctor_params = constructor.as_ref().map(|c| c.params.clone()).unwrap_or_default();
        let ctor_body = constructor.map(|c| c.body).unwrap_or_default();
        let ctor = FunctionDecl { name: Some(name.into()), params: ctor_params, body: ctor_body };
        self.compile_class_constructor(&ctor, &methods)?;
        Ok(())
    }

    fn compile_class_constructor(&mut self, ctor: &FunctionDecl, methods: &[(String, FunctionDecl)]) -> Result<(), String> {
        let saved_method = self.in_method;
        self.in_method = true;

        let name = ctor.name.clone().unwrap_or_else(|| "<class>".into());
        let mut chunk = Chunk::new(&name);
        chunk.arity = (ctor.params.len() + 1) as u8;
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        scope.define_local("this");
        for param in &ctor.params { scope.define_local(param); }

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        for stmt in &ctor.body { self.compile_statement(stmt)?; }

        // Attach methods to this
        for (method_name, method_fn) in methods {
            self.emit_u16(Op::local_get, 1); // this
            self.compile_method(method_fn)?;
            let prop_idx = self.add_string_constant(method_name);
            self.emit_u16(Op::struct_set, prop_idx);
            self.emit(Op::drop);
        }

        self.emit_u16(Op::local_get, 1); // return this
        self.emit(Op::r#return);

        let lc = self.current_scope().next_slot;
        self.chunks[idx].local_count = lc;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved;

        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u16(Op::ref_func, idx as u16, line);
        self.chunks[self.current_chunk_idx].emit(upvalues.len() as u8, line);
        for uv in &upvalues {
            self.chunks[self.current_chunk_idx].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current_chunk_idx].emit(uv.index, line);
        }

        self.in_method = saved_method;
        Ok(())
    }
}

enum VarResolution { Local(u16), Upvalue(u8), Global }
