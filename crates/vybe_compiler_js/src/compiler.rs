use std::rc::Rc;

use vybe_bytecode::{Chunk, Value, Op};
use vybe_parser_js::ast::*;

use crate::scope::Scope;

struct LoopContext {
    start_offset: usize,
    break_patches: Vec<usize>,
    continue_patches: Vec<usize>,
}

/// JS host function indices — set during compilation, must match
/// registration order in the runtime.
#[derive(Clone)]
pub struct HostFnTable {
    pub console_log: u16,
    pub console_error: u16,
    pub console_warn: u16,
    pub js_typeof: u16,
    pub js_to_number: u16,
    pub js_to_string: u16,
    pub js_to_boolean: u16,
    pub js_loose_eq: u16,
    pub js_add: u16,
    // GUI host functions (from vybe_host)
    pub gui: Option<vybe_host::gui::GuiHostFns>,
}

pub struct Compiler {
    chunks: Vec<Chunk>,
    scopes: Vec<Scope>,
    current_chunk_idx: usize,
    loop_stack: Vec<LoopContext>,
    line: u32,
    pub host: HostFnTable,
}

impl Compiler {
    pub fn new(host: HostFnTable) -> Self {
        Compiler {
            chunks: vec![Chunk::new("<script>")],
            scopes: vec![Scope::new()],
            current_chunk_idx: 0,
            loop_stack: Vec::new(),
            line: 1,
            host,
        }
    }

    pub fn compile(mut self, program: &Program) -> Result<Vec<Chunk>, String> {
        for stmt in &program.body {
            self.compile_statement(stmt)?;
        }
        self.emit(Op::PushNull);
        self.emit(Op::Halt);
        let local_count = self.current_scope().next_slot;
        self.chunks[0].local_count = local_count;
        Ok(self.chunks)
    }

    // -- Chunk helpers --

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
        self.emit_u16(Op::Const, idx);
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

    /// Emit CallHost: [CallHost, fn_hi, fn_lo, argc]
    fn emit_host_call(&mut self, fn_idx: u16, argc: u8) {
        let line = self.line;
        let c = &mut self.chunks[self.current_chunk_idx];
        c.emit_op(Op::CallHost, line);
        c.emit((fn_idx >> 8) as u8, line);
        c.emit((fn_idx & 0xff) as u8, line);
        c.emit(argc, line);
    }

    /// Emit JS truthiness conversion: value → Bool
    /// Calls the js_to_boolean host function.
    fn emit_to_bool(&mut self) {
        self.emit_host_call(self.host.js_to_boolean, 1);
    }

    /// Emit JS add: handles string concat + numeric add via host fn.
    fn emit_js_add(&mut self) {
        self.emit_host_call(self.host.js_add, 2);
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
            if let Some(uv_idx) = self.resolve_upvalue(self.scopes.len() - 1, name) {
                return VarResolution::Upvalue(uv_idx);
            }
        }
        VarResolution::Global
    }

    fn resolve_upvalue(&mut self, scope_idx: usize, name: &str) -> Option<u8> {
        if scope_idx == 0 { return None; }
        let parent_idx = scope_idx - 1;
        if let Some(slot) = self.scopes[parent_idx].resolve_local(name) {
            self.scopes[parent_idx].mark_captured(slot);
            return Some(self.scopes[scope_idx].add_upvalue(slot as u8, true));
        }
        if let Some(parent_uv) = self.resolve_upvalue(parent_idx, name) {
            return Some(self.scopes[scope_idx].add_upvalue(parent_uv, false));
        }
        None
    }

    // -- Statements --

    fn compile_statement(&mut self, stmt: &Statement) -> Result<(), String> {
        match stmt {
            Statement::Expression(expr) => {
                self.compile_expression(expr)?;
                self.emit(Op::Pop);
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
                        self.emit(Op::PushNull); // JS undefined → Null in VM
                    }
                    match self.resolve_variable(&decl.name) {
                        VarResolution::Local(_) => {
                            let slot = self.current_scope().resolve_local(&decl.name).unwrap();
                            self.emit_u16(Op::SetLocal, slot);
                            self.emit(Op::Pop);
                        }
                        _ => {
                            if self.scopes.len() == 1 && self.current_scope().depth == 0 && *kind == VarKind::Var {
                                let idx = self.add_string_constant(&decl.name);
                                self.emit_u16(Op::SetGlobal, idx);
                                self.emit(Op::Pop);
                            } else {
                                let slot = self.define_local(&decl.name);
                                self.emit_u16(Op::SetLocal, slot);
                                self.emit(Op::Pop);
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
                        self.emit_u16(Op::SetGlobal, idx);
                        self.emit(Op::Pop);
                    } else {
                        let slot = self.define_local(name);
                        self.emit_u16(Op::SetLocal, slot);
                        self.emit(Op::Pop);
                    }
                }
            }
            Statement::If { test, consequent, alternate } => {
                self.compile_expression(test)?;
                self.emit_to_bool();
                let else_jump = self.emit_jump(Op::JumpIfFalse);
                self.compile_statement(consequent)?;
                if let Some(alt) = alternate {
                    let end_jump = self.emit_jump(Op::Jump);
                    self.patch_jump(else_jump);
                    self.compile_statement(alt)?;
                    self.patch_jump(end_jump);
                } else {
                    self.patch_jump(else_jump);
                }
            }
            Statement::While { test, body } => {
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { start_offset: loop_start, break_patches: vec![], continue_patches: vec![] });
                self.compile_expression(test)?;
                self.emit_to_bool();
                let exit = self.emit_jump(Op::JumpIfFalse);
                self.compile_statement(body)?;
                let cp: Vec<usize> = self.loop_stack.last().unwrap().continue_patches.clone();
                for p in cp { self.patch_jump(p); }
                self.emit_loop(loop_start);
                self.patch_jump(exit);
                let ctx = self.loop_stack.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
            }
            Statement::DoWhile { body, test } => {
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { start_offset: loop_start, break_patches: vec![], continue_patches: vec![] });
                self.compile_statement(body)?;
                let cp: Vec<usize> = self.loop_stack.last().unwrap().continue_patches.clone();
                for p in cp { self.patch_jump(p); }
                self.compile_expression(test)?;
                self.emit_to_bool();
                let exit = self.emit_jump(Op::JumpIfFalse);
                self.emit_loop(loop_start);
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
                        ForInit::Expression(expr) => { self.compile_expression(expr)?; self.emit(Op::Pop); }
                    }
                }
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { start_offset: loop_start, break_patches: vec![], continue_patches: vec![] });
                let exit = if let Some(test) = test {
                    self.compile_expression(test)?;
                    self.emit_to_bool();
                    Some(self.emit_jump(Op::JumpIfFalse))
                } else { None };
                self.compile_statement(body)?;
                let cp: Vec<usize> = self.loop_stack.last().unwrap().continue_patches.clone();
                for p in cp { self.patch_jump(p); }
                if let Some(update) = update { self.compile_expression(update)?; self.emit(Op::Pop); }
                self.emit_loop(loop_start);
                if let Some(e) = exit { self.patch_jump(e); }
                let ctx = self.loop_stack.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
                self.current_scope_mut().end_scope();
            }
            Statement::Return(value) => {
                if let Some(expr) = value { self.compile_expression(expr)?; }
                else { self.emit(Op::PushNull); }
                self.emit(Op::Return);
            }
            Statement::Break(_) => {
                let patch = self.emit_jump(Op::Jump);
                if let Some(ctx) = self.loop_stack.last_mut() { ctx.break_patches.push(patch); }
            }
            Statement::Continue(_) => {
                let patch = self.emit_jump(Op::Jump);
                if let Some(ctx) = self.loop_stack.last_mut() { ctx.continue_patches.push(patch); }
            }
            Statement::Throw(expr) => { self.compile_expression(expr)?; self.emit(Op::Throw); }
            Statement::Try { block, handler, finalizer } => {
                let line = self.line;
                let c = &mut self.chunks[self.current_chunk_idx];
                c.emit_op(Op::TryStart, line);
                c.emit(0, line); c.emit(0, line); c.emit(0, line); c.emit(0, line);
                for s in block { self.compile_statement(s)?; }
                self.emit(Op::TryEnd);
                if let Some(h) = handler { for s in &h.body { self.compile_statement(s)?; } }
                if let Some(f) = finalizer { for s in f { self.compile_statement(s)?; } }
            }
            Statement::Switch { discriminant, cases } => {
                self.compile_expression(discriminant)?;
                self.loop_stack.push(LoopContext { start_offset: 0, break_patches: vec![], continue_patches: vec![] });
                let mut next_case_patch: Option<usize> = None;
                for case in cases {
                    if let Some(patch) = next_case_patch.take() { self.patch_jump(patch); }
                    if let Some(test) = &case.test {
                        self.emit(Op::Dup);
                        self.compile_expression(test)?;
                        self.emit(Op::CmpEq);
                        next_case_patch = Some(self.emit_jump(Op::JumpIfFalse));
                    }
                    for s in &case.consequent { self.compile_statement(s)?; }
                }
                if let Some(patch) = next_case_patch { self.patch_jump(patch); }
                let ctx = self.loop_stack.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
                self.emit(Op::Pop);
            }
            Statement::ClassDeclaration(class) => {
                self.compile_class(class)?;
                if let Some(name) = &class.name {
                    if self.scopes.len() == 1 && self.current_scope().depth == 0 {
                        let idx = self.add_string_constant(name);
                        self.emit_u16(Op::SetGlobal, idx);
                        self.emit(Op::Pop);
                    } else {
                        let slot = self.define_local(name);
                        self.emit_u16(Op::SetLocal, slot);
                        self.emit(Op::Pop);
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
            Expression::Number(n) => {
                self.emit_constant(Value::F64(*n));
            }
            Expression::String(s) => {
                self.emit_constant(Value::String(Rc::from(s.as_str())));
            }
            Expression::Boolean(true) => self.emit(Op::PushTrue),
            Expression::Boolean(false) => self.emit(Op::PushFalse),
            Expression::Null => self.emit(Op::PushNull),
            Expression::Undefined => self.emit(Op::PushNull), // JS undefined → VM Null
            Expression::This => self.emit(Op::PushNull),
            Expression::Identifier(name) => {
                match self.resolve_variable(name) {
                    VarResolution::Local(slot) => self.emit_u16(Op::GetLocal, slot),
                    VarResolution::Upvalue(idx) => self.emit_u8(Op::GetUpvalue, idx),
                    VarResolution::Global => {
                        let idx = self.add_string_constant(name);
                        self.emit_u16(Op::GetGlobal, idx);
                    }
                }
            }
            Expression::Binary { op, left, right } => {
                if *op == BinaryOp::NullishCoalescing {
                    self.compile_expression(left)?;
                    self.emit(Op::Dup);
                    let not_null = self.emit_jump(Op::JumpIfNull);
                    // Not null — keep left, skip right
                    let end = self.emit_jump(Op::Jump);
                    self.patch_jump(not_null);
                    // Was null — discard left, use right
                    self.emit(Op::Pop);
                    self.compile_expression(right)?;
                    self.patch_jump(end);
                    return Ok(());
                }

                self.compile_expression(left)?;
                self.compile_expression(right)?;
                match op {
                    // JS + is special: string concat or numeric add. Delegate to host.
                    BinaryOp::Add => self.emit_js_add(),
                    BinaryOp::Sub => self.emit(Op::SubF),
                    BinaryOp::Mul => self.emit(Op::MulF),
                    BinaryOp::Div => self.emit(Op::DivF),
                    BinaryOp::Mod => self.emit(Op::ModF),
                    BinaryOp::Exp => self.emit(Op::MulF), // TODO
                    BinaryOp::BitAnd => self.emit(Op::BitAnd),
                    BinaryOp::BitOr => self.emit(Op::BitOr),
                    BinaryOp::BitXor => self.emit(Op::BitXor),
                    BinaryOp::Shl => self.emit(Op::Shl),
                    BinaryOp::Shr => self.emit(Op::Shr),
                    BinaryOp::UShr => self.emit(Op::UShr),
                    // JS == is language-specific coercion. Delegate to host.
                    BinaryOp::Eq => self.emit_host_call(self.host.js_loose_eq, 2),
                    BinaryOp::Neq => {
                        self.emit_host_call(self.host.js_loose_eq, 2);
                        self.emit(Op::BoolNot);
                    }
                    // === is same-type equality — that's what CmpEq does
                    BinaryOp::SEq => self.emit(Op::CmpEq),
                    BinaryOp::SNeq => { self.emit(Op::CmpEq); self.emit(Op::BoolNot); }
                    BinaryOp::Lt => self.emit(Op::CmpLtF),
                    BinaryOp::Gt => self.emit(Op::CmpGtF),
                    BinaryOp::Le => self.emit(Op::CmpLeF),
                    BinaryOp::Ge => self.emit(Op::CmpGeF),
                    BinaryOp::InstanceOf => { self.emit(Op::Pop); self.emit(Op::Pop); self.emit(Op::PushFalse); }
                    BinaryOp::In => { self.emit(Op::Pop); self.emit(Op::Pop); self.emit(Op::PushFalse); }
                    BinaryOp::NullishCoalescing => unreachable!(),
                }
            }
            Expression::Logical { op, left, right } => {
                self.compile_expression(left)?;
                match op {
                    LogicalOp::And => {
                        self.emit(Op::Dup);
                        self.emit_to_bool();
                        let end = self.emit_jump(Op::JumpIfFalse);
                        self.emit(Op::Pop);
                        self.compile_expression(right)?;
                        self.patch_jump(end);
                    }
                    LogicalOp::Or => {
                        self.emit(Op::Dup);
                        self.emit_to_bool();
                        let end = self.emit_jump(Op::JumpIfTrue);
                        self.emit(Op::Pop);
                        self.compile_expression(right)?;
                        self.patch_jump(end);
                    }
                }
            }
            Expression::Unary { op, argument } => {
                self.compile_expression(argument)?;
                match op {
                    UnaryOp::Neg => self.emit(Op::NegF),
                    UnaryOp::Pos => {} // no-op
                    UnaryOp::Not => { self.emit_to_bool(); self.emit(Op::BoolNot); }
                    UnaryOp::BitNot => self.emit(Op::BitNot),
                }
            }
            Expression::Update { op, prefix, argument } => {
                if *prefix {
                    self.compile_expression(argument)?;
                    self.emit_constant(Value::F64(1.0));
                    match op { UpdateOp::Increment => self.emit_js_add(), UpdateOp::Decrement => self.emit(Op::SubF) }
                    self.emit(Op::Dup);
                    self.compile_store(argument)?;
                } else {
                    self.compile_expression(argument)?;
                    self.emit(Op::Dup);
                    self.emit_constant(Value::F64(1.0));
                    match op { UpdateOp::Increment => self.emit_js_add(), UpdateOp::Decrement => self.emit(Op::SubF) }
                    self.compile_store(argument)?;
                }
            }
            Expression::Assignment { op, left, right } => {
                if *op == AssignOp::Assign {
                    self.compile_expression(right)?;
                } else {
                    self.compile_expression(left)?;
                    self.compile_expression(right)?;
                    match op {
                        AssignOp::AddAssign => self.emit_js_add(),
                        AssignOp::SubAssign => self.emit(Op::SubF),
                        AssignOp::MulAssign => self.emit(Op::MulF),
                        AssignOp::DivAssign => self.emit(Op::DivF),
                        AssignOp::ModAssign => self.emit(Op::ModF),
                        AssignOp::BitAndAssign => self.emit(Op::BitAnd),
                        AssignOp::BitOrAssign => self.emit(Op::BitOr),
                        AssignOp::BitXorAssign => self.emit(Op::BitXor),
                        AssignOp::ShlAssign => self.emit(Op::Shl),
                        AssignOp::ShrAssign => self.emit(Op::Shr),
                        AssignOp::UShrAssign => self.emit(Op::UShr),
                        _ => {}
                    }
                }
                self.emit(Op::Dup);
                self.compile_store(left)?;
            }
            Expression::Conditional { test, consequent, alternate } => {
                self.compile_expression(test)?;
                self.emit_to_bool();
                let else_j = self.emit_jump(Op::JumpIfFalse);
                self.compile_expression(consequent)?;
                let end_j = self.emit_jump(Op::Jump);
                self.patch_jump(else_j);
                self.compile_expression(alternate)?;
                self.patch_jump(end_j);
            }
            Expression::Member { object, property } => {
                self.compile_expression(object)?;
                let idx = self.add_string_constant(property);
                self.emit_u16(Op::GetProp, idx);
            }
            Expression::ComputedMember { object, property } => {
                self.compile_expression(object)?;
                self.compile_expression(property)?;
                self.emit(Op::GetIndex);
            }
            Expression::Call { callee, arguments } => {
                // Check for known host function calls (console.log, etc.)
                if let Some(host_idx) = self.resolve_host_call(callee) {
                    for arg in arguments { self.compile_expression(arg)?; }
                    self.emit_host_call(host_idx, arguments.len() as u8);
                } else {
                    self.compile_expression(callee)?;
                    for arg in arguments { self.compile_expression(arg)?; }
                    self.emit_u8(Op::Call, arguments.len() as u8);
                }
            }
            Expression::New { callee, arguments } => {
                self.compile_expression(callee)?;
                for arg in arguments { self.compile_expression(arg)?; }
                self.emit_u8(Op::Call, arguments.len() as u8);
            }
            Expression::Array(elements) => {
                for e in elements { self.compile_expression(e)?; }
                self.emit_u16(Op::NewArray, elements.len() as u16);
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
                                VarResolution::Local(slot) => self.emit_u16(Op::GetLocal, slot),
                                VarResolution::Upvalue(idx) => self.emit_u8(Op::GetUpvalue, idx),
                                VarResolution::Global => {
                                    let idx = self.add_string_constant(name);
                                    self.emit_u16(Op::GetGlobal, idx);
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
                self.emit_u16(Op::NewObject, count);
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
                let mut part_count = 0u8;
                for (i, quasi) in quasis.iter().enumerate() {
                    if !quasi.is_empty() || expressions.is_empty() {
                        self.emit_constant(Value::String(Rc::from(quasi.as_str())));
                        part_count += 1;
                    }
                    if i < expressions.len() {
                        self.compile_expression(&expressions[i])?;
                        self.emit_host_call(self.host.js_to_string, 1);
                        part_count += 1;
                    }
                }
                if part_count == 0 {
                    self.emit_constant(Value::String(Rc::from("")));
                } else if part_count > 1 {
                    self.emit_u8(Op::StrConcat, part_count);
                }
            }
            Expression::Typeof(arg) => {
                self.compile_expression(arg)?;
                self.emit_host_call(self.host.js_typeof, 1);
            }
            Expression::Void(arg) => {
                self.compile_expression(arg)?;
                self.emit(Op::Pop);
                self.emit(Op::PushNull);
            }
            Expression::Delete(_) => { self.emit(Op::PushTrue); }
            Expression::Spread(inner) => { self.compile_expression(inner)?; }
            Expression::Sequence(exprs) => {
                for (i, e) in exprs.iter().enumerate() {
                    self.compile_expression(e)?;
                    if i < exprs.len() - 1 { self.emit(Op::Pop); }
                }
            }
        }
        Ok(())
    }

    fn compile_store(&mut self, target: &Expression) -> Result<(), String> {
        match target {
            Expression::Identifier(name) => {
                match self.resolve_variable(name) {
                    VarResolution::Local(slot) => self.emit_u16(Op::SetLocal, slot),
                    VarResolution::Upvalue(idx) => self.emit_u8(Op::SetUpvalue, idx),
                    VarResolution::Global => {
                        let idx = self.add_string_constant(name);
                        self.emit_u16(Op::SetGlobal, idx);
                    }
                }
                self.emit(Op::Pop);
            }
            Expression::Member { object, property } => {
                self.compile_expression(object)?;
                let idx = self.add_string_constant(property);
                self.emit_u16(Op::SetProp, idx);
            }
            Expression::ComputedMember { object, property } => {
                self.compile_expression(object)?;
                self.compile_expression(property)?;
                self.emit(Op::SetIndex);
            }
            _ => { self.emit(Op::Pop); }
        }
        Ok(())
    }

    fn compile_function(&mut self, func: &FunctionDecl) -> Result<(), String> {
        let func_name = func.name.clone().unwrap_or_else(|| "<anonymous>".into());
        let mut func_chunk = Chunk::new(&func_name);
        func_chunk.arity = func.params.len() as u8;
        let func_chunk_idx = self.chunks.len();
        self.chunks.push(func_chunk);

        let mut func_scope = Scope::new_function();
        for param in &func.params { func_scope.define_local(param); }

        let saved_chunk = self.current_chunk_idx;
        self.current_chunk_idx = func_chunk_idx;
        self.scopes.push(func_scope);

        for stmt in &func.body { self.compile_statement(stmt)?; }
        self.emit(Op::PushNull);
        self.emit(Op::Return);

        let local_count = self.current_scope().next_slot;
        self.chunks[func_chunk_idx].local_count = local_count;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved_chunk;

        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u16(Op::Closure, func_chunk_idx as u16, line);
        self.chunks[self.current_chunk_idx].emit(upvalues.len() as u8, line);
        for uv in &upvalues {
            self.chunks[self.current_chunk_idx].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current_chunk_idx].emit(uv.index, line);
        }
        Ok(())
    }

    fn compile_class(&mut self, class: &ClassDecl) -> Result<(), String> {
        let name = class.name.as_deref().unwrap_or("<anonymous>");
        let mut constructor = None;
        for member in &class.body {
            if let ClassMember::Method { value, kind, .. } = member {
                if *kind == MethodKind::Constructor { constructor = Some(value.clone()); }
            }
        }
        let ctor = constructor.unwrap_or(FunctionDecl { name: Some(name.into()), params: vec![], body: vec![] });
        self.compile_function(&ctor)?;
        Ok(())
    }

    /// Check if a call expression is a known host function (e.g. console.log, gui.addControl).
    fn resolve_host_call(&self, callee: &Expression) -> Option<u16> {
        if let Expression::Member { object, property } = callee {
            if let Expression::Identifier(obj_name) = object.as_ref() {
                match obj_name.as_str() {
                    "console" => {
                        return match property.as_str() {
                            "log" => Some(self.host.console_log),
                            "error" => Some(self.host.console_error),
                            "warn" => Some(self.host.console_warn),
                            _ => None,
                        };
                    }
                    "gui" => {
                        if let Some(ref g) = self.host.gui {
                            return match property.as_str() {
                                "createForm" => Some(g.create_form),
                                "addControl" => Some(g.add_control),
                                "setProperty" => Some(g.set_property),
                                "getProperty" => Some(g.get_property),
                                "onEvent" => Some(g.on_event),
                                "showForm" => Some(g.show_form),
                                "runApplication" => Some(g.run_application),
                                "msgBox" => Some(g.msg_box),
                                "closeForm" => Some(g.close_form),
                                _ => None,
                            };
                        }
                    }
                    _ => {}
                }
            }
        }
        None
    }
}

enum VarResolution { Local(u16), Upvalue(u8), Global }
