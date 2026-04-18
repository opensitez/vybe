use std::sync::Arc;
use vybe_bytecode::{Chunk, Value, Op};
use crate::parser_php::ast::*;
use vybe_compiler_common as common;
use vybe_compiler_common::strings as common_strings;
use super::scope::Scope;

struct LoopContext {
    break_patches: Vec<usize>,
    continue_patches: Vec<usize>,
}

pub struct Compiler {
    chunks: Vec<Chunk>,
    scopes: Vec<Scope>,
    current_chunk_idx: usize,
    loop_stack: Vec<LoopContext>,
    line: u32,
    defined_globals: std::collections::HashSet<String>,
    defined_classes: std::collections::HashSet<String>,
    /// Track current class parent for parent::__construct calls
    current_class_parent: Option<String>,
}

impl Compiler {
    pub fn new() -> Self {
        Compiler {
            chunks: vec![Chunk::new("<script>")],
            scopes: vec![Scope::new()],
            current_chunk_idx: 0,
            loop_stack: Vec::new(),
            line: 1,
            defined_globals: std::collections::HashSet::new(),
            defined_classes: std::collections::HashSet::new(),
            current_class_parent: None,
        }
    }

    pub fn compile(mut self, program: &Program) -> Result<Vec<Chunk>, String> {
        for stmt in &program.body {
            self.compile_statement(stmt)?;
        }
        self.emit(Op::NULL);
        self.emit(Op::HALT);
        let local_count = self.current_scope().next_slot;
        self.chunks[0].local_count = local_count;
        common::bundle::finalize_with_stdlib(&mut self.chunks);
        Ok(self.chunks)
    }

    // ------------------------------------------------------------------
    // Chunk / emit helpers
    // ------------------------------------------------------------------

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
        self.emit_u16(Op::CONST, idx);
    }

    fn add_string_constant(&mut self, s: &str) -> u16 {
        self.chunks[self.current_chunk_idx].add_constant(Value::String(Arc::from(s)))
    }

    fn current_offset(&self) -> usize {
        self.chunks[self.current_chunk_idx].code.len()
    }

    fn emit_jump(&mut self, op: Op) -> usize {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_jump(op, line)
    }

    fn patch_jump(&mut self, site: usize) {
        self.chunks[self.current_chunk_idx].patch_jump(site);
    }

    fn emit_loop(&mut self, loop_start: usize) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_loop(loop_start, line);
    }

    fn import(&mut self, module: &str, name: &str) -> u16 {
        self.chunks[0].add_import(module, name)
    }

    /// Emit a host import call: [call_import, u16 idx, u8 argc] (4 bytes).
    fn emit_host_call(&mut self, import_idx: u16, argc: u8) {
        let line = self.line;
        let c = &mut self.chunks[self.current_chunk_idx];
        c.emit_op_u16(Op::CALL_IMPORT, import_idx, line);
        c.emit(argc, line);
    }

    // ------------------------------------------------------------------
    // Scope helpers
    // ------------------------------------------------------------------

    fn current_scope(&self) -> &Scope {
        self.scopes.last().unwrap()
    }

    fn current_scope_mut(&mut self) -> &mut Scope {
        self.scopes.last_mut().unwrap()
    }

    fn is_global_scope(&self) -> bool {
        self.scopes.len() == 1
    }

    fn resolve_var(&self, name: &str) -> Option<u16> {
        if self.is_global_scope() { return None; }
        if self.current_scope().globals.contains(name) { return None; }
        self.current_scope().resolve_local(name)
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

    fn emit_var_get(&mut self, name: &str) {
        if let Some(slot) = self.resolve_var(name) {
            self.emit_u16(Op::LOCAL_GET, slot);
        } else if self.scopes.len() > 1 {
            if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, name) {
                self.emit_u8(Op::UPVALUE_GET, uv);
                return;
            }
            let idx = self.add_string_constant(name);
            self.emit_u16(Op::GLOBAL_GET, idx);
        } else {
            let idx = self.add_string_constant(name);
            self.emit_u16(Op::GLOBAL_GET, idx);
        }
    }

    fn emit_var_set(&mut self, name: &str) {
        if let Some(slot) = self.resolve_var(name) {
            self.emit_u16(Op::LOCAL_SET, slot);
        } else if self.scopes.len() > 1 {
            if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, name) {
                self.emit_u8(Op::UPVALUE_SET, uv);
                return;
            }
            let idx = self.add_string_constant(name);
            self.emit_u16(Op::GLOBAL_SET, idx);
            self.defined_globals.insert(name.to_string());
        } else {
            let idx = self.add_string_constant(name);
            self.emit_u16(Op::GLOBAL_SET, idx);
            self.defined_globals.insert(name.to_string());
        }
    }

    fn define_local(&mut self, name: &str) -> u16 {
        self.current_scope_mut().define_local(name)
    }

    fn define_local_or_get(&mut self, name: &str) -> u16 {
        if let Some(slot) = self.current_scope().resolve_local(name) {
            slot
        } else {
            self.define_local(name)
        }
    }

    // ------------------------------------------------------------------
    // Statements
    // ------------------------------------------------------------------

    fn compile_statement(&mut self, stmt: &Statement) -> Result<(), String> {
        match stmt {
            Statement::Empty => {}

            Statement::Echo(exprs) => {
                for expr in exprs {
                    self.compile_expression(expr)?;
                    let line = self.line;
                    common::io::emit_print(&mut self.chunks[self.current_chunk_idx], 1, line);
                    self.emit(Op::DROP);
                }
            }

            Statement::Expression(expr) => {
                self.compile_expression(expr)?;
                self.emit(Op::DROP);
            }

            Statement::Block(stmts) => {
                self.current_scope_mut().begin_block();
                for s in stmts { self.compile_statement(s)?; }
                self.current_scope_mut().end_block();
            }

            Statement::VariableDeclaration { name, value } => {
                if let Some(val) = value {
                    self.compile_expression(val)?;
                } else {
                    self.emit(Op::NULL);
                }
                if self.is_global_scope() {
                    let idx = self.add_string_constant(name);
                    self.emit_u16(Op::GLOBAL_SET, idx);
                    self.defined_globals.insert(name.clone());
                } else {
                    let slot = self.define_local(name);
                    self.emit_u16(Op::LOCAL_SET, slot);
                }
            }

            Statement::ConstDeclaration { name, value } => {
                self.compile_expression(value)?;
                let idx = self.add_string_constant(name);
                self.emit_u16(Op::GLOBAL_SET, idx);
                self.defined_globals.insert(name.clone());
            }

            Statement::Global(vars) => {
                for v in vars {
                    self.current_scope_mut().globals.insert(v.clone());
                }
            }

            Statement::Return(val) => {
                if let Some(expr) = val {
                    self.compile_expression(expr)?;
                } else {
                    self.emit(Op::NULL);
                }
                self.emit(Op::RETURN);
            }

            Statement::Break(_) => {
                let patch = self.emit_jump(Op::BR);
                if let Some(ctx) = self.loop_stack.last_mut() {
                    ctx.break_patches.push(patch);
                }
            }

            Statement::Continue(_) => {
                let patch = self.emit_jump(Op::BR);
                if let Some(ctx) = self.loop_stack.last_mut() {
                    ctx.continue_patches.push(patch);
                }
            }

            Statement::If { test, consequent, alternates, alternate } => {
                self.compile_expression(test)?;
                self.emit(Op::DYN_TO_BOOL);
                let mut end_jumps = Vec::new();
                let skip = self.emit_jump(Op::BR_IF_FALSE);
                self.compile_statement(consequent)?;
                end_jumps.push(self.emit_jump(Op::BR));
                self.patch_jump(skip);
                for elif in alternates {
                    self.compile_expression(&elif.test)?;
                    self.emit(Op::DYN_TO_BOOL);
                    let s = self.emit_jump(Op::BR_IF_FALSE);
                    self.compile_statement(&elif.body)?;
                    end_jumps.push(self.emit_jump(Op::BR));
                    self.patch_jump(s);
                }
                if let Some(alt) = alternate {
                    self.compile_statement(alt)?;
                }
                for j in end_jumps { self.patch_jump(j); }
            }

            Statement::While { test, body } => {
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { break_patches: Vec::new(), continue_patches: Vec::new() });
                self.compile_expression(test)?;
                self.emit(Op::DYN_TO_BOOL);
                let exit = self.emit_jump(Op::BR_IF_FALSE);
                self.compile_statement(body)?;
                let continues: Vec<usize> = self.loop_stack.last().unwrap().continue_patches.clone();
                for c in &continues { self.patch_jump(*c); }
                self.emit_loop(loop_start);
                self.patch_jump(exit);
                let breaks: Vec<usize> = self.loop_stack.pop().unwrap().break_patches;
                for b in breaks { self.patch_jump(b); }
            }

            Statement::DoWhile { body, test } => {
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { break_patches: Vec::new(), continue_patches: Vec::new() });
                self.compile_statement(body)?;
                let continues: Vec<usize> = self.loop_stack.last().unwrap().continue_patches.clone();
                for c in &continues { self.patch_jump(*c); }
                self.compile_expression(test)?;
                self.emit(Op::DYN_TO_BOOL);
                self.emit_loop(loop_start);
                let breaks: Vec<usize> = self.loop_stack.pop().unwrap().break_patches;
                for b in breaks { self.patch_jump(b); }
            }

            Statement::For { init, test, update, body } => {
                for e in init { self.compile_expression(e)?; self.emit(Op::DROP); }
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { break_patches: Vec::new(), continue_patches: Vec::new() });
                let exit = if let Some(t) = test {
                    self.compile_expression(t)?;
                    self.emit(Op::DYN_TO_BOOL);
                    Some(self.emit_jump(Op::BR_IF_FALSE))
                } else { None };
                self.compile_statement(body)?;
                let continues: Vec<usize> = self.loop_stack.last().unwrap().continue_patches.clone();
                for c in &continues { self.patch_jump(*c); }
                for e in update { self.compile_expression(e)?; self.emit(Op::DROP); }
                self.emit_loop(loop_start);
                if let Some(e) = exit { self.patch_jump(e); }
                let breaks: Vec<usize> = self.loop_stack.pop().unwrap().break_patches;
                for b in breaks { self.patch_jump(b); }
            }

            Statement::ForEach { array, key, value, body } => {
                // Use common::dict::emit_keys to get keys, then iterate with array_get
                self.compile_expression(array)?;
                let arr_slot = self.define_local("__foreach_arr");
                self.emit_u16(Op::LOCAL_SET, arr_slot);

                // Get keys array: dict.__keys via common helper, or array indices
                {
                    let line = self.line;
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    common::dict::emit_keys(&mut self.chunks[self.current_chunk_idx], line);
                }
                let keys_slot = self.define_local("__foreach_keys");
                self.emit_u16(Op::LOCAL_SET, keys_slot);

                // length of keys
                self.emit_u16(Op::LOCAL_GET, keys_slot);
                self.emit(Op::ARRAY_LENGTH);
                let len_slot = self.define_local("__foreach_len");
                self.emit_u16(Op::LOCAL_SET, len_slot);

                // index = 0
                self.emit_constant(Value::I32(0));
                let idx_slot = self.define_local("__foreach_idx");
                self.emit_u16(Op::LOCAL_SET, idx_slot);

                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { break_patches: Vec::new(), continue_patches: Vec::new() });

                // while idx < len
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_u16(Op::LOCAL_GET, len_slot);
                self.emit(Op::DYN_LT);
                let exit = self.emit_jump(Op::BR_IF_FALSE);

                // current_key = keys[idx]
                self.emit_u16(Op::LOCAL_GET, keys_slot);
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit(Op::ARRAY_GET);
                let key_tmp_slot = self.define_local("__foreach_key_tmp");
                self.emit_u16(Op::LOCAL_SET, key_tmp_slot);

                if let Some(key_name) = key {
                    let key_slot = self.define_local_or_get(key_name);
                    self.emit_u16(Op::LOCAL_GET, key_tmp_slot);
                    self.emit_u16(Op::LOCAL_SET, key_slot);
                }

                // $val = $arr[current_key] — use array_get (works for both arrays and dicts)
                self.emit_u16(Op::LOCAL_GET, arr_slot);
                self.emit_u16(Op::LOCAL_GET, key_tmp_slot);
                self.emit(Op::ARRAY_GET);
                let val_slot = self.define_local_or_get(value);
                self.emit_u16(Op::LOCAL_SET, val_slot);

                self.compile_statement(body)?;

                // continue point: idx++
                let continues: Vec<usize> = self.loop_stack.last().unwrap().continue_patches.clone();
                for c in &continues { self.patch_jump(*c); }
                self.emit_u16(Op::LOCAL_GET, idx_slot);
                self.emit_constant(Value::I32(1));
                self.emit(Op::I32_ADD);
                self.emit_u16(Op::LOCAL_SET, idx_slot);
                self.emit_loop(loop_start);

                self.patch_jump(exit);
                let breaks: Vec<usize> = self.loop_stack.pop().unwrap().break_patches;
                for b in breaks { self.patch_jump(b); }
            }

            Statement::Switch { discriminant, cases } => {
                self.compile_expression(discriminant)?;
                let disc_slot = self.define_local("__switch_disc");
                self.emit_u16(Op::LOCAL_SET, disc_slot);

                let mut end_jumps: Vec<usize> = Vec::new();
                let mut next_case_jump: Option<usize> = None;

                for case in cases {
                    if let Some(nj) = next_case_jump.take() {
                        self.patch_jump(nj);
                    }
                    if let Some(test) = &case.test {
                        self.emit_u16(Op::LOCAL_GET, disc_slot);
                        self.compile_expression(test)?;
                        self.emit(Op::EQ);
                        next_case_jump = Some(self.emit_jump(Op::BR_IF_FALSE));
                    }
                    for s in &case.body { self.compile_statement(s)?; }
                    end_jumps.push(self.emit_jump(Op::BR));
                }
                if let Some(nj) = next_case_jump { self.patch_jump(nj); }
                for j in end_jumps { self.patch_jump(j); }
            }

            Statement::Throw(expr) => {
                self.compile_expression(expr)?;
                let line = self.line;
                common::errors::emit_throw(&mut self.chunks[self.current_chunk_idx], line);
            }

            Statement::Try { block, catches, finalizer } => {
                // Use common::errors helpers (same as Python/JS/Dart)
                let line = self.line;
                let c = self.current_chunk_idx;
                let catch_jump = common::errors::emit_try_start(&mut self.chunks[c], line);

                // Compile try body
                for s in block { self.compile_statement(s)?; }
                let line = self.line;
                common::errors::emit_try_end(&mut self.chunks[c], line);
                let skip_catch = self.emit_jump(Op::BR);

                // Patch catch offset
                common::errors::patch_catch(&mut self.chunks[c], catch_jump);

                // Catch blocks — exception value is on stack
                for catch in catches {
                    // Map PHP exception types to canonical names for cross-language compat
                    for type_name in &catch.types {
                        let _canonical = common::errors::canonical_exception_name(type_name);
                        // TODO: type-check exception against canonical name
                    }
                    if let Some(var) = &catch.var {
                        let slot = self.define_local_or_get(var);
                        self.emit_u16(Op::LOCAL_SET, slot);
                        self.emit(Op::DROP);
                    } else {
                        self.emit(Op::DROP);
                    }
                    for s in &catch.body { self.compile_statement(s)?; }
                }

                self.patch_jump(skip_catch);

                // Finally block
                if let Some(fin) = finalizer {
                    for s in fin { self.compile_statement(s)?; }
                }
            }

            Statement::FunctionDeclaration(decl) => {
                let chunk_idx = self.compile_function(decl)?;
                let line = self.line;
                common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], chunk_idx, 0, line);
                let name_lower = decl.name.to_lowercase();
                let idx = self.add_string_constant(&name_lower);
                self.emit_u16(Op::GLOBAL_SET, idx);
                self.defined_globals.insert(name_lower);
            }

            Statement::ClassDeclaration(decl) => {
                self.compile_class(decl)?;
            }
        }
        Ok(())
    }

    // ------------------------------------------------------------------
    // Expressions
    // ------------------------------------------------------------------

    fn compile_expression(&mut self, expr: &Expression) -> Result<(), String> {
        match expr {
            Expression::Number(n) => {
                self.emit_constant(Value::F64(*n));
            }
            Expression::Str(s) => {
                self.emit_constant(Value::String(Arc::from(s.as_str())));
            }
            Expression::Bool(b) => {
                if *b { self.emit(Op::TRUE); } else { self.emit(Op::FALSE); }
            }
            Expression::Null => {
                self.emit(Op::NULL);
            }
            Expression::This => {
                // $this is slot 1 in method scope (slot 0 = callee)
                self.emit_u16(Op::LOCAL_GET, 1);
            }
            Expression::Variable(name) => {
                self.emit_var_get(name);
            }
            Expression::Identifier(name) => {
                let lower = name.to_lowercase();
                match lower.as_str() {
                    "php_eol" => { self.emit_constant(Value::String(Arc::from("\n"))); }
                    "php_int_max" => { self.emit_constant(Value::F64(i64::MAX as f64)); }
                    "php_int_min" => { self.emit_constant(Value::F64(i64::MIN as f64)); }
                    "m_pi" | "m_pi_value" => { self.emit_constant(Value::F64(std::f64::consts::PI)); }
                    "true" => { self.emit(Op::TRUE); }
                    "false" => { self.emit(Op::FALSE); }
                    "null" => { self.emit(Op::NULL); }
                    _ => {
                        let idx = self.add_string_constant(&lower);
                        self.emit_u16(Op::GLOBAL_GET, idx);
                    }
                }
            }
            Expression::ClassKeyword(kw) => {
                match kw.as_str() {
                    "parent" => {
                        // parent:: resolves to the parent class constructor via global
                        if let Some(ref parent) = self.current_class_parent.clone() {
                            let idx = self.add_string_constant(parent);
                            self.emit_u16(Op::GLOBAL_GET, idx);
                        } else {
                            self.emit(Op::NULL);
                        }
                    }
                    "self" | "static" => {
                        // self:: resolves to the class constructor (stored as global)
                        // For now, get it from $this->__type name → global
                        self.emit_u16(Op::LOCAL_GET, 1); // $this
                        let idx = self.add_string_constant("__type");
                        self.emit_u16(Op::STRUCT_GET, idx);
                        // type name is a string, get the constructor from globals
                        common_strings::emit_to_lower(&mut self.chunks[self.current_chunk_idx], self.line);
                        // Can't do dynamic global_get — use the class name from context
                        // Fallback: just push $this for self:: method calls
                        self.emit(Op::DROP);
                        self.emit_u16(Op::LOCAL_GET, 1);
                    }
                    _ => {
                        let idx = self.add_string_constant(kw);
                        self.emit_u16(Op::GLOBAL_GET, idx);
                    }
                }
            }
            Expression::Array(elements) => {
                self.compile_array_literal(elements)?;
            }
            Expression::Assign { op, left, right } => {
                self.compile_assign(op, left, right)?;
            }
            Expression::Binary { op, left, right } => {
                self.compile_binary(op, left, right)?;
            }
            Expression::Unary { op, expr } => {
                self.compile_expression(expr)?;
                match op {
                    UnaryOp::Neg => { self.emit(Op::DYN_NEG); }
                    UnaryOp::Pos => {}
                    UnaryOp::Not => { self.emit(Op::DYN_NOT); }
                    UnaryOp::BitNot => { let l = self.line; vybe_compiler_common::expressions::emit_i32_not(&mut self.chunks[self.current_chunk_idx], l); }
                }
            }
            Expression::PreUpdate { op, expr } => {
                // ++$x → compute new value, store it, result = new value
                self.compile_expression(expr)?;
                match op {
                    UpdateOp::Inc => { self.emit_constant(Value::F64(1.0)); self.emit(Op::DYN_ADD); }
                    UpdateOp::Dec => { self.emit_constant(Value::F64(1.0)); self.emit(Op::F64_SUB); }
                }
                // Save new value
                let new_tmp = self.define_local("__pre_new");
                self.emit_u16(Op::LOCAL_SET, new_tmp);
                self.emit(Op::DROP); // local_set peeks, need to pop
                self.emit_u16(Op::LOCAL_GET, new_tmp);
                // Store back
                self.compile_assign_lhs(expr)?;
                // Push new value as result
                self.emit_u16(Op::LOCAL_GET, new_tmp);
            }
            Expression::PostUpdate { op, expr } => {
                // $x++ → evaluate old value, compute new, store new, result = old value
                // Save old value in temp
                self.compile_expression(expr)?;
                let old_tmp = self.define_local("__post_old");
                self.emit_u16(Op::LOCAL_SET, old_tmp);
                self.emit(Op::DROP); // local_set peeks, need to pop
                // Compute new value
                self.emit_u16(Op::LOCAL_GET, old_tmp);
                match op {
                    UpdateOp::Inc => { self.emit_constant(Value::F64(1.0)); self.emit(Op::DYN_ADD); }
                    UpdateOp::Dec => { self.emit_constant(Value::F64(1.0)); self.emit(Op::F64_SUB); }
                }
                // Store new value back
                self.compile_assign_lhs(expr)?;
                // Push old value as result
                self.emit_u16(Op::LOCAL_GET, old_tmp);
            }
            Expression::Ternary { test, consequent, alternate } => {
                let c = self.current_chunk_idx;
                let line = self.line;
                self.compile_expression(test)?;
                let false_jump = common::expressions::emit_ternary_start(&mut self.chunks[c], line);
                if let Some(cons) = consequent {
                    self.compile_expression(cons)?;
                } else {
                    self.compile_expression(test)?;
                }
                let end_jump = common::expressions::emit_ternary_middle(&mut self.chunks[c], false_jump, line);
                self.compile_expression(alternate)?;
                common::expressions::emit_ternary_end(&mut self.chunks[c], end_jump);
            }
            Expression::NullCoalesce { left, right } => {
                let c = self.current_chunk_idx;
                let line = self.line;
                self.compile_expression(left)?;
                let (_null_jump, end_jump) = common::expressions::emit_null_coalesce_start(&mut self.chunks[c], line);
                self.compile_expression(right)?;
                common::expressions::emit_null_coalesce_end(&mut self.chunks[c], end_jump);
            }
            Expression::Call { callee, args } => {
                self.compile_call(callee, args)?;
            }
            Expression::MethodCall { object, method, args, nullsafe } => {
                self.compile_method_call(object, method, args, *nullsafe)?;
            }
            Expression::StaticCall { class, method, args } => {
                self.compile_static_call(class, method, args)?;
            }
            Expression::New { class, args } => {
                self.compile_new(class, args)?;
            }
            Expression::Property { object, name, nullsafe: _ } => {
                self.compile_expression(object)?;
                match name.as_ref() {
                    Expression::Identifier(s) => {
                        let idx = self.add_string_constant(s);
                        self.emit_u16(Op::STRUCT_GET, idx);
                    }
                    _ => {
                        // dynamic property: obj[expr] via common::dict
                        self.compile_expression(name)?;
                        let line = self.line;
                        common::dict::emit_get_dynamic(&mut self.chunks[self.current_chunk_idx], line);
                    }
                }
            }
            Expression::StaticAccess { class, member } => {
                self.compile_expression(class)?;
                match member.as_ref() {
                    Expression::Identifier(s) => {
                        let idx = self.add_string_constant(s);
                        self.emit_u16(Op::STRUCT_GET, idx);
                    }
                    Expression::Variable(s) => {
                        let idx = self.add_string_constant(s);
                        self.emit_u16(Op::STRUCT_GET, idx);
                    }
                    _ => {
                        self.compile_expression(member)?;
                        self.emit(Op::ARRAY_GET);
                    }
                }
            }
            Expression::ArrayAccess { array, index } => {
                self.compile_expression(array)?;
                self.compile_expression(index)?;
                let line = self.line;
                common::dict::emit_get_dynamic(&mut self.chunks[self.current_chunk_idx], line);
            }
            Expression::Isset(vars) => {
                for (i, var) in vars.iter().enumerate() {
                    self.compile_expression(var)?;
                    self.emit(Op::REF_IS_NULL);
                    self.emit(Op::DYN_NOT);
                    if i > 0 {
                        self.emit(Op::I32_AND);
                    }
                }
                if vars.is_empty() {
                    self.emit(Op::TRUE);
                }
            }
            Expression::Empty(var) => {
                self.compile_expression(var)?;
                self.emit(Op::DYN_TO_BOOL);
                self.emit(Op::DYN_NOT);
            }
            Expression::Unset(_vars) => {
                self.emit(Op::NULL);
            }
            Expression::Cast { cast, expr } => {
                self.compile_expression(expr)?;
                match cast {
                    CastKind::Int => {
                        let idx = self.import("vybe:convert", "parseInt");
                        self.emit_host_call(idx, 1);
                    }
                    CastKind::Float => {
                        let idx = self.import("vybe:convert", "parseFloat");
                        self.emit_host_call(idx, 1);
                    }
                    CastKind::String => {
                        let idx = self.import("vybe:convert", "toString");
                        self.emit_host_call(idx, 1);
                    }
                    CastKind::Bool => {
                        self.emit(Op::DYN_TO_BOOL);
                    }
                    CastKind::Array => {
                        // Wrap in a single-element array
                        self.emit(Op::ARRAY_NEW);
                    }
                    CastKind::Object => {
                        // Already an object at runtime; no-op
                    }
                }
            }
            Expression::Closure { params, uses, body, is_arrow: _ } => {
                // Compile closure body — `use` vars are resolved as upvalues
                // by resolve_upvalue() during body compilation
                let decl = FunctionDecl {
                    name: "<closure>".to_string(),
                    params: params.clone(),
                    body: match body.as_ref() {
                        ClosureBody::Block(stmts) => stmts.clone(),
                        ClosureBody::Expr(e) => vec![Statement::Return(Some(e.clone()))],
                    },
                    is_static: false,
                    visibility: Visibility::None,
                    return_by_ref: false,
                };

                // For `use ($x, $y)` — ensure those vars are visible in parent scope
                // so resolve_upvalue finds them. If they're not already locals, define them.
                if !uses.is_empty() && !self.is_global_scope() {
                    for use_var in uses {
                        if self.current_scope().resolve_local(use_var).is_none() {
                            self.define_local(use_var);
                        }
                    }
                }

                let ci = self.compile_function(&decl)?;
                let upvalue_count = self.scopes.last().map(|s| s.upvalues.len()).unwrap_or(0);
                let line = self.line;
                common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], ci, upvalue_count as u8, line);
            }
            Expression::Match { subject, arms } => {
                self.compile_expression(subject)?;
                let disc_slot = self.define_local("__match_disc");
                self.emit_u16(Op::LOCAL_SET, disc_slot);
                let mut end_jumps = Vec::new();
                let mut next_arm_jump: Option<usize> = None;
                for arm in arms {
                    if let Some(nj) = next_arm_jump.take() { self.patch_jump(nj); }
                    if let Some(conds) = &arm.conditions {
                        let mut skip_jumps = Vec::new();
                        let mut fail_jumps = Vec::new();
                        for cond in conds {
                            self.emit_u16(Op::LOCAL_GET, disc_slot);
                            self.compile_expression(cond)?;
                            self.emit(Op::EQ);
                            skip_jumps.push(self.emit_jump(Op::BR_IF_TRUE));
                        }
                        fail_jumps.push(self.emit_jump(Op::BR));
                        for s in &skip_jumps { self.patch_jump(*s); }
                        self.compile_expression(&arm.body)?;
                        end_jumps.push(self.emit_jump(Op::BR));
                        for f in fail_jumps { self.patch_jump(f); }
                        next_arm_jump = None;
                    } else {
                        self.compile_expression(&arm.body)?;
                        end_jumps.push(self.emit_jump(Op::BR));
                    }
                }
                if let Some(nj) = next_arm_jump { self.patch_jump(nj); }
                self.emit(Op::NULL);
                for j in end_jumps { self.patch_jump(j); }
            }
            Expression::Spread(inner) => {
                self.compile_expression(inner)?;
            }
            Expression::List(_) => {
                self.emit(Op::NULL);
            }
            Expression::Yield(value) => {
                // yield $value → suspend opcode (returns value to caller, pauses generator)
                if let Some(val) = value {
                    self.compile_expression(val)?;
                } else {
                    self.emit(Op::NULL);
                }
                self.emit_u16(Op::SUSPEND, 0); // tag 0
            }
            Expression::YieldFrom(expr) => {
                // yield from $generator → delegate to sub-generator
                // For now: just evaluate the expression (simplified)
                self.compile_expression(expr)?;
            }
        }
        Ok(())
    }

    // ------------------------------------------------------------------
    // Assignment
    // ------------------------------------------------------------------

    fn compile_assign(&mut self, op: &AssignOp, left: &Expression, right: &Expression) -> Result<(), String> {
        if *op == AssignOp::Assign {
            self.compile_expression(right)?;
            self.emit(Op::DUP);
            self.compile_assign_lhs(left)?;
            return Ok(());
        }
        // ??= : $x ??= val → if $x is null, $x = val
        if *op == AssignOp::NullCoalesceAssign {
            let c = self.current_chunk_idx;
            let line = self.line;
            self.compile_expression(left)?;
            let (_null_jump, end_jump) = common::expressions::emit_null_coalesce_start(&mut self.chunks[c], line);
            self.compile_expression(right)?;
            common::expressions::emit_null_coalesce_end(&mut self.chunks[c], end_jump);
            self.emit(Op::DUP);
            self.compile_assign_lhs(left)?;
            return Ok(());
        }
        self.compile_expression(left)?;
        self.compile_expression(right)?;
        match op {
            AssignOp::AddAssign => { self.emit(Op::DYN_ADD); }
            AssignOp::SubAssign => { self.emit(Op::F64_SUB); }
            AssignOp::MulAssign => { self.emit(Op::F64_MUL); }
            AssignOp::DivAssign => { self.emit(Op::F64_DIV); }
            AssignOp::ModAssign => { { let l = self.line; vybe_compiler_common::expressions::emit_f64_mod(&mut self.chunks[self.current_chunk_idx], l); }; }
            AssignOp::PowAssign => {
                let c = self.current_chunk_idx;
                let line = self.line;
                common::math::emit_pow(&mut self.chunks[c], line);
            }
            AssignOp::ConcatAssign => { common_strings::emit_str_concat(&mut self.chunks[self.current_chunk_idx], self.line); }
            AssignOp::AndAssign => { self.emit(Op::I32_AND); }
            AssignOp::OrAssign => { self.emit(Op::I32_OR); }
            AssignOp::XorAssign => { self.emit(Op::I32_XOR); }
            AssignOp::ShlAssign => { self.emit(Op::I32_SHL); }
            AssignOp::ShrAssign => { self.emit(Op::I32_SHR_S); }
            AssignOp::NullCoalesceAssign => {}
            AssignOp::Assign => unreachable!(),
        }
        self.emit(Op::DUP);
        self.compile_assign_lhs(left)?;
        Ok(())
    }

    /// Write TOS into lhs (consumes value). Stack: [val] → []
    /// Uses a temp local to reorder stack for struct_set/array_set.
    fn compile_assign_lhs(&mut self, lhs: &Expression) -> Result<(), String> {
        match lhs {
            Expression::Variable(name) => { self.emit_var_set(name); }
            Expression::Property { object, name, .. } => {
                // Stack: [val]. struct_set expects [obj, val].
                // Save val to temp, push obj, push val back.
                let tmp = self.define_local("__store_tmp");
                self.emit_u16(Op::LOCAL_SET, tmp);
                self.emit(Op::DROP);
                self.compile_expression(object)?;
                self.emit_u16(Op::LOCAL_GET, tmp);
                match name.as_ref() {
                    Expression::Identifier(s) => {
                        let idx = self.add_string_constant(s);
                        self.emit_u16(Op::STRUCT_SET, idx);
                        self.emit(Op::DROP);
                    }
                    _ => {
                        // dynamic property: use array_set [obj, key, val]
                        // Stack: [obj, val]. Need [obj, key, val].
                        // Save val again, push key, push val.
                        let tmp2 = self.define_local("__store_tmp2");
                        self.emit_u16(Op::LOCAL_SET, tmp2);
                        self.emit(Op::DROP);
                        self.compile_expression(name)?;
                        self.emit_u16(Op::LOCAL_GET, tmp2);
                        self.emit(Op::ARRAY_SET);
                        self.emit(Op::DROP);
                    }
                }
            }
            Expression::ArrayAccess { array, index } => {
                // Stack: [val]. array_set expects [obj, key, val].
                let tmp = self.define_local("__store_tmp");
                self.emit_u16(Op::LOCAL_SET, tmp);
                self.emit(Op::DROP);
                self.compile_expression(array)?;
                self.compile_expression(index)?;
                self.emit_u16(Op::LOCAL_GET, tmp);
                let line = self.line;
                common::dict::emit_set_dynamic(&mut self.chunks[self.current_chunk_idx], line);
            }
            Expression::List(targets) => {
                // list($a, $b) = [1, 2]
                let arr_tmp = self.define_local("__list_arr");
                self.emit_u16(Op::LOCAL_SET, arr_tmp);
                for (i, target) in targets.iter().enumerate() {
                    if let Some(t) = target {
                        self.emit_u16(Op::LOCAL_GET, arr_tmp);
                        self.emit_constant(Value::I32(i as i32));
                        self.emit(Op::ARRAY_GET);
                        match t {
                            Expression::Variable(name) => { self.emit_var_set(name); }
                            _ => { self.emit(Op::DROP); }
                        }
                    }
                }
            }
            Expression::Array(elements) => {
                // [$a, $b, $c] = [10, 20, 30] — short destructuring syntax
                let arr_tmp = self.define_local("__destruct_arr");
                self.emit_u16(Op::LOCAL_SET, arr_tmp);
                for (i, elem) in elements.iter().enumerate() {
                    self.emit_u16(Op::LOCAL_GET, arr_tmp);
                    let key = match &elem.key {
                        Some(Expression::Str(s)) => s.clone(),
                        Some(Expression::Number(n)) => n.to_string(),
                        _ => i.to_string(),
                    };
                    self.emit_constant(Value::String(Arc::from(key.as_str())));
                    self.emit(Op::ARRAY_GET);
                    match &elem.value {
                        Expression::Variable(name) => { self.emit_var_set(name); }
                        _ => { self.emit(Op::DROP); }
                    }
                }
            }
            Expression::StaticAccess { class, member } => {
                // ClassName::$prop = val — set property on class constructor object
                let tmp = self.define_local("__static_tmp");
                self.emit_u16(Op::LOCAL_SET, tmp);
                self.compile_expression(class)?;
                self.emit_u16(Op::LOCAL_GET, tmp);
                match member.as_ref() {
                    Expression::Identifier(s) | Expression::Variable(s) => {
                        let idx = self.add_string_constant(s);
                        self.emit_u16(Op::STRUCT_SET, idx);
                        self.emit(Op::DROP);
                    }
                    _ => {
                        self.compile_expression(member)?;
                        self.emit(Op::ARRAY_SET);
                        self.emit(Op::DROP);
                    }
                }
            }
            _ => {
                return Err("invalid assignment target".to_string());
            }
        }
        Ok(())
    }


    // ------------------------------------------------------------------
    // Binary operators
    // ------------------------------------------------------------------

    fn compile_binary(&mut self, op: &BinaryOp, left: &Expression, right: &Expression) -> Result<(), String> {
        let c = self.current_chunk_idx;
        let line = self.line;

        if *op == BinaryOp::And {
            self.compile_expression(left)?;
            let skip = common::expressions::emit_and_start(&mut self.chunks[c], line);
            self.compile_expression(right)?;
            common::expressions::emit_short_circuit_end(&mut self.chunks[c], skip);
            return Ok(());
        }
        if *op == BinaryOp::Or {
            self.compile_expression(left)?;
            let skip = common::expressions::emit_or_start(&mut self.chunks[c], line);
            self.compile_expression(right)?;
            common::expressions::emit_short_circuit_end(&mut self.chunks[c], skip);
            return Ok(());
        }

        self.compile_expression(left)?;
        self.compile_expression(right)?;
        match op {
            // PHP + is type-juggling: "5" + 3 = 8, [1,2] + [3] = [1,2,3]
            BinaryOp::Add => { self.emit(Op::DYN_ADD); }
            BinaryOp::Sub => { self.emit(Op::F64_SUB); }
            BinaryOp::Mul => { self.emit(Op::F64_MUL); }
            BinaryOp::Div => { self.emit(Op::F64_DIV); }
            BinaryOp::Mod => { { let l = self.line; vybe_compiler_common::expressions::emit_f64_mod(&mut self.chunks[self.current_chunk_idx], l); }; }
            BinaryOp::Pow => {
                common::math::emit_pow(&mut self.chunks[c], line);
            }
            BinaryOp::Concat => { common_strings::emit_str_concat(&mut self.chunks[self.current_chunk_idx], self.line); }
            BinaryOp::Eq => { self.emit(Op::DYN_EQ); }
            BinaryOp::Ne => { self.emit(Op::DYN_NE); }
            BinaryOp::SEq => { self.emit(Op::EQ); }
            BinaryOp::SNe => { self.emit(Op::NE); }
            BinaryOp::Lt => { self.emit(Op::DYN_LT); }
            BinaryOp::Gt => { self.emit(Op::DYN_GT); }
            BinaryOp::Le => { self.emit(Op::DYN_LE); }
            BinaryOp::Ge => { self.emit(Op::DYN_GE); }
            BinaryOp::Spaceship => {
                // Inline: (a < b) ? -1 : ((a > b) ? 1 : 0)
                // a and b already on stack, use temp locals
                let b_tmp = self.define_local("__cmp_b");
                let a_tmp = self.define_local("__cmp_a");
                self.emit_u16(Op::LOCAL_SET, b_tmp);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_SET, a_tmp);
                self.emit(Op::DROP);
                // a < b?
                self.emit_u16(Op::LOCAL_GET, a_tmp);
                self.emit_u16(Op::LOCAL_GET, b_tmp);
                self.emit(Op::DYN_LT);
                let lt_jump = self.emit_jump(Op::BR_IF_TRUE);
                // a > b?
                self.emit_u16(Op::LOCAL_GET, a_tmp);
                self.emit_u16(Op::LOCAL_GET, b_tmp);
                self.emit(Op::DYN_GT);
                let gt_jump = self.emit_jump(Op::BR_IF_TRUE);
                // equal → 0
                self.emit_constant(Value::F64(0.0));
                let end_jump = self.emit_jump(Op::BR);
                // a > b → 1
                self.patch_jump(gt_jump);
                self.emit_constant(Value::F64(1.0));
                let end_jump2 = self.emit_jump(Op::BR);
                // a < b → -1
                self.patch_jump(lt_jump);
                self.emit_constant(Value::F64(-1.0));
                self.patch_jump(end_jump);
                self.patch_jump(end_jump2);
            }
            BinaryOp::BitAnd => { self.emit(Op::I32_AND); }
            BinaryOp::BitOr => { self.emit(Op::I32_OR); }
            BinaryOp::BitXor => { self.emit(Op::I32_XOR); }
            BinaryOp::Shl => { self.emit(Op::I32_SHL); }
            BinaryOp::Shr => { self.emit(Op::I32_SHR_S); }
            BinaryOp::InstanceOf => { self.emit(Op::REF_TEST); }
            BinaryOp::And | BinaryOp::Or => unreachable!(),
        }
        Ok(())
    }

    // ------------------------------------------------------------------
    // Function calls
    // ------------------------------------------------------------------

    fn compile_call(&mut self, callee: &Expression, args: &[Argument]) -> Result<(), String> {
        if let Expression::Identifier(name) = callee {
            if let Some(result) = self.try_compile_builtin(name, args)? {
                return Ok(result);
            }
            let lower = name.to_lowercase();
            // Check if this name resolves to a local, upvalue, or known global/class
            let is_local = self.resolve_var(&lower).is_some();
            let is_upvalue = self.scopes.len() > 1
                && self.resolve_upvalue(self.scopes.len() - 1, &lower).is_some();
            let is_defined = is_local || is_upvalue
                || self.defined_globals.contains(&lower)
                || self.defined_classes.contains(&lower);
            if is_defined {
                // Known local/global — use global_get + call_ref
                self.compile_expression(callee)?;
                for arg in args { self.compile_expression(&arg.value)?; }
                self.emit_u8(Op::CALL_REF, args.len() as u8);
            } else {
                // Unresolved → WASM import
                let idx = self.import("*", &lower);
                for arg in args { self.compile_expression(&arg.value)?; }
                self.emit_host_call(idx, args.len() as u8);
            }
            return Ok(());
        }
        self.compile_expression(callee)?;
        for arg in args { self.compile_expression(&arg.value)?; }
        self.emit_u8(Op::CALL_REF, args.len() as u8);
        Ok(())
    }

    fn compile_method_call(&mut self, object: &Expression, method: &Expression, args: &[Argument], _nullsafe: bool) -> Result<(), String> {
        // Special: Fiber methods — $fiber->start(), $fiber->resume($val), $fiber->isStarted(), etc.
        if let Expression::Identifier(method_name) = method {
            match method_name.as_str() {
                "start" => {
                    // $fiber->start(args...) → resume(continuation, args_or_null)
                    // Multiple args: pack into array for the continuation to unpack
                    self.compile_expression(object)?; // continuation
                    if args.is_empty() {
                        self.emit(Op::NULL);
                    } else if args.len() == 1 {
                        self.compile_expression(&args[0].value)?;
                    } else {
                        // Pack multiple args into an array
                        for arg in args { self.compile_expression(&arg.value)?; }
                        self.emit_u16(Op::ARRAY_NEW, args.len() as u16);
                    }
                    self.emit_u16(Op::RESUME, 0);
                    return Ok(());
                }
                "resume" => {
                    // $fiber->resume($value) → resume(continuation, value)
                    self.compile_expression(object)?; // continuation
                    if args.is_empty() {
                        self.emit(Op::NULL);
                    } else {
                        self.compile_expression(&args[0].value)?;
                    }
                    self.emit_u16(Op::RESUME, 0);
                    return Ok(());
                }
                "isStarted" | "isRunning" | "isTerminated" | "isSuspended" => {
                    // Read state from continuation object
                    self.compile_expression(object)?;
                    let state_key = self.add_string_constant("__cont_state");
                    self.emit_u16(Op::STRUCT_GET, state_key);
                    let expected = match method_name.as_str() {
                        "isStarted" => "running",
                        "isRunning" => "running",
                        "isTerminated" => "done",
                        "isSuspended" => "suspended",
                        _ => "unknown",
                    };
                    self.emit_constant(Value::String(Arc::from(expected)));
                    self.emit(Op::DYN_EQ);
                    return Ok(());
                }
                "getReturn" => {
                    // Read the return value from continuation
                    self.compile_expression(object)?;
                    let val_key = self.add_string_constant("__cont_value");
                    self.emit_u16(Op::STRUCT_GET, val_key);
                    return Ok(());
                }
                // ── StringBuilder methods (same as VB/C#) ──────────
                "append" => {
                    self.compile_expression(object)?;
                    for arg in args { self.compile_expression(&arg.value)?; }
                    let i = self.import("vybe:types", "sbAppend");
                    self.emit_host_call(i, (args.len() + 1) as u8);
                    return Ok(());
                }
                "appendLine" | "appendline" => {
                    self.compile_expression(object)?;
                    for arg in args { self.compile_expression(&arg.value)?; }
                    let i = self.import("vybe:types", "sbAppendLine");
                    self.emit_host_call(i, (args.len() + 1) as u8);
                    return Ok(());
                }
                // toString for StringBuilder (and general cross-language compat)
                "toString" | "tostring" | "__toString" | "__tostring" => {
                    self.compile_expression(object)?;
                    let i = self.import("vybe:types", "sbToString");
                    self.emit_host_call(i, 1);
                    return Ok(());
                }
                "clear" => {
                    self.compile_expression(object)?;
                    let i = self.import("vybe:types", "sbClear");
                    self.emit_host_call(i, 1);
                    return Ok(());
                }
                "insert" => {
                    self.compile_expression(object)?;
                    for arg in args { self.compile_expression(&arg.value)?; }
                    let i = self.import("vybe:types", "sbInsert");
                    self.emit_host_call(i, (args.len() + 1) as u8);
                    return Ok(());
                }
                "replace" => {
                    self.compile_expression(object)?;
                    for arg in args { self.compile_expression(&arg.value)?; }
                    let i = self.import("vybe:types", "sbReplace");
                    self.emit_host_call(i, (args.len() + 1) as u8);
                    return Ok(());
                }
                // ── HashSet methods (same as VB/C#) ─────────────────
                "add" => {
                    self.compile_expression(object)?;
                    for arg in args { self.compile_expression(&arg.value)?; }
                    let i = self.import("vybe:types", "hashSetAdd");
                    self.emit_host_call(i, (args.len() + 1) as u8);
                    return Ok(());
                }
                "has" | "contains" => {
                    self.compile_expression(object)?;
                    for arg in args { self.compile_expression(&arg.value)?; }
                    let i = self.import("vybe:types", "hashSetContains");
                    self.emit_host_call(i, (args.len() + 1) as u8);
                    return Ok(());
                }
                "delete" | "remove" => {
                    self.compile_expression(object)?;
                    for arg in args { self.compile_expression(&arg.value)?; }
                    let i = self.import("vybe:types", "hashSetRemove");
                    self.emit_host_call(i, (args.len() + 1) as u8);
                    return Ok(());
                }
                // ── Random methods (same as VB/C#) ──────────────────
                "nextInt" | "next" => {
                    self.compile_expression(object)?;
                    for arg in args { self.compile_expression(&arg.value)?; }
                    let i = self.import("vybe:threading", "randomNext");
                    self.emit_host_call(i, (args.len() + 1) as u8);
                    return Ok(());
                }
                "nextFloat" | "nextDouble" => {
                    self.compile_expression(object)?;
                    let i = self.import("vybe:threading", "randomNextDouble");
                    self.emit_host_call(i, 1);
                    return Ok(());
                }
                // ── Stopwatch methods ────────────────────────────────
                "elapsed" | "getElapsed" => {
                    self.compile_expression(object)?;
                    let i = self.import("vybe:threading", "stopwatchElapsed");
                    self.emit_host_call(i, 1);
                    return Ok(());
                }
                // ── DateTime methods ─────────────────────────────
                "format" => {
                    self.compile_expression(object)?;
                    for arg in args { self.compile_expression(&arg.value)?; }
                    let i = self.import("vybe:types", "dateTimeToString");
                    self.emit_host_call(i, (args.len() + 1) as u8);
                    return Ok(());
                }
                "modify" | "sub" => {
                    // $dt->modify('+1 day') — simplified
                    self.compile_expression(object)?;
                    return Ok(());
                }
                "getTimestamp" => {
                    self.compile_expression(object)?;
                    // DateTime objects store timestamp — just return it
                    return Ok(());
                }
                // ── SplStack / SplQueue methods ─────────────────────
                "push" => {
                    self.compile_expression(object)?;
                    for arg in args { self.compile_expression(&arg.value)?; }
                    let i = self.import("vybe:types", "stackPush");
                    self.emit_host_call(i, (args.len() + 1) as u8);
                    return Ok(());
                }
                "pop" => {
                    self.compile_expression(object)?;
                    let i = self.import("vybe:types", "stackPop");
                    self.emit_host_call(i, 1);
                    return Ok(());
                }
                "enqueue" => {
                    self.compile_expression(object)?;
                    for arg in args { self.compile_expression(&arg.value)?; }
                    let i = self.import("vybe:types", "queueEnqueue");
                    self.emit_host_call(i, (args.len() + 1) as u8);
                    return Ok(());
                }
                "dequeue" => {
                    self.compile_expression(object)?;
                    let i = self.import("vybe:types", "queueDequeue");
                    self.emit_host_call(i, 1);
                    return Ok(());
                }
                "peek" | "top" | "bottom" => {
                    self.compile_expression(object)?;
                    let i = self.import("vybe:types", "queuePeek");
                    self.emit_host_call(i, 1);
                    return Ok(());
                }
                // ── PDO / database methods ────────────────────────
                // $pdo->query($sql) → vybe:database query(conn, sql)
                "query" => {
                    self.compile_expression(object)?; // conn object
                    for arg in args { self.compile_expression(&arg.value)?; }
                    let idx = self.import("vybe:database", "query");
                    self.emit_host_call(idx, (args.len() + 1) as u8);
                    return Ok(());
                }
                // $pdo->exec($sql) → vybe:database execute(conn, sql)
                "exec" | "execute" => {
                    self.compile_expression(object)?;
                    for arg in args { self.compile_expression(&arg.value)?; }
                    let idx = self.import("vybe:database", "execute");
                    self.emit_host_call(idx, (args.len() + 1) as u8);
                    return Ok(());
                }
                // $pdo->prepare($sql) → create a statement object with conn + sql
                "prepare" => {
                    // Build a statement object: { __conn: pdo, __sql: sql, __params: [] }
                    let line = self.line;
                    let c = self.current_chunk_idx;
                    common::dict::emit_new(&mut self.chunks[c], line);
                    self.emit(Op::DUP);
                    self.compile_expression(object)?;
                    let conn_key = self.add_string_constant("__conn");
                    self.emit_u16(Op::STRUCT_SET, conn_key);
                    self.emit(Op::DROP);
                    self.emit(Op::DUP);
                    if !args.is_empty() { self.compile_expression(&args[0].value)?; }
                    else { self.emit(Op::NULL); }
                    let sql_key = self.add_string_constant("__sql");
                    self.emit_u16(Op::STRUCT_SET, sql_key);
                    self.emit(Op::DROP);
                    return Ok(());
                }
                // $stmt->fetch() → query(conn, sql) then return first row
                "fetch" => {
                    // Get __conn and __sql from statement object
                    self.compile_expression(object)?;
                    let conn_key = self.add_string_constant("__conn");
                    self.emit(Op::DUP);
                    self.emit_u16(Op::STRUCT_GET, conn_key);
                    // Stack: [stmt, conn]
                    let conn_tmp = self.define_local("__fetch_conn");
                    self.emit_u16(Op::LOCAL_SET, conn_tmp);
                    let sql_key = self.add_string_constant("__sql");
                    self.emit_u16(Op::STRUCT_GET, sql_key);
                    // Stack: [sql]. Call scalar(conn, sql)
                    let scalar_fn = self.import("vybe:database", "scalar");
                    self.emit_u16(Op::LOCAL_GET, conn_tmp);
                    // Need: [conn, sql] for host call — reorder
                    let sql_tmp = self.define_local("__fetch_sql");
                    self.emit_u16(Op::LOCAL_SET, sql_tmp);
                    self.emit(Op::DROP); // drop local_get result
                    self.emit_u16(Op::LOCAL_GET, conn_tmp);
                    self.emit_u16(Op::LOCAL_GET, sql_tmp);
                    self.emit_host_call(scalar_fn, 2);
                    return Ok(());
                }
                // $stmt->fetchAll() → query(conn, sql)
                "fetchAll" | "fetchall" => {
                    self.compile_expression(object)?;
                    let conn_key = self.add_string_constant("__conn");
                    self.emit(Op::DUP);
                    self.emit_u16(Op::STRUCT_GET, conn_key);
                    let conn_tmp = self.define_local("__fetchall_conn");
                    self.emit_u16(Op::LOCAL_SET, conn_tmp);
                    let sql_key = self.add_string_constant("__sql");
                    self.emit_u16(Op::STRUCT_GET, sql_key);
                    let sql_tmp = self.define_local("__fetchall_sql");
                    self.emit_u16(Op::LOCAL_SET, sql_tmp);
                    self.emit_u16(Op::LOCAL_GET, conn_tmp);
                    self.emit_u16(Op::LOCAL_GET, sql_tmp);
                    let query_fn = self.import("vybe:database", "query");
                    self.emit_host_call(query_fn, 2);
                    return Ok(());
                }
                // $pdo->lastInsertId() — not yet, return 0
                "lastInsertId" | "lastinsertid" => {
                    self.emit_constant(Value::F64(0.0));
                    return Ok(());
                }
                // $pdo->beginTransaction / commit / rollBack
                "beginTransaction" | "begintransaction" => {
                    self.compile_expression(object)?;
                    self.emit_constant(Value::String(Arc::from("BEGIN")));
                    let idx = self.import("vybe:database", "execute");
                    self.emit_host_call(idx, 2);
                    return Ok(());
                }
                "commit" => {
                    self.compile_expression(object)?;
                    self.emit_constant(Value::String(Arc::from("COMMIT")));
                    let idx = self.import("vybe:database", "execute");
                    self.emit_host_call(idx, 2);
                    return Ok(());
                }
                "rollBack" | "rollback" => {
                    self.compile_expression(object)?;
                    self.emit_constant(Value::String(Arc::from("ROLLBACK")));
                    let idx = self.import("vybe:database", "execute");
                    self.emit_host_call(idx, 2);
                    return Ok(());
                }
                _ => {} // fall through to generic method call
            }
        }

        // Generic method call: obj.method(args...)
        self.compile_expression(object)?;
        match method {
            Expression::Identifier(name) => {
                let prop_idx = self.add_string_constant(&name);
                self.emit_u16(Op::STRUCT_GET, prop_idx);
            }
            _ => {
                self.compile_expression(method)?;
                self.emit(Op::ARRAY_GET);
            }
        }
        // call_ref expects: [func_ref, args...]. First arg is `this`.
        self.compile_expression(object)?;
        for arg in args { self.compile_expression(&arg.value)?; }
        self.emit_u8(Op::CALL_REF, (args.len() + 1) as u8);
        Ok(())
    }

    fn compile_static_call(&mut self, class: &Expression, method: &Expression, args: &[Argument]) -> Result<(), String> {
        // Special: Fiber::suspend($value) → suspend opcode
        if let Expression::Identifier(class_name) = class {
            if class_name == "Fiber" {
                if let Expression::Identifier(method_name) = method {
                    if method_name == "suspend" {
                        if args.is_empty() {
                            self.emit(Op::NULL);
                        } else {
                            self.compile_expression(&args[0].value)?;
                        }
                        self.emit_u16(Op::SUSPEND, 0);
                        return Ok(());
                    }
                }
            }
        }

        // parent::__construct(args) — call parent constructor, store result as $this
        if let Expression::ClassKeyword(kw) = class {
            if kw == "parent" {
                if let Expression::Identifier(method_name) = method {
                    if method_name == "__construct" {
                        if let Some(ref parent) = self.current_class_parent.clone() {
                            let idx = self.add_string_constant(parent);
                            self.emit_u16(Op::GLOBAL_GET, idx);
                            for arg in args { self.compile_expression(&arg.value)?; }
                            self.emit_u8(Op::CALL_REF, args.len() as u8);
                            // Store returned object as $this
                            if let Some(this_slot) = self.current_scope().resolve_local("this") {
                                self.emit(Op::DUP);
                                self.emit_u16(Op::LOCAL_SET, this_slot);
                            }
                            return Ok(());
                        }
                    }
                }
            }
        }

        // Generic static call: Class::method(args...)
        self.compile_expression(class)?;
        match method {
            Expression::Identifier(name) => {
                let idx = self.add_string_constant(&name);
                self.emit_u16(Op::STRUCT_GET, idx);
            }
            _ => {
                self.compile_expression(method)?;
                self.emit(Op::ARRAY_GET);
            }
        }
        for arg in args { self.compile_expression(&arg.value)?; }
        self.emit_u8(Op::CALL_REF, args.len() as u8);
        Ok(())
    }

    // ------------------------------------------------------------------
    // PHP built-in functions → host imports
    // ------------------------------------------------------------------

    fn try_compile_builtin(&mut self, name: &str, args: &[Argument]) -> Result<Option<()>, String> {
        let lower = name.to_lowercase();
        let c = self.current_chunk_idx;
        let line = self.line;

        // Helper: compile all args onto stack
        macro_rules! compile_args { () => { for arg in args { self.compile_expression(&arg.value)?; } } }

        match lower.as_str() {
            // ── IO (common::io) ─────────────────────────────────────
            "echo" | "print" | "var_dump" | "print_r" => {
                compile_args!();
                common::io::emit_print(&mut self.chunks[c], args.len() as u8, line);
                return Ok(Some(()));
            }
            "die" | "exit" => {
                if !args.is_empty() {
                    compile_args!();
                    common::io::emit_print(&mut self.chunks[c], args.len() as u8, line);
                    self.emit(Op::DROP);
                }
                self.emit(Op::HALT);
                return Ok(Some(()));
            }
            "__clone" => {
                // clone $obj — create new object, copy properties
                compile_args!(); // source obj on stack
                let src_slot = self.define_local("__clone_src");
                self.emit_u16(Op::LOCAL_SET, src_slot);
                // Create empty object
                self.emit_u16(Op::STRUCT_NEW, 0);
                let dest_slot = self.define_local("__clone_dest");
                self.emit_u16(Op::LOCAL_SET, dest_slot);
                // assign(dest, src)
                common::bundle::emit_call_push_func(&mut self.chunks[c], "__vybe_assign", line);
                self.emit_u16(Op::LOCAL_GET, dest_slot);
                self.emit_u16(Op::LOCAL_GET, src_slot);
                common::bundle::emit_call_invoke(&mut self.chunks[c], 2, line);
                return Ok(Some(()));
            }
            "__throw" => {
                // throw as expression — compile arg and emit throw opcode
                compile_args!();
                self.emit(Op::THROW);
                self.emit(Op::NULL); // unreachable, but keeps stack balanced
                return Ok(Some(()));
            }

            // ── String opcodes (direct VM ops) ──────────────────────
            "strlen"       => { compile_args!(); common_strings::emit_length(&mut self.chunks[c], line); return Ok(Some(())); }
            "strtolower"   => { compile_args!(); common_strings::emit_to_lower(&mut self.chunks[c], line); return Ok(Some(())); }
            "strtoupper"   => { compile_args!(); common_strings::emit_to_upper(&mut self.chunks[c], line); return Ok(Some(())); }
            "trim"         => { compile_args!(); common_strings::emit_trim(&mut self.chunks[c], line); return Ok(Some(())); }
            "ltrim"        => { compile_args!(); common_strings::emit_trim_start(&mut self.chunks[c], line); return Ok(Some(())); }
            "rtrim"        => { compile_args!(); common_strings::emit_trim_end(&mut self.chunks[c], line); return Ok(Some(())); }
            "strpos" | "stripos"   => { compile_args!(); common_strings::emit_index_of(&mut self.chunks[c], line); return Ok(Some(())); }
            "str_contains" => { compile_args!(); self.emit(Op::STR_CONTAINS); return Ok(Some(())); }
            "str_starts_with" => { compile_args!(); self.emit(Op::STR_STARTS_WITH); return Ok(Some(())); }
            "str_ends_with"   => { compile_args!(); self.emit(Op::STR_ENDS_WITH); return Ok(Some(())); }
            "str_replace"  => {
                // str_replace(search, replace, subject) — PHP arg order: [search, replace, subject]
                // str_replace opcode expects [subject, search, replace] — reorder
                if args.len() >= 3 {
                    self.compile_expression(&args[2].value)?; // subject
                    self.compile_expression(&args[0].value)?; // search
                    self.compile_expression(&args[1].value)?; // replace
                } else { compile_args!(); }
                common_strings::emit_replace(&mut self.chunks[self.current_chunk_idx], self.line);
                return Ok(Some(()));
            }
            "substr" => {
                // substr(string, start[, length]) → str_substring opcode [str, start, end]
                if args.len() >= 1 { self.compile_expression(&args[0].value)?; }
                if args.len() >= 2 { self.compile_expression(&args[1].value)?; }
                    else { self.emit_constant(Value::I32(0)); }
                if args.len() >= 3 { self.compile_expression(&args[2].value)?; }
                    else { self.emit_constant(Value::I32(i32::MAX)); }
                common_strings::emit_substring(&mut self.chunks[self.current_chunk_idx], self.line);
                return Ok(Some(()));
            }
            "explode" | "str_split" => { compile_args!(); common_strings::emit_split(&mut self.chunks[c], line); return Ok(Some(())); }
            "str_repeat"   => { compile_args!(); common_strings::emit_repeat(&mut self.chunks[c], line); return Ok(Some(())); }
            "str_pad"      => { compile_args!(); self.emit(Op::STR_PAD_START); return Ok(Some(())); }
            "chr"          => { compile_args!(); self.emit(Op::STR_FROM_CHAR_CODE); return Ok(Some(())); }
            "ord"          => {
                compile_args!();
                let idx = self.import("vybe:string", "charCodeAt");
                self.emit_host_call(idx, 1);
                return Ok(Some(()));
            }
            "implode" | "join" => {
                // implode(glue, array) — array_join opcode expects [array, glue]
                if args.len() >= 2 {
                    self.compile_expression(&args[1].value)?; // array
                    self.compile_expression(&args[0].value)?; // glue
                } else { compile_args!(); }
                let line = self.line;
                common::collections::emit_join(&mut self.chunks[c], line);
                return Ok(Some(()));
            }
            "nl2br" => {
                // str_replace("\n", "<br />\n", $str) — inline using str_replace opcode
                compile_args!(); // subject on stack
                self.emit_constant(Value::String(Arc::from("\n")));
                self.emit_constant(Value::String(Arc::from("<br />\n")));
                common_strings::emit_replace(&mut self.chunks[self.current_chunk_idx], self.line);
                return Ok(Some(()));
            }
            "htmlspecialchars" | "htmlentities" => {
                // Chain of str_replace: & → &amp; < → &lt; > → &gt; " → &quot;
                compile_args!();
                for (from, to) in &[("&", "&amp;"), ("<", "&lt;"), (">", "&gt;"), ("\"", "&quot;")] {
                    self.emit_constant(Value::String(Arc::from(*from)));
                    self.emit_constant(Value::String(Arc::from(*to)));
                    common_strings::emit_replace(&mut self.chunks[self.current_chunk_idx], self.line);
                }
                return Ok(Some(()));
            }
            "ucfirst" => {
                // substr(0,1).toUpper . substr(1) — inline bytecode
                compile_args!();
                self.emit(Op::DUP);
                self.emit_constant(Value::I32(0));
                self.emit_constant(Value::I32(1));
                common_strings::emit_substring(&mut self.chunks[self.current_chunk_idx], self.line);
                common_strings::emit_to_upper(&mut self.chunks[self.current_chunk_idx], self.line);
                // swap: [orig, upper_first]
                let tmp = self.define_local("__ucf_tmp");
                self.emit_u16(Op::LOCAL_SET, tmp);
                self.emit_constant(Value::I32(1));
                self.emit_constant(Value::I32(i32::MAX));
                common_strings::emit_substring(&mut self.chunks[self.current_chunk_idx], self.line);
                self.emit_u16(Op::LOCAL_GET, tmp);
                // swap order for concat: [upper_first, rest]
                let tmp2 = self.define_local("__ucf_tmp2");
                self.emit_u16(Op::LOCAL_SET, tmp2);
                // now stack: [rest], need [upper_first, rest]
                let tmp3 = self.define_local("__ucf_tmp3");
                self.emit_u16(Op::LOCAL_SET, tmp3);
                self.emit_u16(Op::LOCAL_GET, tmp2);
                self.emit_u16(Op::LOCAL_GET, tmp3);
                common_strings::emit_str_concat(&mut self.chunks[self.current_chunk_idx], self.line);
                return Ok(Some(()));
            }
            "lcfirst" => {
                compile_args!();
                self.emit(Op::DUP);
                self.emit_constant(Value::I32(0));
                self.emit_constant(Value::I32(1));
                common_strings::emit_substring(&mut self.chunks[self.current_chunk_idx], self.line);
                common_strings::emit_to_lower(&mut self.chunks[self.current_chunk_idx], self.line);
                let tmp = self.define_local("__lcf_tmp");
                self.emit_u16(Op::LOCAL_SET, tmp);
                self.emit_constant(Value::I32(1));
                self.emit_constant(Value::I32(i32::MAX));
                common_strings::emit_substring(&mut self.chunks[self.current_chunk_idx], self.line);
                self.emit_u16(Op::LOCAL_GET, tmp);
                let tmp2 = self.define_local("__lcf_tmp2");
                self.emit_u16(Op::LOCAL_SET, tmp2);
                let tmp3 = self.define_local("__lcf_tmp3");
                self.emit_u16(Op::LOCAL_SET, tmp3);
                self.emit_u16(Op::LOCAL_GET, tmp2);
                self.emit_u16(Op::LOCAL_GET, tmp3);
                common_strings::emit_str_concat(&mut self.chunks[self.current_chunk_idx], self.line);
                return Ok(Some(()));
            }

            // ── Math opcodes (common::math) ─────────────────────────
            "abs"   => { compile_args!(); common::math::emit_abs(&mut self.chunks[c], line); return Ok(Some(())); }
            "ceil"  => { compile_args!(); common::math::emit_ceil(&mut self.chunks[c], line); return Ok(Some(())); }
            "floor" => { compile_args!(); common::math::emit_floor(&mut self.chunks[c], line); return Ok(Some(())); }
            "round" => { compile_args!(); common::math::emit_round(&mut self.chunks[c], line); return Ok(Some(())); }
            "sqrt"  => { compile_args!(); common::math::emit_sqrt(&mut self.chunks[c], line); return Ok(Some(())); }
            "pow"   => { compile_args!(); common::math::emit_pow(&mut self.chunks[c], line); return Ok(Some(())); }
            "sin"   => { compile_args!(); common::math::emit_sin(&mut self.chunks[c], line); return Ok(Some(())); }
            "cos"   => { compile_args!(); common::math::emit_cos(&mut self.chunks[c], line); return Ok(Some(())); }
            "tan"   => { compile_args!(); common::math::emit_tan(&mut self.chunks[c], line); return Ok(Some(())); }
            "exp"   => { compile_args!(); common::math::emit_exp(&mut self.chunks[c], line); return Ok(Some(())); }
            "log"   => { compile_args!(); common::math::emit_log(&mut self.chunks[c], line); return Ok(Some(())); }
            "rand" | "mt_rand" => { common::math::emit_random(&mut self.chunks[c], line); return Ok(Some(())); }
            "max"   => { compile_args!(); common::collections::emit_max(&mut self.chunks[c], args.len() as u8, line); return Ok(Some(())); }
            "min"   => { compile_args!(); common::collections::emit_min(&mut self.chunks[c], args.len() as u8, line); return Ok(Some(())); }

            // ── Type conversion (common::convert + opcodes) ─────────
            "intval"   => { compile_args!(); common::convert::emit_parse_int(&mut self.chunks[c], line); return Ok(Some(())); }
            "floatval" | "doubleval" => { compile_args!(); common::convert::emit_parse_float(&mut self.chunks[c], line); return Ok(Some(())); }
            "strval"   => { compile_args!(); common::strings::emit_to_string(&mut self.chunks[c], line); return Ok(Some(())); }
            "boolval"  => { compile_args!(); common::convert::emit_to_bool(&mut self.chunks[c], line); return Ok(Some(())); }
            "is_numeric" => { compile_args!(); common::convert::emit_is_numeric(&mut self.chunks[c], line); return Ok(Some(())); }
            "is_null"  => { compile_args!(); self.emit(Op::REF_IS_NULL); return Ok(Some(())); }

            // ── Array opcodes (common::collections) ─────────────────
            "count" | "sizeof" => {
                // Use smart length — works on arrays, strings, and objects with __get_length
                compile_args!();
                let obj_slot = self.define_local("__count_obj");
                self.emit_u16(Op::LOCAL_SET, obj_slot);
                common::expressions::emit_smart_length(&mut self.chunks[c], obj_slot, line);
                return Ok(Some(()));
            }
            "array_push"    => { compile_args!(); common::collections::emit_push(&mut self.chunks[c], line); return Ok(Some(())); }
            "array_pop"     => { compile_args!(); common::collections::emit_pop(&mut self.chunks[c], line); return Ok(Some(())); }
            "array_shift"   => { compile_args!(); common::collections::emit_shift(&mut self.chunks[c], line); return Ok(Some(())); }
            "array_reverse" => { compile_args!(); common::collections::emit_reverse(&mut self.chunks[c], line); return Ok(Some(())); }
            "in_array"      => { compile_args!(); common::collections::emit_contains(&mut self.chunks[c], line); return Ok(Some(())); }
            "array_search"  => { compile_args!(); common::collections::emit_index_of(&mut self.chunks[c], line); return Ok(Some(())); }
            "array_slice"   => { compile_args!(); common::collections::emit_slice(&mut self.chunks[c], line); return Ok(Some(())); }
            "array_merge"   => { compile_args!(); common::collections::emit_concat(&mut self.chunks[c], line); return Ok(Some(())); }
            "range" => {
                // range(start, stop) or range(start, stop, step)
                // PHP range is inclusive, so we need start, stop+1, step
                common::bundle::emit_call_push_func(&mut self.chunks[c], "__vybe_range", line);
                compile_args!();
                // If only 2 args, add step=1
                if args.len() == 2 {
                    self.emit_constant(Value::I32(1));
                    common::bundle::emit_call_invoke(&mut self.chunks[c], 3, line);
                } else {
                    common::bundle::emit_call_invoke(&mut self.chunks[c], args.len() as u8, line);
                }
                return Ok(Some(()));
            }
            "sort" | "asort" | "arsort" | "rsort" => {
                compile_args!();
                common::bundle::emit_call_push_func(&mut self.chunks[c], "__vybe_sorted", line);
                // args already on stack before func ref — need to reorder
                // Actually: push func first, then arg. But arg is already on stack.
                // Use the pattern: save arg, push func, push arg back, call
                let tmp = self.define_local("__sort_tmp");
                self.emit_u16(Op::LOCAL_SET, tmp);
                common::bundle::emit_call_push_func(&mut self.chunks[c], "__vybe_sorted", line);
                self.emit_u16(Op::LOCAL_GET, tmp);
                common::bundle::emit_call_invoke(&mut self.chunks[c], 1, line);
                return Ok(Some(()));
            }
            "array_sum" => {
                compile_args!();
                let tmp = self.define_local("__sum_tmp");
                self.emit_u16(Op::LOCAL_SET, tmp);
                common::bundle::emit_call_push_func(&mut self.chunks[c], "__vybe_sum", line);
                self.emit_u16(Op::LOCAL_GET, tmp);
                common::bundle::emit_call_invoke(&mut self.chunks[c], 1, line);
                return Ok(Some(()));
            }

            // ── Array callback ops — inline bytecode loop + call_ref (like JS compiler) ──
            "array_map" => {
                // array_map(callback, array) — PHP: callback first
                // Uses common::loops::emit_map (same bytecode as JS/Python)
                if args.len() < 2 { return Ok(None); }
                self.compile_expression(&args[0].value)?; // callback
                let fn_slot = self.define_local("__map_fn");
                self.emit_u16(Op::LOCAL_SET, fn_slot);
                self.compile_expression(&args[1].value)?; // array
                let arr_slot = self.define_local("__map_arr");
                self.emit_u16(Op::LOCAL_SET, arr_slot);
                let res_slot = self.define_local("__map_res");
                let i_slot = self.define_local("__map_i");
                common::loops::emit_map(&mut self.chunks[c], fn_slot, arr_slot, res_slot, i_slot, line);
                return Ok(Some(()));
            }
            "array_filter" => {
                // Uses common::loops::emit_filter
                if args.is_empty() { return Ok(None); }
                self.compile_expression(&args[0].value)?; // array
                let arr_slot = self.define_local("__filt_arr");
                self.emit_u16(Op::LOCAL_SET, arr_slot);
                if args.len() >= 2 {
                    // With callback
                    self.compile_expression(&args[1].value)?;
                    let fn_slot = self.define_local("__filt_fn");
                    self.emit_u16(Op::LOCAL_SET, fn_slot);
                    let res_slot = self.define_local("__filt_res");
                    let i_slot = self.define_local("__filt_i");
                    let elem_slot = self.define_local("__filt_elem");
                    common::loops::emit_filter(&mut self.chunks[c], fn_slot, arr_slot, res_slot, i_slot, elem_slot, line);
                    return Ok(Some(()));
                }
                // No callback — filter falsy values inline
                let res_slot = self.define_local("__filt_res");
                let i_slot = self.define_local("__filt_i");
                self.emit_u16(Op::ARRAY_NEW, 0);
                self.emit_u16(Op::LOCAL_SET, res_slot);
                self.emit(Op::DROP);
                let lp = common::loops::emit_for_in_start(&mut self.chunks[c], arr_slot, i_slot, line);
                // element on stack — test truthiness
                let val_slot = self.define_local("__filt_val");
                self.emit_u16(Op::LOCAL_SET, val_slot);
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, val_slot);
                self.emit(Op::DYN_TO_BOOL);
                let skip = self.emit_jump(Op::BR_IF_FALSE);
                self.emit_u16(Op::LOCAL_GET, res_slot);
                self.emit_u16(Op::LOCAL_GET, val_slot);
                self.emit(Op::ARRAY_PUSH);
                self.emit(Op::DROP);
                self.patch_jump(skip);
                common::loops::emit_for_in_end(&mut self.chunks[c], i_slot, lp, line);
                self.emit_u16(Op::LOCAL_GET, res_slot);
                return Ok(Some(()));
            }
            "array_reduce" => {
                // Uses common::loops::emit_reduce
                if args.len() < 2 { return Ok(None); }
                self.compile_expression(&args[0].value)?; // array
                let arr_slot = self.define_local("__red_arr");
                self.emit_u16(Op::LOCAL_SET, arr_slot);
                self.compile_expression(&args[1].value)?; // callback
                let fn_slot = self.define_local("__red_fn");
                self.emit_u16(Op::LOCAL_SET, fn_slot);
                // If initial value provided, set acc to it; otherwise use arr[0]
                if args.len() >= 3 {
                    self.compile_expression(&args[2].value)?;
                    let acc_slot = self.define_local("__red_acc");
                    self.emit_u16(Op::LOCAL_SET, acc_slot);
                    let i_slot = self.define_local("__red_i");
                    // Start from 0 since we have an initial value
                    self.emit_constant(Value::I32(0));
                    self.emit_u16(Op::LOCAL_SET, i_slot);
                    // Inline loop with initial accumulator (structured CF)
                    let lp = common::loops::emit_loop_start(&mut self.chunks[c], line);
                    self.emit_u16(Op::LOCAL_GET, i_slot);
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit(Op::ARRAY_LENGTH);
                    self.emit(Op::DYN_LT);
                    common::loops::emit_loop_cond(&mut self.chunks[c], line);
                    self.emit_u16(Op::LOCAL_GET, fn_slot);
                    self.emit_u16(Op::LOCAL_GET, acc_slot);
                    self.emit_u16(Op::LOCAL_GET, arr_slot);
                    self.emit_u16(Op::LOCAL_GET, i_slot);
                    self.emit(Op::ARRAY_GET);
                    self.emit_u8(Op::CALL_REF, 2);
                    self.emit_u16(Op::LOCAL_SET, acc_slot);
                    // i += 1
                    self.emit_u16(Op::LOCAL_GET, i_slot);
                    self.emit_constant(Value::I32(1));
                    self.emit(Op::I32_ADD);
                    self.emit_u16(Op::LOCAL_SET, i_slot);
                    common::loops::emit_loop_end(&mut self.chunks[c], lp, line);
                    self.emit_u16(Op::LOCAL_GET, acc_slot);
                } else {
                    let acc_slot = self.define_local("__red_acc");
                    let i_slot = self.define_local("__red_i");
                    common::loops::emit_reduce(&mut self.chunks[c], fn_slot, arr_slot, acc_slot, i_slot, line);
                }
                return Ok(Some(()));
            }
            "array_walk" | "array_foreach" => {
                // Uses common::loops::emit_foreach
                if args.len() < 2 { return Ok(None); }
                self.compile_expression(&args[0].value)?; // array
                let arr_slot = self.define_local("__walk_arr");
                self.emit_u16(Op::LOCAL_SET, arr_slot);
                self.compile_expression(&args[1].value)?; // callback
                let fn_slot = self.define_local("__walk_fn");
                self.emit_u16(Op::LOCAL_SET, fn_slot);
                let i_slot = self.define_local("__walk_i");
                common::loops::emit_foreach(&mut self.chunks[c], fn_slot, arr_slot, i_slot, line);
                return Ok(Some(()));
            }
            "usort" => {
                compile_args!();
                let tmp = self.define_local("__usort_tmp");
                self.emit_u16(Op::LOCAL_SET, tmp);
                common::bundle::emit_call_push_func(&mut self.chunks[c], "__vybe_sorted", line);
                self.emit_u16(Op::LOCAL_GET, tmp);
                common::bundle::emit_call_invoke(&mut self.chunks[c], 1, line);
                return Ok(Some(()));
            }

            // ── Dict ops (common::dict) ─────────────────────────────
            "array_keys"   => { compile_args!(); common::dict::emit_keys(&mut self.chunks[c], line); return Ok(Some(())); }
            "array_values" => { compile_args!(); common::dict::emit_values(&mut self.chunks[c], line); return Ok(Some(())); }
            "compact" => {
                let line = self.line;
                common::dict::emit_new(&mut self.chunks[c], line);
                for arg in args {
                    if let Expression::Str(name) = &arg.value {
                        self.emit(Op::DUP);
                        self.emit_var_get(name);
                        let line = self.line;
                        common::dict::emit_set_const_key(&mut self.chunks[c], name, line);
                    }
                }
                return Ok(Some(()));
            }
            "array_key_exists" | "key_exists" => {
                compile_args!();
                let idx = self.import("vybe:object", "hasProperty");
                self.emit_host_call(idx, 2);
                return Ok(Some(()));
            }

            // ── Type checking (opcodes) ─────────────────────────────
            "is_array" => {
                compile_args!();
                let idx = self.import("vybe:array", "isArray");
                self.emit_host_call(idx, 1);
                return Ok(Some(()));
            }
            "is_string" => {
                // Check ref_is_null first, then check if it's a string via str_length (doesn't throw on strings)
                compile_args!();
                self.emit(Op::DUP);
                self.emit(Op::REF_IS_NULL);
                let is_null = self.emit_jump(Op::BR_IF_TRUE);
                // Try str_length — if it doesn't crash, it's a string-like value
                // Simplest: use convert toString and check if unchanged (approximate)
                self.emit(Op::TRUE); // approximate: treat as true for now
                let end = self.emit_jump(Op::BR);
                self.patch_jump(is_null);
                self.emit(Op::DROP);
                self.emit(Op::FALSE);
                self.patch_jump(end);
                return Ok(Some(()));
            }
            "is_int" | "is_integer" | "is_long" => {
                compile_args!();
                common::convert::emit_is_numeric(&mut self.chunks[c], line);
                return Ok(Some(()));
            }
            "is_float" | "is_double" => {
                compile_args!();
                common::convert::emit_is_numeric(&mut self.chunks[c], line);
                return Ok(Some(()));
            }
            "is_bool" => {
                compile_args!();
                // Approximate: check dyn_to_bool produces same value
                self.emit(Op::DUP);
                self.emit(Op::DYN_TO_BOOL);
                self.emit(Op::EQ);
                return Ok(Some(()));
            }
            "is_object" => {
                compile_args!();
                self.emit(Op::DUP);
                self.emit(Op::REF_IS_NULL);
                self.emit(Op::DYN_NOT);
                return Ok(Some(()));
            }
            "gettype" => {
                // Inline: use typeof-style check via opcodes
                // For now, convert to string representation
                compile_args!();
                common::convert::emit_to_string(&mut self.chunks[c], line);
                return Ok(Some(()));
            }
            "define" => {
                if args.len() >= 2 {
                    self.compile_expression(&args[1].value)?;
                    if let Expression::Str(name) = &args[0].value {
                        let idx = self.add_string_constant(name);
                        self.emit_u16(Op::GLOBAL_SET, idx);
                        self.defined_globals.insert(name.clone());
                    }
                }
                self.emit(Op::TRUE);
                return Ok(Some(()));
            }
            "defined" => {
                if let Some(arg) = args.first() {
                    if let Expression::Str(name) = &arg.value {
                        let idx = self.add_string_constant(name);
                        self.emit_u16(Op::GLOBAL_GET, idx);
                        self.emit(Op::REF_IS_NULL);
                        self.emit(Op::DYN_NOT);
                    } else { self.emit(Op::FALSE); }
                } else { self.emit(Op::FALSE); }
                return Ok(Some(()));
            }
            "function_exists" | "class_exists" => {
                if let Some(arg) = args.first() {
                    if let Expression::Str(name) = &arg.value {
                        let idx = self.add_string_constant(&name.to_lowercase());
                        self.emit_u16(Op::GLOBAL_GET, idx);
                    } else {
                        self.compile_expression(&arg.value)?;
                    }
                    self.emit(Op::REF_IS_NULL);
                    self.emit(Op::DYN_NOT);
                } else { self.emit(Op::FALSE); }
                return Ok(Some(()));
            }
            "ob_start" | "ob_end_clean" | "ob_get_clean" => {
                self.emit(Op::NULL);
                return Ok(Some(()));
            }

            // ── Encoding (existing host imports from vybe:convert, same as JS) ──
            "urlencode" | "rawurlencode"   => { compile_args!(); let i = self.import("vybe:convert", "encodeURIComponent"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "urldecode" | "rawurldecode"   => { compile_args!(); let i = self.import("vybe:convert", "decodeURIComponent"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "base64_encode" => { compile_args!(); let i = self.import("vybe:convert", "btoa"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "base64_decode" => { compile_args!(); let i = self.import("vybe:convert", "atob"); self.emit_host_call(i, 1); return Ok(Some(())); }

            // ── JSON (existing host imports, same as JS) ────────────
            "json_encode" => { compile_args!(); let i = self.import("vybe:json", "stringify"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "json_decode" => { compile_args!(); let i = self.import("vybe:json", "parse"); self.emit_host_call(i, 1); return Ok(Some(())); }

            // ── Crypto (existing host imports) ──────────────────────
            "md5"  => { compile_args!(); let i = self.import("vybe:crypto", "md5"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "sha1" | "sha256" => { compile_args!(); let i = self.import("vybe:crypto", "sha256"); self.emit_host_call(i, 1); return Ok(Some(())); }

            // ── Regex (existing host imports, same as JS) ───────────
            "preg_match"     => { compile_args!(); let i = self.import("vybe:regex", "test"); self.emit_host_call(i, args.len().min(2) as u8); return Ok(Some(())); }
            "preg_match_all" => { compile_args!(); let i = self.import("vybe:regex", "matchGroups"); self.emit_host_call(i, args.len().min(2) as u8); return Ok(Some(())); }
            "preg_replace"   => { compile_args!(); let i = self.import("vybe:regex", "replaceAll"); self.emit_host_call(i, args.len().min(3) as u8); return Ok(Some(())); }
            "preg_split"     => { compile_args!(); let i = self.import("vybe:regex", "split"); self.emit_host_call(i, args.len().min(2) as u8); return Ok(Some(())); }

            // (Filesystem/Clocks now in the expanded section above)

            // ── Math host imports (trig, log — not in WASM spec) ────
            "asin"  => { compile_args!(); let i = self.import("vybe:math", "asin"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "acos"  => { compile_args!(); let i = self.import("vybe:math", "acos"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "atan"  => { compile_args!(); let i = self.import("vybe:math", "atan"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "atan2" => { compile_args!(); let i = self.import("vybe:math", "atan2"); self.emit_host_call(i, 2); return Ok(Some(())); }
            "log10" => { compile_args!(); let i = self.import("vybe:math", "log10"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "log2"  => { compile_args!(); let i = self.import("vybe:math", "log2"); self.emit_host_call(i, 1); return Ok(Some(())); }

            // ── String host imports (existing, same as JS uses) ─────
            "sprintf" | "number_format" => {
                compile_args!();
                let i = self.import("vybe:string", "format");
                self.emit_host_call(i, args.len() as u8);
                return Ok(Some(()));
            }

            // ── HTTP (existing wasi:http host) ──
            "file_get_contents" if args.len() == 1 => {
                // Could be file OR URL — check at runtime; for now use readFile
                compile_args!();
                let i = self.import("wasi:filesystem", "readFile");
                self.emit_host_call(i, 1);
                return Ok(Some(()));
            }
            "curl_init" => { self.emit(Op::NULL); return Ok(Some(())); }
            "curl_setopt" => { compile_args!(); self.emit(Op::DROP); self.emit(Op::DROP); self.emit(Op::NULL); return Ok(Some(())); }
            "curl_exec" => {
                compile_args!();
                let i = self.import("wasi:http", "get");
                self.emit_host_call(i, 1);
                return Ok(Some(()));
            }
            "curl_close" => { compile_args!(); self.emit(Op::DROP); self.emit(Op::NULL); return Ok(Some(())); }
            "http_response_code" => { compile_args!(); self.emit(Op::NULL); return Ok(Some(())); }
            "header" => { compile_args!(); self.emit(Op::DROP); self.emit(Op::NULL); return Ok(Some(())); }

            // ── Environment / CLI (existing wasi:cli host) ──
            "getenv" => { compile_args!(); let i = self.import("wasi:cli", "getEnv"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "php_uname" => { let i = self.import("wasi:cli", "platform"); self.emit_host_call(i, 0); return Ok(Some(())); }
            "php_sapi_name" => { self.emit_constant(Value::String(Arc::from("vybe"))); return Ok(Some(())); }
            "phpversion" => { self.emit_constant(Value::String(Arc::from("8.3.0"))); return Ok(Some(())); }
            "getcwd" => { let i = self.import("wasi:cli", "cwd"); self.emit_host_call(i, 0); return Ok(Some(())); }
            "gethostname" => { let i = self.import("wasi:cli", "machineName"); self.emit_host_call(i, 0); return Ok(Some(())); }

            // ── Filesystem (existing wasi:filesystem host) ──
            "file_put_contents" => { compile_args!(); let i = self.import("wasi:filesystem", "writeFile"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            "file_exists" => { compile_args!(); let i = self.import("wasi:filesystem", "exists"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "is_file"  => { compile_args!(); let i = self.import("wasi:filesystem", "isFile"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "is_dir"   => { compile_args!(); let i = self.import("wasi:filesystem", "isDir"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "mkdir"    => { compile_args!(); let i = self.import("wasi:filesystem", "mkdir"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "rmdir" | "unlink" => { compile_args!(); let i = self.import("wasi:filesystem", "remove"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "rename"   => { compile_args!(); let i = self.import("wasi:filesystem", "rename"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            "copy"     => { compile_args!(); let i = self.import("wasi:filesystem", "copy"); self.emit_host_call(i, 2); return Ok(Some(())); }
            "scandir"  => { compile_args!(); let i = self.import("wasi:filesystem", "listDir"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "glob"     => { compile_args!(); let i = self.import("wasi:filesystem", "listDir"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "realpath" => { compile_args!(); let i = self.import("wasi:filesystem", "pathGetFullPath"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "dirname"  => { compile_args!(); let i = self.import("wasi:filesystem", "pathGetDirectory"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "basename" => { compile_args!(); let i = self.import("wasi:filesystem", "pathGetFileName"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "pathinfo" => {
                // Return assoc array with dirname, basename, extension, filename
                compile_args!();
                self.emit(Op::DUP);
                let c = self.current_chunk_idx;
                let line = self.line;
                common::dict::emit_new(&mut self.chunks[c], line);
                // dirname
                self.emit(Op::DUP); self.emit(Op::DUP);
                let di = self.import("wasi:filesystem", "pathGetDirectory");
                self.emit_host_call(di, 1);
                let dk = self.add_string_constant("dirname");
                self.emit_u16(Op::STRUCT_SET, dk); self.emit(Op::DROP);
                // basename
                self.emit(Op::DUP); self.emit(Op::DUP);
                let bi = self.import("wasi:filesystem", "pathGetFileName");
                self.emit_host_call(bi, 1);
                let bk = self.add_string_constant("basename");
                self.emit_u16(Op::STRUCT_SET, bk); self.emit(Op::DROP);
                // extension
                self.emit(Op::DUP); self.emit(Op::DUP);
                let ei = self.import("wasi:filesystem", "pathGetExtension");
                self.emit_host_call(ei, 1);
                let ek = self.add_string_constant("extension");
                self.emit_u16(Op::STRUCT_SET, ek); self.emit(Op::DROP);
                // Drop extra path copy, keep dict
                // Stack management is messy — simplified: just return the dict
                return Ok(Some(()));
            }
            "filesize" => { compile_args!(); let i = self.import("wasi:filesystem", "fileSize"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "tempnam" | "sys_get_temp_dir" => { let i = self.import("wasi:filesystem", "pathGetTempPath"); self.emit_host_call(i, 0); return Ok(Some(())); }
            "file" => {
                // file() reads file into array of lines
                compile_args!();
                let i = self.import("wasi:filesystem", "readFile");
                self.emit_host_call(i, 1);
                // Split by newline
                self.emit_constant(Value::String(Arc::from("\n")));
                common_strings::emit_split(&mut self.chunks[self.current_chunk_idx], self.line);
                return Ok(Some(()));
            }
            "file_get_contents" => { compile_args!(); let i = self.import("wasi:filesystem", "readFile"); self.emit_host_call(i, 1); return Ok(Some(())); }

            // ── Random (existing wasi:random host) ──
            "random_int" => { compile_args!(); let i = self.import("wasi:random", "randomInt"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            "random_bytes" => { compile_args!(); let i = self.import("wasi:random", "randomBytes"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "uniqid" => { let i = self.import("wasi:random", "uuid"); self.emit_host_call(i, 0); return Ok(Some(())); }

            // ── Date/Time (existing wasi:clocks host) ──
            "date" => { compile_args!(); let i = self.import("wasi:clocks", "toISOString"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            "time"      => { let i = self.import("wasi:clocks", "now"); self.emit_host_call(i, 0); return Ok(Some(())); }
            "microtime" => { let i = self.import("wasi:clocks", "hrtime"); self.emit_host_call(i, 0); return Ok(Some(())); }
            "sleep"     => { compile_args!(); let i = self.import("wasi:clocks", "sleep"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "usleep"    => {
                // usleep(microseconds) — convert to ms for sleep
                compile_args!();
                self.emit_constant(Value::F64(1000.0));
                self.emit(Op::F64_DIV);
                let i = self.import("wasi:clocks", "sleep");
                self.emit_host_call(i, 1);
                return Ok(Some(()));
            }

            // ── Process (existing vybe:types host) ──
            "exec" | "shell_exec" | "system" => {
                compile_args!();
                let i = self.import("vybe:types", "processStart");
                self.emit_host_call(i, args.len() as u8);
                return Ok(Some(()));
            }

            // ── XML (existing vybe:xml host) ──
            "simplexml_load_string" => { compile_args!(); let i = self.import("vybe:xml", "parse"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "simplexml_load_file" => {
                // Read file then parse XML
                compile_args!();
                let rf = self.import("wasi:filesystem", "readFile");
                self.emit_host_call(rf, 1);
                let xp = self.import("vybe:xml", "parse");
                self.emit_host_call(xp, 1);
                return Ok(Some(()));
            }

            // ── Sockets (existing vybe:net host) ──
            "fsockopen" => { compile_args!(); let i = self.import("vybe:net", "tcpConnect"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            "fclose" => { compile_args!(); let i = self.import("wasi:filesystem", "closeFile"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "fwrite" | "fputs" => { compile_args!(); let i = self.import("wasi:filesystem", "writeFile_handle"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            "fgets" | "fread" => { compile_args!(); let i = self.import("wasi:filesystem", "lineInput"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }

            // ── GUID (existing vybe:types host) ──
            "uuid_create" | "com_create_guid" => { let i = self.import("wasi:random", "uuid"); self.emit_host_call(i, 0); return Ok(Some(())); }

            // ── Async (common::functions — same as JS await) ──
            "await" => {
                // await($promise) → suspend until resolved
                compile_args!();
                common::functions::emit_await(&mut self.chunks[c], line);
                return Ok(Some(()));
            }

            // ── Convert (existing vybe:convert host) ──
            "dechex" | "bin2hex" => { compile_args!(); let i = self.import("vybe:convert", "hex"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "decoct" => { compile_args!(); let i = self.import("vybe:convert", "oct"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "is_finite" => { compile_args!(); let i = self.import("vybe:convert", "isFinite"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "is_nan" => { compile_args!(); let i = self.import("vybe:convert", "isNaN"); self.emit_host_call(i, 1); return Ok(Some(())); }

            // ── Math (unmapped) ──
            "hypot" => { compile_args!(); let i = self.import("vybe:math", "hypot"); self.emit_host_call(i, 2); return Ok(Some(())); }
            "log1p" => { compile_args!(); common::math::emit_log(&mut self.chunks[c], line); return Ok(Some(())); }
            "pi" => { let i = self.import("vybe:math", "PI"); self.emit_host_call(i, 0); return Ok(Some(())); }
            "m_pi" => { let i = self.import("vybe:math", "PI"); self.emit_host_call(i, 0); return Ok(Some(())); }
            "intdiv" => { compile_args!(); self.emit(Op::F64_DIV); self.emit(Op::F64_TRUNC); return Ok(Some(())); }
            "fmod" => { compile_args!(); { let l = self.line; vybe_compiler_common::expressions::emit_f64_mod(&mut self.chunks[self.current_chunk_idx], l); }; return Ok(Some(())); }
            "fdiv" => { compile_args!(); self.emit(Op::F64_DIV); return Ok(Some(())); }

            // ── Network / Sockets (existing vybe:net host — same as VB/C#) ──
            "dns_get_record" | "gethostbyname" => { compile_args!(); let i = self.import("vybe:net", "dnsResolve"); self.emit_host_call(i, 1); return Ok(Some(())); }
            // TCP server
            "stream_socket_server" | "socket_create_listen" => { compile_args!(); let i = self.import("vybe:net", "tcpListenerNew"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            "stream_socket_accept" | "socket_accept" => { compile_args!(); let i = self.import("vybe:net", "tcpListenerAccept"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "socket_listen" => { compile_args!(); let i = self.import("vybe:net", "tcpListenerStart"); self.emit_host_call(i, 0); return Ok(Some(())); }
            // TCP client
            "socket_connect" => { compile_args!(); let i = self.import("vybe:net", "tcpConnect"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            "socket_close" => { compile_args!(); let i = self.import("vybe:net", "tcpClose"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "socket_read" => { compile_args!(); let i = self.import("vybe:net", "streamReaderReadLine"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            "socket_write" | "socket_send" => { compile_args!(); let i = self.import("vybe:net", "streamWriterWriteLine"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            // Stream wrappers
            "stream_get_contents" => { compile_args!(); let i = self.import("vybe:net", "streamReaderReadLine"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            // UDP
            "socket_create" => { compile_args!(); let i = self.import("vybe:net", "udpNew"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            "socket_sendto" => { compile_args!(); let i = self.import("vybe:net", "udpSend"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            "socket_recvfrom" => { compile_args!(); let i = self.import("vybe:net", "udpReceive"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }

            // ── CLI (existing wasi:cli host) ──
            "readline" | "fgets_stdin" => { let i = self.import("wasi:cli", "readLine"); self.emit_host_call(i, 0); return Ok(Some(())); }
            "error_log" => { compile_args!(); let i = self.import("wasi:cli", "error"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            "trigger_error" => { compile_args!(); let i = self.import("wasi:cli", "warn"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            "php_info" => { self.emit_constant(Value::String(Arc::from("Vybe PHP 8.3 on WASM VM"))); let line = self.line; common::io::emit_print(&mut self.chunks[c], 1, line); return Ok(Some(())); }
            "get_current_user" => { let i = self.import("wasi:cli", "userName"); self.emit_host_call(i, 0); return Ok(Some(())); }

            // ── File handles (additional) ──
            "fopen" => { compile_args!(); let i = self.import("wasi:filesystem", "openFile"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            "feof" => { compile_args!(); self.emit(Op::REF_IS_NULL); return Ok(Some(())); }

            // ── Filesystem (more) ──
            "file_append_contents" | "fopen_append" => { compile_args!(); let i = self.import("wasi:filesystem", "appendFile"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            "stat" | "lstat" => { compile_args!(); let i = self.import("wasi:filesystem", "stat"); self.emit_host_call(i, 1); return Ok(Some(())); }
            "readdir" | "opendir" => { compile_args!(); let i = self.import("wasi:filesystem", "readDirEntries"); self.emit_host_call(i, 1); return Ok(Some(())); }

            // ── HTTP (extended) ──
            "fetch" => { compile_args!(); let i = self.import("wasi:http", "fetch"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }

            // ── SplStack / SplQueue (existing vybe:types host) ──
            // These map to the stack/queue host functions used by VB/C#
            "spl_stack_push" => { compile_args!(); let i = self.import("vybe:types", "stackPush"); self.emit_host_call(i, args.len() as u8); return Ok(Some(())); }
            "spl_stack_pop" => { compile_args!(); let i = self.import("vybe:types", "stackPop"); self.emit_host_call(i, 1); return Ok(Some(())); }

            // ── DateTime (existing vybe:types host — same as C#/VB DateTime) ──
            // new DateTime() maps to vybe:types dateTimeNow
            // $dt->format() maps to vybe:types dateTimeToString

            // ── Database (existing vybe:database host, same as VB/JS) ──
            "mysqli_connect" => {
                compile_args!();
                let i = self.import("vybe:database", "connect");
                self.emit_host_call(i, args.len() as u8);
                return Ok(Some(()));
            }
            "mysqli_query" => {
                compile_args!();
                let i = self.import("vybe:database", "query");
                self.emit_host_call(i, args.len() as u8);
                return Ok(Some(()));
            }
            "mysqli_fetch_all" | "mysqli_fetch_assoc" | "mysqli_fetch_array" => {
                // Result is already an array from query — just return it
                compile_args!();
                return Ok(Some(()));
            }
            "mysqli_close" => {
                compile_args!();
                let i = self.import("vybe:database", "close");
                self.emit_host_call(i, 1);
                return Ok(Some(()));
            }
            "mysqli_num_rows" => {
                compile_args!();
                self.emit(Op::ARRAY_LENGTH);
                return Ok(Some(()));
            }

            // ── Threading (common::threading — same as Python/JS) ──
            "thread_create" => {
                // thread_create(fn) → spawn thread running fn, return handle
                compile_args!();
                common::threading::emit_thread_spawn(&mut self.chunks[c], line);
                return Ok(Some(()));
            }
            "thread_join" => {
                // thread_join($handle) → wait for thread, return result
                compile_args!();
                common::threading::emit_thread_join(&mut self.chunks[c], line);
                return Ok(Some(()));
            }
            "mutex_create" => {
                // mutex_create() → allocate lock word, return address
                let alloc_fn = self.chunks[c].add_import("wasi:thread", "allocLock");
                self.chunks[c].emit_op_u16(Op::CALL_IMPORT, alloc_fn, line);
                self.chunks[c].emit(0, line);
                return Ok(Some(()));
            }
            "mutex_lock" => {
                // mutex_lock($lock) → acquire spinlock
                compile_args!();
                let lock_slot = self.define_local("__lock_addr");
                self.emit_u16(Op::LOCAL_SET, lock_slot);
                common::threading::emit_lock_acquire(&mut self.chunks[c], lock_slot, line);
                self.emit(Op::NULL);
                return Ok(Some(()));
            }
            "mutex_unlock" => {
                // mutex_unlock($lock) → release spinlock
                compile_args!();
                let lock_slot = self.define_local("__unlock_addr");
                self.emit_u16(Op::LOCAL_SET, lock_slot);
                common::threading::emit_lock_release(&mut self.chunks[c], lock_slot, line);
                self.emit(Op::NULL);
                return Ok(Some(()));
            }

            _ => {
                // Check cross-language common imports as fallback
                if let Some((module, func)) = vybe_compiler_common::imports::resolve_common_import(name) {
                    compile_args!();
                    let i = self.import(module, func);
                    self.emit_host_call(i, args.len() as u8);
                    return Ok(Some(()));
                }
                return Ok(None);
            }
        }
    }

    // ------------------------------------------------------------------
    // Array literal — uses common::dict for associative, array_new for indexed
    // ------------------------------------------------------------------

    fn compile_array_literal(&mut self, elements: &[ArrayElement]) -> Result<(), String> {
        // ALL PHP arrays are ordered maps (dicts) — even indexed ones like [1, 2, 3]
        // become {"0": 1, "1": 2, "2": 3}. This ensures type compatibility with
        // JS objects, Python dicts, etc. across the shared VM type system.
        let line = self.line;
        let c = self.current_chunk_idx;
        common::dict::emit_new(&mut self.chunks[c], line);
        for (i, elem) in elements.iter().enumerate() {
            self.emit(Op::DUP); // keep dict on stack
            self.compile_expression(&elem.value)?;
            let key_str = match &elem.key {
                Some(Expression::Str(s)) => s.clone(),
                Some(Expression::Number(n)) => n.to_string(),
                _ => i.to_string(),
            };
            let line = self.line;
            common::dict::emit_set_const_key(&mut self.chunks[c], &key_str, line);
        }
        Ok(())
    }

    // ------------------------------------------------------------------
    // new ClassName(...)
    // ------------------------------------------------------------------

    fn compile_new(&mut self, class: &Expression, args: &[Argument]) -> Result<(), String> {
        if let Expression::Identifier(name) = class {
            // Special: new Fiber(fn) → create continuation via cont_new
            if name == "Fiber" && args.len() >= 1 {
                self.compile_expression(&args[0].value)?;
                self.emit(Op::CONT_NEW);
                return Ok(());
            }
            // Special: new Exception/RuntimeException/etc → cross-language exception object
            if name == "Exception" || name == "RuntimeException" || name == "TypeError"
                || name == "ValueError" || name == "InvalidArgumentException"
                || name == "LogicException" || name == "OutOfRangeException"
                || name == "OverflowException" || name == "UnderflowException"
                || name == "BadMethodCallException" || name == "DomainException"
                || name == "LengthException" || name == "RangeException" {
                let canonical = common::errors::canonical_exception_name(name);
                let this_slot = self.define_local("__exc_this");
                let msg_slot = self.define_local("__exc_msg");
                // Push message arg (or empty string)
                if !args.is_empty() {
                    self.compile_expression(&args[0].value)?;
                } else {
                    self.emit_constant(Value::String(Arc::from("")));
                }
                self.emit_u16(Op::LOCAL_SET, msg_slot);
                let line = self.line;
                let c = self.current_chunk_idx;
                common::errors::emit_exception_constructor(
                    &mut self.chunks[c], this_slot, canonical, msg_slot, line,
                );
                self.emit_u16(Op::LOCAL_GET, this_slot);
                return Ok(());
            }
            // Special: SPL/collections → same host types as VB/C# (cross-language compatible)
            if name == "SplDoublyLinkedList" || name == "ArrayObject" || name == "SplFixedArray" {
                let i = self.import("vybe:types", "listNew");
                self.emit_host_call(i, 0);
                return Ok(());
            }
            if name == "SplPriorityQueue" {
                let i = self.import("vybe:types", "queueNew");
                self.emit_host_call(i, 0);
                return Ok(());
            }
            // new Dictionary / new Map → same as VB Dictionary(Of K,V)
            if name == "Dictionary" || name == "Map" || name == "SplObjectStorage" {
                let i = self.import("vybe:types", "dictNew");
                self.emit_host_call(i, 0);
                return Ok(());
            }
            // new HashSet / new Set
            if name == "HashSet" || name == "Set" {
                let i = self.import("vybe:types", "hashSetNew");
                self.emit_host_call(i, 0);
                return Ok(());
            }
            // new StringBuilder → same as VB/C# StringBuilder
            if name == "StringBuilder" {
                if !args.is_empty() { self.compile_expression(&args[0].value)?; } else { self.emit_constant(Value::String(Arc::from(""))); }
                let i = self.import("vybe:types", "stringBuilderNew");
                self.emit_host_call(i, 1);
                return Ok(());
            }
            // new Random → same as VB/C# Random
            if name == "Random" {
                let i = self.import("vybe:threading", "randomNew");
                self.emit_host_call(i, 0);
                return Ok(());
            }
            // new Stopwatch → same as VB/C# Stopwatch
            if name == "Stopwatch" {
                let i = self.import("vybe:threading", "stopwatchNew");
                self.emit_host_call(i, 0);
                return Ok(());
            }
            // Special: new DateTime() → vybe:types dateTimeNow/dateTimeParse
            if name == "DateTime" || name == "DateTimeImmutable" {
                if args.is_empty() {
                    let i = self.import("vybe:types", "dateTimeNow");
                    self.emit_host_call(i, 0);
                } else {
                    self.compile_expression(&args[0].value)?;
                    let i = self.import("vybe:types", "dateTimeParse");
                    self.emit_host_call(i, 1);
                }
                return Ok(());
            }
            // Special: new SplStack / new SplQueue / new SplFixedArray
            if name == "SplStack" {
                let i = self.import("vybe:types", "stackNew");
                self.emit_host_call(i, 0);
                return Ok(());
            }
            if name == "SplQueue" {
                let i = self.import("vybe:types", "queueNew");
                self.emit_host_call(i, 0);
                return Ok(());
            }
            // Special: new PDO(dsn) → vybe:database connect
            if name == "PDO" || name == "mysqli" {
                // PDO(dsn, [user, pass]) — connect to database
                if args.is_empty() { self.emit(Op::NULL); return Ok(()); }
                self.compile_expression(&args[0].value)?; // DSN string
                let connect = self.import("vybe:database", "connect");
                self.emit_host_call(connect, 1);
                return Ok(());
            }
        }
        if let Expression::Identifier(name) = class {
            let lower = name.to_lowercase();
            let is_local = self.resolve_var(&lower).is_some();
            let is_upvalue = self.scopes.len() > 1
                && self.resolve_upvalue(self.scopes.len() - 1, &lower).is_some();
            let is_defined = is_local || is_upvalue
                || self.defined_globals.contains(&lower)
                || self.defined_classes.contains(&lower);
            if is_defined {
                // Known class — use global_get + call_ref
                self.compile_expression(class)?;
                for arg in args { self.compile_expression(&arg.value)?; }
                self.emit_u8(Op::CALL_REF, args.len() as u8);
            } else {
                // Unresolved class → WASM import
                let idx = self.import("*", &lower);
                for arg in args { self.compile_expression(&arg.value)?; }
                self.emit_host_call(idx, args.len() as u8);
            }
            return Ok(());
        }
        // Dynamic expression (e.g. new $className()) — keep global_get + call_ref
        self.compile_expression(class)?;
        for arg in args { self.compile_expression(&arg.value)?; }
        self.emit_u8(Op::CALL_REF, args.len() as u8);
        Ok(())
    }

    // ------------------------------------------------------------------
    // Function compilation — uses common::functions
    // ------------------------------------------------------------------

    fn compile_function(&mut self, decl: &FunctionDecl) -> Result<usize, String> {
        let chunk_idx = self.chunks.len();
        let chunk = common::functions::create_function_chunk(&decl.name, decl.params.len() as u8);
        self.chunks.push(chunk);

        self.scopes.push(Scope::new_function());
        let saved_chunk = self.current_chunk_idx;
        self.current_chunk_idx = chunk_idx;

        // Define params as locals (slot 0 = callee, slots 1..N = params)
        for param in &decl.params {
            self.define_local(&param.name);
        }

        // Handle default parameter values
        for param in &decl.params {
            if let Some(default) = &param.default {
                if let Some(slot) = self.current_scope().resolve_local(&param.name) {
                    let line = self.line;
                    let skip = common::functions::emit_default_param_start(
                        &mut self.chunks[self.current_chunk_idx], slot, line,
                    );
                    self.compile_expression(default)?;
                    let line = self.line;
                    common::functions::emit_default_param_end(
                        &mut self.chunks[self.current_chunk_idx], slot, skip, line,
                    );
                }
            }
        }

        for stmt in &decl.body {
            self.compile_statement(stmt)?;
        }

        // Epilogue: null + return safety net
        let line = self.line;
        common::functions::emit_function_epilogue(&mut self.chunks[self.current_chunk_idx], line);

        let local_count = self.current_scope().next_slot;
        self.chunks[chunk_idx].local_count = local_count;

        self.scopes.pop();
        self.current_chunk_idx = saved_chunk;
        Ok(chunk_idx)
    }

    // ------------------------------------------------------------------
    // Class compilation — uses common::classes (same pattern as Python)
    // ------------------------------------------------------------------

    fn compile_class(&mut self, decl: &ClassDecl) -> Result<(), String> {
        let class_name = &decl.name;
        let parent_name = decl.parent.as_deref().unwrap_or("").to_string();

        // Track parent for parent::__construct calls
        let saved_parent = self.current_class_parent.take();
        if !parent_name.is_empty() {
            self.current_class_parent = Some(parent_name.clone());
        }

        // Collect methods and their chunk indices
        let mut method_entries: Vec<(String, usize)> = Vec::new();
        let mut static_method_entries: Vec<(String, usize)> = Vec::new();
        let mut init_chunk: Option<usize> = None;
        let mut init_params: Vec<String> = Vec::new();
        let mut field_defaults: Vec<(String, Option<Expression>)> = Vec::new();
        let mut constants: Vec<(String, Expression)> = Vec::new();

        // First pass: compile all methods
        for member in &decl.members {
            match member {
                ClassMember::Method(m) => {
                    let ci = self.compile_method(m, class_name)?;
                    let name_lower = m.name.to_lowercase();
                    if name_lower == "__construct" {
                        init_chunk = Some(ci);
                        init_params = m.params.iter().map(|p| p.name.clone()).collect();
                    }
                    if m.is_static {
                        static_method_entries.push((name_lower, ci));
                    } else {
                        method_entries.push((name_lower, ci));
                    }
                }
                ClassMember::Property { name, default, .. } => {
                    field_defaults.push((name.clone(), default.clone()));
                }
                ClassMember::Constant { name, value } => {
                    constants.push((name.clone(), value.clone()));
                }
            }
        }

        // Build constructor chunk (allocates object, stamps type, binds methods, calls __construct)
        let user_params = init_params.len();
        let ctor_name = class_name.to_lowercase();
        let ctor = common::functions::create_function_chunk(&ctor_name, user_params as u8);
        let ctor_idx = self.chunks.len();
        self.chunks.push(ctor);

        self.scopes.push(Scope::new_function());
        let saved_chunk = self.current_chunk_idx;
        self.current_chunk_idx = ctor_idx;

        // slot 0 = callee (implicit), slots 1..N = user params
        for p in &init_params {
            self.define_local(p);
        }
        let this_slot = self.define_local("__this");

        if parent_name.is_empty() {
            // Base class: create object here
            let line = self.line;
            common::classes::emit_new_typed_object(&mut self.chunks[ctor_idx], this_slot, class_name, line);
        }
        // For child classes: parent::__construct in the body creates the object
        // and stores it in $this via the compile_static_call handler.

        // Mix in trait methods — same pattern as Dart mixins:
        // Create a trait instance, then use __vybe_assign to copy its methods to this.
        // Traits are compiled as regular classes, so calling their constructor
        // returns an object with all methods bound.
        for trait_name in &decl.traits {
            let line = self.line;
            let trait_lower = trait_name.to_lowercase();
            // Get trait constructor, call with 0 args to get prototype with bound methods
            let trait_c = self.chunks[ctor_idx].add_constant(Value::String(Arc::from(trait_lower.as_str())));
            self.chunks[ctor_idx].emit_op_u16(Op::GLOBAL_GET, trait_c, line);
            self.chunks[ctor_idx].emit_op_u8(Op::CALL_REF, 0, line);
            // assign(this, traitPrototype) — copies all methods onto this
            let trait_slot = self.define_local(&format!("__trait_{}", trait_lower));
            self.chunks[ctor_idx].emit_op_u16(Op::LOCAL_SET, trait_slot, line);
            self.chunks[ctor_idx].emit_op(Op::DROP, line);
            common::bundle::emit_call_push_func(&mut self.chunks[ctor_idx], "__vybe_assign", line);
            self.chunks[ctor_idx].emit_op_u16(Op::LOCAL_GET, this_slot, line);
            self.chunks[ctor_idx].emit_op_u16(Op::LOCAL_GET, trait_slot, line);
            common::bundle::emit_call_invoke(&mut self.chunks[ctor_idx], 2, line);
            self.chunks[ctor_idx].emit_op(Op::DROP, line);
        }

        if parent_name.is_empty() {
            // Base class: bind methods first, then call __construct
            for (method_name, method_ci) in &method_entries {
                if method_name == "__construct" { continue; }
                let line = self.line;
                common::classes::emit_bind_method_with_aliases(
                    &mut self.chunks[ctor_idx], this_slot, method_name, *method_ci, line,
                );
            }
        }

        // Set field defaults on the instance (only for base classes — child gets them from parent)
        if parent_name.is_empty() {
            for (field_name, default) in &field_defaults {
                self.emit_u16(Op::LOCAL_GET, this_slot);
                if let Some(val) = default {
                    self.compile_expression(val)?;
                } else {
                    self.emit(Op::NULL);
                }
                let key = self.add_string_constant(field_name);
                self.emit_u16(Op::STRUCT_SET, key);
                self.emit(Op::DROP);
            }
        }

        // Call __construct(this, args...) if present
        if let Some(init_ci) = init_chunk {
            let line = self.line;
            common::functions::emit_ref_func(&mut self.chunks[ctor_idx], init_ci, 0, line);
            self.emit_u16(Op::LOCAL_GET, this_slot); // $this
            for i in 0..user_params {
                self.emit_u16(Op::LOCAL_GET, (i + 1) as u16);
            }
            self.emit_u8(Op::CALL_REF, (user_params + 1) as u8);
            if !parent_name.is_empty() {
                // Child class: __construct called parent::__construct which created
                // the object. __construct returns $this. Use it as our __this.
                self.emit_u16(Op::LOCAL_SET, this_slot);
            } else {
                self.emit(Op::DROP);
            }
        }

        if !parent_name.is_empty() {
            // Child class: bind methods AFTER __construct (object now exists from parent)
            for (method_name, method_ci) in &method_entries {
                if method_name == "__construct" { continue; }
                let line = self.line;
                common::classes::emit_bind_method_with_aliases(
                    &mut self.chunks[ctor_idx], this_slot, method_name, *method_ci, line,
                );
            }
            // Set child field defaults
            for (field_name, default) in &field_defaults {
                self.emit_u16(Op::LOCAL_GET, this_slot);
                if let Some(val) = default {
                    self.compile_expression(val)?;
                } else {
                    self.emit(Op::NULL);
                }
                let key = self.add_string_constant(field_name);
                self.emit_u16(Op::STRUCT_SET, key);
                self.emit(Op::DROP);
            }
        }

        // Stamp __types array for instanceof support
        {
            let line = self.line;
            common::classes::emit_instanceof_chain(&mut self.chunks[ctor_idx], this_slot, class_name, line);
        }

        // Return this
        {
            let line = self.line;
            common::classes::emit_constructor_return(&mut self.chunks[ctor_idx], this_slot, line);
        }

        let local_count = self.current_scope().next_slot;
        self.chunks[ctor_idx].local_count = local_count;
        self.scopes.pop();
        self.current_chunk_idx = saved_chunk;

        // Store constructor as local + global
        let class_local = self.define_local(&class_name.to_lowercase());
        {
            let line = self.line;
            common::classes::emit_store_constructor(
                &mut self.chunks[self.current_chunk_idx], class_name, ctor_idx, class_local, line,
            );
        }
        self.defined_globals.insert(class_name.to_lowercase());
        self.defined_classes.insert(class_name.to_lowercase());

        // Attach static methods to the constructor
        for (sm_name, sm_ci) in &static_method_entries {
            let line = self.line;
            common::classes::emit_attach_static_method(
                &mut self.chunks[self.current_chunk_idx], class_local, sm_name, *sm_ci, line,
            );
        }

        // Set class constants on the constructor object
        // For enum cases: create objects with ->name and ->value accessors
        let _is_enum = !constants.is_empty() && constants.iter().all(|(_, v)| !matches!(v, Expression::Number(_) | Expression::Bool(_)) || true);
        for (const_name, const_val) in &constants {
            self.emit_u16(Op::LOCAL_GET, class_local);
            // Build enum case object: { name: "CaseName", value: val, __type: "ClassName" }
            let line = self.line;
            let c = self.current_chunk_idx;
            common::dict::emit_new(&mut self.chunks[c], line);
            // name property
            self.emit(Op::DUP);
            self.emit_constant(Value::String(Arc::from(const_name.as_str())));
            let name_key = self.add_string_constant("name");
            self.emit_u16(Op::STRUCT_SET, name_key);
            self.emit(Op::DROP);
            // value property
            self.emit(Op::DUP);
            self.compile_expression(const_val)?;
            let val_key = self.add_string_constant("value");
            self.emit_u16(Op::STRUCT_SET, val_key);
            self.emit(Op::DROP);
            // __type property for instanceof
            self.emit(Op::DUP);
            self.emit_constant(Value::String(Arc::from(class_name.as_str())));
            let type_key = self.add_string_constant("__type");
            self.emit_u16(Op::STRUCT_SET, type_key);
            self.emit(Op::DROP);
            // Store on class constructor
            let key = self.add_string_constant(const_name);
            self.emit_u16(Op::STRUCT_SET, key);
            self.emit(Op::DROP);
        }

        // Inherit statics from parent
        if !parent_name.is_empty() {
            self.emit_u16(Op::LOCAL_GET, class_local);
            let line = self.line;
            common::classes::emit_inherit_statics(&mut self.chunks[self.current_chunk_idx], &parent_name, line);
            self.emit(Op::DROP);
        }

        // Register type entry
        let fields: Vec<String> = field_defaults.iter().map(|(n, _)| n.clone()).collect();
        let mut all_methods = method_entries;
        all_methods.extend(static_method_entries);
        common::classes::register_type(
            &mut self.chunks, class_name, &parent_name,
            fields, all_methods, false, Vec::new(), Some(ctor_idx),
        );

        self.current_class_parent = saved_parent;
        Ok(())
    }

    fn compile_method(&mut self, decl: &FunctionDecl, _class_name: &str) -> Result<usize, String> {
        let chunk_idx = self.chunks.len();
        let arity = if decl.is_static {
            decl.params.len() as u8
        } else {
            (decl.params.len() + 1) as u8 // +1 for $this
        };
        let chunk = common::functions::create_function_chunk(&decl.name, arity);
        self.chunks.push(chunk);

        self.scopes.push(Scope::new_function());
        let saved_chunk = self.current_chunk_idx;
        self.current_chunk_idx = chunk_idx;

        if !decl.is_static {
            // Slot 0 = callee (implicit), slot 1 = $this
            self.define_local("this");
        }
        for param in &decl.params {
            self.define_local(&param.name);
        }

        // Handle default parameter values
        for param in &decl.params {
            if let Some(default) = &param.default {
                if let Some(slot) = self.current_scope().resolve_local(&param.name) {
                    let line = self.line;
                    let skip = common::functions::emit_default_param_start(
                        &mut self.chunks[self.current_chunk_idx], slot, line,
                    );
                    self.compile_expression(default)?;
                    let line = self.line;
                    common::functions::emit_default_param_end(
                        &mut self.chunks[self.current_chunk_idx], slot, skip, line,
                    );
                }
            }
        }

        for stmt in &decl.body {
            self.compile_statement(stmt)?;
        }

        // __construct returns $this so child class wrappers can capture the parent-created object
        if decl.name == "__construct" {
            self.emit_u16(Op::LOCAL_GET, 1); // $this at slot 1
            self.emit(Op::RETURN);
        }
        let line = self.line;
        common::functions::emit_function_epilogue(&mut self.chunks[self.current_chunk_idx], line);

        let local_count = self.current_scope().next_slot;
        self.chunks[chunk_idx].local_count = local_count;
        self.scopes.pop();
        self.current_chunk_idx = saved_chunk;
        Ok(chunk_idx)
    }
}
