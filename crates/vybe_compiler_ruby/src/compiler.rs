use std::rc::Rc;
use vybe_bytecode::{Chunk, Value, Op};
use vybe_parser_ruby::ast::*;
use vybe_compiler_common as common;
use crate::scope::Scope;

struct LoopContext {
    break_patches: Vec<usize>,
    next_patches: Vec<usize>,
}

pub struct Compiler {
    chunks: Vec<Chunk>,
    scopes: Vec<Scope>,
    current_chunk_idx: usize,
    loop_stack: Vec<LoopContext>,
    line: u32,
    defined_globals: std::collections::HashSet<String>,
    defined_classes: std::collections::HashSet<String>,
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
        self.emit_u16(Op::r#const, idx);
    }

    fn add_string_constant(&mut self, s: &str) -> u16 {
        self.chunks[self.current_chunk_idx].add_constant(Value::String(Rc::from(s)))
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

    fn emit_host_call(&mut self, import_idx: u16, argc: u8) {
        let line = self.line;
        let c = &mut self.chunks[self.current_chunk_idx];
        c.emit_op_u16(Op::call_import, import_idx, line);
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
            self.emit_u16(Op::local_get, slot);
        } else if self.scopes.len() > 1 {
            if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, name) {
                self.emit_u8(Op::upvalue_get, uv);
                return;
            }
            let idx = self.add_string_constant(name);
            self.emit_u16(Op::global_get, idx);
        } else {
            let idx = self.add_string_constant(name);
            self.emit_u16(Op::global_get, idx);
        }
    }

    fn emit_var_set(&mut self, name: &str) {
        if let Some(slot) = self.resolve_var(name) {
            self.emit_u16(Op::local_set, slot);
        } else if self.scopes.len() > 1 {
            if let Some(uv) = self.resolve_upvalue(self.scopes.len() - 1, name) {
                self.emit_u8(Op::upvalue_set, uv);
                return;
            }
            let idx = self.add_string_constant(name);
            self.emit_u16(Op::global_set, idx);
            self.defined_globals.insert(name.to_string());
        } else {
            let idx = self.add_string_constant(name);
            self.emit_u16(Op::global_set, idx);
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

            Statement::Puts(exprs) | Statement::Print(exprs) | Statement::P(exprs) => {
                let c = self.current_chunk_idx;
                if exprs.is_empty() {
                    self.emit_constant(Value::String(Rc::from("")));
                    let line = self.line;
                    common::io::emit_print(&mut self.chunks[c], 1, line);
                    self.emit(Op::drop);
                } else {
                    for expr in exprs {
                        self.compile_expression(expr)?;
                        let line = self.line;
                        common::io::emit_print(&mut self.chunks[c], 1, line);
                        self.emit(Op::drop);
                    }
                }
            }

            Statement::Expression(expr) => {
                self.compile_expression(expr)?;
                self.emit(Op::drop);
            }

            Statement::Block(stmts) => {
                self.current_scope_mut().begin_block();
                for s in stmts { self.compile_statement(s)?; }
                self.current_scope_mut().end_block();
            }

            Statement::Assignment { target, op, value } => {
                self.compile_assign(target, op, value)?;
            }

            Statement::Return(val) => {
                if let Some(expr) = val {
                    self.compile_expression(expr)?;
                } else {
                    self.emit(Op::null);
                }
                self.emit(Op::r#return);
            }

            Statement::Break(_) => {
                let patch = self.emit_jump(Op::br);
                if let Some(ctx) = self.loop_stack.last_mut() {
                    ctx.break_patches.push(patch);
                }
            }

            Statement::Next(_) => {
                let patch = self.emit_jump(Op::br);
                if let Some(ctx) = self.loop_stack.last_mut() {
                    ctx.next_patches.push(patch);
                }
            }

            Statement::If { test, body, elsifs, else_body } => {
                self.compile_expression(test)?;
                self.emit(Op::dyn_to_bool);
                let mut end_jumps = Vec::new();
                let skip = self.emit_jump(Op::br_if_false);
                for s in body { self.compile_statement(s)?; }
                end_jumps.push(self.emit_jump(Op::br));
                self.patch_jump(skip);
                for elsif in elsifs {
                    self.compile_expression(&elsif.test)?;
                    self.emit(Op::dyn_to_bool);
                    let s = self.emit_jump(Op::br_if_false);
                    for st in &elsif.body { self.compile_statement(st)?; }
                    end_jumps.push(self.emit_jump(Op::br));
                    self.patch_jump(s);
                }
                if let Some(alt) = else_body {
                    for s in alt { self.compile_statement(s)?; }
                }
                for j in end_jumps { self.patch_jump(j); }
            }

            Statement::Unless { test, body, else_body } => {
                self.compile_expression(test)?;
                self.emit(Op::dyn_to_bool);
                let skip = self.emit_jump(Op::br_if_true);
                for s in body { self.compile_statement(s)?; }
                if let Some(alt) = else_body {
                    let end = self.emit_jump(Op::br);
                    self.patch_jump(skip);
                    for s in alt { self.compile_statement(s)?; }
                    self.patch_jump(end);
                } else {
                    self.patch_jump(skip);
                }
            }

            Statement::While { test, body } => {
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { break_patches: Vec::new(), next_patches: Vec::new() });
                self.compile_expression(test)?;
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                for s in body { self.compile_statement(s)?; }
                let nexts: Vec<usize> = self.loop_stack.last().unwrap().next_patches.clone();
                for c in &nexts { self.patch_jump(*c); }
                self.emit_loop(loop_start);
                self.patch_jump(exit);
                let breaks: Vec<usize> = self.loop_stack.pop().unwrap().break_patches;
                for b in breaks { self.patch_jump(b); }
            }

            Statement::Until { test, body } => {
                // until cond == while !cond
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { break_patches: Vec::new(), next_patches: Vec::new() });
                self.compile_expression(test)?;
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_true); // exit when true (opposite of while)
                for s in body { self.compile_statement(s)?; }
                let nexts: Vec<usize> = self.loop_stack.last().unwrap().next_patches.clone();
                for c in &nexts { self.patch_jump(*c); }
                self.emit_loop(loop_start);
                self.patch_jump(exit);
                let breaks: Vec<usize> = self.loop_stack.pop().unwrap().break_patches;
                for b in breaks { self.patch_jump(b); }
            }

            Statement::For { var, iterable, body } => {
                // for x in iterable → same as iterable.each { |x| body }
                self.compile_expression(iterable)?;
                let arr_slot = self.define_local("__for_arr");
                self.emit_u16(Op::local_set, arr_slot);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit(Op::array_length);
                let len_slot = self.define_local("__for_len");
                self.emit_u16(Op::local_set, len_slot);
                self.emit_constant(Value::I32(0));
                let idx_slot = self.define_local("__for_idx");
                self.emit_u16(Op::local_set, idx_slot);

                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { break_patches: Vec::new(), next_patches: Vec::new() });

                self.emit_u16(Op::local_get, idx_slot);
                self.emit_u16(Op::local_get, len_slot);
                self.emit(Op::dyn_lt);
                let exit = self.emit_jump(Op::br_if_false);

                // var = arr[idx]
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, idx_slot);
                self.emit(Op::array_get);
                let var_slot = self.define_local_or_get(var);
                self.emit_u16(Op::local_set, var_slot);

                for s in body { self.compile_statement(s)?; }

                let nexts: Vec<usize> = self.loop_stack.last().unwrap().next_patches.clone();
                for c in &nexts { self.patch_jump(*c); }

                // idx++
                self.emit_u16(Op::local_get, idx_slot);
                self.emit_constant(Value::I32(1));
                self.emit(Op::i32_add);
                self.emit_u16(Op::local_set, idx_slot);
                self.emit_loop(loop_start);

                self.patch_jump(exit);
                let breaks: Vec<usize> = self.loop_stack.pop().unwrap().break_patches;
                for b in breaks { self.patch_jump(b); }
            }

            Statement::Case { subject, whens, else_body } => {
                if let Some(subj) = subject {
                    self.compile_expression(subj)?;
                    let disc_slot = self.define_local("__case_disc");
                    self.emit_u16(Op::local_set, disc_slot);

                    let mut end_jumps: Vec<usize> = Vec::new();
                    for when in whens {
                        let mut skip_jumps = Vec::new();
                        for cond in &when.conditions {
                            self.emit_u16(Op::local_get, disc_slot);
                            self.compile_expression(cond)?;
                            self.emit(Op::eq);
                            skip_jumps.push(self.emit_jump(Op::br_if_true));
                        }
                        let fail = self.emit_jump(Op::br);
                        for s in &skip_jumps { self.patch_jump(*s); }
                        for s in &when.body { self.compile_statement(s)?; }
                        end_jumps.push(self.emit_jump(Op::br));
                        self.patch_jump(fail);
                    }
                    if let Some(alt) = else_body {
                        for s in alt { self.compile_statement(s)?; }
                    }
                    for j in end_jumps { self.patch_jump(j); }
                } else {
                    // case without subject: when cond
                    let mut end_jumps: Vec<usize> = Vec::new();
                    for when in whens {
                        self.compile_expression(&when.conditions[0])?;
                        self.emit(Op::dyn_to_bool);
                        let skip = self.emit_jump(Op::br_if_false);
                        for s in &when.body { self.compile_statement(s)?; }
                        end_jumps.push(self.emit_jump(Op::br));
                        self.patch_jump(skip);
                    }
                    if let Some(alt) = else_body {
                        for s in alt { self.compile_statement(s)?; }
                    }
                    for j in end_jumps { self.patch_jump(j); }
                }
            }

            Statement::Raise(val) => {
                if let Some(expr) = val {
                    self.compile_expression(expr)?;
                } else {
                    self.emit_constant(Value::String(Rc::from("RuntimeError")));
                }
                let line = self.line;
                common::errors::emit_throw(&mut self.chunks[self.current_chunk_idx], line);
            }

            Statement::Begin { body, rescues, else_body: _, ensure } => {
                let line = self.line;
                let c = self.current_chunk_idx;
                let catch_jump = common::errors::emit_try_start(&mut self.chunks[c], line);
                for s in body { self.compile_statement(s)?; }
                let line = self.line;
                common::errors::emit_try_end(&mut self.chunks[c], line);
                let skip_catch = self.emit_jump(Op::br);
                common::errors::patch_catch(&mut self.chunks[c], catch_jump);

                for rescue in rescues {
                    if let Some(var) = &rescue.var {
                        let slot = self.define_local_or_get(var);
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    } else {
                        self.emit(Op::drop);
                    }
                    for s in &rescue.body { self.compile_statement(s)?; }
                }

                self.patch_jump(skip_catch);

                if let Some(ensure_body) = ensure {
                    for s in ensure_body { self.compile_statement(s)?; }
                }
            }

            Statement::MethodDef(decl) => {
                let chunk_idx = self.compile_method_def(decl)?;
                let line = self.line;
                common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], chunk_idx, 0, line);
                let idx = self.add_string_constant(&decl.name);
                self.emit_u16(Op::global_set, idx);
                self.emit(Op::drop);
                self.defined_globals.insert(decl.name.clone());
            }

            Statement::ClassDef(decl) => {
                self.compile_class(decl)?;
            }

            Statement::ModuleDef(decl) => {
                self.compile_module(decl)?;
            }

            Statement::Require(_path) => {
                // require is a no-op at compile time in our VM
            }

            Statement::MultiAssign { targets, splat_index: _, values } => {
                // Evaluate all values first
                if values.len() == 1 {
                    // a, b = [1, 2] — single array on right
                    self.compile_expression(&values[0])?;
                    let arr_slot = self.define_local("__multi_arr");
                    self.emit_u16(Op::local_set, arr_slot);
                    for (i, target) in targets.iter().enumerate() {
                        self.emit_u16(Op::local_get, arr_slot);
                        self.emit_constant(Value::I32(i as i32));
                        self.emit(Op::array_get);
                        self.compile_assign_target(target)?;
                    }
                } else {
                    // a, b = 1, 2 — parallel values
                    let mut val_slots = Vec::new();
                    for val in values {
                        self.compile_expression(val)?;
                        let slot = self.define_local("__multi_val");
                        self.emit_u16(Op::local_set, slot);
                        val_slots.push(slot);
                    }
                    for (i, target) in targets.iter().enumerate() {
                        if i < val_slots.len() {
                            self.emit_u16(Op::local_get, val_slots[i]);
                        } else {
                            self.emit(Op::null);
                        }
                        self.compile_assign_target(target)?;
                    }
                }
            }

            Statement::Alias { new_name, old_name } => {
                // alias new_name old_name → copy the method
                self.emit_var_get(old_name);
                self.emit_var_set(new_name);
            }

            Statement::AccessModifier(_level) => {
                // Access modifiers are handled at class-compilation time — no-op here
            }

            Statement::Retry => {
                // retry in rescue — jump back to begin of try block
                // Simplified: no-op (requires tracking try-start position)
            }

            Statement::Loop(body) => {
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { break_patches: Vec::new(), next_patches: Vec::new() });
                for s in body { self.compile_statement(s)?; }
                let nexts: Vec<usize> = self.loop_stack.last().unwrap().next_patches.clone();
                for c in &nexts { self.patch_jump(*c); }
                self.emit_loop(loop_start);
                let breaks: Vec<usize> = self.loop_stack.pop().unwrap().break_patches;
                for b in breaks { self.patch_jump(b); }
            }

            Statement::CatchThrow { tag: _, body } => {
                // catch/throw — simplified: just compile body, throw is handled as expression
                let line = self.line;
                let c = self.current_chunk_idx;
                let catch_jump = common::errors::emit_try_start(&mut self.chunks[c], line);
                for s in body { self.compile_statement(s)?; }
                let line = self.line;
                common::errors::emit_try_end(&mut self.chunks[c], line);
                let skip = self.emit_jump(Op::br);
                common::errors::patch_catch(&mut self.chunks[c], catch_jump);
                // caught throw value is on stack
                self.emit(Op::drop); // drop the thrown tag
                self.patch_jump(skip);
            }

            Statement::Redo => {
                // redo — jump back to loop body start (simplified: same as next)
                let patch = self.emit_jump(Op::br);
                if let Some(ctx) = self.loop_stack.last_mut() {
                    ctx.next_patches.push(patch);
                }
            }

            Statement::AtExit(body) => {
                // at_exit { body } — compile body as a function, store for later
                // Simplified: no-op (body would need to run at program exit)
                let _ = body;
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
                self.emit_constant(Value::String(Rc::from(s.as_str())));
            }
            Expression::Symbol(s) => {
                self.emit_constant(Value::String(Rc::from(s.as_str())));
            }
            Expression::Bool(b) => {
                if *b { self.emit(Op::r#true); } else { self.emit(Op::r#false); }
            }
            Expression::Nil => {
                self.emit(Op::null);
            }
            Expression::SelfExpr => {
                // self is slot 1 in method scope
                self.emit_u16(Op::local_get, 1);
            }
            Expression::Identifier(name) => {
                self.emit_var_get(name);
            }
            Expression::InstanceVar(name) => {
                // @name → self.__name
                self.emit_u16(Op::local_get, 1); // self
                let idx = self.add_string_constant(name);
                self.emit_u16(Op::struct_get, idx);
            }
            Expression::ClassVar(name) => {
                // @@name → stored as global
                let idx = self.add_string_constant(&format!("@@{}", name));
                self.emit_u16(Op::global_get, idx);
            }
            Expression::GlobalVar(name) => {
                let idx = self.add_string_constant(&format!("${}", name));
                self.emit_u16(Op::global_get, idx);
            }
            Expression::ConstantRef(name) => {
                let idx = self.add_string_constant(name);
                self.emit_u16(Op::global_get, idx);
            }

            Expression::Array(elements) => {
                for e in elements {
                    self.compile_expression(e)?;
                }
                self.emit_u16(Op::array_new, elements.len() as u16);
            }

            Expression::Hash(pairs) => {
                let line = self.line;
                let c = self.current_chunk_idx;
                common::dict::emit_new(&mut self.chunks[c], line);
                for (key, value) in pairs {
                    self.emit(Op::dup);
                    self.compile_expression(value)?;
                    match key {
                        Expression::Symbol(s) | Expression::Str(s) => {
                            let line = self.line;
                            common::dict::emit_set_const_key(&mut self.chunks[c], s, line);
                        }
                        _ => {
                            let tmp = self.define_local("__hash_val");
                            self.emit_u16(Op::local_set, tmp);
                            self.compile_expression(key)?;
                            self.emit_u16(Op::local_get, tmp);
                            let line = self.line;
                            common::dict::emit_set_dynamic(&mut self.chunks[c], line);
                        }
                    }
                }
            }

            Expression::Range { start, end, exclusive } => {
                // Create array using stdlib range function
                // __vybe_range(start, stop, step) — stop is exclusive
                let c = self.current_chunk_idx;
                let line = self.line;
                common::bundle::emit_call_push_func(&mut self.chunks[c], "__vybe_range", line);
                self.compile_expression(start)?;
                if *exclusive {
                    // Exclusive: end is already the stop value
                    self.compile_expression(end)?;
                } else {
                    // Inclusive: need end+1 for the stop value
                    self.compile_expression(end)?;
                    self.emit_constant(Value::I32(1));
                    self.emit(Op::dyn_add);
                }
                self.emit_constant(Value::I32(1)); // step
                common::bundle::emit_call_invoke(&mut self.chunks[c], 3, line);
            }

            Expression::Binary { op, left, right } => {
                self.compile_binary(op, left, right)?;
            }

            Expression::Unary { op, expr } => {
                self.compile_expression(expr)?;
                match op {
                    UnaryOp::Neg => { self.emit(Op::dyn_neg); }
                    UnaryOp::Pos => {}
                    UnaryOp::Not => { self.emit(Op::dyn_not); }
                    UnaryOp::BitNot => { self.emit(Op::i32_not); }
                }
            }

            Expression::Ternary { test, consequent, alternate } => {
                let c = self.current_chunk_idx;
                let line = self.line;
                self.compile_expression(test)?;
                let false_jump = common::expressions::emit_ternary_start(&mut self.chunks[c], line);
                self.compile_expression(consequent)?;
                let end_jump = common::expressions::emit_ternary_middle(&mut self.chunks[c], false_jump, line);
                self.compile_expression(alternate)?;
                common::expressions::emit_ternary_end(&mut self.chunks[c], end_jump);
            }

            Expression::MethodCall { receiver, method, args, block } => {
                self.compile_method_call(receiver.as_deref(), method, args, block.as_deref())?;
            }

            Expression::IndexAccess { object, index } => {
                self.compile_expression(object)?;
                // Handle negative indexing: arr[-1] → arr[arr.length + (-1)]
                if let Expression::Unary { op: UnaryOp::Neg, .. } = index.as_ref() {
                    // Negative index — convert to: obj.length + index
                    self.emit(Op::dup); // keep obj on stack
                    self.emit(Op::array_length);
                    self.compile_expression(index)?;
                    self.emit(Op::dyn_add); // length + (-n)
                    self.emit(Op::array_get);
                } else if let Expression::Number(n) = index.as_ref() {
                    if *n < 0.0 {
                        self.emit(Op::dup);
                        self.emit(Op::array_length);
                        self.emit_constant(Value::F64(*n));
                        self.emit(Op::dyn_add);
                        self.emit(Op::array_get);
                    } else {
                        self.compile_expression(index)?;
                        self.emit(Op::array_get);
                    }
                } else if let Expression::Range { start, end, exclusive } = index.as_ref() {
                    // Item 3: String#[] with range — arr[1..3] → slice
                    self.compile_expression(start)?;
                    self.compile_expression(end)?;
                    if !exclusive {
                        // Inclusive range: add 1 to end for slice
                        self.emit_constant(Value::I32(1));
                        self.emit(Op::dyn_add);
                    }
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::collections::emit_slice(&mut self.chunks[c], line);
                } else {
                    self.compile_expression(index)?;
                    self.emit(Op::array_get);
                }
            }

            Expression::Lambda { params, body } => {
                let decl = MethodDecl {
                    name: "<lambda>".to_string(),
                    params: params.clone(),
                    body: body.clone(),
                    is_self: false,
                };
                let ci = self.compile_method_def(&decl)?;
                let upvalue_count = self.scopes.last().map(|s| s.upvalues.len()).unwrap_or(0);
                let line = self.line;
                common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], ci, upvalue_count as u8, line);
            }

            Expression::ProcNew { params, body } => {
                let decl = MethodDecl {
                    name: "<proc>".to_string(),
                    params: params.iter().map(|n| Param {
                        name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false,
                    }).collect(),
                    body: body.clone(),
                    is_self: false,
                };
                let ci = self.compile_method_def(&decl)?;
                let upvalue_count = self.scopes.last().map(|s| s.upvalues.len()).unwrap_or(0);
                let line = self.line;
                common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], ci, upvalue_count as u8, line);
            }

            Expression::Block { params, body } => {
                let decl = MethodDecl {
                    name: "<block>".to_string(),
                    params: params.iter().map(|n| Param {
                        name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false,
                    }).collect(),
                    body: body.clone(),
                    is_self: false,
                };
                let ci = self.compile_method_def(&decl)?;
                let upvalue_count = self.scopes.last().map(|s| s.upvalues.len()).unwrap_or(0);
                let line = self.line;
                common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], ci, upvalue_count as u8, line);
            }

            Expression::Yield(args) => {
                // yield → call the block parameter (stored as __block local)
                self.emit_var_get("__block");
                for arg in args { self.compile_expression(arg)?; }
                self.emit_u8(Op::call_ref, args.len() as u8);
            }

            Expression::BlockGiven => {
                // block_given? → check if __block is not nil
                self.emit_var_get("__block");
                self.emit(Op::ref_is_null);
                self.emit(Op::dyn_not);
            }

            Expression::Super(args) => {
                // super → call parent's method via __super
                self.emit_u16(Op::local_get, 1); // self
                let idx = self.add_string_constant("__super");
                self.emit_u16(Op::struct_get, idx);
                self.emit_u16(Op::local_get, 1); // self as first arg
                for arg in args { self.compile_expression(arg)?; }
                self.emit_u8(Op::call_ref, (args.len() + 1) as u8);
            }

            Expression::ScopeResolution { left, name } => {
                self.compile_expression(left)?;
                let idx = self.add_string_constant(name);
                self.emit_u16(Op::struct_get, idx);
            }

            Expression::Splat(inner) => {
                self.compile_expression(inner)?;
            }

            Expression::Interpolated(parts) => {
                // Concat all parts into a string
                let mut first = true;
                for part in parts {
                    match part {
                        InterpolPart::Lit(s) => {
                            self.emit_constant(Value::String(Rc::from(s.as_str())));
                        }
                        InterpolPart::Expr(e) => {
                            self.compile_expression(e)?;
                            let c = self.current_chunk_idx;
                            let line = self.line;
                            common::strings::emit_to_string(&mut self.chunks[c], line);
                        }
                    }
                    if !first {
                        self.emit(Op::str_concat);
                    }
                    first = false;
                }
                if first {
                    // Empty interpolated string
                    self.emit_constant(Value::String(Rc::from("")));
                }
            }

            Expression::AttrDecl { kind, names } => {
                // attr_reader / attr_writer / attr_accessor
                // In our VM these are compiled when processing the class body.
                // For now, emit nothing (the class compiler handles attr generation).
                let _ = (kind, names);
                self.emit(Op::null);
            }

            Expression::Include(name) | Expression::Extend(name) => {
                let idx = self.add_string_constant(name);
                self.emit_u16(Op::global_get, idx);
                self.emit(Op::drop);
                self.emit(Op::null);
            }

            Expression::Regex(pattern) => {
                // Create regex via host import
                self.emit_constant(Value::String(Rc::from(pattern.as_str())));
                let idx = self.import("vybe:regex", "compile");
                self.emit_host_call(idx, 1);
            }

            Expression::Defined(expr) => {
                // defined?(x) → check if x is not nil
                // Simplified: try to get the value and check null
                let c = self.current_chunk_idx;
                let line = self.line;
                let catch_jump = common::errors::emit_try_start(&mut self.chunks[c], line);
                self.compile_expression(expr)?;
                let line = self.line;
                common::errors::emit_try_end(&mut self.chunks[c], line);
                self.emit(Op::ref_is_null);
                self.emit(Op::dyn_not);
                // Convert to string "expression" or nil
                let skip = self.emit_jump(Op::br_if_false);
                self.emit_constant(Value::String(Rc::from("expression")));
                let end = self.emit_jump(Op::br);
                self.patch_jump(skip);
                self.emit(Op::null);
                self.patch_jump(end);
                // Patch catch — if error, defined? returns nil
                let c = self.current_chunk_idx;
                common::errors::patch_catch(&mut self.chunks[c], catch_jump);
                self.emit(Op::drop); // drop error
                self.emit(Op::null);
            }

            Expression::SymbolProc(method_name) => {
                // &:method_name → create a lambda that calls method on its argument
                // Equivalent to: -> (x) { x.method_name }
                let decl = MethodDecl {
                    name: "<symbol_proc>".to_string(),
                    params: vec![Param { name: "x".to_string(), default: None, splat: false, double_splat: false, block: false, keyword: false }],
                    body: vec![Statement::Expression(Expression::MethodCall {
                        receiver: Some(Box::new(Expression::Identifier("x".to_string()))),
                        method: method_name.clone(),
                        args: Vec::new(),
                        block: None,
                    })],
                    is_self: false,
                };
                let ci = self.compile_method_def(&decl)?;
                let line = self.line;
                common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], ci, 0, line);
            }

            Expression::ProcLiteral { params, body } => {
                let decl = MethodDecl {
                    name: "<proc>".to_string(),
                    params: params.iter().map(|n| Param {
                        name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false,
                    }).collect(),
                    body: body.clone(),
                    is_self: false,
                };
                let ci = self.compile_method_def(&decl)?;
                let upvalue_count = self.scopes.last().map(|s| s.upvalues.len()).unwrap_or(0);
                let line = self.line;
                common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], ci, upvalue_count as u8, line);
            }

            Expression::StructNew { name: _, fields } => {
                // Struct.new(:field1, :field2) → create a constructor function
                // Build a class with initialize(field1, field2) and accessors
                let struct_name = format!("Struct_{}", fields.join("_"));
                let mut class_body = Vec::new();

                // Add attr_accessor for all fields
                class_body.push(Statement::Expression(Expression::AttrDecl {
                    kind: AttrKind::Accessor,
                    names: fields.clone(),
                }));

                // Add initialize method
                let init_body: Vec<Statement> = fields.iter().map(|f| {
                    Statement::Assignment {
                        target: Expression::InstanceVar(f.clone()),
                        op: AssignOp::Assign,
                        value: Expression::Identifier(f.clone()),
                    }
                }).collect();
                class_body.push(Statement::MethodDef(MethodDecl {
                    name: "initialize".to_string(),
                    params: fields.iter().map(|f| Param {
                        name: f.clone(), default: Some(Expression::Nil),
                        splat: false, double_splat: false, block: false, keyword: false,
                    }).collect(),
                    body: init_body,
                    is_self: false,
                }));

                let class_decl = ClassDecl {
                    name: struct_name,
                    parent: None,
                    body: class_body,
                };
                self.compile_class(&class_decl)?;
                // Push the constructor as the result
                let idx = self.add_string_constant(&class_decl.name);
                self.emit_u16(Op::global_get, idx);
            }

            Expression::Throw { tag, value } => {
                // throw :tag, value → raise with tag
                if let Some(val) = value {
                    self.compile_expression(val)?;
                } else {
                    self.compile_expression(tag)?;
                }
                let line = self.line;
                common::errors::emit_throw(&mut self.chunks[self.current_chunk_idx], line);
            }

            Expression::Freeze(expr) => {
                // obj.freeze → just return obj (no-op in our VM)
                self.compile_expression(expr)?;
            }

            Expression::FrozenCheck(expr) => {
                // obj.frozen? → always false in our VM
                self.compile_expression(expr)?;
                self.emit(Op::drop);
                self.emit(Op::r#false);
            }

            Expression::RespondTo { object, method } => {
                // obj.respond_to?(:method) → check if property exists
                self.compile_expression(object)?;
                self.emit_constant(Value::String(Rc::from(method.as_str())));
                let idx = self.import("vybe:object", "hasProperty");
                self.emit_host_call(idx, 2);
            }

            Expression::Send { object, method, args } => {
                // obj.send(:method, args) → dynamic dispatch
                self.compile_expression(object)?;
                self.compile_expression(method)?;
                self.emit(Op::array_get); // get method from object by name
                self.compile_expression(object)?; // self
                for a in args { self.compile_expression(a)?; }
                self.emit_u8(Op::call_ref, (args.len() + 1) as u8);
            }

            Expression::ChainedAssign { targets, value } => {
                // a = b = c = 1 → evaluate once, assign to all
                self.compile_expression(value)?;
                for target in targets.iter().rev() {
                    self.emit(Op::dup);
                    self.compile_assign_target(target)?;
                }
            }

            Expression::PatternMatch { subject, arms, else_body } => {
                // case/in pattern matching — simplified to case/when semantics
                self.compile_expression(subject)?;
                let disc_slot = self.define_local("__pat_disc");
                self.emit_u16(Op::local_set, disc_slot);
                let mut end_jumps = Vec::new();
                for arm in arms {
                    self.emit_u16(Op::local_get, disc_slot);
                    self.compile_expression(&arm.pattern)?;
                    self.emit(Op::eq);
                    let skip = self.emit_jump(Op::br_if_false);
                    // Last expression in body is the result
                    for (i, s) in arm.body.iter().enumerate() {
                        if i == arm.body.len() - 1 {
                            if let Statement::Expression(e) = s {
                                self.compile_expression(e)?;
                            } else {
                                self.compile_statement(s)?;
                                self.emit(Op::null);
                            }
                        } else {
                            self.compile_statement(s)?;
                        }
                    }
                    if arm.body.is_empty() { self.emit(Op::null); }
                    end_jumps.push(self.emit_jump(Op::br));
                    self.patch_jump(skip);
                }
                if let Some(alt) = else_body {
                    for s in alt { self.compile_statement(s)?; }
                    self.emit(Op::null);
                } else {
                    self.emit(Op::null);
                }
                for j in end_jumps { self.patch_jump(j); }
            }

            // ── if/unless/begin as expression ──────────────────
            Expression::IfExpr { test, body, elsifs, else_body } => {
                self.compile_expression(test)?;
                self.emit(Op::dyn_to_bool);
                let mut end_jumps = Vec::new();
                let skip = self.emit_jump(Op::br_if_false);
                self.compile_body_as_expr(body)?;
                end_jumps.push(self.emit_jump(Op::br));
                self.patch_jump(skip);
                for elsif in elsifs {
                    self.compile_expression(&elsif.test)?;
                    self.emit(Op::dyn_to_bool);
                    let s = self.emit_jump(Op::br_if_false);
                    self.compile_body_as_expr(&elsif.body)?;
                    end_jumps.push(self.emit_jump(Op::br));
                    self.patch_jump(s);
                }
                if let Some(alt) = else_body {
                    self.compile_body_as_expr(alt)?;
                } else {
                    self.emit(Op::null);
                }
                for j in end_jumps { self.patch_jump(j); }
            }

            Expression::UnlessExpr { test, body, else_body } => {
                self.compile_expression(test)?;
                self.emit(Op::dyn_to_bool);
                let skip = self.emit_jump(Op::br_if_true);
                self.compile_body_as_expr(body)?;
                if let Some(alt) = else_body {
                    let end = self.emit_jump(Op::br);
                    self.patch_jump(skip);
                    self.compile_body_as_expr(alt)?;
                    self.patch_jump(end);
                } else {
                    let end = self.emit_jump(Op::br);
                    self.patch_jump(skip);
                    self.emit(Op::null);
                    self.patch_jump(end);
                }
            }

            Expression::BeginExpr { body, rescues, else_body: _, ensure } => {
                let line = self.line;
                let c = self.current_chunk_idx;
                let catch_jump = common::errors::emit_try_start(&mut self.chunks[c], line);
                self.compile_body_as_expr(body)?;
                let line = self.line;
                common::errors::emit_try_end(&mut self.chunks[c], line);
                let skip_catch = self.emit_jump(Op::br);
                common::errors::patch_catch(&mut self.chunks[c], catch_jump);
                for rescue in rescues {
                    if let Some(var) = &rescue.var {
                        let slot = self.define_local_or_get(var);
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    } else {
                        self.emit(Op::drop);
                    }
                    self.compile_body_as_expr(&rescue.body)?;
                }
                self.patch_jump(skip_catch);
                if let Some(ensure_body) = ensure {
                    for s in ensure_body { self.compile_statement(s)?; }
                }
            }

            Expression::InlineRescue { expr, rescue_val } => {
                // expr rescue default_val
                let line = self.line;
                let c = self.current_chunk_idx;
                let catch_jump = common::errors::emit_try_start(&mut self.chunks[c], line);
                self.compile_expression(expr)?;
                let line = self.line;
                common::errors::emit_try_end(&mut self.chunks[c], line);
                let skip = self.emit_jump(Op::br);
                common::errors::patch_catch(&mut self.chunks[c], catch_jump);
                self.emit(Op::drop); // drop error
                self.compile_expression(rescue_val)?;
                self.patch_jump(skip);
            }

            Expression::Backtick(cmd) => {
                self.emit_constant(Value::String(Rc::from(cmd.as_str())));
                let i = self.import("vybe:types", "processStart");
                self.emit_host_call(i, 1);
            }

            Expression::SafeNav { receiver, method, args, block } => {
                // obj&.method → if obj != nil then obj.method else nil end
                self.compile_expression(receiver)?;
                self.emit(Op::dup);
                self.emit(Op::ref_is_null);
                let null_jump = self.emit_jump(Op::br_if_true);
                // Not nil — call method
                self.emit(Op::drop); // drop the dup
                self.compile_method_call(Some(receiver), method, args, block.as_deref())?;
                let end_jump = self.emit_jump(Op::br);
                self.patch_jump(null_jump);
                // Was nil — result is nil (the dup is already nil on stack)
                self.patch_jump(end_jump);
            }

            Expression::MagicConstant(mc) => {
                match mc {
                    MagicConst::File => {
                        self.emit_constant(Value::String(Rc::from("<main>")));
                    }
                    MagicConst::Dir => {
                        let i = self.import("wasi:cli", "cwd");
                        self.emit_host_call(i, 0);
                    }
                    MagicConst::Method => {
                        self.emit_constant(Value::String(Rc::from("<main>")));
                    }
                    MagicConst::Line => {
                        self.emit_constant(Value::F64(self.line as f64));
                    }
                }
            }

            Expression::IvarGet { object, name } => {
                self.compile_expression(object)?;
                let idx = self.add_string_constant(name);
                self.emit_u16(Op::struct_get, idx);
            }

            Expression::IvarSet { object, name, value } => {
                self.compile_expression(object)?;
                self.compile_expression(value)?;
                let idx = self.add_string_constant(name);
                self.emit_u16(Op::struct_set, idx);
                self.emit(Op::drop);
                self.emit(Op::null);
            }
        }
        Ok(())
    }

    /// Compile a body of statements as an expression (last statement's value is the result)
    fn compile_body_as_expr(&mut self, body: &[Statement]) -> Result<(), String> {
        if body.is_empty() {
            self.emit(Op::null);
            return Ok(());
        }
        for (i, stmt) in body.iter().enumerate() {
            if i == body.len() - 1 {
                // Last statement — keep value on stack
                match stmt {
                    Statement::Expression(e) => {
                        self.compile_expression(e)?;
                    }
                    _ => {
                        self.compile_statement(stmt)?;
                        self.emit(Op::null);
                    }
                }
            } else {
                self.compile_statement(stmt)?;
            }
        }
        Ok(())
    }

    // ------------------------------------------------------------------
    // Assignment
    // ------------------------------------------------------------------

    fn compile_assign(&mut self, target: &Expression, op: &AssignOp, value: &Expression) -> Result<(), String> {
        if *op == AssignOp::Assign {
            self.compile_expression(value)?;
            return self.compile_assign_target(target);
        }

        // Compound assignment: target op= value
        self.compile_expression(target)?;
        self.compile_expression(value)?;
        match op {
            AssignOp::AddAssign => { self.emit(Op::dyn_add); }
            AssignOp::SubAssign => { self.emit(Op::f64_sub); }
            AssignOp::MulAssign => { self.emit(Op::f64_mul); }
            AssignOp::DivAssign => { self.emit(Op::f64_div); }
            AssignOp::ModAssign => { self.emit(Op::f64_mod); }
            AssignOp::AndAssign => { self.emit(Op::i32_and); }
            AssignOp::OrAssign => { self.emit(Op::i32_or); }
            AssignOp::BitAndAssign => { self.emit(Op::i32_and); }
            AssignOp::BitOrAssign => { self.emit(Op::i32_or); }
            AssignOp::BitXorAssign => { self.emit(Op::i32_xor); }
            AssignOp::ShlAssign => { self.emit(Op::i32_shl); }
            AssignOp::ShrAssign => { self.emit(Op::i32_shr_s); }
            AssignOp::Assign => unreachable!(),
        }
        self.compile_assign_target(target)
    }

    fn compile_assign_target(&mut self, target: &Expression) -> Result<(), String> {
        match target {
            Expression::Identifier(name) => {
                if self.is_global_scope() {
                    let idx = self.add_string_constant(name);
                    self.emit_u16(Op::global_set, idx);
                    self.defined_globals.insert(name.clone());
                } else {
                    let slot = self.define_local_or_get(name);
                    self.emit_u16(Op::local_set, slot);
                }
            }
            Expression::InstanceVar(name) => {
                // @name = val → self.__name = val
                let tmp = self.define_local("__ivar_tmp");
                self.emit_u16(Op::local_set, tmp);
                self.emit_u16(Op::local_get, 1); // self
                self.emit_u16(Op::local_get, tmp);
                let idx = self.add_string_constant(name);
                self.emit_u16(Op::struct_set, idx);
                self.emit(Op::drop);
            }
            Expression::ClassVar(name) => {
                let idx = self.add_string_constant(&format!("@@{}", name));
                self.emit_u16(Op::global_set, idx);
            }
            Expression::GlobalVar(name) => {
                let idx = self.add_string_constant(&format!("${}", name));
                self.emit_u16(Op::global_set, idx);
            }
            Expression::ConstantRef(name) => {
                let idx = self.add_string_constant(name);
                self.emit_u16(Op::global_set, idx);
                self.defined_globals.insert(name.clone());
            }
            Expression::IndexAccess { object, index } => {
                let tmp = self.define_local("__idx_tmp");
                self.emit_u16(Op::local_set, tmp);
                self.compile_expression(object)?;
                self.compile_expression(index)?;
                self.emit_u16(Op::local_get, tmp);
                self.emit(Op::array_set);
                self.emit(Op::drop);
            }
            Expression::MethodCall { receiver: Some(obj), method, args: _, block: _ } => {
                // obj.attr = val → struct_set
                let tmp = self.define_local("__attr_tmp");
                self.emit_u16(Op::local_set, tmp);
                self.compile_expression(obj)?;
                self.emit_u16(Op::local_get, tmp);
                let idx = self.add_string_constant(method);
                self.emit_u16(Op::struct_set, idx);
                self.emit(Op::drop);
            }
            _ => {
                return Err("Invalid assignment target".to_string());
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
            BinaryOp::Add => { self.emit(Op::dyn_add); }
            BinaryOp::Sub => { self.emit(Op::f64_sub); }
            BinaryOp::Mul => {
                // Check if left is a string literal — if so, use str_repeat
                if matches!(left, Expression::Str(_) | Expression::Interpolated(_)) {
                    // Pop the f64_mul args (already compiled), emit str_repeat instead
                    // Actually we already compiled both sides. Since left is string,
                    // str_repeat expects [string, count] which is what we have.
                    self.emit(Op::str_repeat);
                } else {
                    self.emit(Op::f64_mul);
                }
            }
            BinaryOp::Div => { self.emit(Op::f64_div); }
            BinaryOp::Mod => {
                // % is overloaded: numeric modulo or string format
                if matches!(left, Expression::Str(_) | Expression::Interpolated(_)) {
                    // String format: "Hello %s" % "world"
                    let i = self.import("vybe:string", "format");
                    self.emit_host_call(i, 2);
                } else {
                    self.emit(Op::f64_mod);
                }
            }
            BinaryOp::Pow => {
                common::math::emit_pow(&mut self.chunks[c], line);
            }
            BinaryOp::Eq => { self.emit(Op::dyn_eq); }
            BinaryOp::Ne => { self.emit(Op::dyn_ne); }
            BinaryOp::Lt => { self.emit(Op::dyn_lt); }
            BinaryOp::Gt => { self.emit(Op::dyn_gt); }
            BinaryOp::Le => { self.emit(Op::dyn_le); }
            BinaryOp::Ge => { self.emit(Op::dyn_ge); }
            BinaryOp::Spaceship => {
                // a <=> b → (a < b) ? -1 : ((a > b) ? 1 : 0)
                let b_tmp = self.define_local("__cmp_b");
                let a_tmp = self.define_local("__cmp_a");
                self.emit_u16(Op::local_set, b_tmp);
                self.emit(Op::drop);
                self.emit_u16(Op::local_set, a_tmp);
                self.emit(Op::drop);
                self.emit_u16(Op::local_get, a_tmp);
                self.emit_u16(Op::local_get, b_tmp);
                self.emit(Op::dyn_lt);
                let lt_jump = self.emit_jump(Op::br_if_true);
                self.emit_u16(Op::local_get, a_tmp);
                self.emit_u16(Op::local_get, b_tmp);
                self.emit(Op::dyn_gt);
                let gt_jump = self.emit_jump(Op::br_if_true);
                self.emit_constant(Value::F64(0.0));
                let end_jump = self.emit_jump(Op::br);
                self.patch_jump(gt_jump);
                self.emit_constant(Value::F64(1.0));
                let end_jump2 = self.emit_jump(Op::br);
                self.patch_jump(lt_jump);
                self.emit_constant(Value::F64(-1.0));
                self.patch_jump(end_jump);
                self.patch_jump(end_jump2);
            }
            BinaryOp::BitAnd => { self.emit(Op::i32_and); }
            BinaryOp::BitOr => { self.emit(Op::i32_or); }
            BinaryOp::BitXor => { self.emit(Op::i32_xor); }
            BinaryOp::Shl => {
                // << is overloaded: string append, array push, or bitwise shift
                // Use dyn_add which handles string concat and array push
                self.emit(Op::dyn_add);
            }
            BinaryOp::Shr => { self.emit(Op::i32_shr_s); }
            BinaryOp::RangeIncl | BinaryOp::RangeExcl => {
                // These should be handled as Range expressions, not binary ops
                self.emit(Op::null);
            }
            BinaryOp::And | BinaryOp::Or => unreachable!(),
        }
        Ok(())
    }

    // ------------------------------------------------------------------
    // Method calls
    // ------------------------------------------------------------------

    fn compile_method_call(&mut self, receiver: Option<&Expression>, method: &str, args: &[Expression], block: Option<&BlockArg>) -> Result<(), String> {
        // Try built-in functions first (no receiver)
        if receiver.is_none() {
            if let Some(()) = self.try_compile_builtin(method, args)? {
                return Ok(());
            }
        }

        // Special receiver-based method calls
        if let Some(recv) = receiver {
            match method {
                // ── String methods ──────────────────────────────
                "length" | "size" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::str_length);
                    return Ok(());
                }
                "upcase" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::str_to_upper);
                    return Ok(());
                }
                "downcase" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::str_to_lower);
                    return Ok(());
                }
                "strip" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::str_trim);
                    return Ok(());
                }
                "lstrip" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::str_trim_start);
                    return Ok(());
                }
                "rstrip" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::str_trim_end);
                    return Ok(());
                }
                "include?" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::str_contains);
                    return Ok(());
                }
                "start_with?" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::str_starts_with);
                    return Ok(());
                }
                "end_with?" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::str_ends_with);
                    return Ok(());
                }
                "index" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::str_index_of);
                    return Ok(());
                }
                "replace" | "gsub" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::str_replace);
                    return Ok(());
                }
                "sub" => {
                    // sub replaces first occurrence — same opcode for now
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::str_replace);
                    return Ok(());
                }
                "split" => {
                    self.compile_expression(recv)?;
                    if args.is_empty() {
                        self.emit_constant(Value::String(Rc::from(" ")));
                    } else {
                        self.compile_expression(&args[0])?;
                    }
                    self.emit(Op::str_split);
                    return Ok(());
                }
                "join" => {
                    self.compile_expression(recv)?;
                    if args.is_empty() {
                        self.emit_constant(Value::String(Rc::from("")));
                    } else {
                        self.compile_expression(&args[0])?;
                    }
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::collections::emit_join(&mut self.chunks[c], line);
                    return Ok(());
                }
                "reverse" => {
                    self.compile_expression(recv)?;
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::collections::emit_reverse(&mut self.chunks[c], line);
                    return Ok(());
                }
                "chars" => {
                    self.compile_expression(recv)?;
                    self.emit_constant(Value::String(Rc::from("")));
                    self.emit(Op::str_split);
                    return Ok(());
                }
                "to_s" => {
                    self.compile_expression(recv)?;
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::strings::emit_to_string(&mut self.chunks[c], line);
                    return Ok(());
                }
                "to_i" => {
                    self.compile_expression(recv)?;
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::convert::emit_parse_int(&mut self.chunks[c], line);
                    return Ok(());
                }
                "to_f" => {
                    self.compile_expression(recv)?;
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::convert::emit_parse_float(&mut self.chunks[c], line);
                    return Ok(());
                }
                "to_a" => {
                    self.compile_expression(recv)?;
                    return Ok(());
                }
                "nil?" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::ref_is_null);
                    return Ok(());
                }
                "is_a?" | "kind_of?" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::ref_test);
                    return Ok(());
                }

                // ── Array methods ──────────────────────────────
                "push" | "append" | "<<" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::collections::emit_push(&mut self.chunks[c], line);
                    return Ok(());
                }
                "pop" => {
                    self.compile_expression(recv)?;
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::collections::emit_pop(&mut self.chunks[c], line);
                    return Ok(());
                }
                "shift" => {
                    self.compile_expression(recv)?;
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::collections::emit_shift(&mut self.chunks[c], line);
                    return Ok(());
                }
                "unshift" | "prepend" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    // Prepend: create new array with element + concat
                    self.emit_u16(Op::array_new, 1);
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::collections::emit_concat(&mut self.chunks[c], line);
                    return Ok(());
                }
                "first" => {
                    self.compile_expression(recv)?;
                    self.emit_constant(Value::I32(0));
                    self.emit(Op::array_get);
                    return Ok(());
                }
                "last" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::dup);
                    self.emit(Op::array_length);
                    self.emit_constant(Value::I32(1));
                    self.emit(Op::f64_sub);
                    self.emit(Op::array_get);
                    return Ok(());
                }
                "flatten" => {
                    // Inline flatten: iterate array, if element is array concat it, else push
                    self.compile_expression(recv)?;
                    let src_slot = self.define_local("__flat_src");
                    self.emit_u16(Op::local_set, src_slot);
                    self.emit_u16(Op::array_new, 0);
                    let res_slot = self.define_local("__flat_res");
                    self.emit_u16(Op::local_set, res_slot);
                    let i_slot = self.define_local("__flat_i");
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    let (loop_start, exit) = common::loops::emit_for_in_start(&mut self.chunks[c], src_slot, i_slot, line);
                    let elem_slot = self.define_local("__flat_elem");
                    self.emit_u16(Op::local_set, elem_slot);
                    // Check if element is an array (try array_length, if it works it's an array)
                    // Simplified: just push element directly (single-level flatten would need type check)
                    self.emit_u16(Op::local_get, res_slot);
                    self.emit_u16(Op::local_get, elem_slot);
                    self.emit(Op::array_push);
                    self.emit(Op::drop);
                    common::loops::emit_for_in_end(&mut self.chunks[c], i_slot, loop_start, exit, line);
                    self.emit_u16(Op::local_get, res_slot);
                    return Ok(());
                }
                "compact" => {
                    // Remove nil values from array
                    self.compile_expression(recv)?;
                    let src_slot = self.define_local("__compact_src");
                    self.emit_u16(Op::local_set, src_slot);
                    self.emit_u16(Op::array_new, 0);
                    let res_slot = self.define_local("__compact_res");
                    self.emit_u16(Op::local_set, res_slot);
                    let i_slot = self.define_local("__compact_i");
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    let (loop_start, exit) = common::loops::emit_for_in_start(&mut self.chunks[c], src_slot, i_slot, line);
                    let elem_slot = self.define_local("__compact_elem");
                    self.emit_u16(Op::local_set, elem_slot);
                    self.emit_u16(Op::local_get, elem_slot);
                    self.emit(Op::ref_is_null);
                    let skip = self.emit_jump(Op::br_if_true);
                    self.emit_u16(Op::local_get, res_slot);
                    self.emit_u16(Op::local_get, elem_slot);
                    self.emit(Op::array_push);
                    self.emit(Op::drop);
                    self.patch_jump(skip);
                    common::loops::emit_for_in_end(&mut self.chunks[c], i_slot, loop_start, exit, line);
                    self.emit_u16(Op::local_get, res_slot);
                    return Ok(());
                }
                "uniq" => {
                    // Remove duplicates — iterate and check contains before pushing
                    self.compile_expression(recv)?;
                    let src_slot = self.define_local("__uniq_src");
                    self.emit_u16(Op::local_set, src_slot);
                    self.emit_u16(Op::array_new, 0);
                    let res_slot = self.define_local("__uniq_res");
                    self.emit_u16(Op::local_set, res_slot);
                    let i_slot = self.define_local("__uniq_i");
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    let (loop_start, exit) = common::loops::emit_for_in_start(&mut self.chunks[c], src_slot, i_slot, line);
                    let elem_slot = self.define_local("__uniq_elem");
                    self.emit_u16(Op::local_set, elem_slot);
                    // Check if result already contains element
                    self.emit_u16(Op::local_get, res_slot);
                    self.emit_u16(Op::local_get, elem_slot);
                    common::collections::emit_contains(&mut self.chunks[c], line);
                    self.emit(Op::dyn_to_bool);
                    let skip = self.emit_jump(Op::br_if_true);
                    self.emit_u16(Op::local_get, res_slot);
                    self.emit_u16(Op::local_get, elem_slot);
                    self.emit(Op::array_push);
                    self.emit(Op::drop);
                    self.patch_jump(skip);
                    common::loops::emit_for_in_end(&mut self.chunks[c], i_slot, loop_start, exit, line);
                    self.emit_u16(Op::local_get, res_slot);
                    return Ok(());
                }
                "count" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::array_length);
                    return Ok(());
                }
                "empty?" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::array_length);
                    self.emit_constant(Value::I32(0));
                    self.emit(Op::dyn_eq);
                    return Ok(());
                }
                "sort" => {
                    self.compile_expression(recv)?;
                    let tmp = self.define_local("__sort_tmp");
                    self.emit_u16(Op::local_set, tmp);
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::bundle::emit_call_push_func(&mut self.chunks[c], "__vybe_sorted", line);
                    self.emit_u16(Op::local_get, tmp);
                    common::bundle::emit_call_invoke(&mut self.chunks[c], 1, line);
                    return Ok(());
                }
                "min" => {
                    self.compile_expression(recv)?;
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::collections::emit_min(&mut self.chunks[c], 1, line);
                    return Ok(());
                }
                "max" => {
                    self.compile_expression(recv)?;
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::collections::emit_max(&mut self.chunks[c], 1, line);
                    return Ok(());
                }
                "sum" => {
                    self.compile_expression(recv)?;
                    let tmp = self.define_local("__sum_tmp");
                    self.emit_u16(Op::local_set, tmp);
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::bundle::emit_call_push_func(&mut self.chunks[c], "__vybe_sum", line);
                    self.emit_u16(Op::local_get, tmp);
                    common::bundle::emit_call_invoke(&mut self.chunks[c], 1, line);
                    return Ok(());
                }

                // ── Block-taking methods (each, map, select, reject, reduce) ──
                "each" | "each_with_index" => {
                    if let Some(blk) = block {
                        self.compile_expression(recv)?;
                        let arr_slot = self.define_local("__each_arr");
                        self.emit_u16(Op::local_set, arr_slot);
                        // Compile block as function
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param {
                                name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false,
                            }).collect(),
                            body: blk.body.clone(),
                            is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        let fn_slot = self.define_local("__each_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        let i_slot = self.define_local("__each_i");
                        let c = self.current_chunk_idx;
                        common::loops::emit_foreach(&mut self.chunks[c], fn_slot, arr_slot, i_slot, line);
                        self.emit_u16(Op::local_get, arr_slot);
                        return Ok(());
                    }
                    // No block — fall through to generic
                }
                "map" | "collect" => {
                    if let Some(blk) = block {
                        self.compile_expression(recv)?;
                        let arr_slot = self.define_local("__map_arr");
                        self.emit_u16(Op::local_set, arr_slot);
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param {
                                name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false,
                            }).collect(),
                            body: blk.body.clone(),
                            is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        let fn_slot = self.define_local("__map_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        let res_slot = self.define_local("__map_res");
                        let i_slot = self.define_local("__map_i");
                        let c = self.current_chunk_idx;
                        common::loops::emit_map(&mut self.chunks[c], fn_slot, arr_slot, res_slot, i_slot, line);
                        return Ok(());
                    }
                }
                "select" | "filter" => {
                    if let Some(blk) = block {
                        self.compile_expression(recv)?;
                        let arr_slot = self.define_local("__sel_arr");
                        self.emit_u16(Op::local_set, arr_slot);
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param {
                                name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false,
                            }).collect(),
                            body: blk.body.clone(),
                            is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        let fn_slot = self.define_local("__sel_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        let res_slot = self.define_local("__sel_res");
                        let i_slot = self.define_local("__sel_i");
                        let elem_slot = self.define_local("__sel_elem");
                        let c = self.current_chunk_idx;
                        common::loops::emit_filter(&mut self.chunks[c], fn_slot, arr_slot, res_slot, i_slot, elem_slot, line);
                        return Ok(());
                    }
                }
                "reduce" | "inject" => {
                    if let Some(blk) = block {
                        self.compile_expression(recv)?;
                        let arr_slot = self.define_local("__red_arr");
                        self.emit_u16(Op::local_set, arr_slot);
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param {
                                name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false,
                            }).collect(),
                            body: blk.body.clone(),
                            is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        let fn_slot = self.define_local("__red_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        // If initial value provided as method arg
                        if !args.is_empty() {
                            self.compile_expression(&args[0])?;
                            let acc_slot = self.define_local("__red_acc");
                            self.emit_u16(Op::local_set, acc_slot);
                            let i_slot = self.define_local("__red_i");
                            self.emit_constant(Value::I32(0));
                            self.emit_u16(Op::local_set, i_slot);
                            let c = self.current_chunk_idx;
                            let loop_start = self.current_offset();
                            self.emit_u16(Op::local_get, i_slot);
                            self.emit_u16(Op::local_get, arr_slot);
                            self.emit(Op::array_length);
                            self.emit(Op::dyn_lt);
                            let exit = self.emit_jump(Op::br_if_false);
                            self.emit_u16(Op::local_get, fn_slot);
                            self.emit_u16(Op::local_get, acc_slot);
                            self.emit_u16(Op::local_get, arr_slot);
                            self.emit_u16(Op::local_get, i_slot);
                            self.emit(Op::array_get);
                            self.emit_u8(Op::call_ref, 2);
                            self.emit_u16(Op::local_set, acc_slot);
                            common::loops::emit_for_in_end(&mut self.chunks[c], i_slot, loop_start, exit, line);
                            self.emit_u16(Op::local_get, acc_slot);
                        } else {
                            let acc_slot = self.define_local("__red_acc");
                            let i_slot = self.define_local("__red_i");
                            let c = self.current_chunk_idx;
                            common::loops::emit_reduce(&mut self.chunks[c], fn_slot, arr_slot, acc_slot, i_slot, line);
                        }
                        return Ok(());
                    }
                }
                "any?" => {
                    if let Some(blk) = block {
                        self.compile_block_predicate(recv, blk, true)?;
                        return Ok(());
                    }
                }
                "all?" => {
                    if let Some(blk) = block {
                        self.compile_block_predicate(recv, blk, false)?;
                        return Ok(());
                    }
                }
                "each_with_object" => {
                    if let Some(_blk) = block {
                        self.compile_expression(recv)?;
                        return Ok(());
                    }
                }
                "flat_map" => {
                    if let Some(blk) = block {
                        // flat_map = map + flatten (simplified: just map for now)
                        self.compile_expression(recv)?;
                        let arr_slot = self.define_local("__fmap_arr");
                        self.emit_u16(Op::local_set, arr_slot);
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param {
                                name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false,
                            }).collect(),
                            body: blk.body.clone(),
                            is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        let fn_slot = self.define_local("__fmap_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        let res_slot = self.define_local("__fmap_res");
                        let i_slot = self.define_local("__fmap_i");
                        let c = self.current_chunk_idx;
                        common::loops::emit_map(&mut self.chunks[c], fn_slot, arr_slot, res_slot, i_slot, line);
                        return Ok(());
                    }
                }

                // ── Hash methods ──────────────────────────────
                "keys" => {
                    self.compile_expression(recv)?;
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::dict::emit_keys(&mut self.chunks[c], line);
                    return Ok(());
                }
                "values" => {
                    self.compile_expression(recv)?;
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::dict::emit_values(&mut self.chunks[c], line);
                    return Ok(());
                }
                "has_key?" | "key?" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    let idx = self.import("vybe:object", "hasProperty");
                    self.emit_host_call(idx, 2);
                    return Ok(());
                }
                "merge" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::collections::emit_concat(&mut self.chunks[c], line);
                    return Ok(());
                }
                "delete" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::null);
                    return Ok(());
                }
                "fetch" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::array_get);
                    return Ok(());
                }

                // ── StringBuilder (same as VB/PHP) ─────────────
                "write" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    let i = self.import("vybe:types", "sbAppend");
                    self.emit_host_call(i, (args.len() + 1) as u8);
                    return Ok(());
                }
                "string" => {
                    self.compile_expression(recv)?;
                    let i = self.import("vybe:types", "sbToString");
                    self.emit_host_call(i, 1);
                    return Ok(());
                }

                // ── Fiber / Thread methods ─────────────────────
                "resume" => {
                    self.compile_expression(recv)?; // continuation
                    if args.is_empty() {
                        self.emit(Op::null);
                    } else {
                        self.compile_expression(&args[0])?;
                    }
                    self.emit_u16(Op::resume, 0);
                    return Ok(());
                }
                "alive?" => {
                    self.compile_expression(recv)?;
                    let state_key = self.add_string_constant("__cont_state");
                    self.emit_u16(Op::struct_get, state_key);
                    self.emit_constant(Value::String(Rc::from("done")));
                    self.emit(Op::dyn_ne);
                    return Ok(());
                }

                // ── Proc/Lambda call ───────────────────────────
                "call" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit_u8(Op::call_ref, args.len() as u8);
                    return Ok(());
                }

                // ── Database methods (PDO-compatible) ──────────
                "query" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    let idx = self.import("vybe:database", "query");
                    self.emit_host_call(idx, (args.len() + 1) as u8);
                    return Ok(());
                }
                "execute" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    let idx = self.import("vybe:database", "execute");
                    self.emit_host_call(idx, (args.len() + 1) as u8);
                    return Ok(());
                }
                "close" => {
                    self.compile_expression(recv)?;
                    let idx = self.import("vybe:database", "close");
                    self.emit_host_call(idx, 1);
                    return Ok(());
                }

                // ── IO methods ─────────────────────────────────
                "read" => {
                    self.compile_expression(recv)?;
                    let idx = self.import("wasi:filesystem", "readFile");
                    self.emit_host_call(idx, 1);
                    return Ok(());
                }

                // ── Integer/Numeric methods ─────────────────────
                "times" => {
                    if let Some(blk) = block {
                        self.compile_expression(recv)?;
                        let limit_slot = self.define_local("__times_limit");
                        self.emit_u16(Op::local_set, limit_slot);
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param {
                                name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false,
                            }).collect(),
                            body: blk.body.clone(), is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        let fn_slot = self.define_local("__times_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        let i_slot = self.define_local("__times_i");
                        self.emit_constant(Value::I32(0));
                        self.emit_u16(Op::local_set, i_slot);
                        let loop_start = self.current_offset();
                        self.emit_u16(Op::local_get, i_slot);
                        self.emit_u16(Op::local_get, limit_slot);
                        self.emit(Op::dyn_lt);
                        let exit = self.emit_jump(Op::br_if_false);
                        self.emit_u16(Op::local_get, fn_slot);
                        self.emit_u16(Op::local_get, i_slot);
                        self.emit_u8(Op::call_ref, 1);
                        self.emit(Op::drop);
                        self.emit_u16(Op::local_get, i_slot);
                        self.emit_constant(Value::I32(1));
                        self.emit(Op::i32_add);
                        self.emit_u16(Op::local_set, i_slot);
                        self.emit_loop(loop_start);
                        self.patch_jump(exit);
                        self.emit(Op::null);
                        return Ok(());
                    }
                }
                "upto" => {
                    if let Some(blk) = block {
                        self.compile_expression(recv)?;
                        let start_slot = self.define_local("__upto_start");
                        self.emit_u16(Op::local_set, start_slot);
                        self.compile_expression(&args[0])?;
                        let end_slot = self.define_local("__upto_end");
                        self.emit_u16(Op::local_set, end_slot);
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param { name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false }).collect(),
                            body: blk.body.clone(), is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        let fn_slot = self.define_local("__upto_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        let i_slot = self.define_local("__upto_i");
                        self.emit_u16(Op::local_get, start_slot);
                        self.emit_u16(Op::local_set, i_slot);
                        let loop_start = self.current_offset();
                        self.emit_u16(Op::local_get, i_slot);
                        self.emit_u16(Op::local_get, end_slot);
                        self.emit(Op::dyn_le);
                        let exit = self.emit_jump(Op::br_if_false);
                        self.emit_u16(Op::local_get, fn_slot);
                        self.emit_u16(Op::local_get, i_slot);
                        self.emit_u8(Op::call_ref, 1);
                        self.emit(Op::drop);
                        self.emit_u16(Op::local_get, i_slot);
                        self.emit_constant(Value::I32(1));
                        self.emit(Op::i32_add);
                        self.emit_u16(Op::local_set, i_slot);
                        self.emit_loop(loop_start);
                        self.patch_jump(exit);
                        self.emit(Op::null);
                        return Ok(());
                    }
                }
                "downto" => {
                    if let Some(blk) = block {
                        self.compile_expression(recv)?;
                        let start_slot = self.define_local("__downto_start");
                        self.emit_u16(Op::local_set, start_slot);
                        self.compile_expression(&args[0])?;
                        let end_slot = self.define_local("__downto_end");
                        self.emit_u16(Op::local_set, end_slot);
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param { name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false }).collect(),
                            body: blk.body.clone(), is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        let fn_slot = self.define_local("__downto_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        let i_slot = self.define_local("__downto_i");
                        self.emit_u16(Op::local_get, start_slot);
                        self.emit_u16(Op::local_set, i_slot);
                        let loop_start = self.current_offset();
                        self.emit_u16(Op::local_get, i_slot);
                        self.emit_u16(Op::local_get, end_slot);
                        self.emit(Op::dyn_ge);
                        let exit = self.emit_jump(Op::br_if_false);
                        self.emit_u16(Op::local_get, fn_slot);
                        self.emit_u16(Op::local_get, i_slot);
                        self.emit_u8(Op::call_ref, 1);
                        self.emit(Op::drop);
                        self.emit_u16(Op::local_get, i_slot);
                        self.emit_constant(Value::I32(1));
                        self.emit(Op::f64_sub);
                        self.emit_u16(Op::local_set, i_slot);
                        self.emit_loop(loop_start);
                        self.patch_jump(exit);
                        self.emit(Op::null);
                        return Ok(());
                    }
                }
                "even?" => {
                    self.compile_expression(recv)?;
                    self.emit_constant(Value::F64(2.0));
                    self.emit(Op::f64_mod);
                    self.emit_constant(Value::F64(0.0));
                    self.emit(Op::dyn_eq);
                    return Ok(());
                }
                "odd?" => {
                    self.compile_expression(recv)?;
                    self.emit_constant(Value::F64(2.0));
                    self.emit(Op::f64_mod);
                    self.emit_constant(Value::F64(0.0));
                    self.emit(Op::dyn_ne);
                    return Ok(());
                }
                "between?" => {
                    self.compile_expression(recv)?;
                    let val_slot = self.define_local("__btw_val");
                    self.emit_u16(Op::local_set, val_slot);
                    self.compile_expression(&args[0])?; // min
                    self.emit_u16(Op::local_get, val_slot);
                    self.emit(Op::dyn_le);
                    let skip = self.emit_jump(Op::br_if_false);
                    self.emit_u16(Op::local_get, val_slot);
                    self.compile_expression(&args[1])?; // max
                    self.emit(Op::dyn_le);
                    let end = self.emit_jump(Op::br);
                    self.patch_jump(skip);
                    self.emit(Op::r#false);
                    self.patch_jump(end);
                    return Ok(());
                }
                "clamp" => {
                    // val.clamp(min, max) → max(min, min(val, max_val))
                    self.compile_expression(recv)?;
                    if args.len() >= 2 {
                        // inline: if val < min then min elsif val > max then max else val end
                        let val_slot = self.define_local("__clamp_val");
                        self.emit_u16(Op::local_set, val_slot);
                        self.compile_expression(&args[0])?; // min
                        let min_slot = self.define_local("__clamp_min");
                        self.emit_u16(Op::local_set, min_slot);
                        self.compile_expression(&args[1])?; // max
                        let max_slot = self.define_local("__clamp_max");
                        self.emit_u16(Op::local_set, max_slot);
                        self.emit_u16(Op::local_get, val_slot);
                        self.emit_u16(Op::local_get, min_slot);
                        self.emit(Op::dyn_lt);
                        let use_min = self.emit_jump(Op::br_if_true);
                        self.emit_u16(Op::local_get, val_slot);
                        self.emit_u16(Op::local_get, max_slot);
                        self.emit(Op::dyn_gt);
                        let use_max = self.emit_jump(Op::br_if_true);
                        self.emit_u16(Op::local_get, val_slot);
                        let end1 = self.emit_jump(Op::br);
                        self.patch_jump(use_max);
                        self.emit_u16(Op::local_get, max_slot);
                        let end2 = self.emit_jump(Op::br);
                        self.patch_jump(use_min);
                        self.emit_u16(Op::local_get, min_slot);
                        self.patch_jump(end1);
                        self.patch_jump(end2);
                    }
                    return Ok(());
                }
                "abs" => {
                    self.compile_expression(recv)?;
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::math::emit_abs(&mut self.chunks[c], line);
                    return Ok(());
                }
                "zero?" => {
                    self.compile_expression(recv)?;
                    self.emit_constant(Value::F64(0.0));
                    self.emit(Op::dyn_eq);
                    return Ok(());
                }
                "positive?" => {
                    self.compile_expression(recv)?;
                    self.emit_constant(Value::F64(0.0));
                    self.emit(Op::dyn_gt);
                    return Ok(());
                }
                "negative?" => {
                    self.compile_expression(recv)?;
                    self.emit_constant(Value::F64(0.0));
                    self.emit(Op::dyn_lt);
                    return Ok(());
                }
                "round" => {
                    self.compile_expression(recv)?;
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::math::emit_round(&mut self.chunks[c], line);
                    return Ok(());
                }
                "floor" => {
                    self.compile_expression(recv)?;
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::math::emit_floor(&mut self.chunks[c], line);
                    return Ok(());
                }
                "ceil" => {
                    self.compile_expression(recv)?;
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::math::emit_ceil(&mut self.chunks[c], line);
                    return Ok(());
                }

                // ── More String methods ────────────────────────
                "match" | "match?" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    let i = self.import("vybe:regex", "test");
                    self.emit_host_call(i, 2);
                    return Ok(());
                }
                "scan" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    let i = self.import("vybe:regex", "matchGroups");
                    self.emit_host_call(i, 2);
                    return Ok(());
                }
                "tr" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::str_replace);
                    return Ok(());
                }
                "squeeze" => {
                    self.compile_expression(recv)?;
                    return Ok(());
                }
                "center" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::str_pad_start);
                    return Ok(());
                }
                "ljust" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::str_pad_end);
                    return Ok(());
                }
                "rjust" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::str_pad_start);
                    return Ok(());
                }
                "encode" | "force_encoding" => {
                    self.compile_expression(recv)?;
                    return Ok(());
                }
                "bytes" => {
                    self.compile_expression(recv)?;
                    self.emit_constant(Value::String(Rc::from("")));
                    self.emit(Op::str_split);
                    return Ok(());
                }
                "casecmp" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::str_to_lower);
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::str_to_lower);
                    self.emit(Op::dyn_eq);
                    return Ok(());
                }
                "freeze" => {
                    self.compile_expression(recv)?;
                    return Ok(());
                }
                "frozen?" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::drop);
                    self.emit(Op::r#false);
                    return Ok(());
                }

                // ── More Array methods ─────────────────────────
                "find" | "detect" => {
                    if let Some(blk) = block {
                        self.compile_expression(recv)?;
                        let arr_slot = self.define_local("__find_arr");
                        self.emit_u16(Op::local_set, arr_slot);
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param { name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false }).collect(),
                            body: blk.body.clone(), is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        let fn_slot = self.define_local("__find_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        self.emit(Op::null);
                        let res_slot = self.define_local("__find_res");
                        self.emit_u16(Op::local_set, res_slot);
                        let i_slot = self.define_local("__find_i");
                        let c = self.current_chunk_idx;
                        let (loop_start, exit) = common::loops::emit_for_in_start(&mut self.chunks[c], arr_slot, i_slot, line);
                        let elem_slot = self.define_local("__find_elem");
                        self.emit_u16(Op::local_set, elem_slot);
                        self.emit_u16(Op::local_get, fn_slot);
                        self.emit_u16(Op::local_get, elem_slot);
                        self.emit_u8(Op::call_ref, 1);
                        self.emit(Op::dyn_to_bool);
                        let skip = self.emit_jump(Op::br_if_false);
                        self.emit_u16(Op::local_get, elem_slot);
                        self.emit_u16(Op::local_set, res_slot);
                        self.patch_jump(skip);
                        common::loops::emit_for_in_end(&mut self.chunks[c], i_slot, loop_start, exit, line);
                        self.emit_u16(Op::local_get, res_slot);
                        return Ok(());
                    }
                }
                "find_index" => {
                    if let Some(blk) = block {
                        self.compile_expression(recv)?;
                        let arr_slot = self.define_local("__fi_arr");
                        self.emit_u16(Op::local_set, arr_slot);
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param { name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false }).collect(),
                            body: blk.body.clone(), is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        let fn_slot = self.define_local("__fi_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        self.emit(Op::null);
                        let res_slot = self.define_local("__fi_res");
                        self.emit_u16(Op::local_set, res_slot);
                        let i_slot = self.define_local("__fi_i");
                        let c = self.current_chunk_idx;
                        let (loop_start, exit) = common::loops::emit_for_in_start(&mut self.chunks[c], arr_slot, i_slot, line);
                        let elem_slot = self.define_local("__fi_elem");
                        self.emit_u16(Op::local_set, elem_slot);
                        self.emit_u16(Op::local_get, fn_slot);
                        self.emit_u16(Op::local_get, elem_slot);
                        self.emit_u8(Op::call_ref, 1);
                        self.emit(Op::dyn_to_bool);
                        let skip = self.emit_jump(Op::br_if_false);
                        self.emit_u16(Op::local_get, i_slot);
                        self.emit_u16(Op::local_set, res_slot);
                        self.patch_jump(skip);
                        common::loops::emit_for_in_end(&mut self.chunks[c], i_slot, loop_start, exit, line);
                        self.emit_u16(Op::local_get, res_slot);
                        return Ok(());
                    }
                }
                "sort_by" => {
                    if let Some(_blk) = block {
                        self.compile_expression(recv)?;
                        let tmp = self.define_local("__sortby_tmp");
                        self.emit_u16(Op::local_set, tmp);
                        let c = self.current_chunk_idx;
                        let line = self.line;
                        common::bundle::emit_call_push_func(&mut self.chunks[c], "__vybe_sorted", line);
                        self.emit_u16(Op::local_get, tmp);
                        common::bundle::emit_call_invoke(&mut self.chunks[c], 1, line);
                        return Ok(());
                    }
                }
                "min_by" | "max_by" => {
                    if let Some(_blk) = block {
                        self.compile_expression(recv)?;
                        if method == "min_by" {
                            let c = self.current_chunk_idx;
                            let line = self.line;
                            common::collections::emit_min(&mut self.chunks[c], 1, line);
                        } else {
                            let c = self.current_chunk_idx;
                            let line = self.line;
                            common::collections::emit_max(&mut self.chunks[c], 1, line);
                        }
                        return Ok(());
                    }
                }
                "zip" => {
                    // arr.zip(other) → use stdlib __vybe_zip
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::collections::emit_zip(&mut self.chunks[c], line);
                    return Ok(());
                }
                "group_by" => {
                    if let Some(blk) = block {
                        // group_by { |x| key } → { key => [elements] }
                        self.compile_expression(recv)?;
                        let arr_slot = self.define_local("__gb_arr");
                        self.emit_u16(Op::local_set, arr_slot);
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param { name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false }).collect(),
                            body: blk.body.clone(), is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        let fn_slot = self.define_local("__gb_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        let c = self.current_chunk_idx;
                        common::dict::emit_new(&mut self.chunks[c], line);
                        let hash_slot = self.define_local("__gb_hash");
                        self.emit_u16(Op::local_set, hash_slot);
                        let i_slot = self.define_local("__gb_i");
                        let (loop_start, exit) = common::loops::emit_for_in_start(&mut self.chunks[c], arr_slot, i_slot, line);
                        let elem_slot = self.define_local("__gb_elem");
                        self.emit_u16(Op::local_set, elem_slot);
                        // key = block(elem)
                        self.emit_u16(Op::local_get, fn_slot);
                        self.emit_u16(Op::local_get, elem_slot);
                        self.emit_u8(Op::call_ref, 1);
                        let key_slot = self.define_local("__gb_key");
                        self.emit_u16(Op::local_set, key_slot);
                        // hash[key] ||= []; hash[key].push(elem)
                        self.emit_u16(Op::local_get, hash_slot);
                        self.emit_u16(Op::local_get, key_slot);
                        self.emit(Op::array_get);
                        self.emit(Op::dup);
                        self.emit(Op::ref_is_null);
                        let exists = self.emit_jump(Op::br_if_false);
                        self.emit(Op::drop);
                        self.emit_u16(Op::array_new, 0);
                        self.patch_jump(exists);
                        // Push element
                        self.emit_u16(Op::local_get, elem_slot);
                        self.emit(Op::array_push);
                        self.emit(Op::drop);
                        // Store back
                        let grp_slot = self.define_local("__gb_grp");
                        self.emit_u16(Op::local_set, grp_slot);
                        self.emit_u16(Op::local_get, hash_slot);
                        self.emit_u16(Op::local_get, key_slot);
                        self.emit_u16(Op::local_get, grp_slot);
                        self.emit(Op::array_set);
                        self.emit(Op::drop);
                        common::loops::emit_for_in_end(&mut self.chunks[c], i_slot, loop_start, exit, line);
                        self.emit_u16(Op::local_get, hash_slot);
                        return Ok(());
                    }
                }
                "take" => {
                    self.compile_expression(recv)?;
                    self.emit_constant(Value::I32(0));
                    for a in args { self.compile_expression(a)?; }
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::collections::emit_slice(&mut self.chunks[c], line);
                    return Ok(());
                }
                "drop" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit_constant(Value::I32(i32::MAX));
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::collections::emit_slice(&mut self.chunks[c], line);
                    return Ok(());
                }
                "sample" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::dup);
                    self.emit(Op::array_length);
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::math::emit_random(&mut self.chunks[c], line);
                    self.emit(Op::f64_mul);
                    self.emit(Op::f64_trunc);
                    self.emit(Op::array_get);
                    return Ok(());
                }
                "shuffle" => {
                    self.compile_expression(recv)?;
                    return Ok(());
                }
                "include?" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::collections::emit_contains(&mut self.chunks[c], line);
                    return Ok(());
                }
                "each_slice" | "each_cons" => {
                    if let Some(_blk) = block {
                        self.compile_expression(recv)?;
                        return Ok(());
                    }
                }
                "none?" => {
                    if let Some(blk) = block {
                        self.compile_block_predicate(recv, blk, true)?;
                        self.emit(Op::dyn_not);
                        return Ok(());
                    } else {
                        self.compile_expression(recv)?;
                        self.emit(Op::array_length);
                        self.emit_constant(Value::I32(0));
                        self.emit(Op::dyn_eq);
                        return Ok(());
                    }
                }

                // ── More Hash methods ──────────────────────────
                "each_pair" => {
                    // Same as each for hashes
                    if let Some(blk) = block {
                        self.compile_expression(recv)?;
                        let arr_slot = self.define_local("__ep_arr");
                        self.emit_u16(Op::local_set, arr_slot);
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param { name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false }).collect(),
                            body: blk.body.clone(), is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        let fn_slot = self.define_local("__ep_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        let i_slot = self.define_local("__ep_i");
                        let c = self.current_chunk_idx;
                        common::loops::emit_foreach(&mut self.chunks[c], fn_slot, arr_slot, i_slot, line);
                        self.emit_u16(Op::local_get, arr_slot);
                        return Ok(());
                    }
                }
                "transform_values" => {
                    if let Some(_blk) = block {
                        self.compile_expression(recv)?;
                        return Ok(());
                    }
                }
                "transform_keys" => {
                    if let Some(_blk) = block {
                        self.compile_expression(recv)?;
                        return Ok(());
                    }
                }
                "invert" => {
                    self.compile_expression(recv)?;
                    return Ok(());
                }
                "to_h" => {
                    self.compile_expression(recv)?;
                    return Ok(());
                }
                "each_key" => {
                    if let Some(blk) = block {
                        self.compile_expression(recv)?;
                        let c = self.current_chunk_idx;
                        let line = self.line;
                        common::dict::emit_keys(&mut self.chunks[c], line);
                        let arr_slot = self.define_local("__ek_arr");
                        self.emit_u16(Op::local_set, arr_slot);
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param { name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false }).collect(),
                            body: blk.body.clone(), is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        let fn_slot = self.define_local("__ek_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        let i_slot = self.define_local("__ek_i");
                        let c = self.current_chunk_idx;
                        common::loops::emit_foreach(&mut self.chunks[c], fn_slot, arr_slot, i_slot, line);
                        self.emit(Op::null);
                        return Ok(());
                    }
                }
                "each_value" => {
                    if let Some(blk) = block {
                        self.compile_expression(recv)?;
                        let c = self.current_chunk_idx;
                        let line = self.line;
                        common::dict::emit_values(&mut self.chunks[c], line);
                        let arr_slot = self.define_local("__ev_arr");
                        self.emit_u16(Op::local_set, arr_slot);
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param { name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false }).collect(),
                            body: blk.body.clone(), is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        let fn_slot = self.define_local("__ev_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        let i_slot = self.define_local("__ev_i");
                        let c = self.current_chunk_idx;
                        common::loops::emit_foreach(&mut self.chunks[c], fn_slot, arr_slot, i_slot, line);
                        self.emit(Op::null);
                        return Ok(());
                    }
                }

                // ── Range methods ──────────────────────────────
                "step" => {
                    if let Some(_blk) = block {
                        self.compile_expression(recv)?;
                        return Ok(());
                    }
                }

                // ── Object/introspection ───────────────────────
                "respond_to?" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    let idx = self.import("vybe:object", "hasProperty");
                    self.emit_host_call(idx, 2);
                    return Ok(());
                }
                "send" | "__send__" => {
                    self.compile_expression(recv)?;
                    // First arg is method name
                    if !args.is_empty() {
                        self.compile_expression(&args[0])?;
                        self.emit(Op::array_get);
                        self.compile_expression(recv)?; // self
                        for a in &args[1..] { self.compile_expression(a)?; }
                        self.emit_u8(Op::call_ref, (args.len()) as u8);
                    } else {
                        self.emit(Op::null);
                    }
                    return Ok(());
                }
                "methods" | "instance_methods" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::null); // simplified
                    return Ok(());
                }
                "class" => {
                    self.compile_expression(recv)?;
                    let idx = self.add_string_constant("__type");
                    self.emit_u16(Op::struct_get, idx);
                    return Ok(());
                }
                "dup" | "clone" => {
                    // Shallow copy — simplified: return same object
                    self.compile_expression(recv)?;
                    return Ok(());
                }
                "object_id" => {
                    self.compile_expression(recv)?;
                    self.emit_constant(Value::I32(0));
                    return Ok(());
                }
                "inspect" => {
                    self.compile_expression(recv)?;
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::strings::emit_to_string(&mut self.chunks[c], line);
                    return Ok(());
                }
                "tap" => {
                    if let Some(blk) = block {
                        self.compile_expression(recv)?;
                        self.emit(Op::dup);
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param { name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false }).collect(),
                            body: blk.body.clone(), is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        // swap: [obj, func] → call func(obj)
                        let obj_slot = self.define_local("__tap_obj");
                        let fn_slot = self.define_local("__tap_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        self.emit_u16(Op::local_set, obj_slot);
                        self.emit_u16(Op::local_get, fn_slot);
                        self.emit_u16(Op::local_get, obj_slot);
                        self.emit_u8(Op::call_ref, 1);
                        self.emit(Op::drop);
                        self.emit_u16(Op::local_get, obj_slot);
                        return Ok(());
                    }
                }

                // ── File class methods ─────────────────────────
                // These fire when receiver is File constant
                "exist?" | "exists?" => {
                    for a in args { self.compile_expression(a)?; }
                    let i = self.import("wasi:filesystem", "exists");
                    self.emit_host_call(i, args.len() as u8);
                    return Ok(());
                }
                "directory?" => {
                    for a in args { self.compile_expression(a)?; }
                    let i = self.import("wasi:filesystem", "isDir");
                    self.emit_host_call(i, args.len() as u8);
                    return Ok(());
                }
                "size" if args.len() == 1 => {
                    self.compile_expression(&args[0])?;
                    let i = self.import("wasi:filesystem", "fileSize");
                    self.emit_host_call(i, 1);
                    return Ok(());
                }
                "write" if args.len() >= 2 => {
                    for a in args { self.compile_expression(a)?; }
                    let i = self.import("wasi:filesystem", "writeFile");
                    self.emit_host_call(i, args.len() as u8);
                    return Ok(());
                }
                "readlines" => {
                    self.compile_expression(recv)?;
                    let i = self.import("wasi:filesystem", "readFile");
                    self.emit_host_call(i, 1);
                    self.emit_constant(Value::String(Rc::from("\n")));
                    self.emit(Op::str_split);
                    return Ok(());
                }

                // ── String#* (repeat) ───────────────────────────
                "chr" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::str_from_char_code);
                    return Ok(());
                }
                "ord" => {
                    self.compile_expression(recv)?;
                    let idx = self.import("vybe:string", "charCodeAt");
                    self.emit_host_call(idx, 1);
                    return Ok(());
                }
                "hex" => {
                    self.compile_expression(recv)?;
                    let idx = self.import("vybe:convert", "hex");
                    self.emit_host_call(idx, 1);
                    return Ok(());
                }

                // ── dig (nested access) ────────────────────────
                "dig" => {
                    self.compile_expression(recv)?;
                    for a in args {
                        self.compile_expression(a)?;
                        self.emit(Op::array_get);
                    }
                    return Ok(());
                }

                // ── filter_map ─────────────────────────────────
                "filter_map" => {
                    if let Some(blk) = block {
                        self.compile_expression(recv)?;
                        let arr_slot = self.define_local("__fm_arr");
                        self.emit_u16(Op::local_set, arr_slot);
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param { name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false }).collect(),
                            body: blk.body.clone(), is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        let fn_slot = self.define_local("__fm_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        // result array
                        self.emit_u16(Op::array_new, 0);
                        let res_slot = self.define_local("__fm_res");
                        self.emit_u16(Op::local_set, res_slot);
                        let i_slot = self.define_local("__fm_i");
                        let c = self.current_chunk_idx;
                        let (loop_start, exit) = common::loops::emit_for_in_start(&mut self.chunks[c], arr_slot, i_slot, line);
                        let elem_slot = self.define_local("__fm_elem");
                        self.emit_u16(Op::local_set, elem_slot);
                        // call block
                        self.emit_u16(Op::local_get, fn_slot);
                        self.emit_u16(Op::local_get, elem_slot);
                        self.emit_u8(Op::call_ref, 1);
                        // if result is not nil/false, push to result
                        self.emit(Op::dup);
                        self.emit(Op::ref_is_null);
                        let skip = self.emit_jump(Op::br_if_true);
                        self.emit(Op::dup);
                        self.emit(Op::dyn_to_bool);
                        let skip2 = self.emit_jump(Op::br_if_false);
                        // push to result
                        let val_slot = self.define_local("__fm_val");
                        self.emit_u16(Op::local_set, val_slot);
                        self.emit_u16(Op::local_get, res_slot);
                        self.emit_u16(Op::local_get, val_slot);
                        self.emit(Op::array_push);
                        self.emit(Op::drop);
                        let end = self.emit_jump(Op::br);
                        self.patch_jump(skip);
                        self.emit(Op::drop);
                        let end2 = self.emit_jump(Op::br);
                        self.patch_jump(skip2);
                        self.emit(Op::drop);
                        self.patch_jump(end);
                        self.patch_jump(end2);
                        common::loops::emit_for_in_end(&mut self.chunks[c], i_slot, loop_start, exit, line);
                        self.emit_u16(Op::local_get, res_slot);
                        return Ok(());
                    }
                }
                // ── tally ──────────────────────────────────────
                "tally" => {
                    // Create hash, iterate array, count occurrences
                    self.compile_expression(recv)?;
                    let arr_slot = self.define_local("__tally_arr");
                    self.emit_u16(Op::local_set, arr_slot);
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::dict::emit_new(&mut self.chunks[c], line);
                    let hash_slot = self.define_local("__tally_h");
                    self.emit_u16(Op::local_set, hash_slot);
                    // Iterate and count
                    let i_slot = self.define_local("__tally_i");
                    let c = self.current_chunk_idx;
                    let (loop_start, exit) = common::loops::emit_for_in_start(&mut self.chunks[c], arr_slot, i_slot, line);
                    let elem_slot = self.define_local("__tally_elem");
                    self.emit_u16(Op::local_set, elem_slot);
                    // hash[elem] = (hash[elem] || 0) + 1
                    self.emit_u16(Op::local_get, hash_slot);
                    self.emit_u16(Op::local_get, elem_slot);
                    self.emit(Op::array_get); // current count or null
                    // If null, use 0
                    self.emit(Op::dup);
                    self.emit(Op::ref_is_null);
                    let not_null = self.emit_jump(Op::br_if_false);
                    self.emit(Op::drop);
                    self.emit_constant(Value::F64(0.0));
                    self.patch_jump(not_null);
                    // Add 1
                    self.emit_constant(Value::F64(1.0));
                    self.emit(Op::dyn_add);
                    // Store back
                    let count_slot = self.define_local("__tally_cnt");
                    self.emit_u16(Op::local_set, count_slot);
                    self.emit_u16(Op::local_get, hash_slot);
                    self.emit_u16(Op::local_get, elem_slot);
                    self.emit_u16(Op::local_get, count_slot);
                    self.emit(Op::array_set);
                    self.emit(Op::drop);
                    common::loops::emit_for_in_end(&mut self.chunks[c], i_slot, loop_start, exit, line);
                    self.emit_u16(Op::local_get, hash_slot);
                    return Ok(());
                }
                // ── each_with_object ───────────────────────────
                "each_with_object" => {
                    if let Some(blk) = block {
                        // arr.each_with_object(init) { |elem, obj| ... }
                        if !args.is_empty() {
                            self.compile_expression(&args[0])?;
                        } else {
                            self.emit(Op::null);
                        }
                        let obj_slot = self.define_local("__ewo_obj");
                        self.emit_u16(Op::local_set, obj_slot);
                        self.compile_expression(recv)?;
                        let arr_slot = self.define_local("__ewo_arr");
                        self.emit_u16(Op::local_set, arr_slot);
                        let blk_decl = MethodDecl {
                            name: "<block>".to_string(),
                            params: blk.params.iter().map(|n| Param { name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false }).collect(),
                            body: blk.body.clone(), is_self: false,
                        };
                        let fn_ci = self.compile_method_def(&blk_decl)?;
                        let line = self.line;
                        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
                        let fn_slot = self.define_local("__ewo_fn");
                        self.emit_u16(Op::local_set, fn_slot);
                        let i_slot = self.define_local("__ewo_i");
                        let c = self.current_chunk_idx;
                        let (loop_start, exit) = common::loops::emit_for_in_start(&mut self.chunks[c], arr_slot, i_slot, line);
                        let elem_slot = self.define_local("__ewo_elem");
                        self.emit_u16(Op::local_set, elem_slot);
                        self.emit_u16(Op::local_get, fn_slot);
                        self.emit_u16(Op::local_get, elem_slot);
                        self.emit_u16(Op::local_get, obj_slot);
                        self.emit_u8(Op::call_ref, 2);
                        self.emit(Op::drop);
                        common::loops::emit_for_in_end(&mut self.chunks[c], i_slot, loop_start, exit, line);
                        self.emit_u16(Op::local_get, obj_slot);
                        return Ok(());
                    }
                }
                // ── minmax ─────────────────────────────────────
                "minmax" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::dup);
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::collections::emit_min(&mut self.chunks[c], 1, line);
                    let min_slot = self.define_local("__mm_min");
                    self.emit_u16(Op::local_set, min_slot);
                    self.compile_expression(recv)?;
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::collections::emit_max(&mut self.chunks[c], 1, line);
                    let max_slot = self.define_local("__mm_max");
                    self.emit_u16(Op::local_set, max_slot);
                    self.emit_u16(Op::local_get, min_slot);
                    self.emit_u16(Op::local_get, max_slot);
                    self.emit_u16(Op::array_new, 2);
                    return Ok(());
                }
                // ── Array#rotate ───────────────────────────────
                "rotate" => {
                    self.compile_expression(recv)?;
                    return Ok(());
                }
                "transpose" => {
                    self.compile_expression(recv)?;
                    return Ok(());
                }
                "combination" | "permutation" | "product" => {
                    self.compile_expression(recv)?;
                    return Ok(());
                }
                // ── Integer methods ────────────────────────────
                "gcd" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::null); // simplified
                    return Ok(());
                }
                "lcm" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::null);
                    return Ok(());
                }
                "digits" => {
                    self.compile_expression(recv)?;
                    return Ok(());
                }
                "divmod" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    // [a/b, a%b]
                    let b_slot = self.define_local("__dm_b");
                    let a_slot = self.define_local("__dm_a");
                    self.emit_u16(Op::local_set, b_slot);
                    self.emit_u16(Op::local_set, a_slot);
                    self.emit_u16(Op::local_get, a_slot);
                    self.emit_u16(Op::local_get, b_slot);
                    self.emit(Op::f64_div);
                    self.emit(Op::f64_trunc);
                    self.emit_u16(Op::local_get, a_slot);
                    self.emit_u16(Op::local_get, b_slot);
                    self.emit(Op::f64_mod);
                    self.emit_u16(Op::array_new, 2);
                    return Ok(());
                }
                // ── Regexp.new ─────────────────────────────────
                // (handled via Constant.new in compile_new_call)

                // ── IO/STDIN/STDOUT ────────────────────────────
                "readline" | "gets" => {
                    let i = self.import("wasi:cli", "readLine");
                    self.emit_host_call(i, 0);
                    return Ok(());
                }
                "print_to" | "write_to" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::io::emit_print(&mut self.chunks[c], (args.len()) as u8, line);
                    return Ok(());
                }

                // ── instance_variable_get/set ──────────────────
                "instance_variable_get" => {
                    self.compile_expression(recv)?;
                    if !args.is_empty() {
                        self.compile_expression(&args[0])?;
                    }
                    self.emit(Op::array_get);
                    return Ok(());
                }
                "instance_variable_set" => {
                    self.compile_expression(recv)?;
                    if args.len() >= 2 {
                        self.compile_expression(&args[0])?;
                        self.compile_expression(&args[1])?;
                    }
                    self.emit(Op::array_set);
                    self.emit(Op::drop);
                    self.emit(Op::null);
                    return Ok(());
                }
                "instance_variables" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::null);
                    return Ok(());
                }
                // ── ancestors / superclass ──────────────────────
                "ancestors" | "superclass" => {
                    self.compile_expression(recv)?;
                    self.emit_u16(Op::array_new, 0);
                    return Ok(());
                }
                // ── eql? / equal? ──────────────────────────────
                "eql?" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::eq);
                    return Ok(());
                }
                "equal?" => {
                    self.compile_expression(recv)?;
                    for a in args { self.compile_expression(a)?; }
                    self.emit(Op::eq);
                    return Ok(());
                }
                "hash" => {
                    self.compile_expression(recv)?;
                    self.emit_constant(Value::I32(0));
                    return Ok(());
                }
                // ── encoding ───────────────────────────────────
                "encoding" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::drop);
                    self.emit_constant(Value::String(Rc::from("UTF-8")));
                    return Ok(());
                }
                "valid_encoding?" => {
                    self.compile_expression(recv)?;
                    self.emit(Op::drop);
                    self.emit(Op::r#true);
                    return Ok(());
                }
                // ── lazy ───────────────────────────────────────
                "lazy" => {
                    self.compile_expression(recv)?;
                    return Ok(());
                }
                // ── cycle ──────────────────────────────────────
                "cycle" => {
                    self.compile_expression(recv)?;
                    return Ok(());
                }

                // ── Constructor: Klass.new(args) ───────────────
                "new" => {
                    return self.compile_new_call(recv, args);
                }

                _ => {} // fall through to generic
            }
        }

        // Generic method call
        if let Some(recv) = receiver {
            self.compile_expression(recv)?;
            let prop_idx = self.add_string_constant(method);
            self.emit_u16(Op::struct_get, prop_idx);
            // call_ref: first arg is self
            self.compile_expression(recv)?;
            for a in args { self.compile_expression(a)?; }
            // If block, compile and pass as last arg
            if let Some(blk) = block {
                let blk_decl = MethodDecl {
                    name: "<block>".to_string(),
                    params: blk.params.iter().map(|n| Param {
                        name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false,
                    }).collect(),
                    body: blk.body.clone(),
                    is_self: false,
                };
                let ci = self.compile_method_def(&blk_decl)?;
                let line = self.line;
                common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], ci, 0, line);
                self.emit_u8(Op::call_ref, (args.len() + 2) as u8); // +1 self, +1 block
            } else {
                self.emit_u8(Op::call_ref, (args.len() + 1) as u8);
            }
        } else {
            // Bare function call
            let is_resolved = self.resolve_var(method).is_some()
                || (self.scopes.len() > 1 && self.resolve_upvalue(self.scopes.len() - 1, method).is_some())
                || self.defined_globals.contains(method);
            if is_resolved {
                self.emit_var_get(method);
                for a in args { self.compile_expression(a)?; }
                if let Some(blk) = block {
                    let blk_decl = MethodDecl {
                        name: "<block>".to_string(),
                        params: blk.params.iter().map(|n| Param {
                            name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false,
                        }).collect(),
                        body: blk.body.clone(),
                        is_self: false,
                    };
                    let ci = self.compile_method_def(&blk_decl)?;
                    let line = self.line;
                    common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], ci, 0, line);
                    self.emit_u8(Op::call_ref, (args.len() + 1) as u8);
                } else {
                    self.emit_u8(Op::call_ref, args.len() as u8);
                }
            } else {
                // Unresolved name — emit as host import call
                for a in args { self.compile_expression(a)?; }
                if let Some(blk) = block {
                    let blk_decl = MethodDecl {
                        name: "<block>".to_string(),
                        params: blk.params.iter().map(|n| Param {
                            name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false,
                        }).collect(),
                        body: blk.body.clone(),
                        is_self: false,
                    };
                    let ci = self.compile_method_def(&blk_decl)?;
                    let line = self.line;
                    common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], ci, 0, line);
                    let idx = self.import("*", method);
                    self.emit_host_call(idx, (args.len() + 1) as u8);
                } else {
                    let idx = self.import("*", method);
                    self.emit_host_call(idx, args.len() as u8);
                }
            }
        }
        Ok(())
    }

    fn compile_block_predicate(&mut self, recv: &Expression, blk: &BlockArg, any: bool) -> Result<(), String> {
        self.compile_expression(recv)?;
        let arr_slot = self.define_local("__pred_arr");
        self.emit_u16(Op::local_set, arr_slot);
        let blk_decl = MethodDecl {
            name: "<block>".to_string(),
            params: blk.params.iter().map(|n| Param {
                name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false,
            }).collect(),
            body: blk.body.clone(),
            is_self: false,
        };
        let fn_ci = self.compile_method_def(&blk_decl)?;
        let line = self.line;
        common::functions::emit_ref_func(&mut self.chunks[self.current_chunk_idx], fn_ci, 0, line);
        let fn_slot = self.define_local("__pred_fn");
        self.emit_u16(Op::local_set, fn_slot);

        // result = false/true (any=false, all=true)
        if any { self.emit(Op::r#false); } else { self.emit(Op::r#true); }
        let res_slot = self.define_local("__pred_res");
        self.emit_u16(Op::local_set, res_slot);

        let i_slot = self.define_local("__pred_i");
        let c = self.current_chunk_idx;
        let (loop_start, exit) = common::loops::emit_for_in_start(&mut self.chunks[c], arr_slot, i_slot, line);
        // element on stack
        let elem_slot = self.define_local("__pred_elem");
        self.emit_u16(Op::local_set, elem_slot);
        self.emit_u16(Op::local_get, fn_slot);
        self.emit_u16(Op::local_get, elem_slot);
        self.emit_u8(Op::call_ref, 1);
        self.emit(Op::dyn_to_bool);
        if any {
            let skip = self.emit_jump(Op::br_if_false);
            self.emit(Op::r#true);
            self.emit_u16(Op::local_set, res_slot);
            // Could break early, but simpler to continue
            self.patch_jump(skip);
        } else {
            let skip = self.emit_jump(Op::br_if_true);
            self.emit(Op::r#false);
            self.emit_u16(Op::local_set, res_slot);
            self.patch_jump(skip);
        }
        common::loops::emit_for_in_end(&mut self.chunks[c], i_slot, loop_start, exit, line);
        self.emit_u16(Op::local_get, res_slot);
        Ok(())
    }

    // ------------------------------------------------------------------
    // Built-in functions (no receiver)
    // ------------------------------------------------------------------

    fn try_compile_builtin(&mut self, name: &str, args: &[Expression]) -> Result<Option<()>, String> {
        let c = self.current_chunk_idx;
        let line = self.line;

        macro_rules! compile_args { () => { for arg in args { self.compile_expression(arg)?; } } }

        match name {
            // ── IO ─────────────────────────────────────────────
            "puts" | "print" | "p" => {
                compile_args!();
                common::io::emit_print(&mut self.chunks[c], args.len() as u8, line);
                return Ok(Some(()));
            }

            // ── Math ───────────────────────────────────────────
            "abs" => { compile_args!(); common::math::emit_abs(&mut self.chunks[c], line); return Ok(Some(())); }
            "sqrt" => { compile_args!(); common::math::emit_sqrt(&mut self.chunks[c], line); return Ok(Some(())); }
            "rand" => { common::math::emit_random(&mut self.chunks[c], line); return Ok(Some(())); }

            // ── Type conversion ────────────────────────────────
            "Integer" => { compile_args!(); common::convert::emit_parse_int(&mut self.chunks[c], line); return Ok(Some(())); }
            "Float" => { compile_args!(); common::convert::emit_parse_float(&mut self.chunks[c], line); return Ok(Some(())); }
            "String" => { compile_args!(); common::strings::emit_to_string(&mut self.chunks[c], line); return Ok(Some(())); }

            // ── Array creation ─────────────────────────────────
            "Array" => {
                compile_args!();
                return Ok(Some(()));
            }

            // ── Sleep ──────────────────────────────────────────
            "sleep" => { compile_args!(); let i = self.import("wasi:clocks", "sleep"); self.emit_host_call(i, 1); return Ok(Some(())); }

            // ── Require / load ─────────────────────────────────
            "require" | "require_relative" | "load" => {
                self.emit(Op::null);
                return Ok(Some(()));
            }

            // ── Kernel methods ─────────────────────────────────
            "raise" => {
                compile_args!();
                common::errors::emit_throw(&mut self.chunks[c], line);
                return Ok(Some(()));
            }
            "exit" => {
                self.emit(Op::halt);
                return Ok(Some(()));
            }
            "system" | "exec" => {
                compile_args!();
                let i = self.import("vybe:types", "processStart");
                self.emit_host_call(i, args.len() as u8);
                return Ok(Some(()));
            }
            "gets" => {
                let i = self.import("wasi:cli", "readLine");
                self.emit_host_call(i, 0);
                return Ok(Some(()));
            }

            // ── pp (pretty print) ──────────────────────────────
            "pp" => {
                compile_args!();
                common::io::emit_print(&mut self.chunks[c], args.len() as u8, line);
                return Ok(Some(()));
            }

            // ── sprintf / format ───────────────────────────────
            "sprintf" | "format" => {
                compile_args!();
                let i = self.import("vybe:string", "format");
                self.emit_host_call(i, args.len() as u8);
                return Ok(Some(()));
            }

            // ── open (File.open) ───────────────────────────────
            "open" => {
                compile_args!();
                let i = self.import("wasi:filesystem", "openFile");
                self.emit_host_call(i, args.len() as u8);
                return Ok(Some(()));
            }

            // ── warn ───────────────────────────────────────────
            "warn" => {
                compile_args!();
                let i = self.import("wasi:cli", "warn");
                self.emit_host_call(i, args.len() as u8);
                return Ok(Some(()));
            }

            // ── Kernel methods ─────────────────────────────────
            "at_exit" => {
                self.emit(Op::null);
                return Ok(Some(()));
            }
            "caller" => {
                self.emit_u16(Op::array_new, 0);
                return Ok(Some(()));
            }
            "trap" => {
                compile_args!();
                self.emit(Op::drop);
                self.emit(Op::null);
                return Ok(Some(()));
            }

            // ── File class methods (bare) ──────────────────────
            // File.read etc. handled via receiver method call

            _ => return Ok(None),
        }
    }

    // ------------------------------------------------------------------
    // new (ClassName.new)
    // ------------------------------------------------------------------

    fn compile_new_call(&mut self, class_expr: &Expression, args: &[Expression]) -> Result<(), String> {
        if let Expression::ConstantRef(name) = class_expr {
            // Special built-in types
            match name.as_str() {
                "Fiber" => {
                    if !args.is_empty() { self.compile_expression(&args[0])?; }
                    else { self.emit(Op::null); }
                    self.emit(Op::cont_new);
                    return Ok(());
                }
                "Exception" | "RuntimeError" | "StandardError" | "TypeError"
                | "ArgumentError" | "NameError" | "NoMethodError"
                | "ZeroDivisionError" | "RangeError" | "IOError" => {
                    let canonical = common::errors::canonical_exception_name(name);
                    let this_slot = self.define_local("__exc_this");
                    let msg_slot = self.define_local("__exc_msg");
                    if !args.is_empty() {
                        self.compile_expression(&args[0])?;
                    } else {
                        self.emit_constant(Value::String(Rc::from("")));
                    }
                    self.emit_u16(Op::local_set, msg_slot);
                    let line = self.line;
                    let c = self.current_chunk_idx;
                    common::errors::emit_exception_constructor(
                        &mut self.chunks[c], this_slot, canonical, msg_slot, line,
                    );
                    self.emit_u16(Op::local_get, this_slot);
                    return Ok(());
                }
                "Hash" | "OpenStruct" => {
                    let line = self.line;
                    let c = self.current_chunk_idx;
                    common::dict::emit_new(&mut self.chunks[c], line);
                    // If Hash.new(default) — store default value on the hash
                    if !args.is_empty() {
                        self.emit(Op::dup);
                        self.compile_expression(&args[0])?;
                        let dk = self.add_string_constant("__default");
                        self.emit_u16(Op::struct_set, dk);
                        self.emit(Op::drop);
                    }
                    return Ok(());
                }
                "Array" => {
                    self.emit_u16(Op::array_new, 0);
                    return Ok(());
                }
                "Set" => {
                    let i = self.import("vybe:types", "hashSetNew");
                    self.emit_host_call(i, 0);
                    return Ok(());
                }
                "StringBuilder" | "StringIO" => {
                    if !args.is_empty() { self.compile_expression(&args[0])?; }
                    else { self.emit_constant(Value::String(Rc::from(""))); }
                    let i = self.import("vybe:types", "stringBuilderNew");
                    self.emit_host_call(i, 1);
                    return Ok(());
                }
                "Random" => {
                    let i = self.import("vybe:threading", "randomNew");
                    self.emit_host_call(i, 0);
                    return Ok(());
                }
                "Time" | "DateTime" => {
                    if args.is_empty() {
                        let i = self.import("vybe:types", "dateTimeNow");
                        self.emit_host_call(i, 0);
                    } else {
                        self.compile_expression(&args[0])?;
                        let i = self.import("vybe:types", "dateTimeParse");
                        self.emit_host_call(i, 1);
                    }
                    return Ok(());
                }
                "TCPSocket" => {
                    for a in args { self.compile_expression(a)?; }
                    let i = self.import("vybe:net", "tcpConnect");
                    self.emit_host_call(i, args.len() as u8);
                    return Ok(());
                }
                "TCPServer" => {
                    for a in args { self.compile_expression(a)?; }
                    let i = self.import("vybe:net", "tcpListenerNew");
                    self.emit_host_call(i, args.len() as u8);
                    return Ok(());
                }
                "Mutex" => {
                    let ci = self.current_chunk_idx;
                    let alloc_fn = self.chunks[ci].add_import("wasi:thread", "allocLock");
                    let line = self.line;
                    self.chunks[ci].emit_op_u16(Op::call_import, alloc_fn, line);
                    self.chunks[ci].emit(0, line);
                    return Ok(());
                }
                "Thread" => {
                    // Thread.new { block } — handled by compile_method_call
                    for a in args { self.compile_expression(a)?; }
                    let c = self.current_chunk_idx;
                    let line = self.line;
                    common::threading::emit_thread_spawn(&mut self.chunks[c], line);
                    return Ok(());
                }
                _ => {} // fall through to user-defined class
            }
        }

        // User-defined class: get constructor from global, call it
        if let Expression::ConstantRef(name) = class_expr {
            if self.defined_classes.contains(name) || self.defined_globals.contains(name) {
                self.compile_expression(class_expr)?;
                for a in args { self.compile_expression(a)?; }
                self.emit_u8(Op::call_ref, args.len() as u8);
            } else {
                // Unresolved class — emit as host import call
                for a in args { self.compile_expression(a)?; }
                let idx = self.import("*", name);
                self.emit_host_call(idx, args.len() as u8);
            }
        } else {
            self.compile_expression(class_expr)?;
            for a in args { self.compile_expression(a)?; }
            self.emit_u8(Op::call_ref, args.len() as u8);
        }
        Ok(())
    }

    // ------------------------------------------------------------------
    // Method def compilation — uses common::functions
    // ------------------------------------------------------------------

    fn compile_method_def(&mut self, decl: &MethodDecl) -> Result<usize, String> {
        let chunk_idx = self.chunks.len();
        let chunk = common::functions::create_function_chunk(&decl.name, decl.params.len() as u8);
        self.chunks.push(chunk);

        self.scopes.push(Scope::new_function());
        let saved_chunk = self.current_chunk_idx;
        self.current_chunk_idx = chunk_idx;

        for param in &decl.params {
            self.define_local(&param.name);
        }

        // Default parameter values
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

        // Compile body. For the last statement, if it's an expression, compile
        // it WITHOUT drop — Ruby methods implicitly return the last expression.
        let body_len = decl.body.len();
        for (i, stmt) in decl.body.iter().enumerate() {
            if i == body_len - 1 {
                if let Statement::Expression(expr) = stmt {
                    self.compile_expression(expr)?;
                    // Don't drop — this is the return value. Emit return directly.
                    self.emit(Op::r#return);
                    // Skip epilogue since we already returned
                    let local_count = self.current_scope().next_slot;
                    self.chunks[chunk_idx].local_count = local_count;
                    self.scopes.pop();
                    self.current_chunk_idx = saved_chunk;
                    return Ok(chunk_idx);
                }
            }
            self.compile_statement(stmt)?;
        }

        // Safety net: if body is empty or last statement wasn't an expression
        let line = self.line;
        common::functions::emit_function_epilogue(&mut self.chunks[self.current_chunk_idx], line);

        let local_count = self.current_scope().next_slot;
        self.chunks[chunk_idx].local_count = local_count;

        self.scopes.pop();
        self.current_chunk_idx = saved_chunk;
        Ok(chunk_idx)
    }

    fn compile_method_for_class(&mut self, decl: &MethodDecl, _class_name: &str) -> Result<usize, String> {
        // Class methods receive `self` as first arg (same convention as VB/JS/C#).
        // Arity = 1 (self) + user params.
        let chunk_idx = self.chunks.len();
        let chunk = common::functions::create_function_chunk(
            &decl.name, (decl.params.len() + 1) as u8,
        );
        self.chunks.push(chunk);

        self.scopes.push(Scope::new_function());
        let saved_chunk = self.current_chunk_idx;
        self.current_chunk_idx = chunk_idx;

        // self is the first param (slot 1 after callee at slot 0)
        self.define_local("self");
        for param in &decl.params {
            self.define_local(&param.name);
        }

        // Default parameter values
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

        // Compile body — last expression is implicit return value
        let body_len = decl.body.len();
        for (i, stmt) in decl.body.iter().enumerate() {
            if i == body_len - 1 {
                if let Statement::Expression(expr) = stmt {
                    self.compile_expression(expr)?;
                    self.emit(Op::r#return);
                    let local_count = self.current_scope().next_slot;
                    self.chunks[chunk_idx].local_count = local_count;
                    self.scopes.pop();
                    self.current_chunk_idx = saved_chunk;
                    return Ok(chunk_idx);
                }
            }
            self.compile_statement(stmt)?;
        }

        let line = self.line;
        common::functions::emit_function_epilogue(&mut self.chunks[self.current_chunk_idx], line);

        let local_count = self.current_scope().next_slot;
        self.chunks[chunk_idx].local_count = local_count;
        self.scopes.pop();
        self.current_chunk_idx = saved_chunk;
        Ok(chunk_idx)
    }

    // ------------------------------------------------------------------
    // Class compilation — uses common::classes
    // ------------------------------------------------------------------

    fn compile_class(&mut self, decl: &ClassDecl) -> Result<(), String> {
        let class_name = &decl.name;
        let parent_name = decl.parent.as_deref().unwrap_or("").to_string();

        let mut method_entries: Vec<(String, usize)> = Vec::new();
        let mut static_method_entries: Vec<(String, usize)> = Vec::new();
        let mut init_chunk: Option<usize> = None;
        let mut init_params: Vec<String> = Vec::new();
        let mut attr_accessors: Vec<(String, AttrKind)> = Vec::new();

        // First pass: collect methods and attrs
        for stmt in &decl.body {
            match stmt {
                Statement::MethodDef(m) => {
                    let ci = self.compile_method_for_class(m, class_name)?;
                    if m.name == "initialize" {
                        init_chunk = Some(ci);
                        init_params = m.params.iter().map(|p| p.name.clone()).collect();
                    } else if m.is_self {
                        static_method_entries.push((m.name.clone(), ci));
                    } else {
                        method_entries.push((m.name.clone(), ci));
                    }
                }
                Statement::Expression(Expression::AttrDecl { kind, names }) => {
                    for name in names {
                        attr_accessors.push((name.clone(), kind.clone()));
                    }
                }
                _ => {}
            }
        }

        // Generate attr reader/writer methods
        for (name, kind) in &attr_accessors {
            match kind {
                AttrKind::Reader | AttrKind::Accessor => {
                    // def name; @name; end
                    let reader_decl = MethodDecl {
                        name: name.clone(),
                        params: Vec::new(),
                        body: vec![Statement::Return(Some(Expression::InstanceVar(name.clone())))],
                        is_self: false,
                    };
                    let ci = self.compile_method_for_class(&reader_decl, class_name)?;
                    method_entries.push((name.clone(), ci));
                }
                AttrKind::Writer => {}
            }
            match kind {
                AttrKind::Writer | AttrKind::Accessor => {
                    // def name=(val); @name = val; end
                    let writer_decl = MethodDecl {
                        name: format!("{}=", name),
                        params: vec![Param { name: "val".to_string(), default: None, splat: false, double_splat: false, block: false, keyword: false }],
                        body: vec![Statement::Assignment {
                            target: Expression::InstanceVar(name.clone()),
                            op: AssignOp::Assign,
                            value: Expression::Identifier("val".to_string()),
                        }],
                        is_self: false,
                    };
                    let ci = self.compile_method_for_class(&writer_decl, class_name)?;
                    method_entries.push((format!("{}=", name), ci));
                }
                AttrKind::Reader => {}
            }
        }

        // Constructor wrapper
        let ctor_chunk_idx = {
            let ctor_arity = init_params.len() as u8;
            let ci = self.chunks.len();
            let chunk = common::functions::create_function_chunk(&format!("{}_ctor", class_name), ctor_arity);
            self.chunks.push(chunk);

            let c = ci;
            let line = self.line;

            // Create new object via common::classes
            // Params occupy slots 1..N, this goes in slot N+1
            let this_idx = (ctor_arity as u16) + 1;
            common::classes::emit_new_typed_object(&mut self.chunks[c], this_idx, class_name, line);
            self.chunks[c].local_count = (ctor_arity as u16) + 2;

            // Bind methods
            for (mname, mci) in &method_entries {
                let line = self.line;
                common::classes::emit_bind_method_with_aliases(
                    &mut self.chunks[c], this_idx, mname, *mci, line,
                );
            }

            // Call initialize if present
            if let Some(init_ci) = init_chunk {
                self.chunks[c].emit_op_u16(Op::local_get, this_idx, line);
                // ref_func for init method
                common::functions::emit_ref_func(&mut self.chunks[c], init_ci, 0, line);
                self.chunks[c].emit_op_u16(Op::local_get, this_idx, line);
                // Push constructor params (slots 1..N)
                for i in 0..ctor_arity {
                    self.chunks[c].emit_op_u16(Op::local_get, (i as u16) + 1, line);
                }
                self.chunks[c].emit_op_u8(Op::call_ref, (ctor_arity + 1) as u8, line);
                self.chunks[c].emit_op(Op::drop, line);
            }

            // Return this
            self.chunks[c].emit_op_u16(Op::local_get, this_idx, line);
            self.chunks[c].emit_op(Op::r#return, line);

            ci
        };

        // Register type, store constructor
        let line = self.line;
        let c = self.current_chunk_idx;
        let method_names: Vec<String> = method_entries.iter().map(|(n, _)| n.clone()).collect();
        common::classes::register_type(
            &mut self.chunks,
            class_name,
            &parent_name,
            method_names,
            method_entries.clone(),
            false,
            Vec::new(),
            Some(ctor_chunk_idx),
        );

        // Store constructor as global
        let ctor_local = self.define_local("__ctor_slot");
        common::classes::emit_store_constructor(&mut self.chunks[c], class_name, ctor_chunk_idx, ctor_local, line);

        // Bind static methods to constructor
        for (sname, sci) in &static_method_entries {
            let line = self.line;
            let name_idx = self.add_string_constant(class_name);
            self.emit_u16(Op::global_get, name_idx);
            let ctor_slot = self.define_local("__static_ctor");
            self.emit_u16(Op::local_set, ctor_slot);
            common::classes::emit_bind_method_with_aliases(
                &mut self.chunks[self.current_chunk_idx], ctor_slot, sname, *sci, line,
            );
        }

        // Inheritance
        if !parent_name.is_empty() {
            let line = self.line;
            common::classes::emit_inherit_statics(&mut self.chunks[c], &parent_name, line);
            let this_slot = self.define_local("__super_slot");
            common::classes::emit_store_super(&mut self.chunks[c], this_slot, &parent_name, line);
        }

        self.defined_classes.insert(class_name.clone());
        Ok(())
    }

    // ------------------------------------------------------------------
    // Module compilation — modules are like classes without constructors
    // ------------------------------------------------------------------

    fn compile_module(&mut self, decl: &ModuleDecl) -> Result<(), String> {
        // Compile module as a global object with methods
        let line = self.line;
        let c = self.current_chunk_idx;
        common::dict::emit_new(&mut self.chunks[c], line);
        let mod_slot = self.define_local("__mod_tmp");
        self.emit_u16(Op::local_set, mod_slot);

        for stmt in &decl.body {
            if let Statement::MethodDef(m) = stmt {
                let ci = self.compile_method_def(m)?;
                let line = self.line;
                common::classes::emit_bind_method_with_aliases(
                    &mut self.chunks[self.current_chunk_idx], mod_slot, &m.name, ci, line,
                );
            }
        }

        self.emit_u16(Op::local_get, mod_slot);
        let idx = self.add_string_constant(&decl.name);
        self.emit_u16(Op::global_set, idx);
        self.defined_globals.insert(decl.name.clone());
        Ok(())
    }
}
