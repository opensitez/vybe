use std::rc::Rc;

use vybe_bytecode::{Chunk, Value, Op};
use vybe_parser_dart::*;

use crate::scope::Scope;

struct LoopContext {
    start_offset: usize,
    break_patches: Vec<usize>,
    continue_patches: Vec<usize>,
}

enum VarResolution { Local(u16), Upvalue(u8), Global }

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
        for top in &program.body {
            self.compile_top_level(top)?;
        }
        // Auto-call main if defined
        if self.defined_globals.contains("main") {
            let idx = self.add_string_constant("main");
            self.emit_u16(Op::global_get, idx);
            self.emit_u8(Op::call, 0);
            self.emit(Op::drop);
        }
        self.emit(Op::null);
        self.emit(Op::halt);
        let local_count = self.current_scope().next_slot;
        self.chunks[0].local_count = local_count;
        Ok(self.chunks)
    }

    // ── Emit helpers ──────────────────────────────────────────────────────

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
    fn import(&mut self, module: &str, name: &str) -> u16 {
        self.chunks[0].add_import(module, name)
    }
    fn emit_host_call(&mut self, import_idx: u16, argc: u8) {
        let line = self.line;
        let c = &mut self.chunks[self.current_chunk_idx];
        c.emit_op(Op::call_import, line);
        c.emit((import_idx >> 8) as u8, line);
        c.emit((import_idx & 0xff) as u8, line);
        c.emit(argc, line);
    }
    fn emit_global_set(&mut self, name: &str) {
        let idx = self.add_string_constant(name);
        self.emit_u16(Op::global_set, idx);
        self.defined_globals.insert(name.to_string());
    }
    fn emit_ref_func(&mut self, func_idx: usize, upvalues: &[crate::scope::UpvalueDesc]) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u16(Op::ref_func, func_idx as u16, line);
        self.chunks[self.current_chunk_idx].emit(upvalues.len() as u8, line);
        for uv in upvalues {
            self.chunks[self.current_chunk_idx].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current_chunk_idx].emit(uv.index, line);
        }
    }

    // ── Scope helpers ─────────────────────────────────────────────────────

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

    fn is_known_variable(&self, name: &str) -> bool {
        if self.current_scope().resolve_local(name).is_some() { return true; }
        for scope in self.scopes.iter().rev().skip(1) {
            if scope.resolve_local(name).is_some() { return true; }
        }
        self.defined_globals.contains(name)
    }

    // ── Dart module → host module mapping ─────────────────────────────────

    fn dart_module_alias(name: &str) -> &str {
        match name {
            "print" => "wasi:cli",
            "math" | "Math" => "vybe:math",
            "json" | "JSON" | "jsonDecode" | "jsonEncode" => "vybe:json",
            "http" | "Http" => "wasi:http",
            "File" | "Directory" => "wasi:filesystem",
            "Map" => "vybe:collections",
            "Set" => "vybe:collections",
            _ => name,
        }
    }

    fn resolve_bare_import(&mut self, name: &str) -> Option<u16> {
        match name {
            "print" => Some(self.import("wasi:cli", "log")),
            "int" | "double" => Some(self.import("vybe:convert", "cint")),
            _ => None,
        }
    }

    fn resolve_value_method(&mut self, method: &str) -> Option<u16> {
        let (module, name) = match method {
            "toUpperCase" => ("vybe:string", "toUpperCase"),
            "toLowerCase" => ("vybe:string", "toLowerCase"),
            "trim" => ("vybe:string", "trim"),
            "startsWith" => ("vybe:string", "startsWith"),
            "endsWith" => ("vybe:string", "endsWith"),
            "substring" => ("vybe:string", "substring"),
            "split" => ("vybe:string", "split"),
            "replaceAll" => ("vybe:string", "replaceAll"),
            "contains" => ("vybe:string", "includes"),
            "indexOf" => ("vybe:string", "indexOf"),
            "padLeft" => ("vybe:string", "padStart"),
            "padRight" => ("vybe:string", "padEnd"),
            "add" => ("vybe:array", "push"),
            "removeLast" => ("vybe:array", "pop"),
            "insert" => ("vybe:array", "push"),
            "join" => ("vybe:array", "join"),
            "reversed" => ("vybe:array", "reverse"),
            "sublist" => ("vybe:array", "slice"),
            _ => return None,
        };
        Some(self.import(module, name))
    }

    // ── Top Level ─────────────────────────────────────────────────────────

    fn compile_top_level(&mut self, top: &TopLevel) -> Result<(), String> {
        match top {
            TopLevel::Import(_) => Ok(()), // imports are informational
            TopLevel::Function(f) => {
                self.compile_function_decl(f)?;
                let name = &f.name;
                if self.scopes.len() == 1 && self.current_scope().depth == 0 {
                    self.emit_global_set(name);
                    self.emit(Op::drop);
                } else {
                    let slot = self.define_local(name);
                    self.emit_u16(Op::local_set, slot);
                    self.emit(Op::drop);
                }
                Ok(())
            }
            TopLevel::Class(c) => self.compile_class(c),
            TopLevel::Variable(v) => self.compile_var_decl(v),
            TopLevel::Statement(s) => self.compile_statement(s),
        }
    }

    // ── Statements ────────────────────────────────────────────────────────

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
            Statement::VarDecl(v) => { self.compile_var_decl(v)?; }
            Statement::FunctionDecl(f) => {
                self.compile_function_decl(f)?;
                let name = &f.name;
                if self.scopes.len() == 1 && self.current_scope().depth == 0 {
                    self.emit_global_set(name);
                    self.emit(Op::drop);
                } else {
                    let slot = self.define_local(name);
                    self.emit_u16(Op::local_set, slot);
                    self.emit(Op::drop);
                }
            }
            Statement::If { condition, then_branch, else_branch } => {
                self.compile_expression(condition)?;
                self.emit(Op::dyn_to_bool);
                let else_j = self.emit_jump(Op::br_if_false);
                self.compile_statement(then_branch)?;
                if let Some(alt) = else_branch {
                    let end_j = self.emit_jump(Op::br);
                    self.patch_jump(else_j);
                    self.compile_statement(alt)?;
                    self.patch_jump(end_j);
                } else {
                    self.patch_jump(else_j);
                }
            }
            Statement::While { condition, body } => {
                let start = self.current_offset();
                self.loop_stack.push(LoopContext { start_offset: start, break_patches: vec![], continue_patches: vec![] });
                self.compile_expression(condition)?;
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                self.compile_statement(body)?;
                for p in self.loop_stack.last().unwrap().continue_patches.clone() { self.patch_jump(p); }
                self.emit_loop(start);
                self.patch_jump(exit);
                let ctx = self.loop_stack.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
            }
            Statement::DoWhile { body, condition } => {
                let start = self.current_offset();
                self.loop_stack.push(LoopContext { start_offset: start, break_patches: vec![], continue_patches: vec![] });
                self.compile_statement(body)?;
                for p in self.loop_stack.last().unwrap().continue_patches.clone() { self.patch_jump(p); }
                self.compile_expression(condition)?;
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                self.emit_loop(start);
                self.patch_jump(exit);
                let ctx = self.loop_stack.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
            }
            Statement::For(for_stmt) => {
                self.current_scope_mut().begin_scope();
                if let Some(init) = &for_stmt.init {
                    match init {
                        ForInit::VarDecl(v) => { self.compile_var_decl(v)?; }
                        ForInit::Expression(e) => { self.compile_expression(e)?; self.emit(Op::drop); }
                    }
                }
                let start = self.current_offset();
                self.loop_stack.push(LoopContext { start_offset: start, break_patches: vec![], continue_patches: vec![] });
                let exit = if let Some(cond) = &for_stmt.condition {
                    self.compile_expression(cond)?;
                    self.emit(Op::dyn_to_bool);
                    Some(self.emit_jump(Op::br_if_false))
                } else { None };
                self.compile_statement(&for_stmt.body)?;
                for p in self.loop_stack.last().unwrap().continue_patches.clone() { self.patch_jump(p); }
                for upd in &for_stmt.update {
                    self.compile_expression(upd)?;
                    self.emit(Op::drop);
                }
                self.emit_loop(start);
                if let Some(e) = exit { self.patch_jump(e); }
                let ctx = self.loop_stack.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
                self.current_scope_mut().end_scope();
            }
            Statement::ForIn { var_name, iterable, body, .. } => {
                self.current_scope_mut().begin_scope();
                // __arr = iterable
                self.compile_expression(iterable)?;
                let arr_slot = self.define_local("__for_in_arr");
                self.emit_u16(Op::local_set, arr_slot);
                self.emit(Op::drop);
                // __i = 0
                self.emit_constant(Value::F64(0.0));
                let i_slot = self.define_local("__for_in_i");
                self.emit_u16(Op::local_set, i_slot);
                self.emit(Op::drop);
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { start_offset: loop_start, break_patches: vec![], continue_patches: vec![] });
                // __i < __arr.length
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit(Op::str_length);
                self.emit(Op::dyn_lt);
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                // var = __arr[__i]
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::array_get);
                let var_slot = self.define_local(var_name);
                self.emit_u16(Op::local_set, var_slot);
                self.emit(Op::drop);
                self.compile_statement(body)?;
                for p in self.loop_stack.last().unwrap().continue_patches.clone() { self.patch_jump(p); }
                // __i++
                self.emit_u16(Op::local_get, i_slot);
                self.emit_constant(Value::F64(1.0));
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot);
                self.emit(Op::drop);
                self.emit_loop(loop_start);
                self.patch_jump(exit);
                let ctx = self.loop_stack.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
                self.current_scope_mut().end_scope();
            }
            Statement::Switch { expr, cases } => {
                self.compile_expression(expr)?;
                self.loop_stack.push(LoopContext { start_offset: 0, break_patches: vec![], continue_patches: vec![] });
                let mut test_jumps: Vec<(usize, usize)> = Vec::new();
                let mut default_idx: Option<usize> = None;
                for (i, case) in cases.iter().enumerate() {
                    if let Some(lbl) = &case.label {
                        self.emit(Op::dup);
                        self.compile_expression(lbl)?;
                        self.emit(Op::eq);
                        let body_jump = self.emit_jump(Op::br_if_true);
                        test_jumps.push((i, body_jump));
                    } else {
                        default_idx = Some(i);
                    }
                }
                let to_default_or_end = self.emit_jump(Op::br);
                let mut body_offsets: Vec<usize> = Vec::new();
                for case in cases {
                    body_offsets.push(self.current_offset());
                    for s in &case.body { self.compile_statement(s)?; }
                }
                let _end = self.current_offset();
                for (case_idx, jump) in &test_jumps {
                    if *case_idx < body_offsets.len() {
                        let target = body_offsets[*case_idx];
                        let jump_ip = *jump + 2;
                        let offset = target as i16 - jump_ip as i16;
                        let c = &mut self.chunks[self.current_chunk_idx];
                        c.code[*jump] = (offset >> 8) as u8;
                        c.code[*jump + 1] = (offset & 0xff) as u8;
                    }
                }
                if let Some(di) = default_idx {
                    if di < body_offsets.len() {
                        let target = body_offsets[di];
                        let jump_ip = to_default_or_end + 2;
                        let offset = target as i16 - jump_ip as i16;
                        let c = &mut self.chunks[self.current_chunk_idx];
                        c.code[to_default_or_end] = (offset >> 8) as u8;
                        c.code[to_default_or_end + 1] = (offset & 0xff) as u8;
                    }
                } else {
                    self.patch_jump(to_default_or_end);
                }
                let ctx = self.loop_stack.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
                self.emit(Op::drop);
            }
            Statement::Return(val) => {
                if let Some(e) = val { self.compile_expression(e)?; }
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
            Statement::Throw(expr) => {
                self.compile_expression(expr)?;
                self.emit(Op::throw);
            }
            Statement::Try { body, catches, finally } => {
                let try_start_pos = self.current_offset();
                let line = self.line;
                let c = &mut self.chunks[self.current_chunk_idx];
                c.emit_op(Op::try_start, line);
                c.emit(0, line); c.emit(0, line);
                c.emit(0, line); c.emit(0, line);
                for s in body { self.compile_statement(s)?; }
                self.emit(Op::try_end);
                let skip_catch = self.emit_jump(Op::br);
                let catch_pos = self.current_offset();
                let ip_after = try_start_pos + 5;
                let catch_offset = catch_pos as i16 - ip_after as i16;
                let c = &mut self.chunks[self.current_chunk_idx];
                c.code[try_start_pos + 1] = (catch_offset >> 8) as u8;
                c.code[try_start_pos + 2] = (catch_offset & 0xff) as u8;
                if let Some(catch) = catches.first() {
                    self.current_scope_mut().begin_scope();
                    if let Some(ref var) = catch.var_name {
                        let slot = self.define_local(var);
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    } else {
                        self.emit(Op::drop);
                    }
                    for s in &catch.body { self.compile_statement(s)?; }
                    self.current_scope_mut().end_scope();
                } else {
                    self.emit(Op::drop);
                }
                self.patch_jump(skip_catch);
                if let Some(fin) = finally {
                    for s in fin { self.compile_statement(s)?; }
                }
            }
            Statement::Assert(cond, msg) => {
                self.compile_expression(cond)?;
                self.emit(Op::dyn_to_bool);
                let ok = self.emit_jump(Op::br_if_true);
                if let Some(m) = msg {
                    self.compile_expression(m)?;
                } else {
                    self.emit_constant(Value::String(Rc::from("Assertion failed")));
                }
                self.emit(Op::throw);
                self.patch_jump(ok);
            }
            Statement::Empty => {}
        }
        Ok(())
    }

    // ── Variable declarations ─────────────────────────────────────────────

    fn compile_var_decl(&mut self, v: &VarDecl) -> Result<(), String> {
        if let Some(init) = &v.initializer {
            self.compile_expression(init)?;
        } else {
            self.emit(Op::null);
        }
        let name = &v.name;
        if self.scopes.len() == 1 && self.current_scope().depth == 0 {
            self.emit_global_set(name);
            self.emit(Op::drop);
        } else {
            let slot = self.define_local(name);
            self.emit_u16(Op::local_set, slot);
            self.emit(Op::drop);
        }
        Ok(())
    }

    // ── Function compilation ──────────────────────────────────────────────

    fn compile_function_decl(&mut self, f: &FunctionDecl) -> Result<(), String> {
        let name = &f.name;
        let arity = f.params.positional.len() + f.params.optional_pos.len() + f.params.named.len();
        let mut chunk = Chunk::new(name);
        chunk.arity = arity as u8;
        let idx = self.chunks.len();
        self.chunks.push(chunk);
        let mut scope = Scope::new_function();
        for p in &f.params.positional { scope.define_local(&p.name); }
        for p in &f.params.optional_pos { scope.define_local(&p.name); }
        for p in &f.params.named { scope.define_local(&p.name); }
        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);
        match &f.body {
            FunctionBody::Block(stmts) => {
                for s in stmts { self.compile_statement(s)?; }
            }
            FunctionBody::Expression(expr) => {
                self.compile_expression(expr)?;
                self.emit(Op::r#return);
                let lc = self.current_scope().next_slot;
                self.chunks[idx].local_count = lc;
                let upvalues = self.current_scope().upvalues.clone();
                self.scopes.pop();
                self.current_chunk_idx = saved;
                self.emit_ref_func(idx, &upvalues);
                return Ok(());
            }
            FunctionBody::Empty => {}
        }
        self.emit(Op::null);
        self.emit(Op::r#return);
        let lc = self.current_scope().next_slot;
        self.chunks[idx].local_count = lc;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved;
        self.emit_ref_func(idx, &upvalues);
        Ok(())
    }

    // ── Class compilation ─────────────────────────────────────────────────

    fn compile_class(&mut self, class: &ClassDecl) -> Result<(), String> {
        // Class → constructor function that creates objects with methods bound
        let class_name = &class.name;
        // Compile constructor function
        let mut chunk = Chunk::new(class_name);
        // Find constructor to get arity
        let ctor = class.members.iter().find(|m| matches!(m, ClassMember::Constructor { .. }));
        let ctor_arity = match ctor {
            Some(ClassMember::Constructor { params, .. }) => {
                params.positional.len() + params.optional_pos.len() + params.named.len()
            }
            _ => 0,
        };
        chunk.arity = ctor_arity as u8;
        let idx = self.chunks.len();
        self.chunks.push(chunk);
        let mut scope = Scope::new_function();
        // Define constructor params
        if let Some(ClassMember::Constructor { params, .. }) = ctor {
            for p in &params.positional { scope.define_local(&p.name); }
            for p in &params.optional_pos { scope.define_local(&p.name); }
            for p in &params.named { scope.define_local(&p.name); }
        }
        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        // Create `this` object
        self.emit_u16(Op::struct_new, 0);
        let this_slot = self.define_local("this");
        self.emit_u16(Op::local_set, this_slot);
        self.emit(Op::drop);

        // Set __type on the instance
        self.emit_u16(Op::local_get, this_slot);
        self.emit_constant(Value::String(Rc::from(class_name.as_str())));
        let type_idx = self.add_string_constant("__type");
        self.emit_u16(Op::struct_set, type_idx);
        self.emit(Op::drop);

        // Handle this.field params
        if let Some(ClassMember::Constructor { params, .. }) = ctor {
            for p in &params.positional {
                if p.is_this {
                    if let Some(slot) = self.current_scope().resolve_local(&p.name) {
                        self.emit_u16(Op::local_get, this_slot);
                        self.emit_u16(Op::local_get, slot);
                        let prop_idx = self.add_string_constant(&p.name);
                        self.emit_u16(Op::struct_set, prop_idx);
                        self.emit(Op::drop);
                    }
                }
            }
        }

        // Initialize fields
        for member in &class.members {
            if let ClassMember::Field { name, initializer, .. } = member {
                self.emit_u16(Op::local_get, this_slot);
                if let Some(init) = initializer {
                    self.compile_expression(init)?;
                } else {
                    self.emit(Op::null);
                }
                let prop_idx = self.add_string_constant(name);
                self.emit_u16(Op::struct_set, prop_idx);
                self.emit(Op::drop);
            }
        }

        // Constructor body
        if let Some(ClassMember::Constructor { body, .. }) = ctor {
            if let Some(stmts) = body {
                for s in stmts { self.compile_statement(s)?; }
            }
        }

        // Bind methods to the instance
        for member in &class.members {
            if let ClassMember::Method { decl, is_static, .. } = member {
                if *is_static { continue; }
                // Compile the method as a nested function
                self.compile_function_decl(decl)?;
                // Set it on `this`
                self.emit_u16(Op::local_get, this_slot);
                // Stack: [closure, this] — need [this, closure] for struct_set
                // Swap using temp
                let tmp = self.define_local(&format!("__m_{}", decl.name));
                // closure is below this. Actually ref_func pushed closure, then local_get pushed this.
                // Stack: [..., closure, this]
                self.emit_u16(Op::local_set, tmp); // save this
                self.emit(Op::drop);
                // Stack: [..., closure]
                let tmp2 = self.define_local(&format!("__mc_{}", decl.name));
                self.emit_u16(Op::local_set, tmp2); // save closure
                self.emit(Op::drop);
                // Now push: this, closure
                self.emit_u16(Op::local_get, tmp); // this
                self.emit_u16(Op::local_get, tmp2); // closure
                let method_idx = self.add_string_constant(&decl.name);
                self.emit_u16(Op::struct_set, method_idx);
                self.emit(Op::drop);
            }
        }

        // Return this
        self.emit_u16(Op::local_get, this_slot);
        self.emit(Op::r#return);
        let lc = self.current_scope().next_slot;
        self.chunks[idx].local_count = lc;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved;
        self.emit_ref_func(idx, &upvalues);
        // Set as global
        self.defined_classes.insert(class_name.clone());
        self.emit_global_set(class_name);
        self.emit(Op::drop);
        Ok(())
    }

    // ── Expressions ───────────────────────────────────────────────────────

    fn compile_expression(&mut self, expr: &Expression) -> Result<(), String> {
        match expr {
            Expression::Int(n) => {
                if *n >= 0 && *n <= i32::MAX as i64 {
                    self.emit_constant(Value::I32(*n as i32));
                } else {
                    self.emit_constant(Value::I64(*n));
                }
            }
            Expression::Double(n) => { self.emit_constant(Value::F64(*n)); }
            Expression::Bool(true) => { self.emit(Op::r#true); }
            Expression::Bool(false) => { self.emit(Op::r#false); }
            Expression::Null => { self.emit(Op::null); }
            Expression::String(s) => {
                match s {
                    StringExpr::Simple(text) => {
                        self.emit_constant(Value::String(Rc::from(text.as_str())));
                    }
                    StringExpr::Interpolated(parts) => {
                        let count = parts.len();
                        for part in parts {
                            match part {
                                StringPart::Literal(lit) => {
                                    self.emit_constant(Value::String(Rc::from(lit.as_str())));
                                }
                                StringPart::Expr(e) => {
                                    self.compile_expression(e)?;
                                }
                            }
                        }
                        if count > 1 {
                            self.emit_u8(Op::str_concat_n, count as u8);
                        }
                    }
                }
            }
            Expression::Identifier(name) => {
                // Check for bare host function calls handled at call sites
                match self.resolve_variable(name) {
                    VarResolution::Local(slot) => { self.emit_u16(Op::local_get, slot); }
                    VarResolution::Upvalue(idx) => { self.emit_u8(Op::upvalue_get, idx); }
                    VarResolution::Global => {
                        let idx = self.add_string_constant(name);
                        self.emit_u16(Op::global_get, idx);
                    }
                }
            }
            Expression::This => {
                if let Some(slot) = self.current_scope().resolve_local("this") {
                    self.emit_u16(Op::local_get, slot);
                } else {
                    self.emit(Op::null);
                }
            }
            Expression::Super => { self.emit(Op::null); }
            Expression::List { elements, .. } => {
                for e in elements { self.compile_expression(e)?; }
                self.emit_u16(Op::array_new, elements.len() as u16);
            }
            Expression::Map { entries, .. } => {
                // Build map via host call
                let ctor = self.import("vybe:collections", "Map");
                self.emit_host_call(ctor, 0);
                for (key, val) in entries {
                    self.emit(Op::dup);
                    self.compile_expression(key)?;
                    self.compile_expression(val)?;
                    let set_idx = self.import("vybe:collections", "mapSet");
                    self.emit_host_call(set_idx, 3);
                    self.emit(Op::drop);
                }
            }
            Expression::Set { elements, .. } => {
                let ctor = self.import("vybe:collections", "Set");
                self.emit_host_call(ctor, 0);
                for e in elements {
                    self.emit(Op::dup);
                    self.compile_expression(e)?;
                    let add_idx = self.import("vybe:collections", "setAdd");
                    self.emit_host_call(add_idx, 2);
                    self.emit(Op::drop);
                }
            }
            Expression::Binary { op, left, right } => {
                self.compile_expression(left)?;
                self.compile_expression(right)?;
                match op {
                    BinOp::Add => self.emit(Op::dyn_add),
                    BinOp::Sub => self.emit(Op::f64_sub),
                    BinOp::Mul => self.emit(Op::f64_mul),
                    BinOp::Div => self.emit(Op::f64_div),
                    BinOp::IntDiv => { self.emit(Op::f64_div); self.emit(Op::f64_floor); }
                    BinOp::Mod => self.emit(Op::f64_mod),
                    BinOp::Eq => self.emit(Op::dyn_eq),
                    BinOp::NotEq => self.emit(Op::dyn_ne),
                    BinOp::Lt => self.emit(Op::dyn_lt),
                    BinOp::Gt => self.emit(Op::dyn_gt),
                    BinOp::Le => self.emit(Op::dyn_le),
                    BinOp::Ge => self.emit(Op::dyn_ge),
                    BinOp::And => {
                        // Short-circuit: already compiled both — but for proper short-circuit
                        // we'd need to restructure. For now, use logical and.
                        self.emit(Op::i32_and);
                    }
                    BinOp::Or => { self.emit(Op::i32_or); }
                    BinOp::BitAnd => self.emit(Op::i32_and),
                    BinOp::BitOr => self.emit(Op::i32_or),
                    BinOp::BitXor => self.emit(Op::i32_xor),
                    BinOp::Shl => self.emit(Op::i32_shl),
                    BinOp::Shr => self.emit(Op::i32_shr_s),
                    BinOp::UShr => self.emit(Op::i32_shr_u),
                }
            }
            Expression::Unary { op, expr: inner } => {
                match op {
                    UnaryOp::Neg => {
                        self.compile_expression(inner)?;
                        self.emit(Op::dyn_neg);
                    }
                    UnaryOp::Not => {
                        self.compile_expression(inner)?;
                        self.emit(Op::dyn_not);
                    }
                    UnaryOp::BitNot => {
                        self.compile_expression(inner)?;
                        self.emit(Op::i32_not);
                    }
                    UnaryOp::PreInc => {
                        self.compile_expression(inner)?;
                        self.emit_constant(Value::F64(1.0));
                        self.emit(Op::dyn_add);
                        self.compile_store(inner)?;
                        self.compile_expression(inner)?;
                    }
                    UnaryOp::PreDec => {
                        self.compile_expression(inner)?;
                        self.emit_constant(Value::F64(1.0));
                        self.emit(Op::f64_sub);
                        self.compile_store(inner)?;
                        self.compile_expression(inner)?;
                    }
                }
            }
            Expression::PostfixUnary { op, expr: inner } => {
                self.compile_expression(inner)?;
                self.emit(Op::dup); // keep original value
                self.emit_constant(Value::F64(1.0));
                match op {
                    PostfixOp::PostInc => self.emit(Op::dyn_add),
                    PostfixOp::PostDec => self.emit(Op::f64_sub),
                }
                self.compile_store(inner)?;
                // Original value is still on stack from dup
            }
            Expression::Assign { op, left, right } => {
                match op {
                    AssignOp::Assign => {
                        self.compile_expression(right)?;
                    }
                    _ => {
                        self.compile_expression(left)?;
                        self.compile_expression(right)?;
                        match op {
                            AssignOp::AddAssign => self.emit(Op::dyn_add),
                            AssignOp::SubAssign => self.emit(Op::f64_sub),
                            AssignOp::MulAssign => self.emit(Op::f64_mul),
                            AssignOp::DivAssign => self.emit(Op::f64_div),
                            AssignOp::ModAssign => self.emit(Op::f64_mod),
                            _ => self.emit(Op::dyn_add), // fallback
                        }
                    }
                }
                self.emit(Op::dup); // assignment is an expression, keep value
                self.compile_store(left)?;
            }
            Expression::Ternary { cond, then, else_ } => {
                self.compile_expression(cond)?;
                self.emit(Op::dyn_to_bool);
                let else_j = self.emit_jump(Op::br_if_false);
                self.compile_expression(then)?;
                let end_j = self.emit_jump(Op::br);
                self.patch_jump(else_j);
                self.compile_expression(else_)?;
                self.patch_jump(end_j);
            }
            Expression::NullCoalesce { left, right } => {
                self.compile_expression(left)?;
                self.emit(Op::dup);
                let skip = self.emit_jump(Op::br_if_null);
                // left is not null — skip right
                let end = self.emit_jump(Op::br);
                self.patch_jump(skip);
                self.emit(Op::drop); // drop null
                self.compile_expression(right)?;
                self.patch_jump(end);
            }
            Expression::Member { object, member, .. } => {
                self.compile_expression(object)?;
                // Check for .length
                if member == "length" {
                    self.emit(Op::str_length);
                } else {
                    let prop_idx = self.add_string_constant(member);
                    self.emit_u16(Op::struct_get, prop_idx);
                }
            }
            Expression::Index { object, index } => {
                self.compile_expression(object)?;
                self.compile_expression(index)?;
                self.emit(Op::array_get);
            }
            Expression::Call { callee, args, .. } => {
                self.compile_call(callee, args)?;
            }
            Expression::New { class, args, .. } => {
                // Call class constructor
                let idx = self.add_string_constant(class);
                self.emit_u16(Op::global_get, idx);
                for arg in args { self.compile_expression(&arg.value)?; }
                self.emit_u8(Op::call, args.len() as u8);
            }
            Expression::Const { class, args, .. } => {
                let idx = self.add_string_constant(class);
                self.emit_u16(Op::global_get, idx);
                for arg in args { self.compile_expression(&arg.value)?; }
                self.emit_u8(Op::call, args.len() as u8);
            }
            Expression::Lambda { params, body, .. } => {
                let arity = params.positional.len() + params.optional_pos.len() + params.named.len();
                let mut chunk = Chunk::new("<lambda>");
                chunk.arity = arity as u8;
                let idx = self.chunks.len();
                self.chunks.push(chunk);
                let mut scope = Scope::new_function();
                for p in &params.positional { scope.define_local(&p.name); }
                for p in &params.optional_pos { scope.define_local(&p.name); }
                for p in &params.named { scope.define_local(&p.name); }
                let saved = self.current_chunk_idx;
                self.current_chunk_idx = idx;
                self.scopes.push(scope);
                match body.as_ref() {
                    FunctionBody::Block(stmts) => {
                        for s in stmts { self.compile_statement(s)?; }
                        self.emit(Op::null);
                        self.emit(Op::r#return);
                    }
                    FunctionBody::Expression(e) => {
                        self.compile_expression(e)?;
                        self.emit(Op::r#return);
                    }
                    FunctionBody::Empty => {
                        self.emit(Op::null);
                        self.emit(Op::r#return);
                    }
                }
                let lc = self.current_scope().next_slot;
                self.chunks[idx].local_count = lc;
                let upvalues = self.current_scope().upvalues.clone();
                self.scopes.pop();
                self.current_chunk_idx = saved;
                self.emit_ref_func(idx, &upvalues);
            }
            Expression::Is { expr: inner, type_ann, negated } => {
                self.compile_expression(inner)?;
                let type_idx = self.add_string_constant(&type_ann.name);
                self.emit_u16(Op::ref_test, type_idx);
                if *negated { self.emit(Op::bool_not); }
            }
            Expression::As { expr: inner, .. } => {
                // Type cast — in our dynamic VM, just pass through
                self.compile_expression(inner)?;
            }
            Expression::Await(inner) => {
                self.compile_expression(inner)?;
                self.emit(Op::r#await);
            }
            Expression::Spread(inner) => {
                self.compile_expression(inner)?;
                self.emit(Op::spread);
            }
            Expression::IfNull { left, right } => {
                self.compile_expression(left)?;
                self.emit(Op::dup);
                let not_null = self.emit_jump(Op::br_if_null);
                let end = self.emit_jump(Op::br);
                self.patch_jump(not_null);
                self.emit(Op::drop);
                self.compile_expression(right)?;
                self.patch_jump(end);
            }
            Expression::Cascade { object, ops } => {
                self.compile_expression(object)?;
                for op in ops {
                    self.emit(Op::dup);
                    match op {
                        CascadeOp::Method(name, args) => {
                            for a in args { self.compile_expression(&a.value)?; }
                            let prop_idx = self.add_string_constant(name);
                            self.emit_u16(Op::struct_get, prop_idx);
                            self.emit_u8(Op::call, args.len() as u8);
                            self.emit(Op::drop);
                        }
                        CascadeOp::Assign(field, val) => {
                            self.compile_expression(val)?;
                            let prop_idx = self.add_string_constant(field);
                            self.emit_u16(Op::struct_set, prop_idx);
                            self.emit(Op::drop);
                        }
                        CascadeOp::Field(name) => {
                            let prop_idx = self.add_string_constant(name);
                            self.emit_u16(Op::struct_get, prop_idx);
                            self.emit(Op::drop);
                        }
                        CascadeOp::Index(idx_expr) => {
                            self.compile_expression(idx_expr)?;
                            self.emit(Op::array_get);
                            self.emit(Op::drop);
                        }
                    }
                }
            }
        }
        Ok(())
    }

    // ── Call compilation ──────────────────────────────────────────────────

    fn compile_call(&mut self, callee: &Expression, args: &[Argument]) -> Result<(), String> {
        // Handle print() as a bare host call
        if let Expression::Identifier(name) = callee {
            if name == "print" {
                for a in args { self.compile_expression(&a.value)?; }
                let idx = self.import("wasi:cli", "log");
                self.emit_host_call(idx, args.len() as u8);
                return Ok(());
            }
            // Other bare imports
            if let Some(imp) = self.resolve_bare_import(name) {
                for a in args { self.compile_expression(&a.value)?; }
                self.emit_host_call(imp, args.len() as u8);
                return Ok(());
            }
        }

        // Handle method calls: obj.method(args) 
        if let Expression::Member { object, member, .. } = callee {
            // Check if it's a namespace call (e.g. math.sqrt)
            if let Expression::Identifier(obj_name) = object.as_ref() {
                if !self.is_known_variable(obj_name) {
                    let module = Self::dart_module_alias(obj_name);
                    for a in args { self.compile_expression(&a.value)?; }
                    let import_idx = self.import(module, member);
                    self.emit_host_call(import_idx, args.len() as u8);
                    return Ok(());
                }
            }

            // Instance method: check for known value methods
            if let Some(imp) = self.resolve_value_method(member) {
                // Push object as first arg, then remaining args
                self.compile_expression(object)?;
                for a in args { self.compile_expression(&a.value)?; }
                self.emit_host_call(imp, (args.len() + 1) as u8);
                return Ok(());
            }

            // toString() method
            if member == "toString" {
                self.compile_expression(object)?;
                let idx = self.import("vybe:convert", "toString");
                self.emit_host_call(idx, 1);
                return Ok(());
            }

            // Generic method call: obj.method(...) → struct_get + call
            self.compile_expression(object)?;
            let prop_idx = self.add_string_constant(member);
            self.emit_u16(Op::struct_get, prop_idx);
            for a in args { self.compile_expression(&a.value)?; }
            self.emit_u8(Op::call, args.len() as u8);
            return Ok(());
        }

        // Generic function call
        self.compile_expression(callee)?;
        for a in args { self.compile_expression(&a.value)?; }
        self.emit_u8(Op::call, args.len() as u8);
        Ok(())
    }

    // ── Store helpers ─────────────────────────────────────────────────────

    fn compile_store(&mut self, target: &Expression) -> Result<(), String> {
        match target {
            Expression::Identifier(name) => {
                match self.resolve_variable(name) {
                    VarResolution::Local(slot) => {
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    }
                    VarResolution::Upvalue(idx) => {
                        self.emit_u8(Op::upvalue_set, idx);
                        self.emit(Op::drop);
                    }
                    VarResolution::Global => {
                        self.emit_global_set(name);
                        self.emit(Op::drop);
                    }
                }
            }
            Expression::Member { object, member, .. } => {
                self.compile_expression(object)?;
                // Stack: [value, obj] — struct_set expects [obj, val]
                // Value was pushed before this call. We need to swap.
                // Use temp local
                let tmp = self.define_local("__store_tmp");
                self.emit_u16(Op::local_set, tmp); // save obj
                self.emit(Op::drop);
                let tmp2 = self.define_local("__store_val");
                self.emit_u16(Op::local_set, tmp2); // save val
                self.emit(Op::drop);
                self.emit_u16(Op::local_get, tmp); // push obj
                self.emit_u16(Op::local_get, tmp2); // push val
                let prop_idx = self.add_string_constant(member);
                self.emit_u16(Op::struct_set, prop_idx);
                self.emit(Op::drop);
            }
            Expression::Index { object, index } => {
                let tmp = self.define_local("__idx_val");
                self.emit_u16(Op::local_set, tmp);
                self.emit(Op::drop);
                self.compile_expression(object)?;
                self.compile_expression(index)?;
                self.emit_u16(Op::local_get, tmp);
                self.emit(Op::array_set);
                self.emit(Op::drop);
            }
            _ => {} // can't store to other expression types
        }
        Ok(())
    }
}
