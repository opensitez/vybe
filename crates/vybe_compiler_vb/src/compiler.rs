use std::rc::Rc;
use std::collections::HashSet;

use vybe_bytecode::{Chunk, Value, Op};
use vybe_parser_basic::ast::*;

use crate::scope::Scope;

pub struct Compiler {
    pub(crate) chunks: Vec<Chunk>,
    pub(crate) scopes: Vec<Scope>,
    pub(crate) current_chunk_idx: usize,
    pub(crate) line: u32,
    pub(crate) defined_globals: HashSet<String>,
    pub(crate) defined_classes: HashSet<String>,
    pub(crate) function_name_stack: Vec<String>,
}

#[derive(Clone, Copy)]
pub(crate) enum VarResolution { Local(u16), Global }

impl Compiler {
    pub fn new() -> Self {
        Compiler {
            chunks: vec![Chunk::new("<script>")],
            scopes: vec![Scope::new()],
            current_chunk_idx: 0,
            line: 1,
            defined_globals: HashSet::new(),
            defined_classes: HashSet::new(),
            function_name_stack: Vec::new(),
        }
    }

    pub fn compile(mut self, program: &Program) -> Result<Vec<Chunk>, String> {
        for decl in &program.declarations {
            self.compile_declaration(decl)?;
        }
        for stmt in &program.statements {
            self.compile_statement(stmt)?;
        }
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

    // ---- Emit helpers ----

    pub(crate) fn import(&mut self, module: &str, name: &str) -> u16 {
        self.chunks[0].add_import(module, name)
    }
    pub(crate) fn emit(&mut self, op: Op) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op(op, line);
    }
    pub(crate) fn emit_u16(&mut self, op: Op, operand: u16) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u16(op, operand, line);
    }
    pub(crate) fn emit_u8(&mut self, op: Op, operand: u8) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u8(op, operand, line);
    }
    pub(crate) fn emit_constant(&mut self, value: Value) {
        let idx = self.chunks[self.current_chunk_idx].add_constant(value);
        self.emit_u16(Op::r#const, idx);
    }
    pub(crate) fn emit_jump(&mut self, op: Op) -> usize {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_jump(op, line)
    }
    pub(crate) fn patch_jump(&mut self, offset: usize) {
        self.chunks[self.current_chunk_idx].patch_jump(offset);
    }
    pub(crate) fn current_offset(&self) -> usize {
        self.chunks[self.current_chunk_idx].current_offset()
    }
    pub(crate) fn emit_loop(&mut self, target: usize) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_loop(target, line);
    }
    pub(crate) fn add_string_constant(&mut self, s: &str) -> u16 {
        self.chunks[self.current_chunk_idx].add_constant(Value::String(Rc::from(s)))
    }
    pub(crate) fn emit_host_call(&mut self, import_idx: u16, argc: u8) {
        let line = self.line;
        let c = &mut self.chunks[self.current_chunk_idx];
        c.emit_op(Op::call_import, line);
        c.emit((import_idx >> 8) as u8, line);
        c.emit((import_idx & 0xff) as u8, line);
        c.emit(argc, line);
    }
    pub(crate) fn emit_global_set(&mut self, name: &str) {
        let idx = self.add_string_constant(name);
        self.emit_u16(Op::global_set, idx);
        self.defined_globals.insert(name.to_lowercase());
    }

    // ---- Scope ----

    pub(crate) fn current_scope(&self) -> &Scope { self.scopes.last().unwrap() }
    pub(crate) fn current_scope_mut(&mut self) -> &mut Scope { self.scopes.last_mut().unwrap() }
    pub(crate) fn define_local(&mut self, name: &str) -> u16 { self.current_scope_mut().define_local(name) }

    pub(crate) fn resolve_variable(&self, name: &str) -> VarResolution {
        let lower = name.to_lowercase();
        if let Some(slot) = self.current_scope().resolve_local(&lower) {
            return VarResolution::Local(slot);
        }
        for scope in self.scopes.iter().rev().skip(1) {
            if scope.resolve_local(&lower).is_some() {
                return VarResolution::Local(scope.resolve_local(&lower).unwrap());
            }
        }
        VarResolution::Global
    }

    pub(crate) fn is_namespace(&self, name: &str) -> bool {
        matches!(name,
            "math" | "console" | "convert" | "strings" | "array"
            | "window" | "file" | "io" | "directory"
            | "vybe" | "system" | "application"
            | "environment" | "thread" | "json" | "color"
            | "datetime" | "stringbuilder" | "process"
        )
    }

    pub(crate) fn is_namespace_expr(&self, expr: &Expression) -> bool {
        match expr {
            Expression::Variable(name) => self.is_namespace(&name.as_str().to_lowercase()),
            Expression::MemberAccess(inner, _) => self.is_namespace_expr(inner),
            _ => false,
        }
    }

    // ---- Declarations ----

    pub(crate) fn compile_declaration(&mut self, decl: &Declaration) -> Result<(), String> {
        match decl {
            Declaration::Sub(sub) => {
                self.compile_sub(sub)?;
                let name = sub.name.as_str().to_lowercase();
                self.emit_global_set(&name);
                self.emit(Op::drop);
            }
            Declaration::Function(func) => {
                self.compile_function(func)?;
                let name = func.name.as_str().to_lowercase();
                self.emit_global_set(&name);
                self.emit(Op::drop);
            }
            Declaration::Class(class) => {
                self.compile_class(class)?;
                let name = class.name.as_str().to_lowercase();
                self.defined_classes.insert(name.clone());
                self.emit_global_set(&name);
                self.emit(Op::drop);
            }
            Declaration::Variable(vars) => {
                for var in vars {
                    if let Some(ref init) = var.initializer {
                        self.compile_expression(init)?;
                    } else {
                        self.emit(Op::null);
                    }
                    let name = var.name.as_str().to_lowercase();
                    if self.scopes.len() == 1 && self.current_scope().depth == 0 {
                        self.emit_global_set(&name);
                        self.emit(Op::drop);
                    } else {
                        let slot = self.define_local(&name);
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    }
                }
            }
            Declaration::Constant(c) => {
                self.compile_expression(&c.value)?;
                let name = c.name.as_str().to_lowercase();
                if self.scopes.len() == 1 {
                    self.emit_global_set(&name);
                    self.emit(Op::drop);
                } else {
                    let slot = self.define_local(&name);
                    self.emit_u16(Op::local_set, slot);
                    self.emit(Op::drop);
                }
            }
            Declaration::Imports(_) | Declaration::Namespace(_) |
            Declaration::Enum(_) | Declaration::Interface(_) |
            Declaration::Structure(_) | Declaration::Delegate(_) |
            Declaration::Event(_) => {
                // TODO: implement these
            }
        }
        Ok(())
    }

    // ---- Sub / Function / Class compilation ----

    pub(crate) fn compile_store_ident(&mut self, target: &Identifier) -> Result<(), String> {
        let name = target.as_str().to_lowercase();
        if let Some(func_name) = self.function_name_stack.last() {
            if name == *func_name {
                let rv_slot = self.current_scope().resolve_local("__return_val").unwrap();
                self.emit_u16(Op::local_set, rv_slot);
                self.emit(Op::drop);
                return Ok(());
            }
        }
        match self.resolve_variable(&name) {
            VarResolution::Local(slot) => {
                self.emit_u16(Op::local_set, slot);
                self.emit(Op::drop);
            }
            VarResolution::Global => {
                self.emit_global_set(&name);
                self.emit(Op::drop);
            }
        }
        Ok(())
    }

    fn compile_sub(&mut self, sub: &SubDecl) -> Result<(), String> {
        let name = sub.name.as_str();
        let mut chunk = Chunk::new(name);
        chunk.arity = sub.parameters.len() as u8;
        let idx = self.chunks.len();
        self.chunks.push(chunk);
        let mut scope = Scope::new_function();
        for param in &sub.parameters { scope.define_local(&param.name.as_str().to_lowercase()); }
        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);
        for stmt in &sub.body { self.compile_statement(stmt)?; }
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

    fn compile_function(&mut self, func: &FunctionDecl) -> Result<(), String> {
        self.compile_sub_like(&func.name, &func.parameters, &func.body, Some(&func.name))
    }

    fn compile_sub_like(&mut self, name: &Identifier, params: &[Parameter], body: &[Statement], return_var: Option<&Identifier>) -> Result<(), String> {
        let fname = name.as_str();
        let mut chunk = Chunk::new(fname);
        chunk.arity = params.len() as u8;
        let idx = self.chunks.len();
        self.chunks.push(chunk);
        let mut scope = Scope::new_function();
        for param in params { scope.define_local(&param.name.as_str().to_lowercase()); }
        if return_var.is_some() { scope.define_local("__return_val"); }
        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);
        if let Some(rv) = return_var { self.function_name_stack.push(rv.as_str().to_lowercase()); }
        if return_var.is_some() {
            self.emit(Op::null);
            let rv_slot = self.current_scope().resolve_local("__return_val").unwrap();
            self.emit_u16(Op::local_set, rv_slot);
            self.emit(Op::drop);
        }
        for stmt in body { self.compile_statement(stmt)?; }
        if return_var.is_some() { self.function_name_stack.pop(); }
        if return_var.is_some() {
            let rv_slot = self.current_scope().resolve_local("__return_val").unwrap();
            self.emit_u16(Op::local_get, rv_slot);
        } else {
            self.emit(Op::null);
        }
        self.emit(Op::r#return);
        let lc = self.current_scope().next_slot;
        self.chunks[idx].local_count = lc;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved;
        self.emit_ref_func(idx, &upvalues);
        Ok(())
    }

    // compile_class is in classes.rs

    pub(crate) fn emit_ref_func(&mut self, func_idx: usize, upvalues: &[crate::scope::UpvalueDesc]) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u16(Op::ref_func, func_idx as u16, line);
        self.chunks[self.current_chunk_idx].emit(upvalues.len() as u8, line);
        for uv in upvalues {
            self.chunks[self.current_chunk_idx].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current_chunk_idx].emit(uv.index, line);
        }
    }
}
