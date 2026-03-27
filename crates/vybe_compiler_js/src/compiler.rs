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
    /// Track names that have been set as globals (class declarations, function declarations, var).
    defined_globals: std::collections::HashSet<String>,
    /// Track names that are class constructors (for static method dispatch).
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
            in_method: false,
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
        Ok(self.chunks)
    }

    /// Emit a global set and track the name as defined.
    fn emit_global_set(&mut self, name: &str) {
        let idx = self.add_string_constant(name);
        self.emit_u16(Op::global_set, idx);
        self.defined_globals.insert(name.to_string());
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
        for scope in self.scopes.iter().rev().skip(1) {
            if scope.resolve_local(name).is_some() {
                return true;
            }
        }
        // Also check previously defined globals
        if self.defined_globals.contains(name) {
            return true;
        }
        false
    }

    /// JS-to-module alias table. Maps JS namespace objects to VSI module names.
    /// If a JS identifier isn't a variable and isn't in this table, it's passed
    /// through as-is (e.g. "MyModule" → ("MyModule", "method")).
    fn js_module_alias(obj_name: &str) -> &str {
        match obj_name {
            // JS standard objects → WASI system modules
            "console" => "wasi:cli",
            "Date"    => "wasi:clocks",
            // JS standard objects → vybe language runtime
            "Math"    => "vybe:math",
            "JSON"    => "vybe:json",
            "Object"  => "vybe:object",
            "RegExp"  => "vybe:regex",
            "Array"   => "vybe:array",
            "String"  => "vybe:string",
            "Number"  => "vybe:convert",
            "Map"     => "vybe:collections",
            "Set"     => "vybe:collections",
            // Platform modules — WASI names for system I/O
            "fs"      => "wasi:filesystem",
            "clock"   => "wasi:clocks",
            "env"     => "wasi:cli",
            "random"  => "wasi:random",
            "http"    => "wasi:http",
            // Platform modules — vybe names for non-WASI
            "gui"     => "vybe:gui",
            "db"      => "vybe:database",
            "Promise" => "vybe:runtime",
            // Unknown → pass through as-is
            _ => obj_name,
        }
    }

    /// JS-specific name remapping. Most methods pass through as-is.
    /// Only remap when JS convention differs from VSI naming.
    fn js_remap<'a>(module: &'a str, method: &'a str) -> (&'a str, &'a str) {
        match (module, method) {
            ("vybe:math", "random") => ("wasi:random", "random"),
            ("vybe:runtime", "resolve") => ("vybe:runtime", "promiseResolve"),
            ("vybe:runtime", "reject") => ("vybe:runtime", "promiseReject"),
            ("vybe:runtime", "all") => ("vybe:runtime", "promiseAll"),
            _ => (module, method),
        }
    }

    /// Resolve bare global calls that are host imports, not user functions.
    /// Each language compiler defines its own set of bare imports.
    /// Resolve builtin constructors: new Map(), new Set(), new RegExp()
    fn resolve_builtin_constructor(&mut self, name: &str) -> Option<u16> {
        match name {
            "Map" => Some(self.import("vybe:collections", "Map")),
            "Set" => Some(self.import("vybe:collections", "Set")),
            "Error" => Some(self.import("vybe:runtime", "Error")),
            "TypeError" => Some(self.import("vybe:runtime", "TypeError")),
            "RangeError" => Some(self.import("vybe:runtime", "RangeError")),
            _ => None,
        }
    }

    fn resolve_bare_import(&mut self, name: &str) -> Option<u16> {
        match name {
            // JS standard globals — encoding/decoding
            // setTimeout is handled specially in compile_call
            // "setTimeout" => handled below,
            "btoa" => return Some(self.import("vybe:convert", "btoa")),
            "atob" => return Some(self.import("vybe:convert", "atob")),
            "encodeURIComponent" => return Some(self.import("vybe:convert", "encodeURIComponent")),
            "decodeURIComponent" => return Some(self.import("vybe:convert", "decodeURIComponent")),
            // JS standard globals — type conversion
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
            "charAt" | "substring" | "split" | "replace" | "replaceAll" |
            "charCodeAt" | "repeat" | "padStart" | "padEnd" => ("vybe:string", method),
            // Array methods
            "push" | "pop" | "shift" | "join" | "reverse" | "concat" |
            "fill" | "flat" => ("vybe:array", method),
            // Shared — host dispatches by type at runtime
            "slice" => ("vybe:array", "slice"),
            "indexOf" => ("vybe:string", "indexOf"),
            "includes" => ("vybe:string", "includes"),
            // Note: Map/Set methods (set, get, has, delete, add, keys, values)
            // are NOT listed here because they conflict with user class methods.
            // They go through the regular method call path (obj.method → struct_get + call).
            // The collections module attaches these methods on the instance during construction.
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
                    // Value is on stack — bind it to the pattern
                    self.compile_binding(&decl.pattern, *kind)?;
                }
            }
            Statement::FunctionDeclaration(func) => {
                self.compile_function(func)?;
                if let Some(name) = &func.name {
                    if self.scopes.len() == 1 && self.current_scope().depth == 0 {
                        self.emit_global_set(name);
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
                // Emit try_start with placeholder catch offset
                let try_start_pos = self.current_offset();
                let line = self.line;
                let c = &mut self.chunks[self.current_chunk_idx];
                c.emit_op(Op::try_start, line);
                c.emit(0, line); c.emit(0, line); // catch offset placeholder
                c.emit(0, line); c.emit(0, line); // finally offset placeholder

                // Compile try block
                for s in block { self.compile_statement(s)?; }
                self.emit(Op::try_end);
                let skip_catch = self.emit_jump(Op::br); // jump over catch block

                // Patch catch offset: relative from IP after try_start reads its operands
                // IP after try_start = try_start_pos + 5 (1 opcode + 2 catch + 2 finally)
                let catch_pos = self.current_offset();
                let ip_after_try_start = try_start_pos + 5;
                let catch_offset = catch_pos as i16 - ip_after_try_start as i16;
                let c = &mut self.chunks[self.current_chunk_idx];
                c.code[try_start_pos + 1] = (catch_offset >> 8) as u8;
                c.code[try_start_pos + 2] = (catch_offset & 0xff) as u8;

                // Catch block — exception value is on stack
                if let Some(h) = handler {
                    self.current_scope_mut().begin_scope();
                    if let Some(ref param) = h.param {
                        // Bind the exception to the catch parameter
                        let slot = self.define_local(param);
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    } else {
                        self.emit(Op::drop); // discard exception if no param
                    }
                    for s in &h.body { self.compile_statement(s)?; }
                    self.current_scope_mut().end_scope();
                } else {
                    self.emit(Op::drop); // discard exception
                }

                self.patch_jump(skip_catch);

                // Finally block
                if let Some(f) = finalizer {
                    for s in f { self.compile_statement(s)?; }
                }
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
                    self.defined_classes.insert(name.clone());
                    if self.scopes.len() == 1 && self.current_scope().depth == 0 {
                        self.emit_global_set(name);
                        self.emit(Op::drop);
                    } else {
                        let slot = self.define_local(name);
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    }
                }
            }
            Statement::ForOf { left, right, body } => {
                // Desugar: for (let x of arr) → { let __arr = arr; let __i = 0; while (__i < __arr.length) { let x = __arr[__i]; body; __i++; } }
                self.current_scope_mut().begin_scope();

                // __arr = right
                self.compile_expression(right)?;
                let arr_slot = self.define_local("__for_of_arr");
                self.emit_u16(Op::local_set, arr_slot);
                self.emit(Op::drop);

                // __i = 0
                self.emit_constant(Value::F64(0.0));
                let i_slot = self.define_local("__for_of_i");
                self.emit_u16(Op::local_set, i_slot);
                self.emit(Op::drop);

                // Loop start: __i < __arr.length
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { start_offset: loop_start, break_patches: vec![], continue_patches: vec![] });

                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                let len_idx = self.add_string_constant("length");
                self.emit_u16(Op::struct_get, len_idx);
                self.emit(Op::dyn_lt);
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);

                // let x = __arr[__i]
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::array_get);
                let var_name = match left {
                    ForInTarget::VarDecl(_, name) => name.clone(),
                    ForInTarget::Identifier(name) => name.clone(),
                };
                let var_slot = self.define_local(&var_name);
                self.emit_u16(Op::local_set, var_slot);
                self.emit(Op::drop);

                // body
                self.compile_statement(body)?;

                // continue target
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
            Statement::ForIn { left, right, body } => {
                // Desugar: for (let k in obj) → { let __keys = Object.keys(obj); for (let k of __keys) { body; } }
                self.current_scope_mut().begin_scope();

                // __keys = host call to get object keys
                self.compile_expression(right)?;
                let keys_idx = self.import("vybe:object", "keys");
                self.emit_host_call(keys_idx, 1);
                let keys_slot = self.define_local("__for_in_keys");
                self.emit_u16(Op::local_set, keys_slot);
                self.emit(Op::drop);

                // __i = 0
                self.emit_constant(Value::F64(0.0));
                let i_slot = self.define_local("__for_in_i");
                self.emit_u16(Op::local_set, i_slot);
                self.emit(Op::drop);

                // Loop: __i < __keys.length
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { start_offset: loop_start, break_patches: vec![], continue_patches: vec![] });

                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, keys_slot);
                let len_idx = self.add_string_constant("length");
                self.emit_u16(Op::struct_get, len_idx);
                self.emit(Op::dyn_lt);
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);

                // let k = __keys[__i]
                self.emit_u16(Op::local_get, keys_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::array_get);
                let var_name = match left {
                    ForInTarget::VarDecl(_, name) => name.clone(),
                    ForInTarget::Identifier(name) => name.clone(),
                };
                let var_slot = self.define_local(&var_name);
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
            Statement::Labeled { body, .. } => { self.compile_statement(body)?; }
            Statement::Empty => {}

            // -- Modules --
            Statement::Import { specifiers, source } => {
                // Host modules (vybe:*): import binds names to host function calls.
                // User modules (./file.js): handled by the module loader before compilation.
                // At this stage, user module imports have already been resolved and
                // their exports injected as globals. So we just bind the names.
                for spec in specifiers {
                    match spec {
                        ImportSpecifier::Named { name, alias } => {
                            let local_name = alias.as_ref().unwrap_or(name);
                            // For host modules: create a namespace object or bind directly
                            if source.starts_with("vybe:") {
                                // Import from host module — we don't need to do anything here.
                                // The compiler resolves host calls by namespace at call sites.
                                // But we store the module name so the namespace resolver knows
                                // that this identifier maps to a host module.
                                // For now: set a global to null as placeholder.
                                self.emit(Op::null);
                                let idx = self.add_string_constant(local_name);
                                self.emit_u16(Op::global_set, idx);
                                self.emit(Op::drop);
                            } else {
                                // User module: the export should already be in globals
                                // (set by the module loader). Just create a local alias if needed.
                                if alias.is_some() {
                                    let src_idx = self.add_string_constant(name);
                                    self.emit_u16(Op::global_get, src_idx);
                                    let dst_idx = self.add_string_constant(local_name);
                                    self.emit_u16(Op::global_set, dst_idx);
                                    self.emit(Op::drop);
                                }
                            }
                        }
                        ImportSpecifier::Namespace(name) => {
                            // import * as name — create namespace object
                            // For host modules, the namespace already works via resolve_namespace_call
                            self.emit(Op::null);
                            let idx = self.add_string_constant(name);
                            self.emit_u16(Op::global_set, idx);
                            self.emit(Op::drop);
                        }
                        ImportSpecifier::Default(name) => {
                            // import defaultName from "mod"
                            // Look for "__default" export in globals
                            let src = format!("{}.__default", source);
                            let src_idx = self.add_string_constant(&src);
                            self.emit_u16(Op::global_get, src_idx);
                            let dst_idx = self.add_string_constant(name);
                            self.emit_u16(Op::global_set, dst_idx);
                            self.emit(Op::drop);
                        }
                    }
                }
            }
            Statement::Export { declaration, specifiers, default } => {
                // Compile the declaration if any
                if let Some(decl) = declaration {
                    self.compile_statement(decl)?;
                    // The declaration already set the global/local.
                    // Nothing extra needed — exports are just globals that other modules can see.
                }
                // export { a, b } — these are already globals, nothing to do
                // The module loader reads the export specifiers when linking.

                // export default expr
                if let Some(expr) = default {
                    self.compile_expression(expr)?;
                    let idx = self.add_string_constant("__default");
                    self.emit_u16(Op::global_set, idx);
                    self.emit(Op::drop);
                }
            }
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
            Expression::Super => {
                // super is used in two contexts:
                // 1. super() — call parent constructor (handled in Call)
                // 2. super.method() — call parent method (handled in Member)
                // As a standalone expression, push null placeholder
                self.emit(Op::null);
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
                    Expression::Member { object, property, .. } => {
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
            Expression::Member { object, property, optional } => {
                // Check for namespace constants: Math.PI, Math.E, Number.MAX_VALUE, etc.
                if let Expression::Identifier(obj_name) = object.as_ref() {
                    if !self.is_known_variable(obj_name) {
                        let module = Self::js_module_alias(obj_name);
                        let (module, name) = Self::js_remap(module, property);
                        // Try as a zero-arg host call (for constants like Math.PI)
                        let idx = self.import(module, name);
                        self.emit_host_call(idx, 0);
                        return Ok(());
                    }
                }
                self.compile_expression(object)?;
                if *optional {
                    // obj?.prop — if obj is null, short-circuit to null
                    self.emit(Op::dup);
                    let skip = self.emit_jump(Op::br_if_null);
                    let idx = self.add_string_constant(property);
                    self.emit_u16(Op::struct_get, idx);
                    let end = self.emit_jump(Op::br);
                    self.patch_jump(skip);
                    // null was already on stack from dup+branch, just leave it
                    self.patch_jump(end);
                } else {
                    let idx = self.add_string_constant(property);
                    self.emit_u16(Op::struct_get, idx);
                }
            }
            Expression::ComputedMember { object, property } => {
                self.compile_expression(object)?;
                self.compile_expression(property)?;
                self.emit(Op::array_get);
            }
            Expression::Call { callee, arguments, .. } => {
                self.compile_call(callee, arguments)?;
            }
            Expression::New { callee, arguments } => {
                // Check for builtin constructors (Map, Set, etc.)
                if let Expression::Identifier(name) = callee.as_ref() {
                    if let Some(host_idx) = self.resolve_builtin_constructor(name) {
                        // Create empty object, call host constructor with it
                        self.emit_u16(Op::struct_new, 0);
                        for arg in arguments { self.compile_expression(arg)?; }
                        self.emit_host_call(host_idx, (arguments.len() + 1) as u8);
                        return Ok(());
                    }
                }
                // User class: push constructor, create empty obj, call
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
                        PropertyDef::Computed { key, value } => {
                            // Key is an expression — evaluate it to get the string key
                            self.compile_expression(key)?;
                            self.compile_expression(value)?;
                            count += 1;
                        }
                        PropertyDef::Spread(_) => {}
                    }
                }
                self.emit_u16(Op::struct_new, count);
            }
            Expression::Function(func) => { self.compile_function(func)?; }
            Expression::ArrowFunction { params, body, is_async } => {
                let func = match body {
                    ArrowBody::Block(stmts) => FunctionDecl { name: None, params: params.clone(), body: stmts.clone(), is_async: *is_async },
                    ArrowBody::Expression(expr) => FunctionDecl { name: None, params: params.clone(), body: vec![Statement::Return(Some(*expr.clone()))], is_async: *is_async },
                };
                self.compile_function(&func)?;
            }
            Expression::Await(inner) => {
                // For synchronous promises: await just evaluates the expression.
                // Emit the expression, then the await opcode.
                // The VM handles Promise checking and fiber suspension.
                self.compile_expression(inner)?;
                self.emit(Op::r#await);
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
        if let Expression::Member { object, property, .. } = callee {
            // obj.method() pattern
            if let Expression::Identifier(obj_name) = object.as_ref() {
                if !self.is_known_variable(obj_name) {
                    // Direct WASM opcodes for Math functions (no host call overhead)
                    if obj_name == "Math" {
                        if let Some(()) = self.try_math_intrinsic(property, arguments)? {
                            return Ok(());
                        }
                    }
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

            // Array higher-order methods — desugar to forEach + push pattern
            // These methods need VM callbacks (calling JS functions from loops).
            // We desugar them in the compiler to bytecode loops that use `call`.
            if matches!(property.as_str(), "map" | "filter" | "forEach" | "find" | "reduce" | "sort") {
                if self.compile_array_callback_method(object, property, arguments)? {
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

            // Method call on object. Could be a user class method or a builtin
            // collection method. Try runtime dispatch first — it returns Null for
            // non-builtins, in which case we fall through to struct_get + call.
            //
            // For builtins (Map, Set): callMethod handles it directly.
            // For user objects: regular struct_get + call with this binding.
            //
            // We use callMethod(obj, "methodName", ...args) for all method calls.
            // The runtime checks __type and dispatches accordingly.
            // If it returns Null and the method isn't a builtin, we do regular call.
            self.compile_expression(object)?;
            self.emit_constant(Value::String(Rc::from(property.as_str())));
            for arg in arguments { self.compile_expression(arg)?; }
            let cm_idx = self.import("vybe:runtime", "callMethod");
            self.emit_host_call(cm_idx, (arguments.len() + 2) as u8);

            // Check if callMethod returned Null (not a builtin) — then do regular call
            self.emit(Op::dup);
            self.emit(Op::ref_is_null);
            let done = self.emit_jump(Op::br_if_false); // not null = result, skip
            self.emit(Op::drop); // drop the null
            // Regular method call: obj.method(args)
            self.compile_expression(object)?;
            let prop_idx = self.add_string_constant(property);
            self.emit_u16(Op::struct_get, prop_idx);
            // If calling on a class name (static call), don't pass this
            let is_static = if let Expression::Identifier(obj_name) = object.as_ref() {
                self.defined_classes.contains(obj_name)
            } else { false };
            if is_static {
                for arg in arguments { self.compile_expression(arg)?; }
                self.emit_u8(Op::call, arguments.len() as u8);
            } else {
                self.compile_expression(object)?; // this
                for arg in arguments { self.compile_expression(arg)?; }
                self.emit_u8(Op::call, (arguments.len() + 1) as u8);
            }
            self.patch_jump(done);
            return Ok(());
        }

        // Bare function call: func(args)
        // Direct WASM opcodes for common JS globals
        if let Expression::Identifier(name) = callee {
            if let Some(()) = self.try_bare_intrinsic(name, arguments)? {
                return Ok(());
            }
            if let Some(idx) = self.resolve_bare_import(name) {
                for arg in arguments { self.compile_expression(arg)?; }
                self.emit_host_call(idx, arguments.len() as u8);
                return Ok(());
            }
        }

        // super() — call parent constructor with this + args
        if matches!(callee, Expression::Super) {
            if self.in_method {
                // Get this.__super (set during class compilation)
                self.emit_u16(Op::local_get, 1); // this
                let super_idx = self.add_string_constant("__super");
                self.emit_u16(Op::struct_get, super_idx);
                // Push this as first arg, then user args
                self.emit_u16(Op::local_get, 1); // this
                for arg in arguments { self.compile_expression(arg)?; }
                self.emit_u8(Op::call, (arguments.len() + 1) as u8);
                return Ok(());
            }
            // Outside method — just push null
            self.emit(Op::null);
            return Ok(());
        }

        // Special builtins that need compiler-level handling
        if let Expression::Identifier(name) = callee {
            if name == "setTimeout" && arguments.len() >= 1 {
                // setTimeout(fn, ms) → set_timer opcode (VM handles event loop)
                self.compile_expression(&arguments[0])?; // callback
                if arguments.len() >= 2 {
                    self.compile_expression(&arguments[1])?; // ms
                } else {
                    self.emit_constant(Value::F64(0.0));
                }
                self.emit(Op::set_timer);
                return Ok(());
            }
            if name == "setInterval" && arguments.len() >= 1 {
                // setInterval — same as setTimeout for now (TODO: repeating)
                self.compile_expression(&arguments[0])?;
                if arguments.len() >= 2 {
                    self.compile_expression(&arguments[1])?;
                } else {
                    self.emit_constant(Value::F64(0.0));
                }
                self.emit(Op::set_timer);
                return Ok(());
            }
        }

        // Regular function call — check for spread
        self.compile_expression(callee)?;
        let argc = self.compile_args_with_spread(arguments)?;
        self.emit_u8(Op::call, argc);
        Ok(())
    }

    /// Compile function arguments, handling spread.
    /// Returns the argument count to use in the call opcode.
    fn compile_args_with_spread(&mut self, arguments: &[Expression]) -> Result<u8, String> {
        let has_spread = arguments.iter().any(|a| matches!(a, Expression::Spread(_)));

        if !has_spread {
            // No spread — compile normally
            for arg in arguments { self.compile_expression(arg)?; }
            return Ok(arguments.len() as u8);
        }

        // Case 1: single spread of array literal: f(...[1,2,3]) → f(1,2,3)
        if arguments.len() == 1 {
            if let Expression::Spread(inner) = &arguments[0] {
                if let Expression::Array(elems) = inner.as_ref() {
                    for elem in elems { self.compile_expression(elem)?; }
                    return Ok(elems.len() as u8);
                }
            }
        }

        // Case 2: expand spread inline — count non-spread args + expand known arrays
        let mut total = 0u8;
        for arg in arguments {
            match arg {
                Expression::Spread(inner) => {
                    if let Expression::Array(elems) = inner.as_ref() {
                        // Known array literal — inline elements
                        for elem in elems { self.compile_expression(elem)?; }
                        total += elems.len() as u8;
                    } else {
                        // Dynamic spread: pass the array as a single arg
                        // The callee receives it as one value — not ideal but functional
                        self.compile_expression(inner)?;
                        total += 1;
                    }
                }
                _ => {
                    self.compile_expression(arg)?;
                    total += 1;
                }
            }
        }
        Ok(total)
    }

    /// Bind the value on top of stack to a pattern (var/let/const declaration).
    fn compile_binding(&mut self, pattern: &BindingPattern, kind: VarKind) -> Result<(), String> {
        match pattern {
            BindingPattern::Identifier(name) => {
                match self.resolve_variable(name) {
                    VarResolution::Local(_) => {
                        let slot = self.current_scope().resolve_local(name).unwrap();
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    }
                    _ => {
                        if self.scopes.len() == 1 && self.current_scope().depth == 0 && kind == VarKind::Var {
                            self.emit_global_set(name);
                            self.emit(Op::drop);
                        } else {
                            let slot = self.define_local(name);
                            self.emit_u16(Op::local_set, slot);
                            self.emit(Op::drop);
                        }
                    }
                }
            }
            BindingPattern::Object(props) => {
                // Stack has the object value. For each prop, dup + get property + bind.
                for prop in props {
                    self.emit(Op::dup); // keep object on stack
                    let key_idx = self.add_string_constant(&prop.key);
                    self.emit_u16(Op::struct_get, key_idx);
                    // If there's a default and the value is null, use default
                    if let Some(ref default_expr) = prop.default {
                        self.emit(Op::dup);
                        self.emit(Op::ref_is_null);
                        let skip = self.emit_jump(Op::br_if_false);
                        self.emit(Op::drop); // drop null
                        self.compile_expression(default_expr)?;
                        self.patch_jump(skip);
                    }
                    // Bind to the target name
                    let target_name = match &prop.value {
                        Some(BindingPattern::Identifier(n)) => n.clone(),
                        None => prop.key.clone(), // shorthand
                        _ => {
                            // Nested destructuring — recurse
                            if let Some(ref nested) = prop.value {
                                self.compile_binding(nested, kind)?;
                                continue;
                            }
                            prop.key.clone()
                        }
                    };
                    let slot = self.define_local(&target_name);
                    self.emit_u16(Op::local_set, slot);
                    self.emit(Op::drop);
                }
                self.emit(Op::drop); // drop the object
            }
            BindingPattern::Array(elems) => {
                // Stack has the array value. For each elem, dup + get index + bind.
                for (i, elem) in elems.iter().enumerate() {
                    match elem {
                        ArrayPatternElem::Pattern(pat, default) => {
                            self.emit(Op::dup);
                            self.emit_constant(Value::F64(i as f64));
                            self.emit(Op::array_get);
                            if let Some(default_expr) = default {
                                self.emit(Op::dup);
                                self.emit(Op::ref_is_null);
                                let skip = self.emit_jump(Op::br_if_false);
                                self.emit(Op::drop);
                                self.compile_expression(default_expr)?;
                                self.patch_jump(skip);
                            }
                            self.compile_binding(pat, kind)?;
                        }
                        ArrayPatternElem::Rest(name) => {
                            // ...rest — slice remaining elements
                            self.emit(Op::dup);
                            self.emit_constant(Value::F64(i as f64));
                            let slice_idx = self.import("vybe:array", "slice");
                            self.emit_host_call(slice_idx, 2);
                            let slot = self.define_local(name);
                            self.emit_u16(Op::local_set, slot);
                            self.emit(Op::drop);
                        }
                        ArrayPatternElem::Hole => {
                            // Skip this index
                        }
                    }
                }
                self.emit(Op::drop); // drop the array
            }
        }
        Ok(())
    }

    /// Compile arr.map(fn), arr.filter(fn), arr.forEach(fn), arr.find(fn),
    /// arr.reduce(fn, init), arr.sort(fn) as inline bytecode loops.
    /// Uses call (same as WASM call_indirect) for the callback.
    /// Returns true if the method was handled.
    fn compile_array_callback_method(
        &mut self,
        object: &Expression,
        method: &str,
        arguments: &[Expression],
    ) -> Result<bool, String> {
        match method {
            "map" => {
                // arr.map(fn) → { let __r=[]; for i in arr { __r.push(fn(arr[i],i,arr)); } __r }
                self.current_scope_mut().begin_scope();
                // __arr = object
                self.compile_expression(object)?;
                let arr_slot = self.define_local("__cb_arr");
                self.emit_u16(Op::local_set, arr_slot);
                self.emit(Op::drop);
                // __fn = callback
                if arguments.is_empty() { return Err("map requires a callback".into()); }
                self.compile_expression(&arguments[0])?;
                let fn_slot = self.define_local("__cb_fn");
                self.emit_u16(Op::local_set, fn_slot);
                self.emit(Op::drop);
                // __result = []
                self.emit_u16(Op::array_new, 0);
                let result_slot = self.define_local("__cb_result");
                self.emit_u16(Op::local_set, result_slot);
                self.emit(Op::drop);
                // __i = 0
                self.emit_constant(Value::F64(0.0));
                let i_slot = self.define_local("__cb_i");
                self.emit_u16(Op::local_set, i_slot);
                self.emit(Op::drop);
                // loop: while __i < __arr.length
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                let len_idx = self.add_string_constant("length");
                self.emit_u16(Op::struct_get, len_idx);
                self.emit(Op::dyn_lt);
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                // val = fn(arr[i], i, arr) — use call (WASM call_indirect)
                self.emit_u16(Op::local_get, fn_slot);   // push callback
                self.emit_u16(Op::local_get, arr_slot);   // arr
                self.emit_u16(Op::local_get, i_slot);     // i
                self.emit(Op::array_get);                  // arr[i]
                self.emit_u16(Op::local_get, i_slot);     // i
                self.emit_u16(Op::local_get, arr_slot);   // arr
                self.emit_u8(Op::call, 3);                // fn(arr[i], i, arr) → val on stack
                // Store val, then push(result, val)
                let val_slot = self.define_local("__cb_val");
                self.emit_u16(Op::local_set, val_slot);    // val_slot = val (TOS)
                self.emit(Op::drop);                        // pop val from stack
                let push_idx = self.import("vybe:array", "push");
                self.emit_u16(Op::local_get, result_slot); // push result_arr
                self.emit_u16(Op::local_get, val_slot);    // push val
                self.emit_host_call(push_idx, 2);          // push(result_arr, val)
                self.emit(Op::drop);                        // discard push return
                // i++
                self.emit_u16(Op::local_get, i_slot);
                self.emit_constant(Value::F64(1.0));
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot);
                self.emit(Op::drop);
                self.emit_loop(loop_start);
                self.patch_jump(exit);
                // push result
                self.emit_u16(Op::local_get, result_slot);
                self.current_scope_mut().end_scope();
                Ok(true)
            }
            "filter" => {
                self.current_scope_mut().begin_scope();
                self.compile_expression(object)?;
                let arr_slot = self.define_local("__cb_arr");
                self.emit_u16(Op::local_set, arr_slot); self.emit(Op::drop);
                self.compile_expression(&arguments[0])?;
                let fn_slot = self.define_local("__cb_fn");
                self.emit_u16(Op::local_set, fn_slot); self.emit(Op::drop);
                self.emit_u16(Op::array_new, 0);
                let result_slot = self.define_local("__cb_result");
                self.emit_u16(Op::local_set, result_slot); self.emit(Op::drop);
                self.emit_constant(Value::F64(0.0));
                let i_slot = self.define_local("__cb_i");
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                let len_idx = self.add_string_constant("length");
                self.emit_u16(Op::struct_get, len_idx);
                self.emit(Op::dyn_lt); self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                // elem = arr[i]
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::array_get);
                let elem_slot = self.define_local("__cb_elem");
                self.emit_u16(Op::local_set, elem_slot); self.emit(Op::drop);
                // if fn(elem, i, arr) → push elem
                self.emit_u16(Op::local_get, fn_slot);
                self.emit_u16(Op::local_get, elem_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u8(Op::call, 3);
                self.emit(Op::dyn_to_bool);
                let skip = self.emit_jump(Op::br_if_false);
                let push_idx = self.import("vybe:array", "push");
                self.emit_u16(Op::local_get, result_slot);
                self.emit_u16(Op::local_get, elem_slot);
                self.emit_host_call(push_idx, 2);
                self.emit(Op::drop);
                self.patch_jump(skip);
                // i++
                self.emit_u16(Op::local_get, i_slot);
                self.emit_constant(Value::F64(1.0));
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                self.emit_loop(loop_start);
                self.patch_jump(exit);
                self.emit_u16(Op::local_get, result_slot);
                self.current_scope_mut().end_scope();
                Ok(true)
            }
            "forEach" => {
                self.current_scope_mut().begin_scope();
                self.compile_expression(object)?;
                let arr_slot = self.define_local("__cb_arr");
                self.emit_u16(Op::local_set, arr_slot); self.emit(Op::drop);
                self.compile_expression(&arguments[0])?;
                let fn_slot = self.define_local("__cb_fn");
                self.emit_u16(Op::local_set, fn_slot); self.emit(Op::drop);
                self.emit_constant(Value::F64(0.0));
                let i_slot = self.define_local("__cb_i");
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                let len_idx = self.add_string_constant("length");
                self.emit_u16(Op::struct_get, len_idx);
                self.emit(Op::dyn_lt); self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                self.emit_u16(Op::local_get, fn_slot);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::array_get);
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u8(Op::call, 3);
                self.emit(Op::drop); // discard return
                // i++
                self.emit_u16(Op::local_get, i_slot);
                self.emit_constant(Value::F64(1.0));
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                self.emit_loop(loop_start);
                self.patch_jump(exit);
                self.emit(Op::null); // forEach returns undefined
                self.current_scope_mut().end_scope();
                Ok(true)
            }
            "find" => {
                self.current_scope_mut().begin_scope();
                self.compile_expression(object)?;
                let arr_slot = self.define_local("__cb_arr");
                self.emit_u16(Op::local_set, arr_slot); self.emit(Op::drop);
                self.compile_expression(&arguments[0])?;
                let fn_slot = self.define_local("__cb_fn");
                self.emit_u16(Op::local_set, fn_slot); self.emit(Op::drop);
                self.emit_constant(Value::F64(0.0));
                let i_slot = self.define_local("__cb_i");
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                self.emit(Op::null);
                let result_slot = self.define_local("__cb_result");
                self.emit_u16(Op::local_set, result_slot); self.emit(Op::drop);
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                let len_idx = self.add_string_constant("length");
                self.emit_u16(Op::struct_get, len_idx);
                self.emit(Op::dyn_lt); self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                // elem = arr[i]
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::array_get);
                let elem_slot = self.define_local("__cb_elem");
                self.emit_u16(Op::local_set, elem_slot); self.emit(Op::drop);
                self.emit_u16(Op::local_get, fn_slot);
                self.emit_u16(Op::local_get, elem_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u8(Op::call, 3);
                self.emit(Op::dyn_to_bool);
                let skip = self.emit_jump(Op::br_if_false);
                // found — store and break
                self.emit_u16(Op::local_get, elem_slot);
                self.emit_u16(Op::local_set, result_slot); self.emit(Op::drop);
                let done = self.emit_jump(Op::br);
                self.patch_jump(skip);
                // i++
                self.emit_u16(Op::local_get, i_slot);
                self.emit_constant(Value::F64(1.0));
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                self.emit_loop(loop_start);
                self.patch_jump(exit);
                self.patch_jump(done);
                self.emit_u16(Op::local_get, result_slot);
                self.current_scope_mut().end_scope();
                Ok(true)
            }
            "reduce" => {
                self.current_scope_mut().begin_scope();
                self.compile_expression(object)?;
                let arr_slot = self.define_local("__cb_arr");
                self.emit_u16(Op::local_set, arr_slot); self.emit(Op::drop);
                self.compile_expression(&arguments[0])?;
                let fn_slot = self.define_local("__cb_fn");
                self.emit_u16(Op::local_set, fn_slot); self.emit(Op::drop);
                // init value
                if arguments.len() > 1 {
                    self.compile_expression(&arguments[1])?;
                } else {
                    self.emit(Op::null);
                }
                let acc_slot = self.define_local("__cb_acc");
                self.emit_u16(Op::local_set, acc_slot); self.emit(Op::drop);
                self.emit_constant(Value::F64(0.0));
                let i_slot = self.define_local("__cb_i");
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                let len_idx = self.add_string_constant("length");
                self.emit_u16(Op::struct_get, len_idx);
                self.emit(Op::dyn_lt); self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                // acc = fn(acc, arr[i], i, arr)
                self.emit_u16(Op::local_get, fn_slot);
                self.emit_u16(Op::local_get, acc_slot);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::array_get);
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u8(Op::call, 4);
                self.emit_u16(Op::local_set, acc_slot); self.emit(Op::drop);
                // i++
                self.emit_u16(Op::local_get, i_slot);
                self.emit_constant(Value::F64(1.0));
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                self.emit_loop(loop_start);
                self.patch_jump(exit);
                self.emit_u16(Op::local_get, acc_slot);
                self.current_scope_mut().end_scope();
                Ok(true)
            }
            "sort" => {
                // Bubble sort with optional comparator
                self.current_scope_mut().begin_scope();
                self.compile_expression(object)?;
                let arr_slot = self.define_local("__cb_arr");
                self.emit_u16(Op::local_set, arr_slot); self.emit(Op::drop);
                // comparator (optional)
                let has_fn = !arguments.is_empty();
                let fn_slot = if has_fn {
                    self.compile_expression(&arguments[0])?;
                    let s = self.define_local("__cb_fn");
                    self.emit_u16(Op::local_set, s); self.emit(Op::drop);
                    s
                } else { 0 };
                // len
                self.emit_u16(Op::local_get, arr_slot);
                let len_idx = self.add_string_constant("length");
                self.emit_u16(Op::struct_get, len_idx);
                let len_slot = self.define_local("__cb_len");
                self.emit_u16(Op::local_set, len_slot); self.emit(Op::drop);
                // outer loop: i
                self.emit_constant(Value::F64(0.0));
                let i_slot = self.define_local("__cb_i");
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                let outer_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, len_slot);
                self.emit(Op::dyn_lt); self.emit(Op::dyn_to_bool);
                let outer_exit = self.emit_jump(Op::br_if_false);
                // inner loop: j
                self.emit_constant(Value::F64(0.0));
                let j_slot = self.define_local("__cb_j");
                self.emit_u16(Op::local_set, j_slot); self.emit(Op::drop);
                let inner_start = self.current_offset();
                // j < len - i - 1
                self.emit_u16(Op::local_get, j_slot);
                self.emit_u16(Op::local_get, len_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::dyn_add); // len + i... wait, need len - i - 1
                // Hmm, dyn_add won't subtract. Let me use f64_sub.
                // len - i - 1
                self.emit(Op::drop); // drop the bad add
                self.emit_u16(Op::local_get, j_slot);
                self.emit_u16(Op::local_get, len_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::f64_sub);
                self.emit_constant(Value::F64(1.0));
                self.emit(Op::f64_sub);
                self.emit(Op::dyn_lt); self.emit(Op::dyn_to_bool);
                let inner_exit = self.emit_jump(Op::br_if_false);
                // compare arr[j] vs arr[j+1]
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, j_slot);
                self.emit(Op::array_get); // arr[j]
                let a_slot = self.define_local("__sort_a");
                self.emit_u16(Op::local_set, a_slot); self.emit(Op::drop);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, j_slot);
                self.emit_constant(Value::F64(1.0));
                self.emit(Op::dyn_add);
                self.emit(Op::array_get); // arr[j+1]
                let b_slot = self.define_local("__sort_b");
                self.emit_u16(Op::local_set, b_slot); self.emit(Op::drop);
                // cmp
                if has_fn {
                    self.emit_u16(Op::local_get, fn_slot);
                    self.emit_u16(Op::local_get, a_slot);
                    self.emit_u16(Op::local_get, b_slot);
                    self.emit_u8(Op::call, 2);
                    // comparator returns number: >0 means swap
                    self.emit_constant(Value::F64(0.0));
                    self.emit(Op::dyn_gt);
                } else {
                    // default: a > b
                    self.emit_u16(Op::local_get, a_slot);
                    self.emit_u16(Op::local_get, b_slot);
                    self.emit(Op::dyn_gt);
                }
                self.emit(Op::dyn_to_bool);
                let no_swap = self.emit_jump(Op::br_if_false);
                // swap: arr[j] = b, arr[j+1] = a
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, j_slot);
                self.emit_constant(Value::String(Rc::from(""))); // key placeholder — we need to use array_set properly
                // Actually array_set wants [obj, key, val]. Let me use the host push approach.
                // Simpler: use computed member assignment
                // arr[j] = b
                self.emit(Op::drop); // drop the "" placeholder
                // I'll just use struct_set/array_set directly
                // arr obj is on stack... actually this is getting messy. Let me use host functions.
                let set_idx = self.import("vybe:array", "setAt");
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, j_slot);
                self.emit_u16(Op::local_get, b_slot);
                self.emit_host_call(set_idx, 3);
                self.emit(Op::drop);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, j_slot);
                self.emit_constant(Value::F64(1.0));
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_get, a_slot);
                self.emit_host_call(set_idx, 3);
                self.emit(Op::drop);
                self.patch_jump(no_swap);
                // j++
                self.emit_u16(Op::local_get, j_slot);
                self.emit_constant(Value::F64(1.0));
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, j_slot); self.emit(Op::drop);
                self.emit_loop(inner_start);
                self.patch_jump(inner_exit);
                // i++
                self.emit_u16(Op::local_get, i_slot);
                self.emit_constant(Value::F64(1.0));
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                self.emit_loop(outer_start);
                self.patch_jump(outer_exit);
                self.emit_u16(Op::local_get, arr_slot);
                self.current_scope_mut().end_scope();
                Ok(true)
            }
            "some" | "every" | "findIndex" | "flatMap" => {
                // TODO: implement these similarly
                Ok(false) // fall through to regular method call
            }
            _ => Ok(false),
        }
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
            Expression::Member { object, property, .. } => {
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
        for param in &func.params { scope.define_local(&param.name); }

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        // Emit default parameter initialization
        for param in &func.params {
            if let Some(ref default_expr) = param.default {
                // if (param is null) { param = default; }
                let slot = self.current_scope().resolve_local(&param.name).unwrap();
                self.emit_u16(Op::local_get, slot);
                self.emit(Op::ref_is_null);
                let skip = self.emit_jump(Op::br_if_false); // if NOT null, skip
                self.compile_expression(default_expr)?;
                self.emit_u16(Op::local_set, slot);
                self.emit(Op::drop);
                self.patch_jump(skip);
            }
        }

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
        for param in &func.params { scope.define_local(&param.name); }

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
        let mut instance_methods: Vec<(String, FunctionDecl, MethodKind)> = Vec::new();
        let mut static_methods: Vec<(String, FunctionDecl)> = Vec::new();
        let mut static_props: Vec<(String, Option<Expression>)> = Vec::new();

        for member in &class.body {
            match member {
                ClassMember::Method { key, value, kind, is_static } => {
                    if *kind == MethodKind::Constructor {
                        constructor = Some(value.clone());
                    } else if *is_static {
                        static_methods.push((key.clone(), value.clone()));
                    } else {
                        instance_methods.push((key.clone(), value.clone(), kind.clone()));
                    }
                }
                ClassMember::Property { key, value, is_static } => {
                    if *is_static {
                        static_props.push((key.clone(), value.clone()));
                    }
                    // Instance fields collected below
                }
            }
        }

        // If extends, compile parent class reference and store for super
        if let Some(ref super_expr) = class.super_class {
            self.compile_expression(super_expr)?;
            let parent_slot = self.define_local(&format!("__parent_{}", name));
            self.emit_u16(Op::local_set, parent_slot);
            self.emit(Op::drop);
        }

        // Separate regular methods from getters/setters
        let mut regular_methods: Vec<(String, FunctionDecl)> = Vec::new();
        let mut getters: Vec<(String, FunctionDecl)> = Vec::new();
        let mut setters: Vec<(String, FunctionDecl)> = Vec::new();
        for (key, value, kind) in &instance_methods {
            match kind {
                MethodKind::Get => getters.push((key.clone(), value.clone())),
                MethodKind::Set => setters.push((key.clone(), value.clone())),
                _ => regular_methods.push((key.clone(), value.clone())),
            }
        }

        // Inject instance field initializers into constructor body
        let mut field_init_stmts: Vec<Statement> = Vec::new();
        for member in &class.body {
            if let ClassMember::Property { key, value: Some(val_expr), is_static: false } = member {
                field_init_stmts.push(Statement::Expression(Expression::Assignment {
                    op: AssignOp::Assign,
                    left: Box::new(Expression::Member {
                        object: Box::new(Expression::This),
                        property: key.clone(),
                        optional: false,
                    }),
                    right: Box::new(val_expr.clone()),
                }));
            }
        }

        let ctor_params = constructor.as_ref().map(|c| c.params.clone()).unwrap_or_default();
        let mut ctor_body = constructor.map(|c| c.body).unwrap_or_default();
        // Prepend field initializers before the constructor body
        for (i, stmt) in field_init_stmts.into_iter().enumerate() {
            ctor_body.insert(i, stmt);
        }
        let ctor = FunctionDecl { name: Some(name.into()), params: ctor_params, body: ctor_body, is_async: false };

        self.compile_class_constructor_full(&ctor, &regular_methods, &getters, &setters, &class.super_class)?;
        // Constructor closure is now on the stack.
        // Attach static methods and properties to the constructor object.
        for (method_name, method_fn) in &static_methods {
            self.emit(Op::dup); // keep constructor on stack
            self.compile_function(method_fn)?;
            let prop_idx = self.add_string_constant(method_name);
            self.emit_u16(Op::struct_set, prop_idx);
            self.emit(Op::drop);
        }
        for (prop_name, prop_value) in &static_props {
            self.emit(Op::dup);
            if let Some(expr) = prop_value {
                self.compile_expression(expr)?;
            } else {
                self.emit(Op::null);
            }
            let prop_idx = self.add_string_constant(prop_name);
            self.emit_u16(Op::struct_set, prop_idx);
            self.emit(Op::drop);
        }
        Ok(())
    }

    fn compile_class_constructor_full(
        &mut self,
        ctor: &FunctionDecl,
        methods: &[(String, FunctionDecl)],
        getters: &[(String, FunctionDecl)],
        setters: &[(String, FunctionDecl)],
        super_class: &Option<Box<Expression>>,
    ) -> Result<(), String> {
        let saved_method = self.in_method;
        self.in_method = true;

        let name = ctor.name.clone().unwrap_or_else(|| "<class>".into());
        let mut chunk = Chunk::new(&name);
        chunk.arity = (ctor.params.len() + 1) as u8;
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        scope.define_local("this");
        for param in &ctor.params { scope.define_local(&param.name); }

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        // If extends: set this.__super = parent constructor
        if let Some(super_expr) = super_class {
            self.emit_u16(Op::local_get, 1); // this
            self.compile_expression(super_expr)?; // parent class (constructor fn)
            let super_idx = self.add_string_constant("__super");
            self.emit_u16(Op::struct_set, super_idx);
            self.emit(Op::drop);
        }

        // Compile constructor body (may contain super() calls)
        for stmt in &ctor.body { self.compile_statement(stmt)?; }

        // Attach regular methods to this
        for (method_name, method_fn) in methods {
            self.emit_u16(Op::local_get, 1); // this
            self.compile_method(method_fn)?;
            let prop_idx = self.add_string_constant(method_name);
            self.emit_u16(Op::struct_set, prop_idx);
            self.emit(Op::drop);
        }

        // Attach getters as __get_name methods
        for (getter_name, getter_fn) in getters {
            self.emit_u16(Op::local_get, 1);
            self.compile_method(getter_fn)?;
            let prop_name = format!("__get_{}", getter_name);
            let prop_idx = self.add_string_constant(&prop_name);
            self.emit_u16(Op::struct_set, prop_idx);
            self.emit(Op::drop);
        }

        // Attach setters as __set_name methods
        for (setter_name, setter_fn) in setters {
            self.emit_u16(Op::local_get, 1);
            self.compile_method(setter_fn)?;
            let prop_name = format!("__set_{}", setter_name);
            let prop_idx = self.add_string_constant(&prop_name);
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

    /// Emit direct WASM opcodes for Math.* functions instead of host calls.
    fn try_math_intrinsic(&mut self, method: &str, args: &[Expression]) -> Result<Option<()>, String> {
        // Single-argument Math functions
        if args.len() == 1 {
            let op = match method {
                "abs"   => Some(Op::f64_abs),
                "floor" => Some(Op::f64_floor),
                "ceil"  => Some(Op::f64_ceil),
                "sqrt"  => Some(Op::f64_sqrt),
                "trunc" => Some(Op::f64_trunc),
                "round" => Some(Op::f64_nearest),
                "sign"  => None, // multi-opcode, fall through to host
                _ => None,
            };
            if let Some(op) = op {
                self.compile_expression(&args[0])?;
                self.emit(op);
                return Ok(Some(()));
            }
        }
        // Two-argument Math functions
        if args.len() == 2 {
            let op = match method {
                "min" => Some(Op::f64_min),
                "max" => Some(Op::f64_max),
                _ => None,
            };
            if let Some(op) = op {
                self.compile_expression(&args[0])?;
                self.compile_expression(&args[1])?;
                self.emit(op);
                return Ok(Some(()));
            }
        }
        Ok(None)
    }

    /// Emit direct WASM opcodes for bare JS globals.
    /// Note: parseInt/parseFloat must stay as host calls because they parse STRINGS.
    fn try_bare_intrinsic(&mut self, _name: &str, _args: &[Expression]) -> Result<Option<()>, String> {
        Ok(None)
    }
}

enum VarResolution { Local(u16), Upvalue(u8), Global }
