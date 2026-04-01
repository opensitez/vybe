use std::rc::Rc;

use vybe_bytecode::{Chunk, Value, Op};
use vybe_bytecode::chunk::TypeEntry;
use vybe_compiler_common::classes as common_classes;
use vybe_compiler_common::expressions as common_expr;
use vybe_compiler_common::functions as common_fn;
use vybe_parser_js::ast::*;

use crate::scope::Scope;

struct LoopContext {
    _start_offset: usize,
    break_patches: Vec<usize>,
    continue_patches: Vec<usize>,
    label: Option<String>,
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
    /// Track variable names known to hold class instances (from `let x = new ClassName()`).
    /// Used to avoid short-circuiting method calls like push/pop to array intrinsics.
    class_instances: std::collections::HashSet<String>,
    /// ESM Integration: .wasm files referenced by import statements.
    /// The CLI/runtime should load these modules before execution.
    pub wasm_imports: Vec<String>,
    /// Label for the next loop (set by Labeled statement)
    pending_label: Option<String>,
    /// WASM GC type table entries — one per class declaration.
    type_entries: Vec<TypeEntry>,
    /// Class name → index into type_entries (for set_type_id at construction sites).
    class_type_ids: std::collections::HashMap<String, usize>,
}

impl Compiler {
    pub fn new() -> Self {
        Compiler {
            chunks: vec![Chunk::new("<script>")],
            scopes: vec![Scope::new()],
            current_chunk_idx: 0,
            loop_stack: Vec::new(),
            line: 1,
            pending_label: None,
            in_method: false,
            defined_globals: std::collections::HashSet::new(),
            defined_classes: std::collections::HashSet::new(),
            class_instances: std::collections::HashSet::new(),
            wasm_imports: Vec::new(),
            type_entries: Vec::new(),
            class_type_ids: std::collections::HashMap::new(),
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
        self.chunks[0].types = self.type_entries;
        vybe_compiler_common::bundle::finalize_with_stdlib(&mut self.chunks);
        Ok(self.chunks)
    }

    /// Compile and return both chunks and any .wasm ESM imports found.
    pub fn compile_with_imports(mut self, program: &Program) -> Result<(Vec<Chunk>, Vec<String>), String> {
        for stmt in &program.body {
            self.compile_statement(stmt)?;
        }
        self.emit(Op::null);
        self.emit(Op::halt);
        let local_count = self.current_scope().next_slot;
        self.chunks[0].local_count = local_count;
        self.chunks[0].types = self.type_entries;
        vybe_compiler_common::bundle::finalize_with_stdlib(&mut self.chunks);
        Ok((self.chunks, self.wasm_imports))
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
                    // Track if this variable is a class instance (for method dispatch)
                    let is_class_new = if let Some(init) = &decl.init {
                        matches!(init, Expression::New { callee, .. }
                            if matches!(callee.as_ref(), Expression::Identifier(name) if self.defined_classes.contains(name)))
                    } else { false };

                    if let Some(init) = &decl.init {
                        self.compile_expression(init)?;
                    } else {
                        self.emit(Op::null);
                    }
                    // Value is on stack — bind it to the pattern
                    if is_class_new {
                        if let BindingPattern::Identifier(name) = &decl.pattern {
                            self.class_instances.insert(name.clone());
                        }
                    }
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
                self.loop_stack.push(LoopContext { _start_offset: start, break_patches: vec![], continue_patches: vec![], label: self.pending_label.take() });
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
                self.loop_stack.push(LoopContext { _start_offset: start, break_patches: vec![], continue_patches: vec![], label: self.pending_label.take() });
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
                self.loop_stack.push(LoopContext { _start_offset: start, break_patches: vec![], continue_patches: vec![], label: self.pending_label.take() });
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
            Statement::Break(label) => {
                let p = self.emit_jump(Op::br);
                // Find the loop context matching the label (or innermost if no label)
                if let Some(lbl) = label {
                    for ctx in self.loop_stack.iter_mut().rev() {
                        if ctx.label.as_deref() == Some(lbl) {
                            ctx.break_patches.push(p);
                            break;
                        }
                    }
                } else if let Some(ctx) = self.loop_stack.last_mut() {
                    ctx.break_patches.push(p);
                }
            }
            Statement::Continue(label) => {
                let p = self.emit_jump(Op::br);
                if let Some(lbl) = label {
                    for ctx in self.loop_stack.iter_mut().rev() {
                        if ctx.label.as_deref() == Some(lbl) {
                            ctx.continue_patches.push(p);
                            break;
                        }
                    }
                } else if let Some(ctx) = self.loop_stack.last_mut() {
                    ctx.continue_patches.push(p);
                }
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
                // JS switch with fallthrough: once a case matches, execute all subsequent
                // case bodies until a break is hit.
                self.compile_expression(discriminant)?;
                self.loop_stack.push(LoopContext { _start_offset: 0, break_patches: vec![], continue_patches: vec![], label: self.pending_label.take() });

                // Phase 1: emit all case tests, jumping to their body positions
                let mut test_jumps: Vec<(usize, usize)> = Vec::new(); // (case_idx, jump_to_body)
                let mut default_idx: Option<usize> = None;
                let mut next_test_patch: Option<usize> = None;
                for (i, case) in cases.iter().enumerate() {
                    if let Some(p) = next_test_patch.take() { self.patch_jump(p); }
                    if let Some(test) = &case.test {
                        self.emit(Op::dup);
                        self.compile_expression(test)?;
                        self.emit(Op::eq);
                        let body_jump = self.emit_jump(Op::br_if_true);
                        test_jumps.push((i, body_jump));
                        next_test_patch = None;
                    } else {
                        default_idx = Some(i);
                    }
                }
                // If no case matched, jump to default or end
                let to_default_or_end = self.emit_jump(Op::br);

                // Phase 2: emit all case bodies sequentially (fallthrough order)
                let mut body_offsets: Vec<usize> = Vec::new();
                for (_i, case) in cases.iter().enumerate() {
                    body_offsets.push(self.current_offset());
                    for s in &case.consequent { self.compile_statement(s)?; }
                    // No implicit break — falls through to next case body
                }
                let _end_offset = self.current_offset();

                // Phase 3: patch test jumps to body positions
                for (case_idx, jump) in &test_jumps {
                    if *case_idx < body_offsets.len() {
                        let target = body_offsets[*case_idx];
                        let jump_ip = *jump + 2; // after the br_if_true + offset
                        let offset = target as i16 - jump_ip as i16;
                        let c = &mut self.chunks[self.current_chunk_idx];
                        c.code[*jump] = (offset >> 8) as u8;
                        c.code[*jump + 1] = (offset & 0xff) as u8;
                    }
                }

                // Patch default/end jump
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
            Statement::ClassDeclaration(class) => {
                self.compile_class(class)?;
                if let Some(name) = &class.name {
                    // Set name property on the constructor for instanceof support
                    self.emit(Op::dup);
                    self.emit_constant(Value::String(Rc::from(name.as_str())));
                    let name_idx = self.add_string_constant("name");
                    self.emit_u16(Op::struct_set, name_idx);
                    self.emit(Op::drop);
                    // Set __parent to parent constructor for instanceof chain
                    if let Some(ref super_expr) = class.super_class {
                        self.emit(Op::dup);
                        self.compile_expression(super_expr)?;
                        let parent_idx = self.add_string_constant("__parent");
                        self.emit_u16(Op::struct_set, parent_idx);
                        self.emit(Op::drop);
                    }
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
                self.emit(Op::i32_const_0);
                let i_slot = self.define_local("__for_of_i");
                self.emit_u16(Op::local_set, i_slot);
                self.emit(Op::drop);

                // Loop start: __i < __arr.length
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { _start_offset: loop_start, break_patches: vec![], continue_patches: vec![], label: self.pending_label.take() });

                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit(Op::str_length);
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
                self.emit(Op::i32_const_1);
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
                self.emit(Op::i32_const_0);
                let i_slot = self.define_local("__for_in_i");
                self.emit_u16(Op::local_set, i_slot);
                self.emit(Op::drop);

                // Loop: __i < __keys.length
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { _start_offset: loop_start, break_patches: vec![], continue_patches: vec![], label: self.pending_label.take() });

                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, keys_slot);
                self.emit(Op::str_length);
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
                self.emit(Op::i32_const_1);
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot);
                self.emit(Op::drop);

                self.emit_loop(loop_start);
                self.patch_jump(exit);
                let ctx = self.loop_stack.pop().unwrap();
                for p in ctx.break_patches { self.patch_jump(p); }
                self.current_scope_mut().end_scope();
            }
            Statement::Labeled { label, body } => {
                // Compile body — if it's a loop, the loop will push a LoopContext
                // We set a pending label so the next loop picks it up
                self.pending_label = Some(label.clone());
                self.compile_statement(body)?;
                self.pending_label = None;
            }
            Statement::Empty => {}

            // -- Modules --
            Statement::Import { specifiers, source } => {
                // ESM Integration (Source Phase Imports):
                // - "vybe:*" → host module, resolve at call sites
                // - "./*.wasm" → WASM module, load exports as globals
                // - "./*.js" / "./*.vb" → user module, pre-resolved by loader
                //
                // For .wasm: we store metadata so the CLI can load the module
                // and register its exports before execution.
                if source.ends_with(".wasm") {
                    // WASM ESM import: record the source + requested exports
                    // The CLI/runtime will use ModuleResolver to load these
                    for spec in specifiers {
                        match spec {
                            ImportSpecifier::Named { name, alias } => {
                                let local = alias.as_ref().unwrap_or(name);
                                // Emit: global_get(name) — the runtime pre-loads WASM exports as globals
                                let idx = self.add_string_constant(&name.to_lowercase());
                                self.emit_u16(Op::global_get, idx);
                                let dst = self.add_string_constant(&local.to_lowercase());
                                self.emit_u16(Op::global_set, dst);
                                self.emit(Op::drop);
                            }
                            ImportSpecifier::Namespace(name) => {
                                // import * as math from "./math.wasm"
                                // All exports bundled as an object — handled by runtime
                                self.emit(Op::null);
                                let idx = self.add_string_constant(&name.to_lowercase());
                                self.emit_u16(Op::global_set, idx);
                                self.emit(Op::drop);
                            }
                            ImportSpecifier::Default(name) => {
                                let idx = self.add_string_constant(&name.to_lowercase());
                                self.emit_u16(Op::global_get, idx);
                                self.emit(Op::drop);
                            }
                        }
                    }
                    // Store the wasm source for the CLI to resolve
                    self.wasm_imports.push(source.clone());
                    return Ok(());
                }

                for spec in specifiers {
                    match spec {
                        ImportSpecifier::Named { name, alias } => {
                            let local_name = alias.as_ref().unwrap_or(name);
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
            Statement::Export { declaration, specifiers: _, default } => {
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
            Expression::Number(n) => {
                let v = *n;
                if v.fract() == 0.0 && v >= i32::MIN as f64 && v <= i32::MAX as f64 {
                    self.emit_constant(Value::I32(v as i32));
                } else {
                    self.emit_constant(Value::F64(v));
                }
            }
            Expression::String(s) => { self.emit_constant(Value::String(Rc::from(s.as_str()))); }
            Expression::Boolean(true) => self.emit(Op::r#true),
            Expression::Boolean(false) => self.emit(Op::r#false),
            Expression::Null => self.emit(Op::null),
            Expression::Undefined => self.emit(Op::undefined),
            Expression::This => {
                // Resolve "this" as a variable — in a method it's local slot 1,
                // in an arrow function inside a method it becomes an upvalue.
                match self.resolve_variable("this") {
                    VarResolution::Local(slot) => self.emit_u16(Op::local_get, slot),
                    VarResolution::Upvalue(idx) => self.emit_u8(Op::upvalue_get, idx),
                    VarResolution::Global => self.emit(Op::null),
                }
            }
            Expression::Super => {
                // super is used in two contexts:
                // 1. super() — call parent constructor (handled in Call)
                // 2. super.method() — call parent method (handled in Member)
                // As a standalone expression, push null placeholder
                self.emit(Op::null);
            }
            Expression::Identifier(name) => {
                // JS built-in globals
                match name.as_str() {
                    "NaN" => { self.emit_constant(Value::F64(f64::NAN)); }
                    "Infinity" => { self.emit_constant(Value::F64(f64::INFINITY)); }
                    "undefined" => { self.emit(Op::undefined); }
                    _ => match self.resolve_variable(name) {
                        VarResolution::Local(slot) => self.emit_u16(Op::local_get, slot),
                        VarResolution::Upvalue(idx) => self.emit_u8(Op::upvalue_get, idx),
                        VarResolution::Global => {
                            let idx = self.add_string_constant(name);
                            self.emit_u16(Op::global_get, idx);
                        }
                    }
                }
            }
            Expression::Binary { op, left, right } => {
                if *op == BinaryOp::NullishCoalescing {
                    self.compile_expression(left)?;
                    let chunk = &mut self.chunks[self.current_chunk_idx];
                    let (_null_jump, end_jump) = common_expr::emit_null_coalesce_start(chunk, self.line);
                    self.compile_expression(right)?;
                    common_expr::emit_null_coalesce_end(&mut self.chunks[self.current_chunk_idx], end_jump);
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
                    BinaryOp::Exp => {
                        let idx = self.import("vybe:math", "pow");
                        self.emit_host_call(idx, 2);
                    }
                    BinaryOp::BitAnd => self.emit(Op::i32_and),
                    BinaryOp::BitOr => self.emit(Op::i32_or),
                    BinaryOp::BitXor => self.emit(Op::i32_xor),
                    BinaryOp::Shl => self.emit(Op::i32_shl),
                    BinaryOp::Shr => self.emit(Op::i32_shr_s),
                    BinaryOp::UShr => self.emit(Op::i32_shr_u),
                    BinaryOp::Eq => self.emit(Op::dyn_eq),   // == (loose, with coercion)
                    BinaryOp::Neq => self.emit(Op::dyn_ne),  // != (loose)
                    BinaryOp::SEq => self.emit(Op::eq),      // === (strict, no coercion)
                    BinaryOp::SNeq => self.emit(Op::ne),     // !== (strict)
                    BinaryOp::Lt => self.emit(Op::dyn_lt),
                    BinaryOp::Gt => self.emit(Op::dyn_gt),
                    BinaryOp::Le => self.emit(Op::dyn_le),
                    BinaryOp::Ge => self.emit(Op::dyn_ge),
                    BinaryOp::InstanceOf => {
                        // a instanceof B → use ref_test opcode when B is a known identifier.
                        // ref_test uses the TypeRegistry for proper subtype checking
                        // (including WASM GC-style inheritance chains).
                        // For dynamic expressions, fall back to host call.
                        //
                        // At this point both left and right are on the stack.
                        // ref_test only needs the left value + a static type name.
                        // So: drop the right, emit ref_test with the type name.
                        if let Expression::Identifier(class_name) = right.as_ref() {
                            self.emit(Op::drop); // drop the constructor object
                            let type_idx = self.add_string_constant(&class_name.to_lowercase());
                            self.emit_u16(Op::ref_test, type_idx);
                        } else {
                            // Dynamic right-hand side — fall back to host call
                            let idx = self.import("vybe:object", "instanceOf");
                            self.emit_host_call(idx, 2);
                        }
                    }
                    BinaryOp::In => {
                        // "key" in obj → host call hasProperty(key, obj)
                        let idx = self.import("vybe:object", "hasProperty");
                        self.emit_host_call(idx, 2);
                    }
                    BinaryOp::NullishCoalescing => unreachable!(),
                }
            }
            Expression::Logical { op, left, right } => {
                self.compile_expression(left)?;
                match op {
                    LogicalOp::And => {
                        let jump = common_expr::emit_and_start(&mut self.chunks[self.current_chunk_idx], self.line);
                        self.compile_expression(right)?;
                        common_expr::emit_short_circuit_end(&mut self.chunks[self.current_chunk_idx], jump);
                    }
                    LogicalOp::Or => {
                        let jump = common_expr::emit_or_start(&mut self.chunks[self.current_chunk_idx], self.line);
                        self.compile_expression(right)?;
                        common_expr::emit_short_circuit_end(&mut self.chunks[self.current_chunk_idx], jump);
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
                    self.emit(Op::i32_const_1);
                    match op { UpdateOp::Increment => self.emit_js_add(), UpdateOp::Decrement => self.emit(Op::f64_sub) }
                    self.emit(Op::dup);
                    self.compile_store(argument)?;
                } else {
                    self.compile_expression(argument)?;
                    self.emit(Op::dup);
                    self.emit(Op::i32_const_1);
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
                self.compile_expression(test)?;
                let false_jump = common_expr::emit_ternary_start(&mut self.chunks[self.current_chunk_idx], self.line);
                self.compile_expression(consequent)?;
                let end_jump = common_expr::emit_ternary_middle(&mut self.chunks[self.current_chunk_idx], false_jump, self.line);
                self.compile_expression(alternate)?;
                common_expr::emit_ternary_end(&mut self.chunks[self.current_chunk_idx], end_jump);
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
                } else if property == "length" {
                    // .length → direct opcode (works on strings and arrays)
                    self.emit(Op::str_length);
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
                // User class: push constructor, create empty obj with __type, call
                self.compile_expression(callee)?;
                self.emit_u16(Op::struct_new, 0);
                // Set __type on the new object for instanceof support
                if let Expression::Identifier(class_name) = callee.as_ref() {
                    self.emit(Op::dup);
                    self.emit_constant(Value::String(Rc::from(class_name.as_str())));
                    let type_idx = self.add_string_constant("__type");
                    self.emit_u16(Op::struct_set, type_idx);
                    self.emit(Op::drop);
                }
                for arg in arguments { self.compile_expression(arg)?; }
                self.emit_u8(Op::call, (arguments.len() + 1) as u8);
            }
            Expression::Array(elements) => {
                for e in elements { self.compile_expression(e)?; }
                self.emit_u16(Op::array_new, elements.len() as u16);
            }
            Expression::Object(properties) => {
                let has_spread = properties.iter().any(|p| matches!(p, PropertyDef::Spread(_)));
                if has_spread {
                    // Object with spread: build incrementally via struct_set
                    {
                        let line = self.line;
                        vybe_compiler_common::dict::emit_new(&mut self.chunks[self.current_chunk_idx], line);
                    }
                    for prop in properties {
                        match prop {
                            PropertyDef::KeyValue { key, value } => {
                                self.emit(Op::dup);
                                self.compile_expression(value)?;
                                let line = self.line;
                                vybe_compiler_common::dict::emit_set_const_key(&mut self.chunks[self.current_chunk_idx], key, line);
                            }
                            PropertyDef::Shorthand(name) => {
                                self.emit(Op::dup);
                                match self.resolve_variable(name) {
                                    VarResolution::Local(slot) => self.emit_u16(Op::local_get, slot),
                                    VarResolution::Upvalue(idx) => self.emit_u8(Op::upvalue_get, idx),
                                    VarResolution::Global => { let idx = self.add_string_constant(name); self.emit_u16(Op::global_get, idx); }
                                }
                                let idx = self.add_string_constant(name);
                                self.emit_u16(Op::struct_set, idx);
                                self.emit(Op::drop);
                            }
                            PropertyDef::Method { key, value } => {
                                self.emit(Op::dup);
                                self.compile_method(value)?;
                                let idx = self.add_string_constant(key);
                                self.emit_u16(Op::struct_set, idx);
                                self.emit(Op::drop);
                            }
                            PropertyDef::Spread(src) => {
                                self.compile_expression(src)?;
                                let idx = self.import("vybe:object", "assign");
                                self.emit_host_call(idx, 2);
                            }
                            PropertyDef::Computed { key, value } => {
                                // TODO: computed + spread combo
                                self.emit(Op::dup);
                                self.compile_expression(value)?;
                                // Would need dynamic struct_set
                                self.emit(Op::drop);
                                self.emit(Op::drop);
                                let _ = key;
                            }
                        }
                    }
                } else {
                    // No spread: use efficient struct_new with k/v pairs on stack,
                    // then attach __keys array for enumeration tracking.
                    let mut count = 0u16;
                    let mut string_keys: Vec<String> = Vec::new();
                    for prop in properties {
                        match prop {
                            PropertyDef::KeyValue { key, value } => {
                                self.emit_constant(Value::String(Rc::from(key.as_str())));
                                self.compile_expression(value)?;
                                string_keys.push(key.clone());
                                count += 1;
                            }
                            PropertyDef::Shorthand(name) => {
                                self.emit_constant(Value::String(Rc::from(name.as_str())));
                                match self.resolve_variable(name) {
                                    VarResolution::Local(slot) => self.emit_u16(Op::local_get, slot),
                                    VarResolution::Upvalue(idx) => self.emit_u8(Op::upvalue_get, idx),
                                    VarResolution::Global => { let idx = self.add_string_constant(name); self.emit_u16(Op::global_get, idx); }
                                }
                                string_keys.push(name.clone());
                                count += 1;
                            }
                            PropertyDef::Method { key, value } => {
                                self.emit_constant(Value::String(Rc::from(key.as_str())));
                                self.compile_method(value)?; // method, not function — adds `this` as local 0
                                string_keys.push(key.clone());
                                count += 1;
                            }
                            PropertyDef::Computed { key, value } => {
                                self.compile_expression(key)?;
                                self.compile_expression(value)?;
                                // Computed keys can't be tracked statically in __keys
                                count += 1;
                            }
                            PropertyDef::Spread(_) => unreachable!(),
                        }
                    }
                    self.emit_u16(Op::struct_new, count);
                    // Attach __keys array for dict enumeration
                    {
                        let line = self.line;
                        let chunk = &mut self.chunks[self.current_chunk_idx];
                        chunk.emit_op(Op::dup, line);
                        // Build __keys array with all string keys
                        for k in &string_keys {
                            let idx = chunk.add_constant(Value::String(Rc::from(k.as_str())));
                            chunk.emit_op_u16(Op::r#const, idx, line);
                        }
                        chunk.emit_op_u16(Op::array_new, string_keys.len() as u16, line);
                        let keys_idx = chunk.add_constant(Value::String(Rc::from("__keys")));
                        chunk.emit_op_u16(Op::struct_set, keys_idx, line);
                        chunk.emit_op(Op::drop, line);
                    }
                }
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
                self.compile_expression(inner)?;
                common_fn::emit_await(&mut self.chunks[self.current_chunk_idx], self.line);
            }
            Expression::TemplateLiteral { quasis, expressions } => {
                // str_concat_n handles type coercion (format!("{}", v)) so no toString call needed
                let mut count = 0u8;
                for (i, quasi) in quasis.iter().enumerate() {
                    if !quasi.is_empty() || expressions.is_empty() {
                        self.emit_constant(Value::String(Rc::from(quasi.as_str())));
                        count += 1;
                    }
                    if i < expressions.len() {
                        self.compile_expression(&expressions[i])?;
                        count += 1;
                    }
                }
                if count == 0 { self.emit_constant(Value::String(Rc::from(""))); }
                else if count > 1 { self.emit_u8(Op::str_concat_n, count); }
            }
            Expression::Typeof(arg) => {
                self.compile_expression(arg)?;
                self.emit(Op::ref_typeof); // direct opcode replaces host call
            }
            Expression::Void(arg) => {
                self.compile_expression(arg)?;
                self.emit(Op::drop);
                self.emit(Op::null);
            }
            Expression::Delete(target) => {
                // delete obj.prop → remove property from object
                match target.as_ref() {
                    Expression::Member { object, property, .. } => {
                        self.compile_expression(object)?;
                        let prop = property.to_lowercase();
                        let idx = self.import("vybe:object", "deleteProperty");
                        self.emit_constant(Value::String(std::rc::Rc::from(prop.as_str())));
                        self.emit_host_call(idx, 2);
                    }
                    Expression::ComputedMember { object, property } => {
                        self.compile_expression(object)?;
                        self.compile_expression(property)?;
                        let idx = self.import("vybe:object", "deleteProperty");
                        self.emit_host_call(idx, 2);
                    }
                    _ => { self.emit(Op::r#true); }
                }
            }
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
        // super.method(args) → this.__base_method(this, args)
        if let Expression::Member { object, property, .. } = callee {
            if matches!(object.as_ref(), Expression::Super) {
                let base_name = format!("__base_{}", property);
                self.emit_u16(Op::local_get, 1); // this
                let prop_idx = self.add_string_constant(&base_name);
                self.emit_u16(Op::struct_get, prop_idx);
                self.emit_u16(Op::local_get, 1); // this as first arg
                for arg in arguments { self.compile_expression(arg)?; }
                self.emit_u8(Op::call, (arguments.len() + 1) as u8);
                return Ok(());
            }
        }

        if let Expression::Member { object, property, .. } = callee {
            // obj.method() pattern
            if let Expression::Identifier(obj_name) = object.as_ref() {
                if !self.is_known_variable(obj_name) {
                    // Direct WASM opcodes for namespace functions
                    if obj_name == "Math" {
                        if let Some(()) = self.try_math_intrinsic(property, arguments)? {
                            return Ok(());
                        }
                    }
                    if let Some(()) = self.try_namespace_intrinsic(obj_name, property, arguments)? {
                        return Ok(());
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

            // Function.prototype.call/apply/bind
            match property.as_str() {
                "call" => {
                    // fn.call(thisArg, arg1, arg2, ...) → call fn with args (skip thisArg for non-methods)
                    self.compile_expression(object)?; // fn
                    // Skip thisArg (first argument), pass the rest
                    for arg in arguments.iter().skip(1) { self.compile_expression(arg)?; }
                    let argc = if arguments.is_empty() { 0 } else { (arguments.len() - 1) as u8 };
                    self.emit_u8(Op::call, argc);
                    return Ok(());
                }
                "apply" => {
                    // fn.apply(thisArg, argsArray) → call fn with thisArg
                    // Simplified: call fn(thisArg) — args array expansion not trivial
                    self.compile_expression(object)?;
                    if let Some(this_arg) = arguments.first() {
                        self.compile_expression(this_arg)?;
                    } else {
                        self.emit(Op::null);
                    }
                    self.emit_u8(Op::call, 1);
                    return Ok(());
                }
                "bind" => {
                    // fn.bind(thisArg) → return fn (simplified: no partial application)
                    self.compile_expression(object)?;
                    if !arguments.is_empty() { self.emit(Op::drop); }
                    for arg in arguments { self.compile_expression(arg)?; self.emit(Op::drop); }
                    self.compile_expression(object)?;
                    return Ok(());
                }
                "hasOwnProperty" => {
                    // obj.hasOwnProperty(key) → hasProperty(key, obj)
                    if let Some(key) = arguments.first() {
                        self.compile_expression(key)?;
                        self.compile_expression(object)?;
                        let idx = self.import("vybe:object", "hasProperty");
                        self.emit_host_call(idx, 2);
                        return Ok(());
                    }
                }
                _ => {}
            }

            // Array higher-order methods — desugar to forEach + push pattern
            // These methods need VM callbacks (calling JS functions from loops).
            // We desugar them in the compiler to bytecode loops that use `call`.
            if matches!(property.as_str(), "map" | "filter" | "forEach" | "find" | "reduce" | "sort" | "some" | "every" | "findIndex") {
                if self.compile_array_callback_method(object, property, arguments)? {
                    return Ok(());
                }
            }

            // obj IS a variable — try direct WASM string/array opcodes first
            if let Some(()) = self.try_value_method_intrinsic(object, property, arguments)? {
                return Ok(());
            }

            // Fallback to host call for value methods — but skip array-specific
            // methods on known class instances (class methods would be shadowed)
            let is_class_instance = matches!(object.as_ref(), Expression::Identifier(name)
                if self.class_instances.contains(name));
            let is_array_only = matches!(property.as_str(),
                "push" | "pop" | "shift" | "join" | "reverse" | "concat" | "fill" | "flat");
            if !(is_class_instance && is_array_only) {
                if let Some(idx) = self.resolve_value_method(property) {
                    self.compile_expression(object)?;
                    for arg in arguments { self.compile_expression(arg)?; }
                    self.emit_host_call(idx, (arguments.len() + 1) as u8);
                    return Ok(());
                }
            }

            // Method call on object. Could be a user class method or a builtin
            // collection method. Try runtime dispatch first — it returns Undefined for
            // non-builtins, in which case we fall through to struct_get + call.
            //
            // Save object to a temp local to avoid re-evaluating chained expressions.
            self.compile_expression(object)?;
            let obj_tmp = self.define_local("__obj_tmp");
            self.emit(Op::dup);
            self.emit_u16(Op::local_set, obj_tmp);
            self.emit(Op::drop);
            // callMethod(obj, "methodName", ...args)
            self.emit_constant(Value::String(Rc::from(property.as_str())));
            for arg in arguments { self.compile_expression(arg)?; }
            let cm_idx = self.import("vybe:runtime", "callMethod");
            self.emit_host_call(cm_idx, (arguments.len() + 2) as u8);

            // Check if callMethod returned Undefined (not a builtin) — then do regular call
            self.emit(Op::dup);
            self.emit(Op::ref_is_null);
            let done = self.emit_jump(Op::br_if_false); // not null/undefined = result, skip
            self.emit(Op::drop); // drop the undefined
            // Regular method call: obj.method(args) using saved temp
            self.emit_u16(Op::local_get, obj_tmp);
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
                self.emit_u16(Op::local_get, obj_tmp); // this
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
                    self.emit(Op::i32_const_0);
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
                    self.emit(Op::i32_const_0);
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
                self.emit(Op::i32_const_0);
                let i_slot = self.define_local("__cb_i");
                self.emit_u16(Op::local_set, i_slot);
                self.emit(Op::drop);
                // loop: while __i < __arr.length
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit(Op::str_length);
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
                self.emit_u16(Op::local_get, result_slot); // push result_arr
                self.emit_u16(Op::local_get, val_slot);    // push val
                self.emit(Op::array_push);                  // direct opcode
                self.emit(Op::drop);                        // discard push return
                // i++
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::i32_const_1);
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
                self.emit(Op::i32_const_0);
                let i_slot = self.define_local("__cb_i");
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit(Op::str_length);
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
                self.emit_u16(Op::local_get, result_slot);
                self.emit_u16(Op::local_get, elem_slot);
                self.emit(Op::array_push);
                self.emit(Op::drop);
                self.patch_jump(skip);
                // i++
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::i32_const_1);
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
                self.emit(Op::i32_const_0);
                let i_slot = self.define_local("__cb_i");
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit(Op::str_length);
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
                self.emit(Op::i32_const_1);
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
                self.emit(Op::i32_const_0);
                let i_slot = self.define_local("__cb_i");
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                self.emit(Op::null);
                let result_slot = self.define_local("__cb_result");
                self.emit_u16(Op::local_set, result_slot); self.emit(Op::drop);
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit(Op::str_length);
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
                self.emit(Op::i32_const_1);
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
                self.emit(Op::i32_const_0);
                let i_slot = self.define_local("__cb_i");
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit(Op::str_length);
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
                self.emit(Op::i32_const_1);
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
                self.emit(Op::str_length);
                let len_slot = self.define_local("__cb_len");
                self.emit_u16(Op::local_set, len_slot); self.emit(Op::drop);
                // outer loop: i
                self.emit(Op::i32_const_0);
                let i_slot = self.define_local("__cb_i");
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                let outer_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, len_slot);
                self.emit(Op::dyn_lt); self.emit(Op::dyn_to_bool);
                let outer_exit = self.emit_jump(Op::br_if_false);
                // inner loop: j
                self.emit(Op::i32_const_0);
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
                self.emit(Op::i32_const_1);
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
                self.emit(Op::i32_const_1);
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
                    self.emit(Op::i32_const_0);
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
                self.emit(Op::i32_const_1);
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_get, a_slot);
                self.emit_host_call(set_idx, 3);
                self.emit(Op::drop);
                self.patch_jump(no_swap);
                // j++
                self.emit_u16(Op::local_get, j_slot);
                self.emit(Op::i32_const_1);
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, j_slot); self.emit(Op::drop);
                self.emit_loop(inner_start);
                self.patch_jump(inner_exit);
                // i++
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::i32_const_1);
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                self.emit_loop(outer_start);
                self.patch_jump(outer_exit);
                self.emit_u16(Op::local_get, arr_slot);
                self.current_scope_mut().end_scope();
                Ok(true)
            }
            "some" => {
                // arr.some(fn) → loop, return true if any fn(elem) is truthy
                self.current_scope_mut().begin_scope();
                let cb = &arguments[0];
                let arr_slot = self.define_local("__some_arr");
                let i_slot = self.define_local("__some_i");
                let len_slot = self.define_local("__some_len");
                self.compile_expression(object)?;
                self.emit(Op::dup);
                self.emit_u16(Op::local_set, arr_slot); self.emit(Op::drop);
                self.emit(Op::str_length); // array length
                self.emit_u16(Op::local_set, len_slot); self.emit(Op::drop);
                self.emit(Op::i32_const_0);
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                let loop_start = self.current_offset();
                // i < len
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, len_slot);
                self.emit(Op::dyn_lt);
                let exit = self.emit_jump(Op::br_if_false);
                // callback(arr[i])
                self.compile_expression(cb)?;
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::array_get);
                self.emit_u8(Op::call, 1);
                // if truthy, return true
                self.emit(Op::dyn_to_bool);
                let found = self.emit_jump(Op::br_if_true);
                // i++
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::i32_const_1);
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                self.emit_loop(loop_start);
                // exit lands here (loop condition failed) — not found
                self.patch_jump(exit);
                self.emit(Op::r#false);
                let end = self.emit_jump(Op::br);
                // found lands here — callback was truthy
                self.patch_jump(found);
                self.emit(Op::r#true);
                self.patch_jump(end);
                self.current_scope_mut().end_scope();
                Ok(true)
            }
            "every" => {
                // arr.every(fn) → loop, return false if any fn(elem) is falsy
                self.current_scope_mut().begin_scope();
                let cb = &arguments[0];
                let arr_slot = self.define_local("__every_arr");
                let i_slot = self.define_local("__every_i");
                let len_slot = self.define_local("__every_len");
                self.compile_expression(object)?;
                self.emit(Op::dup);
                self.emit_u16(Op::local_set, arr_slot); self.emit(Op::drop);
                self.emit(Op::str_length);
                self.emit_u16(Op::local_set, len_slot); self.emit(Op::drop);
                self.emit(Op::i32_const_0);
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, len_slot);
                self.emit(Op::dyn_lt);
                let exit = self.emit_jump(Op::br_if_false);
                self.compile_expression(cb)?;
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::array_get);
                self.emit_u8(Op::call, 1);
                self.emit(Op::dyn_to_bool);
                self.emit(Op::dyn_not);
                let failed = self.emit_jump(Op::br_if_true);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::i32_const_1);
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                self.emit_loop(loop_start);
                // exit lands here (loop finished) — all passed
                self.patch_jump(exit);
                self.emit(Op::r#true);
                let end = self.emit_jump(Op::br);
                // failed lands here — callback was falsy
                self.patch_jump(failed);
                self.emit(Op::r#false);
                self.patch_jump(end);
                self.current_scope_mut().end_scope();
                Ok(true)
            }
            "findIndex" => {
                // arr.findIndex(fn) → loop, return index if fn(elem) truthy, else -1
                self.current_scope_mut().begin_scope();
                let cb = &arguments[0];
                let arr_slot = self.define_local("__fi_arr");
                let i_slot = self.define_local("__fi_i");
                let len_slot = self.define_local("__fi_len");
                self.compile_expression(object)?;
                self.emit(Op::dup);
                self.emit_u16(Op::local_set, arr_slot); self.emit(Op::drop);
                self.emit(Op::str_length);
                self.emit_u16(Op::local_set, len_slot); self.emit(Op::drop);
                self.emit(Op::i32_const_0);
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                let loop_start = self.current_offset();
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, len_slot);
                self.emit(Op::dyn_lt);
                let exit = self.emit_jump(Op::br_if_false);
                self.compile_expression(cb)?;
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::array_get);
                self.emit_u8(Op::call, 1);
                self.emit(Op::dyn_to_bool);
                let found = self.emit_jump(Op::br_if_true);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::i32_const_1);
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot); self.emit(Op::drop);
                self.emit_loop(loop_start);
                // exit lands here (loop finished) — not found
                self.patch_jump(exit);
                self.emit_constant(Value::F64(-1.0));
                let end = self.emit_jump(Op::br);
                // found lands here — callback was truthy
                self.patch_jump(found);
                self.emit_u16(Op::local_get, i_slot); // found index
                self.patch_jump(end);
                self.current_scope_mut().end_scope();
                Ok(true)
            }
            "flatMap" => {
                Ok(false) // fall through to host call
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
            AssignOp::ExpAssign => {
                let idx = self.import("vybe:math", "pow");
                self.emit_host_call(idx, 2);
            }
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
                // Stack: [value]. Need [obj, value] for struct_set.
                // Save value to temp, push obj, push value back.
                let tmp = self.define_local("__store_tmp");
                self.emit_u16(Op::local_set, tmp); self.emit(Op::drop);
                self.compile_expression(object)?;
                self.emit_u16(Op::local_get, tmp);
                let idx = self.add_string_constant(property);
                self.emit_u16(Op::struct_set, idx);
                self.emit(Op::drop); // drop struct_set result
            }
            Expression::ComputedMember { object, property } => {
                let tmp = self.define_local("__store_tmp2");
                self.emit_u16(Op::local_set, tmp); self.emit(Op::drop);
                self.compile_expression(object)?;
                self.compile_expression(property)?;
                self.emit_u16(Op::local_get, tmp);
                self.emit(Op::array_set);
            }
            _ => { self.emit(Op::drop); }
        }
        Ok(())
    }

    fn compile_function(&mut self, func: &FunctionDecl) -> Result<(), String> {
        let name = func.name.clone().unwrap_or_else(|| "<anonymous>".into());
        let chunk = common_fn::create_function_chunk(&name, func.params.len() as u8);
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        for param in &func.params { scope.define_local(&param.name); }

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        // Emit default parameter initialization: if param is null/undefined, use default
        for param in &func.params {
            if let Some(ref default_expr) = param.default {
                let slot = self.current_scope().resolve_local(&param.name).unwrap();
                self.emit_u16(Op::local_get, slot);
                // Check for both null and undefined (missing args are padded with Null)
                self.emit(Op::dup);
                self.emit(Op::ref_is_null);
                let is_null = self.emit_jump(Op::br_if_true);
                // Check undefined
                self.emit(Op::undefined);
                self.emit(Op::eq);
                let is_undef = self.emit_jump(Op::br_if_true);
                // Not null/undefined — skip default
                let skip = self.emit_jump(Op::br);
                self.patch_jump(is_null);
                self.emit(Op::drop); // drop the dup
                self.patch_jump(is_undef);
                self.compile_expression(default_expr)?;
                self.emit_u16(Op::local_set, slot);
                self.emit(Op::drop);
                self.patch_jump(skip);
            }
        }

        for stmt in &func.body { self.compile_statement(stmt)?; }
        common_fn::emit_function_epilogue(&mut self.chunks[idx], self.line);

        let lc = self.current_scope().next_slot;
        self.chunks[idx].local_count = lc;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved;

        let line = self.line;
        common_fn::emit_ref_func(&mut self.chunks[self.current_chunk_idx], idx, upvalues.len() as u8, line);
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
        let chunk = common_fn::create_function_chunk(&name, (func.params.len() + 1) as u8); // +1 for this

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
        common_fn::emit_ref_func(&mut self.chunks[self.current_chunk_idx], idx, upvalues.len() as u8, line);
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
        let had_constructor = constructor.is_some();
        let mut ctor_body = constructor.map(|c| c.body).unwrap_or_default();
        // If derived class has no explicit constructor, auto-insert super() call
        if !had_constructor && class.super_class.is_some() {
            ctor_body.insert(0, Statement::Expression(Expression::Call {
                callee: Box::new(Expression::Super),
                arguments: vec![],
                optional: false,
            }));
        }
        // Prepend field initializers before the constructor body
        for (i, stmt) in field_init_stmts.into_iter().enumerate() {
            ctor_body.insert(i, stmt);
        }
        let ctor = FunctionDecl { name: Some(name.into()), params: ctor_params, body: ctor_body, is_async: false };

        self.compile_class_constructor_full(&ctor, &regular_methods, &getters, &setters, &class.super_class)?;
        // Constructor closure is now on the stack.
        // If extends, copy parent's static methods to this constructor via Object.assign
        if let Some(ref super_expr) = class.super_class {
            self.emit(Op::dup);
            self.compile_expression(super_expr)?;
            let idx = self.import("vybe:object", "assign");
            self.emit_host_call(idx, 2);
            self.emit(Op::drop);
        }
        // Register class name as a known global *before* compiling static methods,
        // so that references like `Counter.count` inside static methods resolve as
        // global_get + struct_get instead of being treated as module imports.
        if let Some(ref class_name) = class.name {
            self.defined_globals.insert(class_name.clone());
        }
        // Attach own static methods and properties (overwrite inherited if same name)
        for (method_name, method_fn) in &static_methods {
            self.emit(Op::dup);
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

        // Default parameter initialization
        for param in &ctor.params {
            if let Some(ref default_expr) = param.default {
                let slot = self.current_scope().resolve_local(&param.name).unwrap();
                self.emit_u16(Op::local_get, slot);
                self.emit(Op::dup);
                self.emit(Op::ref_is_null);
                let is_null = self.emit_jump(Op::br_if_true);
                self.emit(Op::undefined);
                self.emit(Op::eq);
                let is_undef = self.emit_jump(Op::br_if_true);
                let skip = self.emit_jump(Op::br);
                self.patch_jump(is_null);
                self.emit(Op::drop);
                self.patch_jump(is_undef);
                self.compile_expression(default_expr)?;
                self.emit_u16(Op::local_set, slot);
                self.emit(Op::drop);
                self.patch_jump(skip);
            }
        }

        // If extends: set this.__super = parent constructor
        if let Some(super_expr) = super_class {
            self.emit_u16(Op::local_get, 1); // this
            self.compile_expression(super_expr)?; // parent class (constructor fn)
            let super_idx = self.add_string_constant("__super");
            self.emit_u16(Op::struct_set, super_idx);
            self.emit(Op::drop);
        }

        // Split constructor body at super() call: emit super() first, then attach
        // methods (so parent methods are overridden by child), then rest of constructor.
        // If no super(), methods are attached before the entire body.
        let mut super_stmts: Vec<&Statement> = Vec::new();
        let mut rest_stmts: Vec<&Statement> = Vec::new();
        let mut found_super = false;
        for stmt in &ctor.body {
            if !found_super {
                let is_super = matches!(stmt,
                    Statement::Expression(Expression::Call { callee, .. })
                    if matches!(callee.as_ref(), Expression::Super)
                );
                super_stmts.push(stmt);
                if is_super { found_super = true; }
            } else {
                rest_stmts.push(stmt);
            }
        }
        // If no super() found, all statements go to rest (methods attached first)
        if !found_super {
            rest_stmts = super_stmts;
            super_stmts = Vec::new();
        }
        // Emit super() and pre-super statements
        for stmt in &super_stmts { self.compile_statement(stmt)?; }

        // Save parent methods as __base_name before child overrides them
        if super_class.is_some() {
            let line = self.line;
            for (method_name, _) in methods.iter() {
                common_classes::emit_save_base_method(
                    &mut self.chunks[self.current_chunk_idx], 1, method_name, line,
                );
            }
        }

        // Attach child methods (overwrite parent's)
        let mut method_entries: Vec<(String, usize)> = Vec::new();
        for (method_name, method_fn) in methods {
            self.emit_u16(Op::local_get, 1); // this
            self.compile_method(method_fn)?;
            let method_chunk_idx = self.chunks.len() - 1;
            method_entries.push((method_name.clone(), method_chunk_idx));
            let prop_idx = self.add_string_constant(method_name);
            self.emit_u16(Op::struct_set, prop_idx);
            self.emit(Op::drop);
            // Cross-language aliases so JS classes are callable from Python/VB/C#
            let line = self.line;
            common_classes::emit_cross_language_aliases(
                &mut self.chunks[self.current_chunk_idx], 1, method_name, method_chunk_idx, line,
            );
        }

        // Attach getters as __get_name methods
        for (getter_name, getter_fn) in getters {
            self.emit_u16(Op::local_get, 1);
            self.compile_method(getter_fn)?;
            let getter_chunk_idx = self.chunks.len() - 1;
            let prop_name = format!("__get_{}", getter_name);
            method_entries.push((prop_name.clone(), getter_chunk_idx));
            let prop_idx = self.add_string_constant(&prop_name);
            self.emit_u16(Op::struct_set, prop_idx);
            self.emit(Op::drop);
            // Cross-language aliases for getters (e.g. __get_length → __len__)
            let line = self.line;
            common_classes::emit_cross_language_aliases(
                &mut self.chunks[self.current_chunk_idx], 1, &prop_name, getter_chunk_idx, line,
            );
        }

        // Attach setters as __set_name methods
        for (setter_name, setter_fn) in setters {
            self.emit_u16(Op::local_get, 1);
            self.compile_method(setter_fn)?;
            let setter_chunk_idx = self.chunks.len() - 1;
            let prop_name = format!("__set_{}", setter_name);
            method_entries.push((prop_name.clone(), setter_chunk_idx));
            let prop_idx = self.add_string_constant(&prop_name);
            self.emit_u16(Op::struct_set, prop_idx);
            self.emit(Op::drop);
            // Cross-language aliases for setters
            let line = self.line;
            common_classes::emit_cross_language_aliases(
                &mut self.chunks[self.current_chunk_idx], 1, &prop_name, setter_chunk_idx, line,
            );
        }

        // Compile remaining constructor body after methods are attached
        for stmt in &rest_stmts { self.compile_statement(stmt)?; }

        // Push class name to this.__types array for instanceof chain support
        // After super() + constructor body, this.__types has parent types. Append ours.
        {
            let types_idx = self.add_string_constant("__types");
            self.emit_u16(Op::local_get, 1); // this
            self.emit_u16(Op::struct_get, types_idx); // this.__types (array or null)
            // If null/undefined, create empty array
            self.emit(Op::dup);
            self.emit(Op::ref_is_null);
            let has_types = self.emit_jump(Op::br_if_false);
            self.emit(Op::drop); // drop null
            self.emit_u16(Op::array_new, 0); // create []
            self.patch_jump(has_types);
            // Push class name
            self.emit_constant(Value::String(Rc::from(ctor.name.as_deref().unwrap_or("<class>"))));
            self.emit(Op::array_push);
            // Store back
            self.emit_u16(Op::local_get, 1);
            // swap: [array, this] → need [this, array] for struct_set
            let tmp = self.define_local("__types_tmp");
            self.emit_u16(Op::local_set, tmp); self.emit(Op::drop); // save this
            // stack: [array], need [this, array]
            // Actually array is top, this is in tmp
            self.emit_u16(Op::local_get, tmp); // push this
            // swap positions: emit as struct_set expects [obj, val]
            // But array is below this now... Let me restructure
            // stack after local_get tmp: [array, this]
            // struct_set pops val then obj: val=this, obj=array — WRONG
            // Need: [this, array] for struct_set("__types")
            // Use another temp
            let arr_tmp = self.define_local("__arr_tmp");
            self.emit(Op::drop); // drop this that we just pushed
            // stack: [array]
            self.emit_u16(Op::local_set, arr_tmp); self.emit(Op::drop); // save array
            self.emit_u16(Op::local_get, 1); // this
            self.emit_u16(Op::local_get, arr_tmp); // array
            self.emit_u16(Op::struct_set, types_idx);
            self.emit(Op::drop);
        }

        // Stamp type_id on this (WASM GC).
        // The type_id is resolved at VM load time from the type table.
        {
            let tid_name = format!("__tid_{}", name.to_lowercase());
            let tid_idx = self.add_string_constant(&tid_name);
            self.emit_u16(Op::local_get, 1); // this
            self.emit_u16(Op::global_get, tid_idx);
            self.emit(Op::set_type_id);
        }

        // Return `this` from constructor
        {
            let line = self.line;
            common_classes::emit_constructor_return(
                &mut self.chunks[self.current_chunk_idx], 1, line,
            );
        }

        let lc = self.current_scope().next_slot;
        self.chunks[idx].local_count = lc;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved;

        // --- WASM GC: Register type entry in compile-time type table ---
        let parent_name = if super_class.is_some() {
            match super_class.as_ref().unwrap().as_ref() {
                Expression::Identifier(n) => n.to_lowercase(),
                _ => String::new(),
            }
        } else {
            String::new()
        };
        let field_names: Vec<String> = Vec::new(); // JS classes don't declare fields statically
        let type_entry_idx = self.type_entries.len();
        self.type_entries.push(TypeEntry {
            name: name.to_lowercase(),
            parent: parent_name,
            fields: field_names,
            methods: method_entries,
            is_interface: false,
            implements: Vec::new(),
            constructor_chunk: Some(idx),
        });
        self.class_type_ids.insert(name.to_lowercase(), type_entry_idx);

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
    fn try_bare_intrinsic(&mut self, _name: &str, _args: &[Expression]) -> Result<Option<()>, String> {
        Ok(None)
    }

    /// Emit direct WASM opcodes for namespace.method() calls (String.*, Array.*, Number.*).
    fn try_namespace_intrinsic(&mut self, ns: &str, method: &str, args: &[Expression]) -> Result<Option<()>, String> {
        match (ns, method, args.len()) {
            // String.fromCharCode(n) → str_from_char_code
            ("String", "fromCharCode", 1) => {
                self.compile_expression(&args[0])?;
                self.emit(Op::i32_from_f64);
                self.emit(Op::str_from_char_code);
                Ok(Some(()))
            }
            // Array.isArray(x) → ref_is_object check + array kind
            // Simplified: use host call (proper isArray needs kind check)
            // Number.isInteger(x) → trunc check
            ("Number", "isInteger", 1) => {
                self.compile_expression(&args[0])?;
                self.emit(Op::dup);
                self.emit(Op::f64_trunc);
                self.emit(Op::dyn_eq);
                Ok(Some(()))
            }
            // Number.isNaN(x) → x != x
            ("Number", "isNaN", 1) => {
                self.compile_expression(&args[0])?;
                self.emit(Op::dup);
                self.emit(Op::dyn_eq); // NaN != NaN → false
                self.emit(Op::dyn_not);
                Ok(Some(()))
            }
            // Array.isArray(x) → ref_is_array
            ("Array", "isArray", 1) => {
                self.compile_expression(&args[0])?;
                self.emit(Op::ref_is_array);
                Ok(Some(()))
            }
            // Array.from(iterable) → convert to array (simplified: clone if already array)
            ("Array", "from", 1) => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:array", "from");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            // Object.keys/values/entries/assign/freeze/fromEntries/hasOwn
            ("Object", "keys", 1) => { self.compile_expression(&args[0])?; let idx = self.import("vybe:object", "keys"); self.emit_host_call(idx, 1); Ok(Some(())) }
            ("Object", "values", 1) => { self.compile_expression(&args[0])?; let idx = self.import("vybe:object", "values"); self.emit_host_call(idx, 1); Ok(Some(())) }
            ("Object", "entries", 1) => { self.compile_expression(&args[0])?; let idx = self.import("vybe:object", "entries"); self.emit_host_call(idx, 1); Ok(Some(())) }
            ("Object", "freeze", 1) => { self.compile_expression(&args[0])?; let idx = self.import("vybe:object", "freeze"); self.emit_host_call(idx, 1); Ok(Some(())) }
            ("Object", "fromEntries", 1) => { self.compile_expression(&args[0])?; let idx = self.import("vybe:object", "fromEntries"); self.emit_host_call(idx, 1); Ok(Some(())) }
            ("Object", "hasOwn", 2) => { self.compile_expression(&args[0])?; self.compile_expression(&args[1])?; let idx = self.import("vybe:object", "hasOwn"); self.emit_host_call(idx, 2); Ok(Some(())) }
            ("Object", "assign", _) => {
                for a in args { self.compile_expression(a)?; }
                let idx = self.import("vybe:object", "assign");
                self.emit_host_call(idx, args.len() as u8);
                Ok(Some(()))
            }
            ("Object", "create", 1) => { self.compile_expression(&args[0])?; let idx = self.import("vybe:object", "create"); self.emit_host_call(idx, 1); Ok(Some(())) }
            ("Array", "isArray", 1) => { self.compile_expression(&args[0])?; self.emit(Op::ref_is_array); Ok(Some(())) }
            ("Array", "from", 1) => { self.compile_expression(&args[0])?; let idx = self.import("vybe:array", "from"); self.emit_host_call(idx, 1); Ok(Some(())) }
            ("Number", "isNaN", 1) => { self.compile_expression(&args[0])?; self.emit(Op::dup); self.emit(Op::dyn_ne); Ok(Some(())) }
            ("Number", "isFinite", 1) => { self.compile_expression(&args[0])?; self.emit(Op::f64_abs); self.emit_constant(Value::F64(f64::MAX)); self.emit(Op::dyn_le); Ok(Some(())) }
            ("Number", "parseInt", _) => {
                for a in args { self.compile_expression(a)?; }
                let idx = self.import("vybe:convert", "cint");
                self.emit_host_call(idx, args.len() as u8);
                Ok(Some(()))
            }
            ("Number", "parseFloat", 1) => {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:convert", "cdbl");
                self.emit_host_call(idx, 1);
                Ok(Some(()))
            }
            ("JSON", "parse", 1) => { self.compile_expression(&args[0])?; let idx = self.import("vybe:json", "parse"); self.emit_host_call(idx, 1); Ok(Some(())) }
            ("JSON", "stringify", 1) => { self.compile_expression(&args[0])?; let idx = self.import("vybe:json", "stringify"); self.emit_host_call(idx, 1); Ok(Some(())) }
            _ => Ok(None),
        }
    }

    /// Emit direct WASM string/array opcodes for value method calls.
    /// obj.method(args) where obj is a variable.
    fn try_value_method_intrinsic(&mut self, object: &Expression, method: &str, args: &[Expression]) -> Result<Option<()>, String> {
        // Skip array-specific intrinsics when the object is a known class instance —
        // class methods named "push", "pop", etc. would be shadowed by array intrinsics.
        let is_class_instance = matches!(object, Expression::Identifier(name)
            if self.class_instances.contains(name));

        // Zero-arg methods
        if args.is_empty() {
            let op = match method {
                // String methods — always safe (never clash with class method names)
                "toUpperCase" => Some(Op::str_to_upper),
                "toLowerCase" => Some(Op::str_to_lower),
                "trim" => Some(Op::str_trim),
                "trimStart" => Some(Op::str_trim_start),
                "trimEnd" => Some(Op::str_trim_end),
                "toString" => {
                    self.compile_expression(object)?;
                    let idx = self.import("vybe:convert", "toString");
                    self.emit_host_call(idx, 1);
                    return Ok(Some(()));
                }
                "join" if !is_class_instance => {
                    // join() with no args → join with ","
                    self.compile_expression(object)?;
                    self.emit_constant(Value::String(Rc::from(",")));
                    self.emit(Op::array_join);
                    return Ok(Some(()));
                }
                // Array methods — only safe on non-identifier objects (member accesses, etc.)
                "pop" if !is_class_instance => Some(Op::array_pop),
                "shift" if !is_class_instance => Some(Op::array_shift),
                "reverse" if !is_class_instance => Some(Op::array_reverse),
                "lastIndexOf" => Some(Op::str_last_index_of),
                _ => None,
            };
            if let Some(op) = op {
                self.compile_expression(object)?;
                self.emit(op);
                return Ok(Some(()));
            }
        }

        // One-arg methods: obj.method(arg)
        if args.len() == 1 {
            let op = match method {
                // String methods
                "charAt" => Some(Op::str_char_at),
                "charCodeAt" => Some(Op::str_char_code_at),
                "startsWith" => Some(Op::str_starts_with),
                "endsWith" => Some(Op::str_ends_with),
                "split" => Some(Op::str_split),
                "repeat" => Some(Op::str_repeat),
                // Array methods — only safe on non-identifier objects
                "push" if !is_class_instance => Some(Op::array_push),
                "join" if !is_class_instance => Some(Op::array_join),
                "concat" if !is_class_instance => Some(Op::array_concat),
                "indexOf" => Some(Op::str_index_of),
                "lastIndexOf" => Some(Op::str_last_index_of),
                // "includes" is polymorphic (string + array) — handled by existing dispatch
                "fill" if !is_class_instance => Some(Op::array_fill),
                _ => None,
            };
            if let Some(op) = op {
                self.compile_expression(object)?;
                self.compile_expression(&args[0])?;
                self.emit(op);
                return Ok(Some(()));
            }
        }

        // Two-arg string-only methods
        if args.len() == 2 {
            let op = match method {
                "substring" => Some(Op::str_substring),
                "padStart" => Some(Op::str_pad_start),
                "padEnd" => Some(Op::str_pad_end),
                "replace" | "replaceAll" => Some(Op::str_replace),
                // slice is polymorphic (string + array) — keep as host call
                _ => None,
            };
            if let Some(op) = op {
                self.compile_expression(object)?;
                self.compile_expression(&args[0])?;
                self.compile_expression(&args[1])?;
                self.emit(op);
                return Ok(Some(()));
            }
        }

        // ── Special multi-arg methods ──
        match method {
            "at" => {
                // arr.at(i) / str.at(i) — supports negative indexing
                if args.len() == 1 {
                    self.compile_expression(object)?;
                    self.compile_expression(&args[0])?;
                    // Negative index: handled by array_get in VM (already supports negative)
                    self.emit(Op::array_get);
                    return Ok(Some(()));
                }
            }
            "splice" if !is_class_instance => {
                if args.len() >= 2 {
                    self.compile_expression(object)?;
                    for a in args { self.compile_expression(a)?; }
                    let splice_idx = self.import("vybe:array", "splice");
                    self.emit_host_call(splice_idx, (args.len() + 1) as u8);
                    return Ok(Some(()));
                }
            }
            "unshift" if !is_class_instance => {
                if args.len() >= 1 {
                    self.compile_expression(object)?;
                    self.emit_constant(Value::I32(0));
                    self.emit_constant(Value::I32(0));
                    for a in args { self.compile_expression(a)?; }
                    let splice_idx = self.import("vybe:array", "splice");
                    self.emit_host_call(splice_idx, (args.len() + 3) as u8);
                    return Ok(Some(()));
                }
            }
            "entries" if !is_class_instance => {
                // arr.entries() → [[0, arr[0]], [1, arr[1]], ...]
                // Use enumerate pattern
                self.compile_expression(object)?;
                let enumerate_fn = self.import("vybe:array", "enumerate");
                self.emit_host_call(enumerate_fn, 1);
                return Ok(Some(()));
            }
            "flat" if !is_class_instance && args.len() <= 1 => {
                self.compile_expression(object)?;
                if args.len() == 1 {
                    self.compile_expression(&args[0])?;
                } else {
                    self.emit_constant(Value::I32(1)); // default depth = 1
                }
                let flat_fn = self.import("vybe:array", "flat");
                self.emit_host_call(flat_fn, 2);
                return Ok(Some(()));
            }
            _ => {}
        }

        Ok(None)
    }
}

enum VarResolution { Local(u16), Upvalue(u8), Global }
