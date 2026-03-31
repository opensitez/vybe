use std::rc::Rc;

use vybe_bytecode::{Chunk, Value, Op};
use vybe_parser_dart::*;

use crate::scope::Scope;

struct LoopContext {
    _start_offset: usize,
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
    classes: std::collections::HashMap<String, ClassDecl>,
    enums: std::collections::HashMap<String, EnumDecl>,
    extensions: Vec<ExtensionDecl>,
    current_class: Option<String>,
    type_entries: Vec<vybe_bytecode::chunk::TypeEntry>,
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
            classes: std::collections::HashMap::new(),
            enums: std::collections::HashMap::new(),
            extensions: Vec::new(),
            current_class: None,
            type_entries: Vec::new(),
        }
    }

    pub fn compile(mut self, program: &Program) -> Result<Vec<Chunk>, String> {
        // Pre-collect definitions
        for top in &program.body {
            match top {
                TopLevel::Extension(ext) => self.extensions.push(ext.clone()),
                TopLevel::Class(class) => { self.classes.insert(class.name.clone(), class.clone()); }
                TopLevel::Enum(en) => { self.enums.insert(en.name.clone(), en.clone()); }
                _ => {}
            }
        }

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
        // Attach type entries to script chunk
        self.chunks[0].types = self.type_entries;
        vybe_compiler_common::bundle::finalize_with_stdlib(&mut self.chunks);
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
    fn define_local(&mut self, name: &str, is_final: bool, is_const: bool) -> u16 { self.current_scope_mut().define_local(name, is_final, is_const) }

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
            "gui" => "vybe:gui",
            "db" => "vybe:database",
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
                    let slot = self.define_local(name, false, false);
                    self.emit_u16(Op::local_set, slot);
                    self.emit(Op::drop);
                }
                Ok(())
            }
            TopLevel::Class(c) => {
                self.compile_class(c)?;
                Ok(())
            }
            TopLevel::Enum(en) => self.compile_enum(en),
            TopLevel::Typedef(t) => {
                // Typedefs are currently used for clarity and metadata
                // We'll just register the name for potential resolution
                self.defined_classes.insert(t.name.clone());
                Ok(())
            }
            TopLevel::Variable(v) => self.compile_var_decl(v),
            TopLevel::Extension(ext) => self.compile_extension(ext),
            TopLevel::Statement(s) => self.compile_statement(s),
        }
    }

    fn compile_enum(&mut self, en: &EnumDecl) -> Result<(), String> {
        let enum_name = &en.name;
        self.defined_classes.insert(enum_name.clone());
        
        for (i, val) in en.values.iter().enumerate() {
            // Create a singleton instance
            self.emit_u16(Op::struct_new, 0); 
            let slot = self.define_local("__enum_tmp", true, true);
            self.emit_u16(Op::local_set, slot);
            self.emit(Op::drop);

            // Set __type
            self.emit_u16(Op::local_get, slot);
            self.emit_constant(Value::String(Rc::from(enum_name.as_str())));
            let type_idx = self.add_string_constant("__type");
            self.emit_u16(Op::struct_set, type_idx);
            self.emit(Op::drop);

            // Set index
            self.emit_u16(Op::local_get, slot);
            self.emit_constant(Value::I64(i as i64));
            let index_idx = self.add_string_constant("index");
            self.emit_u16(Op::struct_set, index_idx);
            self.emit(Op::drop);

            // Set name
            self.emit_u16(Op::local_get, slot);
            self.emit_constant(Value::String(Rc::from(val.as_str())));
            let name_idx = self.add_string_constant("name");
            self.emit_u16(Op::struct_set, name_idx);
            self.emit(Op::drop);

            // Export as EnumName.Value
            let full_name = format!("{}.{}", enum_name, val);
            self.emit_u16(Op::local_get, slot);
            self.emit_global_set(&full_name);
            self.emit(Op::drop);
        }
        Ok(())
    }

    fn compile_extension(&mut self, ext: &ExtensionDecl) -> Result<(), String> {
        let ext_name = ext.name.as_deref().unwrap_or("Extension");
        for member in &ext.members {
            if let ClassMember::Method { decl, .. } = member {
                let mut f_clone = decl.clone();
                f_clone.name = format!("{}_{}", ext_name, decl.name);
                // In a real implementation, we'd add the receiver as the first parameter.
                // For our simplified compiler, we'll assume the resolver handles the stack.
                self.compile_function_decl(&f_clone)?;
                let name = f_clone.name.clone();
                self.emit_global_set(&name);
                self.emit(Op::drop);
            }
        }
        Ok(())
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
                    let slot = self.define_local(name, false, false);
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
                self.loop_stack.push(LoopContext { _start_offset: start, break_patches: vec![], continue_patches: vec![] });
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
                self.loop_stack.push(LoopContext { _start_offset: start, break_patches: vec![], continue_patches: vec![] });
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
                self.loop_stack.push(LoopContext { _start_offset: start, break_patches: vec![], continue_patches: vec![] });
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
                let arr_slot = self.define_local("__for_in_arr", true, false);
                self.emit_u16(Op::local_set, arr_slot);
                self.emit(Op::drop);
                // __i = 0
                self.emit_constant(Value::F64(0.0));
                let i_slot = self.define_local("__for_in_i", false, false);
                self.emit_u16(Op::local_set, i_slot);
                self.emit(Op::drop);
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext { _start_offset: loop_start, break_patches: vec![], continue_patches: vec![] });
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
                let var_slot = self.define_local(var_name, false, false);
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
                self.loop_stack.push(LoopContext { _start_offset: 0, break_patches: vec![], continue_patches: vec![] });
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
                let try_chunk = self.current_chunk_idx;
                let line = self.line;
                let c = &mut self.chunks[self.current_chunk_idx];
                c.emit_op(Op::try_start, line);
                c.emit(0, line); c.emit(0, line);
                c.emit(0, line); c.emit(0, line);
                
                for s in body { self.compile_statement(s)?; }
                self.emit(Op::try_end);
                
                // Emit a jump that will skip all catches; record the chunk it was emitted into
                let skip_chunk = self.current_chunk_idx;
                let skip_offset = self.emit_jump(Op::br);

                // Ensure subsequent catch emission happens in the same chunk as the try
                let saved_chunk = self.current_chunk_idx;
                self.current_chunk_idx = try_chunk;
                let catch_entry_pos = self.current_offset();

                // Patch try_start with catch_entry_pos (uses the original try's chunk)
                let ip_after = try_start_pos + 5; // try_start(1) + offset(2) + finally(2)
                let catch_offset = catch_entry_pos as i16 - ip_after as i16;
                let c = &mut self.chunks[try_chunk];
                c.code[try_start_pos + 1] = (catch_offset >> 8) as u8;
                c.code[try_start_pos + 2] = (catch_offset & 0xff) as u8;

                // Optional debug dump: if `VYBE_DART_DEBUG_CATCH` is set in the
                // environment, print a short slice of the compiled code and
                // constants around this try/catch for offline inspection.
                if std::env::var("VYBE_DART_DEBUG_CATCH").is_ok() {
                    let c = &self.chunks[self.current_chunk_idx];
                    let start = try_start_pos;
                    let end = std::cmp::min(catch_entry_pos + 16, c.code.len());
                    eprintln!("--- DART CATCH BYTECODE DUMP (chunk {}) ---", self.current_chunk_idx);
                    eprintln!("code[{}..{}]: {:?}", start, end, &c.code[start..end]);
                    eprintln!("constants (len={}):", c.constants.len());
                    for (i, cons) in c.constants.iter().enumerate() {
                        eprintln!("  [{}] {:?}", i, cons);
                        if i > 50 { break; }
                    }
                    // Full dump if requested
                    if std::env::var("VYBE_DART_DEBUG_CATCH_FULL").is_ok() {
                        eprintln!("--- FULL CHUNK DUMP ---");
                        eprintln!("code (len={}): {:?}", c.code.len(), c.code);
                        eprintln!("constants (len={}): {:?}", c.constants.len(), c.constants);
                        eprintln!("--- END FULL DUMP ---");
                    }
                }

                // Handle catches
                let mut end_jumps: Vec<(usize, usize)> = Vec::new();
                for (i, catch) in catches.iter().enumerate() {
                    let is_last = i == catches.len() - 1;
                    
                    if let Some(ref type_name) = catch.on_type {
                        self.emit(Op::dup);
                        let type_name_norm = self.runtime_type_name(type_name);
                        // Use ref_test with normalized runtime type name for consistent behavior.
                        let type_idx = self.add_string_constant(&type_name_norm);
                        self.emit_u16(Op::ref_test, type_idx);
                        let next_catch = self.emit_jump(Op::br_if_false);

                        // Optional runtime trace (stack-safe): log a label and the exception
                        // value, then drop the host-call result so the exception value remains
                        // on the stack for the upcoming local_set.
                        if std::env::var("VYBE_DART_DEBUG_CATCH_TRACE").is_ok() {
                            let log_imp = self.import("wasi:cli", "log");
                            self.emit_constant(Value::String(Rc::from("[dart-catch] exception:")));
                            self.emit_host_call(log_imp, 1);
                            self.emit(Op::drop);
                            self.emit(Op::dup);
                            self.emit_host_call(log_imp, 1);
                            self.emit(Op::drop);
                        }

                        // This block matched
                        self.current_scope_mut().begin_scope();
                        if let Some(ref var) = catch.var_name {
                            let slot = self.define_local(var, false, false);
                            self.emit_u16(Op::local_set, slot);
                            self.emit(Op::drop);
                        } else {
                            self.emit(Op::drop);
                        }
                        for s in &catch.body { self.compile_statement(s)?; }
                        self.current_scope_mut().end_scope();
                        
                        if !is_last {
                            // record the chunk and offset for later patching, because
                            // compiling the catch body may temporarily switch current_chunk_idx
                            end_jumps.push((self.current_chunk_idx, self.emit_jump(Op::br)));
                        }
                        self.patch_jump(next_catch);
                    } else {
                        // Bare catch - always matches
                        self.current_scope_mut().begin_scope();
                        if let Some(ref var) = catch.var_name {
                            let slot = self.define_local(var, false, false);
                            self.emit_u16(Op::local_set, slot);
                            self.emit(Op::drop);
                        } else {
                            self.emit(Op::drop);
                        }
                        for s in &catch.body { self.compile_statement(s)?; }
                        self.current_scope_mut().end_scope();
                        break; // further catches are unreachable
                    }
                }
                
                // If we fell through (no match), and it's not a rethrow, we should probably rethrow or drop
                if !catches.is_empty() && catches.last().unwrap().on_type.is_some() {
                    self.emit(Op::throw); 
                } else if catches.is_empty() {
                    self.emit(Op::drop);
                }

                for (ch, off) in end_jumps { self.chunks[ch].patch_jump(off); }
                self.chunks[skip_chunk].patch_jump(skip_offset);

                // restore previously selected chunk
                self.current_chunk_idx = saved_chunk;

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
            let slot = self.define_local(name, v.is_final, v.is_const);
            self.emit_u16(Op::local_set, slot);
            self.emit(Op::drop);
        }
        Ok(())
    }

    // ── Function compilation ──────────────────────────────────────────────

    fn compile_function_decl(&mut self, f: &FunctionDecl) -> Result<(), String> {
        let name = &f.name;
        let positional_arity = f.params.positional.len() + f.params.optional_pos.len();
        let has_named = !f.params.named.is_empty();
        
        let mut chunk = Chunk::new(name);
        chunk.arity = (positional_arity + if has_named { 1 } else { 0 }) as u8;
        let idx = self.chunks.len();
        self.chunks.push(chunk);
        
        let mut scope = Scope::new_function();
        // If we're compiling a method inside a class, provide a `this` local
        if self.current_class.is_some() {
            scope.define_local("this", true, false);
        }
        for p in &f.params.positional { scope.define_local(&p.name, false, false); }
        for p in &f.params.optional_pos { scope.define_local(&p.name, false, false); }
        
        let named_args_slot = if has_named {
            Some(scope.define_local("__named_args", true, false))
        } else {
            None
        };
        
        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        if f.is_async {
            // Compile the actual body into a sub-chunk (the "worker")
            let body_chunk = Chunk::new(format!("{}$async", name));
            let body_idx = self.chunks.len();
            self.chunks.push(body_chunk);
            
            let saved_chunk = self.current_chunk_idx;
            self.current_chunk_idx = body_idx;
            
            // Compile locals and body into the sub-chunk
            // (Simplified: in a real implementation we'd pass all current locals as upvalues)
            match &f.body {
                FunctionBody::Block(stmts) => {
                    for s in stmts { self.compile_statement(s)?; }
                }
                FunctionBody::Expression(expr) => {
                    self.compile_expression(expr)?;
                }
                FunctionBody::Empty => {}
            }
            self.emit(Op::r#return);
            self.chunks[body_idx].local_count = self.current_scope().next_slot;
            
            self.current_chunk_idx = saved_chunk;
            
            // In the main function, just call the async runner
            let runner = self.import("vybe:runtime", "asyncStart");
            self.emit_ref_func(body_idx, &[]); // 0 upvalues for now
            self.emit_host_call(runner, 1);
            self.emit(Op::r#return);
            
            self.scopes.pop();
            self.current_chunk_idx = saved;
            return Ok(());
        }

        // Handle default values for optional positional parameters
        for p in &f.params.optional_pos {
            if let Some(dv) = &p.default_value {
                let slot = self.current_scope().resolve_local(&p.name).unwrap();
                self.emit_u16(Op::local_get, slot);
                self.emit(Op::dup);
                let is_null = self.emit_jump(Op::br_if_null);
                
                // Not null path: keep the dup
                let done = self.emit_jump(Op::br);
                
                self.patch_jump(is_null);
                self.emit(Op::drop); // drop the extra null
                self.compile_expression(dv)?;
                
                self.patch_jump(done);
                self.emit_u16(Op::local_set, slot);
                self.emit(Op::drop); // clean up stack
            }
        }

        // Handle named parameters
        if let Some(map_slot) = named_args_slot {
            for p in &f.params.named {
                let target_slot = self.define_local(&p.name, false, false);
                
                // Get from the named args Map: mapValue = map[name]
                self.emit_u16(Op::local_get, map_slot);
                self.emit_constant(Value::String(Rc::from(p.name.as_str())));
                let get_idx = self.import("vybe:collections", "mapGet");
                self.emit_host_call(get_idx, 2);
                
                // If Null/Undefined and we have a default value, apply it
                if let Some(dv) = &p.default_value {
                    self.emit(Op::dup);
                    let is_null = self.emit_jump(Op::br_if_null);
                    
                    let done = self.emit_jump(Op::br);
                    
                    self.patch_jump(is_null);
                    self.emit(Op::drop); // drop the extra null
                    self.compile_expression(dv)?;
                    
                    self.patch_jump(done);
                }
                
                self.emit_u16(Op::local_set, target_slot);
                self.emit(Op::drop); // cleanup stack
            }
        }

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
        let saved_class = self.current_class.clone();
        self.current_class = Some(class.name.clone());
        
        // Class → constructor function that creates objects with methods bound
        let class_name = &class.name;
        
        // Collect all members (inheritance + mixins)
        let members = self.collect_all_members(class);
        
        let mut ctors = Vec::new();
        for member in &class.members {
            if let ClassMember::Constructor { name, .. } = member {
                ctors.push((name.clone(), member.clone()));
            }
        }
        
        // If NO constructors defined, provide a default one
        if ctors.is_empty() {
             ctors.push((None, ClassMember::Constructor {
                 name: None,
                 params: Params { positional: vec![], optional_pos: vec![], named: vec![] },
                 initializers: vec![],
                 body: None,
                 is_const: false,
                 is_factory: false,
             }));
        }

        for (ctor_name, ctor_member) in ctors {
            let full_name = if let Some(n) = &ctor_name { format!("{}.{}", class_name, n) } else { class_name.clone() };
            let mut chunk = Chunk::new(&full_name);
            
            if let ClassMember::Constructor { ref params, .. } = ctor_member {
                let positional_count = params.positional.len() + params.optional_pos.len();
                chunk.arity = positional_count as u8; // Simplification for VM arity check
                
                let idx = self.chunks.len();
                self.chunks.push(chunk);
                let mut scope = Scope::new_function();
                for p in &params.positional { scope.define_local(&p.name, false, false); }
                for p in &params.optional_pos { scope.define_local(&p.name, false, false); }
                for p in &params.named { scope.define_local(&p.name, false, false); }
                
                let saved = self.current_chunk_idx;
                self.current_chunk_idx = idx;
                self.scopes.push(scope);

                // Create `this` object
                self.emit_u16(Op::struct_new, 0);
                let this_slot = self.define_local("this", true, false);
                self.emit_u16(Op::local_set, this_slot);
                self.emit(Op::drop);

                // Set __type
                self.emit_u16(Op::local_get, this_slot);
                self.emit_constant(Value::String(Rc::from(class_name.as_str())));
                let type_idx = self.add_string_constant("__type");
                self.emit_u16(Op::struct_set, type_idx);
                self.emit(Op::drop);

                // 1. Initialize default fields FIRST
                for member in &members {
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

                // 2. Initializing formals (this.field) overwrites defaults
                for p in params.positional.iter().chain(&params.optional_pos).chain(&params.named) {
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

                // 3. Constructor initializations ( : x = 1, y = 2 )
                if let ClassMember::Constructor { ref initializers, ref body, .. } = ctor_member {
                    for init in initializers {
                        match init {
                            CtorInitializer::FieldInit(name, expr) => {
                                self.emit_u16(Op::local_get, this_slot);
                                self.compile_expression(&expr)?;
                                let prop_idx = self.add_string_constant(&name);
                                self.emit_u16(Op::struct_set, prop_idx);
                                self.emit(Op::drop);
                            }
                            CtorInitializer::SuperCall(args) => {
                                for arg in args { self.compile_expression(&arg.value)?; self.emit(Op::drop); }
                            }
                            CtorInitializer::RedirectingCall(r_name, args) => {
                                // Call the target constructor and replace `this` with its return value.
                                let target_name = if let Some(n) = r_name { format!("{}.{}", class_name, n) } else { class_name.clone() };
                                let t_idx = self.add_string_constant(&target_name);
                                self.emit_u16(Op::global_get, t_idx); // push constructor function
                                let count = self.emit_args(&args)?;
                                self.emit_u8(Op::call, count); // call -> returns constructed object

                                // overwrite this_slot with returned object
                                self.emit_u16(Op::local_set, this_slot);
                                // don't drop — local_set consumed the value
                            }
                            CtorInitializer::AssertInit(expr) => {
                                self.compile_expression(&expr)?;
                                let is_true = self.emit_jump(Op::br_if_true);
                                self.emit_constant(Value::String(Rc::from("Assertion failed")));
                                self.emit(Op::throw);
                                self.patch_jump(is_true);
                            }
                        }
                    }
                    if let Some(stmts) = body {
                        for s in stmts { self.compile_statement(&s)?; }
                    }
                }

                // 4. Bind methods
                for member in &members {
                    if let ClassMember::Method { decl, is_static, .. } = member {
                        if *is_static { continue; }
                        // compile_function_decl leaves the function value on the stack
                        self.compile_function_decl(decl)?;
                        // store the function into a temp local
                        let fn_tmp = self.define_local(&format!("__m_{}", decl.name), true, false);
                        self.emit_u16(Op::local_set, fn_tmp);

                        // set the method on `this`: this.<name> = fn_tmp
                        self.emit_u16(Op::local_get, this_slot);
                        self.emit_u16(Op::local_get, fn_tmp);
                        let m_idx = self.add_string_constant(&decl.name);
                        self.emit_u16(Op::struct_set, m_idx);
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
                self.emit_global_set(&full_name);
                self.emit(Op::drop);
            }
        }
        self.current_class = saved_class;
        self.defined_classes.insert(class_name.clone());

        // Register type entry for cross-language interop
        let field_names: Vec<String> = class.members.iter().filter_map(|m| {
            if let ClassMember::Field { name, .. } = m { Some(name.to_lowercase()) } else { None }
        }).collect();
        let method_names: Vec<(String, usize)> = class.members.iter().filter_map(|m| {
            if let ClassMember::Method { decl, is_static, .. } = m {
                if !is_static { Some((decl.name.to_lowercase(), 0usize)) } else { None }
            } else { None }
        }).collect();
        let first_ctor_chunk = self.chunks.iter().position(|c| c.name == *class_name || c.name.starts_with(&format!("{}.", class_name)));
        self.type_entries.push(vybe_bytecode::chunk::TypeEntry {
            name: class_name.to_lowercase(),
            parent: class.extends.as_ref().map(|s| s.to_lowercase()).unwrap_or_default(),
            fields: field_names,
            methods: method_names,
            is_interface: class.is_abstract,
            implements: class.implements.iter().map(|s| s.to_lowercase()).collect(),
            constructor_chunk: first_ctor_chunk,
        });

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
                                    let to_str = self.import("vybe:convert", "toString");
                                    self.emit_host_call(to_str, 1);
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
            Expression::List { elements, type_arg } => {
                for e in elements { self.compile_expression(e)?; }
                if let Some(t) = type_arg {
                    // Emit typed array creation hint
                    let type_idx = self.add_string_constant(&t.name);
                    self.emit_u16(Op::r#const, type_idx);
                    self.emit_u16(Op::array_new, elements.len() as u16 + 1); // +1 for type hint
                } else {
                    self.emit_u16(Op::array_new, elements.len() as u16);
                }
            }
            Expression::Map { entries, type_args } => {
                // Build map via host call
                let ctor = self.import("vybe:collections", "Map");
                let argc = if let Some((k, v)) = type_args {
                    self.emit_constant(Value::String(Rc::from(k.name.as_str())));
                    self.emit_constant(Value::String(Rc::from(v.name.as_str())));
                    2
                } else { 0 };
                self.emit_host_call(ctor, argc);
                for (key, val) in entries {
                    self.emit(Op::dup);
                    self.compile_expression(key)?;
                    self.compile_expression(val)?;
                    let set_idx = self.import("vybe:collections", "mapSet");
                    self.emit_host_call(set_idx, 3);
                    self.emit(Op::drop);
                }
            }
            Expression::Set { elements, type_arg } => {
                let ctor = self.import("vybe:collections", "Set");
                let argc = if let Some(t) = type_arg {
                    self.emit_constant(Value::String(Rc::from(t.name.as_str())));
                    1
                } else { 0 };
                self.emit_host_call(ctor, argc);
                for el in elements {
                    self.emit(Op::dup);
                    self.compile_expression(el)?;
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
                // Enforcement: check if left is final/const
                if let Expression::Identifier(name) = left.as_ref() {
                    if let Some(local) = self.current_scope().resolve_local_full(name) {
                        if local.is_final || local.is_const {
                            return Err(format!("Cannot assign to final/const variable '{}'", name));
                        }
                    }
                }
                
                match op {
                    AssignOp::Assign => {
                        self.compile_expression(right)?;
                    }
                    AssignOp::NullAssign => {
                        self.compile_expression(left)?;
                        self.emit(Op::dup);
                        let not_null = self.emit_jump(Op::br_if_null);
                        // Case: NOT null. Result is original value.
                        let end = self.emit_jump(Op::br);
                        
                        self.patch_jump(not_null);
                        // Case: IS null. Evaluate right and store.
                        self.emit(Op::drop);
                        self.compile_expression(right)?;
                        self.emit(Op::dup);
                        self.compile_store(left)?;
                        self.patch_jump(end);
                        return Ok(());
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
            Expression::Member { object, member, null_safe } => {
                self.compile_expression(object)?;
                if *null_safe {
                    self.emit(Op::dup);
                    let skip = self.emit_jump(Op::br_if_null);
                    self.compile_member_access(member)?;
                    let end = self.emit_jump(Op::br);
                    self.patch_jump(skip);
                    self.patch_jump(end);
                } else {
                    self.compile_member_access(member)?;
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
                let count = self.emit_args(args)?;
                self.emit_u8(Op::call, count);
            }
            Expression::Const { class, args, .. } => {
                let idx = self.add_string_constant(class);
                self.emit_u16(Op::global_get, idx);
                let count = self.emit_args(args)?;
                self.emit_u8(Op::call, count);
            }
            Expression::Cascade { object, ops, null_safe } => {
                self.compile_expression(object)?;
                let slot = self.define_local("__cascade_obj", true, false);
                self.emit_u16(Op::local_set, slot);
                self.emit(Op::drop);
                
                let mut end_cascade = None;
                if *null_safe {
                    self.emit_u16(Op::local_get, slot);
                    let is_null = self.emit_jump(Op::br_if_null);
                    end_cascade = Some(is_null);
                }

                for op in ops {
                    match op {
                        CascadeOp::Method(name, args) => {
                            let prop_idx = self.add_string_constant(name);
                            self.emit_u16(Op::local_get, slot);
                            self.emit_u16(Op::struct_get, prop_idx); // push callee
                            self.emit_u16(Op::local_get, slot);       // push receiver as first arg
                            let count = self.emit_args(args)?;
                            self.emit_u8(Op::call, count + 1);
                            self.emit(Op::drop);
                        }
                        CascadeOp::Field(name) => {
                            let prop_idx = self.add_string_constant(name);
                            self.emit_u16(Op::local_get, slot);
                            self.emit_u16(Op::struct_get, prop_idx);
                            self.emit(Op::drop);
                        }
                        CascadeOp::Assign(name, val) => {
                            self.emit_u16(Op::local_get, slot);
                            self.compile_expression(val)?;
                            let prop_idx = self.add_string_constant(name);
                            self.emit_u16(Op::struct_set, prop_idx);
                            self.emit(Op::drop);
                        }
                        CascadeOp::Index(idx_expr) => {
                            self.emit_u16(Op::local_get, slot);
                            self.compile_expression(idx_expr)?;
                            self.emit(Op::array_get);
                            self.emit(Op::drop);
                        }
                    }
                }

                if let Some(label) = end_cascade {
                    self.patch_jump(label);
                }
                self.emit_u16(Op::local_get, slot);
            }
            Expression::Switch { expr, cases } => {
                self.compile_switch_expression(expr, cases)?;
            }
            Expression::Lambda { params, body, .. } => {
                let arity = params.positional.len() + params.optional_pos.len() + params.named.len();
                let mut chunk = Chunk::new("<lambda>");
                chunk.arity = arity as u8;
                let idx = self.chunks.len();
                self.chunks.push(chunk);
                let mut scope = Scope::new_function();
                for p in &params.positional { scope.define_local(&p.name, false, false); }
                for p in &params.optional_pos { scope.define_local(&p.name, false, false); }
                for p in &params.named { scope.define_local(&p.name, false, false); }
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
                let type_idx = self.add_string_constant(&self.runtime_type_name(&type_ann.name));
                self.emit_u16(Op::ref_test, type_idx);
                if *negated { self.emit(Op::bool_not); }
            }
            Expression::As { expr: inner, type_ann } => {
                self.compile_expression(inner)?;
                let type_idx = self.add_string_constant(&self.runtime_type_name(&type_ann.name));
                self.emit_u16(Op::ref_cast, type_idx);
            }
            Expression::Await(inner) => {
                self.compile_expression(inner)?;
                self.emit(Op::r#await);
            }
            Expression::Spread(inner) => {
                self.compile_expression(inner)?;
                let spread_idx = self.import("vybe:collections", "spread");
                self.emit_host_call(spread_idx, 1);
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
            Expression::Record { elements } => {
                for el in elements {
                    self.compile_expression(&el.value)?;
                }
                self.emit_u16(Op::struct_new, elements.len() as u16);
            }
        }
        Ok(())
    }

    fn compile_member_access(&mut self, member: &str) -> Result<(), String> {
        // Special case for primitives
        if member == "length" {
            self.emit(Op::str_length);
            return Ok(());
        }

        // Try to find an extension method
        let extensions = self.extensions.clone();
        for ext in &extensions {
            for m in &ext.members {
                match m {
                    ClassMember::Method { decl, .. } if decl.name == member => {
                        // Found an extension method!
                        // Desugar to: Extension_method(receiver)
                        
                        // Stack reordering: [receiver] -> [func, receiver]
                        let tmp = self.define_local("__ext_tmp", true, false);
                        self.emit_u16(Op::local_set, tmp);
                        
                        let idx = self.add_string_constant(&format!("{}_{}", ext.name.as_deref().unwrap_or("Extension"), member));
                        self.emit_u16(Op::global_get, idx);
                        self.emit_u16(Op::local_get, tmp);
                        self.emit_u8(Op::call, 1);
                        return Ok(());
                    }
                    _ => {}
                }
            }
        }

        let prop_idx = self.add_string_constant(member);
        self.emit_u16(Op::struct_get, prop_idx);
        Ok(())
    }

    // ── Call compilation ──────────────────────────────────────────────────

    fn emit_args(&mut self, args: &[Argument]) -> Result<u8, String> {
        let mut positional_count = 0;
        let mut named_args = Vec::new();

        for arg in args {
            if let Some(label) = &arg.label {
                named_args.push((label.clone(), &arg.value));
            } else {
                self.compile_expression(&arg.value)?;
                positional_count += 1;
            }
        }

        if !named_args.is_empty() {
            // Build Map for named args
            let ctor = self.import("vybe:collections", "Map");
            self.emit_host_call(ctor, 0);
            for (label, value) in named_args {
                self.emit(Op::dup);
                self.emit_constant(Value::String(Rc::from(label.as_str())));
                self.compile_expression(value)?;
                let set_idx = self.import("vybe:collections", "mapSet");
                self.emit_host_call(set_idx, 3);
                self.emit(Op::drop);
            }
            Ok(positional_count + 1)
        } else {
            Ok(positional_count)
        }
    }

    fn compile_call(&mut self, callee: &Expression, args: &[Argument]) -> Result<(), String> {
        // Handle print() as a bare host call
        if let Expression::Identifier(name) = callee {
            if name == "print" {
                let count = self.emit_args(args)?;
                let idx = self.import("wasi:cli", "log");
                self.emit_host_call(idx, count);
                return Ok(());
            }
            // Other bare imports
            if let Some(imp) = self.resolve_bare_import(name) {
                let count = self.emit_args(args)?;
                self.emit_host_call(imp, count);
                return Ok(());
            }
        }

        // Handle method calls: obj.method(args) 
        if let Expression::Member { object, member, .. } = callee {
            // Check if it's a Class named constructor call: Class.named()
            if let Expression::Identifier(obj_name) = object.as_ref() {
                if self.classes.contains_key(obj_name) {
                    let full_name = format!("{}.{}", obj_name, member);
                    let idx = self.add_string_constant(&full_name);
                    self.emit_u16(Op::global_get, idx);
                    let count = self.emit_args(args)?;
                    self.emit_u8(Op::call, count);
                    return Ok(());
                }
            }

            // Check if it's a namespace call (e.g. math.sqrt)
            if let Expression::Identifier(obj_name) = object.as_ref() {
                if !self.is_known_variable(obj_name) {
                    let module = Self::dart_module_alias(obj_name);
                    let count = self.emit_args(args)?;
                    let import_idx = self.import(module, member);
                    self.emit_host_call(import_idx, count);
                    return Ok(());
                }
            }

            // Instance method: check for known value methods
            if let Some(imp) = self.resolve_value_method(member) {
                // Push object as first arg, then remaining args
                self.compile_expression(object)?;
                let count = self.emit_args(args)?;
                self.emit_host_call(imp, count + 1);
                return Ok(());
            }

            // Check for EXTENSION methods
            let extensions = self.extensions.clone();
            for ext in &extensions {
                for m in &ext.members {
                    if let ClassMember::Method { decl, .. } = m {
                        if decl.name == *member {
                            // extension method found!
                            let idx = self.add_string_constant(&format!("{}_{}", ext.name.as_deref().unwrap_or("Extension"), member));
                            self.emit_u16(Op::global_get, idx);
                            // Push receiver (object) as 1st arg
                            self.compile_expression(object)?;
                            // Push remaining args
                            let count = self.emit_args(args)?;
                            self.emit_u8(Op::call, count + 1);
                            return Ok(());
                        }
                    }
                }
            }

            // toString() method
            if member == "toString" {
                self.compile_expression(object)?;
                let idx = self.import("vybe:convert", "toString");
                self.emit_host_call(idx, 1);
                return Ok(());
            }

            // Generic method call: obj.method(...) → preserve receiver and call with receiver as first arg
            self.compile_expression(object)?; // pushes receiver
            let obj_tmp = self.define_local("__call_obj", true, false);
            self.emit_u16(Op::local_set, obj_tmp); // pop receiver -> stored
            self.emit(Op::drop);

            self.emit_u16(Op::local_get, obj_tmp); // push receiver for struct_get
            let prop_idx = self.add_string_constant(member);
            self.emit_u16(Op::struct_get, prop_idx); // pushes function

            self.emit_u16(Op::local_get, obj_tmp); // push receiver as first arg
            let count = self.emit_args(args)?;
            self.emit_u8(Op::call, count + 1);
            return Ok(());
        }

        // Generic function call
        self.compile_expression(callee)?;
        let count = self.emit_args(args)?;
        self.emit_u8(Op::call, count);
        Ok(())
    }

    // ── Store helpers ─────────────────────────────────────────────────────

    fn compile_store(&mut self, target: &Expression) -> Result<(), String> {
        match target {
            Expression::Identifier(name) => {
                match self.resolve_variable(name) {
                    VarResolution::Local(slot) => {
                        if let Some(local) = self.current_scope().resolve_local_full(name) {
                            if local.is_final || local.is_const {
                                return Err(format!("Cannot assign to final/const variable '{}'", name));
                            }
                        }
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    }
                    VarResolution::Upvalue(idx) => {
                        self.emit_u8(Op::upvalue_set, idx);
                        self.emit(Op::drop);
                    }
                    VarResolution::Global => {
                        // If the name is a known variable (global or previously defined), set it.
                        if self.is_known_variable(name) {
                            self.emit_global_set(name);
                            self.emit(Op::drop);
                        } else if let Some(class_name) = &self.current_class {
                            // Fallback: maybe this identifier is a field on the current class
                            if let Some(cls) = self.classes.get(class_name).cloned() {
                                let is_field = self.collect_all_members(&cls).iter().any(|m| match m {
                                    ClassMember::Field { name: n, .. } => n == name,
                                    _ => false,
                                });
                                if is_field {
                                    if let Some(this_slot) = self.current_scope().resolve_local("this") {
                                        // [value] -> [this, value]
                                        let tmp = self.define_local("__tmp_as", true, false);
                                        self.emit_u16(Op::local_set, tmp);
                                        self.emit(Op::drop);

                                        self.emit_u16(Op::local_get, this_slot);
                                        self.emit_u16(Op::local_get, tmp);
                                        let prop_idx = self.add_string_constant(name);
                                        self.emit_u16(Op::struct_set, prop_idx);
                                        self.emit(Op::drop);
                                    } else {
                                        return Err(format!("Identifier '{}' matches field but 'this' not found", name));
                                    }
                                } else {
                                    return Err(format!("Undefined identifier: {}", name));
                                }
                            } else {
                                return Err(format!("Undefined identifier: {}", name));
                            }
                        } else {
                            return Err(format!("Undefined identifier: {}", name));
                        }
                    }
                }
            }
            Expression::Member { object, member, .. } => {
                self.compile_expression(object)?;
                // Stack: [value, obj] — struct_set expects [obj, val]
                // Value was pushed before this call. We need to swap.
                // Use temp local
                let tmp = self.define_local("__store_tmp", true, false);
                self.emit_u16(Op::local_set, tmp); // save obj
                self.emit(Op::drop);
                let tmp2 = self.define_local("__store_val", true, false);
                self.emit_u16(Op::local_set, tmp2); // save val
                self.emit(Op::drop);
                self.emit_u16(Op::local_get, tmp); // push obj
                self.emit_u16(Op::local_get, tmp2); // push val
                let prop_idx = self.add_string_constant(member);
                self.emit_u16(Op::struct_set, prop_idx);
                self.emit(Op::drop);
            }
            Expression::Index { object, index } => {
                let tmp = self.define_local("__idx_val", true, false);
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

    fn compile_switch_expression(&mut self, expr: &Expression, cases: &[SwitchExpressionCase]) -> Result<(), String> {
        self.compile_expression(expr)?;
        let val_slot = self.define_local("__matched_val", true, false);
        self.emit_u16(Op::local_set, val_slot);
        self.emit(Op::drop);

        let mut end_jumps = Vec::new();
        
        for case in cases {
            // Check pattern
            self.emit_u16(Op::local_get, val_slot);
            let next_case = self.compile_pattern(&case.pattern)?;
            
            // Check guard if present
            if let Some(guard) = &case.guard {
                self.compile_expression(guard)?;
                self.emit(Op::dyn_to_bool);
                let skip_guard = self.emit_jump(Op::br_if_false);
                
                // Pattern matched AND guard passed
                self.compile_expression(&case.result)?;
                end_jumps.push(self.emit_jump(Op::br));
                
                self.patch_jump(skip_guard);
            } else {
                // Pattern matched, no guard
                self.compile_expression(&case.result)?;
                end_jumps.push(self.emit_jump(Op::br));
            }
            
            self.patch_jump(next_case);
        }
        
        // Default: throw error
        let msg_idx = self.add_string_constant("Switch expression not exhaustive");
        self.emit_u16(Op::r#const, msg_idx); // this is wrong, should use emit_constant but okay for now
        self.emit(Op::throw);
        
        for j in end_jumps {
            self.patch_jump(j);
        }
        
        Ok(())
    }

    fn compile_pattern(&mut self, pattern: &Pattern) -> Result<usize, String> {
        match pattern {
            Pattern::Constant(e) => {
                self.compile_expression(e)?;
                self.emit(Op::dyn_eq);
                let skip = self.emit_jump(Op::br_if_false);
                Ok(skip)
            }
            Pattern::Wildcard => {
                self.emit(Op::drop);
                self.emit(Op::r#true);
                let skip = self.emit_jump(Op::br_if_false);
                Ok(skip)
            }
            Pattern::Variable(name) => {
                let slot = self.define_local(name, false, false);
                self.emit_u16(Op::local_set, slot);
                self.emit(Op::drop);
                self.emit(Op::r#true);
                let skip = self.emit_jump(Op::br_if_false);
                Ok(skip)
            }
            Pattern::Type(type_name) => {
                let type_name_norm = self.runtime_type_name(type_name);
                let type_idx = self.add_string_constant(&type_name_norm);
                self.emit_u16(Op::ref_test, type_idx);
                let skip = self.emit_jump(Op::br_if_false);
                Ok(skip)
            }
            Pattern::Relational { op, val } => {
                self.compile_expression(val)?;
                match op.as_str() {
                    ">" => self.emit(Op::dyn_gt),
                    "<" => self.emit(Op::dyn_lt),
                    ">=" => self.emit(Op::dyn_ge),
                    "<=" => self.emit(Op::dyn_le),
                    "==" => self.emit(Op::dyn_eq),
                    "!=" => self.emit(Op::dyn_ne),
                    _ => unreachable!(),
                }
                let skip = self.emit_jump(Op::br_if_false);
                Ok(skip)
            }
            Pattern::List(patterns) => {
                // Check if it's an array and length matches
                self.emit(Op::dup);
                self.emit(Op::ref_is_array);
                let not_array = self.emit_jump(Op::br_if_false);
                
                self.emit(Op::dup);
                self.emit(Op::array_length);
                self.emit_constant(Value::I64(patterns.len() as i64));
                self.emit(Op::eq);
                let wrong_len = self.emit_jump(Op::br_if_false);
                
                let mut p_skips = Vec::new();
                for (i, p) in patterns.iter().enumerate() {
                    self.emit(Op::dup);
                    self.emit_constant(Value::I64(i as i64));
                    self.emit(Op::array_get);
                    p_skips.push(self.compile_pattern(p)?);
                }
                
                let success = self.emit_jump(Op::br);
                self.patch_jump(not_array);
                self.patch_jump(wrong_len);
                for s in p_skips { self.patch_jump(s); }
                let fail = self.emit_jump(Op::br_if_true); // this is a hack to skip
                self.patch_jump(success);
                Ok(fail)
            }
            Pattern::Map(entries) => {
                // Check if it's an object/map
                self.emit(Op::dup);
                self.emit(Op::ref_is_object);
                let skip = self.emit_jump(Op::br_if_false);
                
                let mut e_skips = Vec::new();
                for (key_expr, val_pat) in entries {
                    self.emit(Op::dup);
                    self.compile_expression(key_expr)?;
                    self.emit(Op::array_get);
                    e_skips.push(self.compile_pattern(val_pat)?);
                }
                
                let success = self.emit_jump(Op::br);
                self.patch_jump(skip);
                for s in e_skips { self.patch_jump(s); }
                let fail = self.emit_jump(Op::br_if_true);
                self.patch_jump(success);
                Ok(fail)
            }
            Pattern::Record(elements) => {
                self.emit(Op::dup);
                self.emit(Op::ref_is_object);
                let skip = self.emit_jump(Op::br_if_false);
                
                let mut e_skips = Vec::new();
                for (i, el) in elements.iter().enumerate() {
                    self.emit(Op::dup);
                    let prop_idx = if let Some(label) = &el.label {
                        self.add_string_constant(label)
                    } else {
                        self.add_string_constant(&i.to_string())
                    };
                    self.emit_u16(Op::struct_get, prop_idx);
                    e_skips.push(self.compile_pattern(&el.pattern)?);
                }
                
                let success = self.emit_jump(Op::br);
                self.patch_jump(skip);
                for s in e_skips { self.patch_jump(s); }
                let fail = self.emit_jump(Op::br_if_true);
                self.patch_jump(success);
                Ok(fail)
            }
            Pattern::Object { class_name, fields } => {
                let type_idx = self.add_string_constant(class_name);
                self.emit(Op::dup);
                self.emit_u16(Op::ref_test, type_idx);
                let skip = self.emit_jump(Op::br_if_false);
                
                let mut f_skips = Vec::new();
                for (name, pat) in fields {
                    self.emit(Op::dup);
                    let prop_idx = self.add_string_constant(name);
                    self.emit_u16(Op::struct_get, prop_idx);
                    f_skips.push(self.compile_pattern(pat)?);
                }
                
                let success = self.emit_jump(Op::br);
                self.patch_jump(skip);
                for s in f_skips { self.patch_jump(s); }
                let fail = self.emit_jump(Op::br_if_true);
                self.patch_jump(success);
                Ok(fail)
            }
            Pattern::Logical(left, right, is_or) => {
                self.emit(Op::dup);
                let l_skip = self.compile_pattern(left)?;
                if *is_or {
                    let success = self.emit_jump(Op::br);
                    self.patch_jump(l_skip);
                    let r_skip = self.compile_pattern(right)?;
                    self.patch_jump(success);
                    Ok(r_skip)
                } else {
                    let r_skip = self.compile_pattern(right)?;
                    Ok(r_skip)
                }
            }
        }
    }

    fn collect_all_members(&self, class: &ClassDecl) -> Vec<ClassMember> {
        let mut all_members = std::collections::HashMap::new();

        // 1. Inherit from superclass (extends)
        if let Some(sup) = &class.extends {
            if let Some(sup_decl) = self.classes.get(sup) {
                for m in self.collect_all_members(sup_decl) {
                    let name = self.member_name(&m);
                    all_members.insert(name, m);
                }
            }
        }

        // 2. Mix in members (with)
        for mixin in &class.mixins {
            if let Some(mixin_decl) = self.classes.get(mixin) {
                for m in self.collect_all_members(mixin_decl) {
                    let name = self.member_name(&m);
                    all_members.insert(name, m);
                }
            }
        }

        // 3. Current class members (overrides previous ones)
        for m in &class.members {
            let name = self.member_name(m);
            all_members.insert(name, m.clone());
        }

        all_members.into_values().collect()
    }

    fn member_name(&self, m: &ClassMember) -> String {
        match m {
            ClassMember::Field { name, .. } => name.clone(),
            ClassMember::Method { decl, .. } => decl.name.clone(),
            ClassMember::Constructor { name, .. } => name.as_deref().unwrap_or("constructor").to_string(),
        }
    }

    // Normalize Dart type names to the runtime's expected type strings for
    // `ref_test` / `ref_cast`. E.g. `String` -> `string`, `int` -> `integer`.
    fn runtime_type_name(&self, name: &str) -> String {
        match name.to_lowercase().as_str() {
            "string" => "string".to_string(),
            "int" | "integer" | "i32" => "integer".to_string(),
            "double" | "f64" | "float" => "double".to_string(),
            "num" | "number" => "number".to_string(),
            "bool" | "boolean" => "boolean".to_string(),
            other => other.to_string(),
        }
    }
}
