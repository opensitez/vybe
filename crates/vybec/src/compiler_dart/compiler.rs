use std::sync::Arc;

use vybe_bytecode::{Chunk, Value, Op};
use vybe_compiler_common::classes as common_classes;
use vybe_compiler_common::collections as common_collections;
use vybe_compiler_common::expressions as common_expr;
use vybe_compiler_common::functions as common_fn;
use vybe_compiler_common::io as common_io;
use vybe_compiler_common::loops as common_loops;
use vybe_compiler_common::strings as common_strings;
use vybe_compiler_common::threading as common_thread;
use vybe_compiler_common::convert as common_convert;
use vybe_compiler_common::errors as common_errors;
use crate::parser_dart::*;

use super::scope::Scope;

struct LoopContext {
    break_label_depth: u32,
    continue_label_depth: u32,
}

enum VarResolution { Local(u16), Upvalue(u8), Global }

pub struct Compiler {
    chunks: Vec<Chunk>,
    scopes: Vec<Scope>,
    current_chunk_idx: usize,
    loop_stack: Vec<LoopContext>,
    label_depth: u32,
    line: u32,
    defined_globals: std::collections::HashSet<String>,
    defined_classes: std::collections::HashSet<String>,
    classes: std::collections::HashMap<String, ClassDecl>,
    enums: std::collections::HashMap<String, EnumDecl>,
    extensions: Vec<ExtensionDecl>,
    current_class: Option<String>,
}

impl Compiler {
    pub fn new() -> Self {
        Compiler {
            chunks: vec![Chunk::new("<script>")],
            scopes: vec![Scope::new()],
            current_chunk_idx: 0,
            loop_stack: Vec::new(),
            label_depth: 0,
            line: 1,
            defined_globals: std::collections::HashSet::new(),
            defined_classes: std::collections::HashSet::new(),
            classes: std::collections::HashMap::new(),
            enums: std::collections::HashMap::new(),
            extensions: Vec::new(),
            current_class: None,
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
            self.emit_u16(Op::GLOBAL_GET, idx);
            self.emit_u8(Op::CALL, 0);
            self.emit(Op::DROP);
        }
        self.emit(Op::NULL);
        self.emit(Op::HALT);
        let local_count = self.current_scope().next_slot;
        self.chunks[0].local_count = local_count;
        vybe_compiler_common::bundle::finalize_with_stdlib(&mut self.chunks);
        Ok(self.chunks)
    }

    // ── Emit helpers ──────────────────────────────────────────────────────

    fn chunk_mut(&mut self) -> &mut Chunk {
        &mut self.chunks[self.current_chunk_idx]
    }

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
    fn chunk(&mut self) -> &mut Chunk {
        &mut self.chunks[self.current_chunk_idx]
    }
    /// Forward-only conditional jump (for expressions, try/catch, assertions, pattern matching).
    /// Loop/break/continue use structured CF via emit_br/emit_block/emit_loop_s.
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
    fn add_string_constant(&mut self, s: &str) -> u16 {
        self.chunks[self.current_chunk_idx].add_constant(Value::String(Arc::from(s)))
    }
    fn import(&mut self, module: &str, name: &str) -> u16 {
        self.chunks[0].add_import(module, name)
    }
    fn emit_host_call(&mut self, import_idx: u16, argc: u8) {
        let line = self.line;
        let c = &mut self.chunks[self.current_chunk_idx];
        c.emit_op(Op::CALL_IMPORT, line);
        c.emit((import_idx >> 8) as u8, line);
        c.emit((import_idx & 0xff) as u8, line);
        c.emit(argc, line);
    }
    fn emit_global_set(&mut self, name: &str) {
        let idx = self.add_string_constant(name);
        self.emit_u16(Op::GLOBAL_SET, idx);
        self.defined_globals.insert(name.to_string());
    }
    fn emit_ref_func(&mut self, func_idx: usize, upvalues: &[super::scope::UpvalueDesc]) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u16(Op::REF_FUNC, func_idx as u16, line);
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
            "http" | "Http" | "HttpClient" => "wasi:http",
            "File" | "Directory" | "FileSystemEntity" | "Path" => "wasi:filesystem",
            "Map" => "vybe:collections",
            "Set" => "vybe:collections",
            "gui" => "vybe:gui",
            "db" => "vybe:database",
            // dart:io — sockets, networking (same as VB TcpClient, Python socket, JS net)
            "Socket" | "ServerSocket" | "RawDatagramSocket" | "InternetAddress" => "vybe:net",
            // dart:crypto (package:crypto) — same as VB SHA256, Python hashlib, JS crypto
            "sha256" | "sha1" | "md5" | "Hmac" => "vybe:crypto",
            // dart:core types — same as VB DateTime, StringBuilder
            "DateTime" => "vybe:types",
            "StringBuffer" | "StringBuilder" => "vybe:types",
            "Duration" => "vybe:types",
            "Stopwatch" => "vybe:threading",
            "Random" => "vybe:threading",
            "RegExp" => "vybe:regex",
            // dart:async — same as JS Promise, Python asyncio
            "Future" | "Stream" | "Completer" => "vybe:runtime",
            // dart:convert
            "utf8" | "base64" | "Encoding" => "vybe:convert",
            // dart:isolate — threading (same as Python threading, JS Worker)
            "Isolate" | "ReceivePort" | "SendPort" => "vybe:threading",
            // xml
            "XmlDocument" | "XmlElement" => "vybe:xml",
            // Process
            "Process" | "ProcessResult" => "vybe:types",
            _ => name,
        }
    }

    fn resolve_bare_import(&mut self, name: &str) -> Option<u16> {
        // Check cross-language common imports first
        if let Some((module, func)) = vybe_compiler_common::imports::resolve_common_import(name) {
            return Some(self.import(module, func));
        }
        match name {
            "print" => Some(self.import("wasi:cli", "log")),
            "int" | "double" => Some(self.import("vybe:convert", "cint")),
            _ => None,
        }
    }

    /// Compile string/array methods as direct WASM opcodes (same as Python/JS/VB/C#).
    fn try_compile_opcode_method(&mut self, object: &Expression, method: &str, args: &[Argument]) -> Result<Option<()>, String> {
        match method {
            "toUpperCase" => { self.compile_expression(object)?; common_strings::emit_to_upper(&mut self.chunks[self.current_chunk_idx], self.line); Ok(Some(())) }
            "toLowerCase" => { self.compile_expression(object)?; common_strings::emit_to_lower(&mut self.chunks[self.current_chunk_idx], self.line); Ok(Some(())) }
            "trim" => { self.compile_expression(object)?; common_strings::emit_trim(&mut self.chunks[self.current_chunk_idx], self.line); Ok(Some(())) }
            "trimLeft" => { self.compile_expression(object)?; common_strings::emit_trim_start(&mut self.chunks[self.current_chunk_idx], self.line); Ok(Some(())) }
            "trimRight" => { self.compile_expression(object)?; common_strings::emit_trim_end(&mut self.chunks[self.current_chunk_idx], self.line); Ok(Some(())) }
            "startsWith" => {
                self.compile_expression(object)?;
                if let Some(a) = args.first() { self.compile_expression(&a.value)?; }
                self.emit(Op::STR_STARTS_WITH);
                Ok(Some(()))
            }
            "endsWith" => {
                self.compile_expression(object)?;
                if let Some(a) = args.first() { self.compile_expression(&a.value)?; }
                self.emit(Op::STR_ENDS_WITH);
                Ok(Some(()))
            }
            "contains" => {
                self.compile_expression(object)?;
                if let Some(a) = args.first() { self.compile_expression(&a.value)?; }
                self.emit(Op::STR_CONTAINS);
                Ok(Some(()))
            }
            "indexOf" => {
                self.compile_expression(object)?;
                if let Some(a) = args.first() { self.compile_expression(&a.value)?; }
                common_strings::emit_index_of(&mut self.chunks[self.current_chunk_idx], self.line);
                Ok(Some(()))
            }
            "lastIndexOf" => {
                self.compile_expression(object)?;
                if let Some(a) = args.first() { self.compile_expression(&a.value)?; }
                common_strings::emit_last_index_of(&mut self.chunks[self.current_chunk_idx], self.line);
                Ok(Some(()))
            }
            "replaceAll" => {
                self.compile_expression(object)?;
                for a in args { self.compile_expression(&a.value)?; }
                common_strings::emit_replace(&mut self.chunks[self.current_chunk_idx], self.line);
                Ok(Some(()))
            }
            "split" => {
                self.compile_expression(object)?;
                for a in args { self.compile_expression(&a.value)?; }
                common_strings::emit_split(&mut self.chunks[self.current_chunk_idx], self.line);
                Ok(Some(()))
            }
            "substring" => {
                self.compile_expression(object)?;
                for a in args { self.compile_expression(&a.value)?; }
                common_strings::emit_substring(&mut self.chunks[self.current_chunk_idx], self.line);
                Ok(Some(()))
            }
            "padLeft" => {
                self.compile_expression(object)?;
                for a in args { self.compile_expression(&a.value)?; }
                self.emit(Op::STR_PAD_START);
                Ok(Some(()))
            }
            "padRight" => {
                self.compile_expression(object)?;
                for a in args { self.compile_expression(&a.value)?; }
                self.emit(Op::STR_PAD_END);
                Ok(Some(()))
            }
            // Array methods
            "add" => {
                self.compile_expression(object)?;
                for a in args { self.compile_expression(&a.value)?; }
                common_collections::emit_push(&mut self.chunks[self.current_chunk_idx], self.line);
                Ok(Some(()))
            }
            "removeLast" => {
                self.compile_expression(object)?;
                common_collections::emit_pop(&mut self.chunks[self.current_chunk_idx], self.line);
                Ok(Some(()))
            }
            "reversed" => {
                self.compile_expression(object)?;
                common_collections::emit_reverse(&mut self.chunks[self.current_chunk_idx], self.line);
                Ok(Some(()))
            }
            "join" => {
                self.compile_expression(object)?;
                if let Some(a) = args.first() {
                    self.compile_expression(&a.value)?;
                } else {
                    self.emit_constant(Value::String(Arc::from(",")));
                }
                common_collections::emit_join(&mut self.chunks[self.current_chunk_idx], self.line);
                Ok(Some(()))
            }
            // List higher-order methods (inline loops, same pattern as JS/Python)
            "map" => {
                if args.len() == 1 {
                    return self.compile_list_map(object, args);
                }
                Ok(None)
            }
            "where" => {
                if args.len() == 1 {
                    return self.compile_list_where(object, args);
                }
                Ok(None)
            }
            "forEach" => {
                if args.len() == 1 {
                    return self.compile_list_foreach(object, args);
                }
                Ok(None)
            }
            "reduce" => {
                if args.len() == 1 {
                    return self.compile_list_reduce(object, args);
                }
                Ok(None)
            }
            "any" => {
                if args.len() == 1 {
                    return self.compile_list_any_every(object, args, true);
                }
                Ok(None)
            }
            "every" => {
                if args.len() == 1 {
                    return self.compile_list_any_every(object, args, false);
                }
                Ok(None)
            }
            "toList" => {
                // .toList() is identity for lists (already a list)
                self.compile_expression(object)?;
                Ok(Some(()))
            }
            "toSet" => {
                self.compile_expression(object)?;
                let pyset = self.import("vybe:array", "pyset");
                self.emit_host_call(pyset, 1);
                Ok(Some(()))
            }
            _ => Ok(None),
        }
    }

    /// list.map(fn) → [fn(x) for x in list]
    fn compile_list_map(&mut self, object: &Expression, args: &[Argument]) -> Result<Option<()>, String> {
        let fn_slot = self.define_local("__map_fn", true, false);
        let arr_slot = self.define_local("__map_arr", true, false);
        let result_slot = self.define_local("__map_res", true, false);
        let idx_slot = self.define_local("__map_i", true, false);

        self.compile_expression(&args[0].value)?;
        self.emit_u16(Op::LOCAL_SET, fn_slot); self.emit(Op::DROP);
        self.compile_expression(object)?;
        self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);

        common_loops::emit_map(&mut self.chunks[self.current_chunk_idx], fn_slot, arr_slot, result_slot, idx_slot, self.line);
        Ok(Some(()))
    }

    /// list.where(fn) → [x for x in list if fn(x)]
    fn compile_list_where(&mut self, object: &Expression, args: &[Argument]) -> Result<Option<()>, String> {
        let fn_slot = self.define_local("__wh_fn", true, false);
        let arr_slot = self.define_local("__wh_arr", true, false);
        let result_slot = self.define_local("__wh_res", true, false);
        let idx_slot = self.define_local("__wh_i", true, false);
        let elem_slot = self.define_local("__wh_e", true, false);

        self.compile_expression(&args[0].value)?;
        self.emit_u16(Op::LOCAL_SET, fn_slot); self.emit(Op::DROP);
        self.compile_expression(object)?;
        self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);

        common_loops::emit_filter(&mut self.chunks[self.current_chunk_idx], fn_slot, arr_slot, result_slot, idx_slot, elem_slot, self.line);
        Ok(Some(()))
    }

    /// list.forEach(fn) → calls fn(x) for each x, returns null
    fn compile_list_foreach(&mut self, object: &Expression, args: &[Argument]) -> Result<Option<()>, String> {
        let fn_slot = self.define_local("__fe_fn", true, false);
        let arr_slot = self.define_local("__fe_arr", true, false);
        let idx_slot = self.define_local("__fe_i", true, false);

        self.compile_expression(&args[0].value)?;
        self.emit_u16(Op::LOCAL_SET, fn_slot); self.emit(Op::DROP);
        self.compile_expression(object)?;
        self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);

        common_loops::emit_foreach(&mut self.chunks[self.current_chunk_idx], fn_slot, arr_slot, idx_slot, self.line);
        Ok(Some(()))
    }

    /// list.reduce(fn) → fn(fn(fn(list[0], list[1]), list[2]), ...)
    fn compile_list_reduce(&mut self, object: &Expression, args: &[Argument]) -> Result<Option<()>, String> {
        let fn_slot = self.define_local("__rd_fn", true, false);
        let arr_slot = self.define_local("__rd_arr", true, false);
        let acc_slot = self.define_local("__rd_acc", true, false);
        let idx_slot = self.define_local("__rd_i", true, false);

        self.compile_expression(&args[0].value)?;
        self.emit_u16(Op::LOCAL_SET, fn_slot); self.emit(Op::DROP);
        self.compile_expression(object)?;
        self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);

        common_loops::emit_reduce(&mut self.chunks[self.current_chunk_idx], fn_slot, arr_slot, acc_slot, idx_slot, self.line);
        Ok(Some(()))
    }

    /// list.any(fn) / list.every(fn)
    fn compile_list_any_every(&mut self, object: &Expression, args: &[Argument], is_any: bool) -> Result<Option<()>, String> {
        let fn_slot = self.define_local("__ae_fn", true, false);
        let arr_slot = self.define_local("__ae_arr", true, false);
        let idx_slot = self.define_local("__ae_i", true, false);

        self.compile_expression(&args[0].value)?;
        self.emit_u16(Op::LOCAL_SET, fn_slot); self.emit(Op::DROP);
        self.compile_expression(object)?;
        self.emit_u16(Op::LOCAL_SET, arr_slot); self.emit(Op::DROP);

        common_loops::emit_any_every(&mut self.chunks[self.current_chunk_idx], fn_slot, arr_slot, idx_slot, is_any, self.line);
        Ok(Some(()))
    }

    fn resolve_value_method(&mut self, method: &str) -> Option<u16> {
        // Most string/array methods are now opcodes (try_compile_opcode_method).
        // Only methods without opcode equivalents remain here as host calls.
        let (module, name) = match method {
            "insert" => ("vybe:array", "splice"),
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
                    self.emit(Op::DROP);
                } else {
                    let slot = self.define_local(name, false, false);
                    self.emit_u16(Op::LOCAL_SET, slot);
                    self.emit(Op::DROP);
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
            self.emit_u16(Op::STRUCT_NEW, 0); 
            let slot = self.define_local("__enum_tmp", true, true);
            self.emit_u16(Op::LOCAL_SET, slot);
            self.emit(Op::DROP);

            // Set __type
            self.emit_u16(Op::LOCAL_GET, slot);
            self.emit_constant(Value::String(Arc::from(enum_name.as_str())));
            let type_idx = self.add_string_constant("__type");
            self.emit_u16(Op::STRUCT_SET, type_idx);
            self.emit(Op::DROP);

            // Set index
            self.emit_u16(Op::LOCAL_GET, slot);
            self.emit_constant(Value::I64(i as i64));
            let index_idx = self.add_string_constant("index");
            self.emit_u16(Op::STRUCT_SET, index_idx);
            self.emit(Op::DROP);

            // Set name
            self.emit_u16(Op::LOCAL_GET, slot);
            self.emit_constant(Value::String(Arc::from(val.as_str())));
            let name_idx = self.add_string_constant("name");
            self.emit_u16(Op::STRUCT_SET, name_idx);
            self.emit(Op::DROP);

            // Export as EnumName.Value
            let full_name = format!("{}.{}", enum_name, val);
            self.emit_u16(Op::LOCAL_GET, slot);
            self.emit_global_set(&full_name);
            self.emit(Op::DROP);
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
                self.emit(Op::DROP);
            }
        }
        Ok(())
    }

    // ── Statements ────────────────────────────────────────────────────────

    fn compile_statement(&mut self, stmt: &Statement) -> Result<(), String> {
        match stmt {
            Statement::Expression(expr) => {
                self.compile_expression(expr)?;
                self.emit(Op::DROP);
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
                    self.emit(Op::DROP);
                } else {
                    let slot = self.define_local(name, false, false);
                    self.emit_u16(Op::LOCAL_SET, slot);
                    self.emit(Op::DROP);
                }
            }
            Statement::If { condition, then_branch, else_branch } => {
                let line = self.line;
                let outer = self.chunk().emit_block(line); self.label_depth += 1;
                let then_block = self.chunk().emit_block(line); self.label_depth += 1;
                self.compile_expression(condition)?;
                { let line = self.line; common_convert::emit_to_bool(self.chunk_mut(), line); }
                self.emit(Op::DYN_NOT);
                self.chunk().emit_br_if(0, line);
                self.compile_statement(then_branch)?;
                self.chunk().emit_br(1, line);
                self.chunk().emit_end(line); self.chunk().patch_block(then_block); self.label_depth -= 1;
                if let Some(alt) = else_branch {
                    self.compile_statement(alt)?;
                }
                self.chunk().emit_end(line); self.chunk().patch_block(outer); self.label_depth -= 1;
            }
            Statement::While { condition, body } => {
                let line = self.line;
                let lp = common_loops::emit_loop_start(self.chunk(), line);
                self.label_depth += 2;
                self.loop_stack.push(LoopContext { break_label_depth: self.label_depth - 1, continue_label_depth: self.label_depth });
                self.compile_expression(condition)?;
                common_loops::emit_loop_cond(self.chunk(), line);
                self.compile_statement(body)?;
                self.loop_stack.pop();
                common_loops::emit_loop_end(self.chunk(), lp, line);
                self.label_depth -= 2;
            }
            Statement::DoWhile { body, condition } => {
                let line = self.line;
                let lp = common_loops::emit_do_loop_start(self.chunk(), line);
                self.label_depth += 2;
                self.loop_stack.push(LoopContext { break_label_depth: self.label_depth - 1, continue_label_depth: self.label_depth });
                self.compile_statement(body)?;
                self.compile_expression(condition)?;
                common_loops::emit_do_loop_end(self.chunk(), lp, false, line);
                self.label_depth -= 2;
                self.loop_stack.pop();
            }
            Statement::For(for_stmt) => {
                let line = self.line;
                self.current_scope_mut().begin_scope();
                if let Some(init) = &for_stmt.init {
                    match init {
                        ForInit::VarDecl(v) => { self.compile_var_decl(v)?; }
                        ForInit::Expression(e) => { self.compile_expression(e)?; self.emit(Op::DROP); }
                    }
                }
                let block_p = self.chunk().emit_block(line);
                let (loop_p, _) = self.chunk().emit_loop_s(line);
                self.label_depth += 2;
                if let Some(cond) = &for_stmt.condition {
                    self.compile_expression(cond)?;
                    common_loops::emit_loop_cond(self.chunk(), line);
                }
                let has_update = !for_stmt.update.is_empty();
                let body_block_p = if has_update {
                    let bp = self.chunk().emit_block(line); self.label_depth += 1; Some(bp)
                } else { None };
                let break_depth = self.label_depth - (if has_update { 2 } else { 1 });
                let continue_depth = self.label_depth;
                self.loop_stack.push(LoopContext { break_label_depth: break_depth, continue_label_depth: continue_depth });
                self.compile_statement(&for_stmt.body)?;
                self.loop_stack.pop();
                if let Some(bp) = body_block_p {
                    self.chunk().emit_end(line); self.chunk().patch_block(bp); self.label_depth -= 1;
                }
                for upd in &for_stmt.update {
                    self.compile_expression(upd)?;
                    self.emit(Op::DROP);
                }
                self.chunk().emit_br(0, line);
                self.chunk().emit_end(line); self.chunk().patch_loop(loop_p); self.label_depth -= 1;
                self.chunk().emit_end(line); self.chunk().patch_block(block_p); self.label_depth -= 1;
                self.current_scope_mut().end_scope();
            }
            Statement::ForIn { var_name, iterable, body, .. } => {
                self.current_scope_mut().begin_scope();
                self.compile_expression(iterable)?;
                let arr_slot = self.define_local("__for_in_arr", true, false);
                self.emit_u16(Op::LOCAL_SET, arr_slot);
                self.emit(Op::DROP);
                let i_slot = self.define_local("__for_in_i", false, false);

                let line = self.line;
                let lp = common_loops::emit_for_in_start(self.chunk_mut(), arr_slot, i_slot, line);
                let break_depth = self.label_depth + 1;
                let continue_depth = self.label_depth + 3;
                self.label_depth += 3;
                self.loop_stack.push(LoopContext { break_label_depth: break_depth, continue_label_depth: continue_depth });

                let var_slot = self.define_local(var_name, false, false);
                self.emit_u16(Op::LOCAL_SET, var_slot);
                self.emit(Op::DROP);
                self.compile_statement(body)?;

                self.loop_stack.pop();
                let line = self.line;
                common_loops::emit_for_in_end(self.chunk_mut(), i_slot, lp, line);
                self.label_depth -= 3;
                self.current_scope_mut().end_scope();
            }
            Statement::Switch { expr, cases } => {
                let line = self.line;
                self.compile_expression(expr)?;
                let disc_slot = self.define_local("__switch_disc", false, false);
                self.emit_u16(Op::LOCAL_SET, disc_slot);
                self.emit(Op::DROP);

                let outer = self.chunk().emit_block(line); self.label_depth += 1;
                self.loop_stack.push(LoopContext { break_label_depth: self.label_depth, continue_label_depth: self.label_depth });

                for case in cases {
                    let arm_block = self.chunk().emit_block(line); self.label_depth += 1;
                    if let Some(lbl) = &case.label {
                        let match_block = self.chunk().emit_block(line); self.label_depth += 1;
                        self.emit_u16(Op::LOCAL_GET, disc_slot);
                        self.compile_expression(lbl)?;
                        self.emit(Op::EQ);
                        self.emit(Op::DYN_TO_BOOL);
                        self.chunk().emit_br_if(0, line); // match → body
                        self.chunk().emit_br(1, line); // no match → skip arm
                        self.chunk().emit_end(line); self.chunk().patch_block(match_block); self.label_depth -= 1;
                    }
                    for s in &case.body { self.compile_statement(s)?; }
                    self.chunk().emit_br(1, line);
                    self.chunk().emit_end(line); self.chunk().patch_block(arm_block); self.label_depth -= 1;
                }
                self.loop_stack.pop();
                self.chunk().emit_end(line); self.chunk().patch_block(outer); self.label_depth -= 1;
            }
            Statement::Return(val) => {
                if let Some(e) = val { self.compile_expression(e)?; }
                else { self.emit(Op::NULL); }
                self.emit(Op::RETURN);
            }
            Statement::Break(_) => {
                let line = self.line;
                if let Some(ctx) = self.loop_stack.last() {
                    let depth = (self.label_depth - ctx.break_label_depth) as u8;
                    self.chunk().emit_br(depth, line);
                }
            }
            Statement::Continue(_) => {
                let line = self.line;
                if let Some(ctx) = self.loop_stack.last() {
                    let depth = (self.label_depth - ctx.continue_label_depth) as u8;
                    self.chunk().emit_br(depth, line);
                }
            }
            Statement::Throw(expr) => {
                self.compile_expression(expr)?;
                common_errors::emit_throw(&mut self.chunks[self.current_chunk_idx], self.line);
            }
            Statement::Try { body, catches, finally } => {
                let try_chunk = self.current_chunk_idx;
                let line = self.line;
                let catch_jump = common_errors::emit_try_start(&mut self.chunks[self.current_chunk_idx], line);

                for s in body { self.compile_statement(s)?; }
                common_errors::emit_try_end(&mut self.chunks[self.current_chunk_idx], self.line);

                // Emit a jump that will skip all catches; record the chunk it was emitted into
                let skip_chunk = self.current_chunk_idx;
                let skip_offset = self.emit_jump(Op::BR);

                // Ensure subsequent catch emission happens in the same chunk as the try
                let saved_chunk = self.current_chunk_idx;
                self.current_chunk_idx = try_chunk;

                // Patch catch offset
                common_errors::patch_catch(&mut self.chunks[try_chunk], catch_jump);

                // Optional debug dump: if `VYBE_DART_DEBUG_CATCH` is set in the
                // environment, print a short slice of the compiled code and
                // constants around this try/catch for offline inspection.
                if std::env::var("VYBE_DART_DEBUG_CATCH").is_ok() {
                    let c = &self.chunks[self.current_chunk_idx];
                    let start = if catch_jump > 0 { catch_jump - 1 } else { 0 };
                    let catch_entry = c.current_offset();
                    let end = std::cmp::min(catch_entry + 16, c.code.len());
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
                        self.emit(Op::DUP);
                        // Normalize for cross-language compat (Dart FormatException → ValueError etc)
                        let canonical = common_errors::canonical_exception_name(type_name);
                        let type_name_norm = self.runtime_type_name(canonical);
                        let type_idx = self.add_string_constant(&type_name_norm);
                        self.emit_u16(Op::REF_TEST, type_idx);
                        let next_catch = self.emit_jump(Op::BR_IF_FALSE);

                        // Optional runtime trace (stack-safe): log a label and the exception
                        // value, then drop the host-call result so the exception value remains
                        // on the stack for the upcoming local_set.
                        if std::env::var("VYBE_DART_DEBUG_CATCH_TRACE").is_ok() {
                            let log_imp = self.import("wasi:cli", "log");
                            self.emit_constant(Value::String(Arc::from("[dart-catch] exception:")));
                            self.emit_host_call(log_imp, 1);
                            self.emit(Op::DROP);
                            self.emit(Op::DUP);
                            self.emit_host_call(log_imp, 1);
                            self.emit(Op::DROP);
                        }

                        // This block matched
                        self.current_scope_mut().begin_scope();
                        if let Some(ref var) = catch.var_name {
                            let slot = self.define_local(var, false, false);
                            self.emit_u16(Op::LOCAL_SET, slot);
                            self.emit(Op::DROP);
                        } else {
                            self.emit(Op::DROP);
                        }
                        for s in &catch.body { self.compile_statement(s)?; }
                        self.current_scope_mut().end_scope();
                        
                        if !is_last {
                            // record the chunk and offset for later patching, because
                            // compiling the catch body may temporarily switch current_chunk_idx
                            end_jumps.push((self.current_chunk_idx, self.emit_jump(Op::BR)));
                        }
                        self.patch_jump(next_catch);
                    } else {
                        // Bare catch - always matches
                        self.current_scope_mut().begin_scope();
                        if let Some(ref var) = catch.var_name {
                            let slot = self.define_local(var, false, false);
                            self.emit_u16(Op::LOCAL_SET, slot);
                            self.emit(Op::DROP);
                        } else {
                            self.emit(Op::DROP);
                        }
                        for s in &catch.body { self.compile_statement(s)?; }
                        self.current_scope_mut().end_scope();
                        break; // further catches are unreachable
                    }
                }
                
                // If we fell through (no match), and it's not a rethrow, we should probably rethrow or drop
                if !catches.is_empty() && catches.last().unwrap().on_type.is_some() {
                    common_errors::emit_throw(&mut self.chunks[self.current_chunk_idx], self.line); 
                } else if catches.is_empty() {
                    self.emit(Op::DROP);
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
                { let line = self.line; common_convert::emit_to_bool(self.chunk_mut(), line); }
                let ok = self.emit_jump(Op::BR_IF_TRUE);
                if let Some(m) = msg {
                    self.compile_expression(m)?;
                } else {
                    self.emit_constant(Value::String(Arc::from("Assertion failed")));
                }
                common_errors::emit_throw(&mut self.chunks[self.current_chunk_idx], self.line);
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
            self.emit(Op::NULL);
        }
        let name = &v.name;
        if self.scopes.len() == 1 && self.current_scope().depth == 0 {
            self.emit_global_set(name);
            self.emit(Op::DROP);
        } else {
            let slot = self.define_local(name, v.is_final, v.is_const);
            self.emit_u16(Op::LOCAL_SET, slot);
            self.emit(Op::DROP);
        }
        Ok(())
    }

    // ── Function compilation ──────────────────────────────────────────────

    fn compile_function_decl(&mut self, f: &FunctionDecl) -> Result<(), String> {
        let name = &f.name;
        let positional_arity = f.params.positional.len() + f.params.optional_pos.len();
        let has_named = !f.params.named.is_empty();
        
        let arity = (positional_arity + if has_named { 1 } else { 0 }) as u8;
        let chunk = common_fn::create_function_chunk(name, arity);
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
            // Async functions compile like normal functions — the body runs
            // synchronously until it hits an `await` expression, which emits
            // Op::r#await. The VM suspends the fiber at that point and resumes
            // when the promise resolves. No special wrapper needed.
            //
            // This is the same approach across all languages (Python, JS, Dart, C#).
            // The VM's fiber system + event loop handle suspension/resumption.
            match &f.body {
                FunctionBody::Block(stmts) => {
                    for s in stmts { self.compile_statement(s)?; }
                }
                FunctionBody::Expression(expr) => {
                    self.compile_expression(expr)?;
                    self.emit(Op::RETURN);
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
            common_fn::emit_function_epilogue(&mut self.chunks[idx], self.line);
            let lc = self.current_scope().next_slot;
            self.chunks[idx].local_count = lc;
            let _upvalues = self.current_scope().upvalues.clone();
            self.scopes.pop();
            self.current_chunk_idx = saved;
            return Ok(());
        }

        // Handle default values for optional positional parameters
        for p in &f.params.optional_pos {
            if let Some(dv) = &p.default_value {
                let slot = self.current_scope().resolve_local(&p.name).unwrap();
                let line = self.line;
                let skip = common_fn::emit_default_param_start(
                    &mut self.chunks[self.current_chunk_idx], slot, line,
                );
                self.compile_expression(dv)?;
                let line = self.line;
                common_fn::emit_default_param_end(
                    &mut self.chunks[self.current_chunk_idx], slot, skip, line,
                );
            }
        }

        // Handle named parameters
        if let Some(map_slot) = named_args_slot {
            for p in &f.params.named {
                let target_slot = self.define_local(&p.name, false, false);
                
                // Get from the named args map: map[name] via array_get (standard WASM)
                self.emit_u16(Op::LOCAL_GET, map_slot);
                self.emit_constant(Value::String(Arc::from(p.name.as_str())));
                common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
                
                // If Null/Undefined and we have a default value, apply it
                if let Some(dv) = &p.default_value {
                    self.emit(Op::DUP);
                    let is_null = self.emit_jump(Op::BR_IF_NULL);
                    
                    let done = self.emit_jump(Op::BR);
                    
                    self.patch_jump(is_null);
                    self.emit(Op::DROP); // drop the extra null
                    self.compile_expression(dv)?;
                    
                    self.patch_jump(done);
                }
                
                self.emit_u16(Op::LOCAL_SET, target_slot);
                self.emit(Op::DROP); // cleanup stack
            }
        }

        match &f.body {
            FunctionBody::Block(stmts) => {
                for s in stmts { self.compile_statement(s)?; }
            }
            FunctionBody::Expression(expr) => {
                self.compile_expression(expr)?;
                self.emit(Op::RETURN);
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
        let line = self.line;
        common_fn::emit_function_epilogue(&mut self.chunks[idx], line);
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

                // Create `this` object with type stamps
                let this_slot = self.define_local("this", true, false);
                common_classes::emit_new_typed_object(
                    &mut self.chunks[idx], this_slot, class_name, self.line,
                );

                // Store __super if class extends another
                if let Some(ref parent) = class.extends {
                    let parent_lower = parent.to_lowercase();
                    let is_framework = parent_lower.starts_with("system.")
                        || matches!(parent_lower.as_str(), "object" | "exception");
                    if !is_framework {
                        common_classes::emit_store_super(
                            &mut self.chunks[idx], this_slot, &parent_lower, self.line,
                        );
                    }
                }

                // 1. Initialize default fields FIRST
                for member in &members {
                    if let ClassMember::Field { name, initializer, .. } = member {
                        self.emit_u16(Op::LOCAL_GET, this_slot);
                        if let Some(init) = initializer {
                            self.compile_expression(init)?;
                        } else {
                            self.emit(Op::NULL);
                        }
                        let prop_idx = self.add_string_constant(name);
                        self.emit_u16(Op::STRUCT_SET, prop_idx);
                        self.emit(Op::DROP);
                    }
                }

                // 2. Initializing formals (this.field) overwrites defaults
                for p in params.positional.iter().chain(&params.optional_pos).chain(&params.named) {
                    if p.is_this {
                        if let Some(slot) = self.current_scope().resolve_local(&p.name) {
                            self.emit_u16(Op::LOCAL_GET, this_slot);
                            self.emit_u16(Op::LOCAL_GET, slot);
                            let prop_idx = self.add_string_constant(&p.name);
                            self.emit_u16(Op::STRUCT_SET, prop_idx);
                            self.emit(Op::DROP);
                        }
                    }
                }

                // 3. Constructor initializations ( : x = 1, y = 2 )
                if let ClassMember::Constructor { ref initializers, ref body, .. } = ctor_member {
                    for init in initializers {
                        match init {
                            CtorInitializer::FieldInit(name, expr) => {
                                self.emit_u16(Op::LOCAL_GET, this_slot);
                                self.compile_expression(&expr)?;
                                let prop_idx = self.add_string_constant(&name);
                                self.emit_u16(Op::STRUCT_SET, prop_idx);
                                self.emit(Op::DROP);
                            }
                            CtorInitializer::SuperCall(args) => {
                                // Call parent constructor with this + args
                                if let Some(ref parent) = class.extends {
                                    let parent_lower = parent.to_lowercase();
                                    let parent_idx = self.add_string_constant(&parent_lower);
                                    self.emit_u16(Op::GLOBAL_GET, parent_idx);
                                    self.emit_u16(Op::LOCAL_GET, this_slot);
                                    for arg in args { self.compile_expression(&arg.value)?; }
                                    self.emit_u8(Op::CALL, (args.len() + 1) as u8);
                                    // Parent constructor returns this (already stamped), update our this
                                    self.emit_u16(Op::LOCAL_SET, this_slot);
                                }
                            }
                            CtorInitializer::RedirectingCall(r_name, args) => {
                                // Call the target constructor and replace `this` with its return value.
                                let target_name = if let Some(n) = r_name { format!("{}.{}", class_name, n) } else { class_name.clone() };
                                let t_idx = self.add_string_constant(&target_name);
                                self.emit_u16(Op::GLOBAL_GET, t_idx); // push constructor function
                                let count = self.emit_args(&args)?;
                                self.emit_u8(Op::CALL, count); // call -> returns constructed object

                                // overwrite this_slot with returned object
                                self.emit_u16(Op::LOCAL_SET, this_slot);
                                // don't drop — local_set consumed the value
                            }
                            CtorInitializer::AssertInit(expr) => {
                                self.compile_expression(&expr)?;
                                let is_true = self.emit_jump(Op::BR_IF_TRUE);
                                self.emit_constant(Value::String(Arc::from("Assertion failed")));
                                common_errors::emit_throw(&mut self.chunks[self.current_chunk_idx], self.line);
                                self.patch_jump(is_true);
                            }
                        }
                    }
                    if let Some(stmts) = body {
                        for s in stmts { self.compile_statement(&s)?; }
                    }
                }

                // 3b. Save base methods before child overrides (for super.method() calls)
                if class.extends.is_some() {
                    // Only save methods that the child class actually overrides
                    let child_method_names: std::collections::HashSet<String> = class.members.iter().filter_map(|m| {
                        if let ClassMember::Method { decl, is_static, .. } = m {
                            if !is_static { Some(decl.name.clone()) } else { None }
                        } else { None }
                    }).collect();
                    for member in &members {
                        if let ClassMember::Method { decl, is_static, .. } = member {
                            if !is_static && child_method_names.contains(&decl.name) {
                                common_classes::emit_save_base_method(
                                    &mut self.chunks[idx], this_slot, &decl.name, self.line,
                                );
                            }
                        }
                    }
                }

                // 4. Bind instance methods + getters/setters
                let mut static_methods_to_bind: Vec<(String, usize)> = Vec::new();
                for member in &members {
                    match member {
                        ClassMember::Method { decl, is_static, kind, .. } => {
                            if *is_static {
                                // Compile static methods, record for later attachment to constructor
                                self.compile_function_decl(decl)?;
                                let fn_tmp = self.define_local(&format!("__sm_{}", decl.name), true, false);
                                self.emit_u16(Op::LOCAL_SET, fn_tmp);
                                self.emit(Op::DROP);
                                static_methods_to_bind.push((decl.name.clone(), self.chunks.len() - 1));
                                continue;
                            }
                            match kind {
                                MethodKind::Getter => {
                                    // Compile getter, bind as __get_<name>
                                    self.compile_function_decl(decl)?;
                                    let fn_tmp = self.define_local(&format!("__g_{}", decl.name), true, false);
                                    self.emit_u16(Op::LOCAL_SET, fn_tmp);
                                    self.emit(Op::DROP);
                                    let get_name = format!("__get_{}", decl.name);
                                    self.emit_u16(Op::LOCAL_GET, this_slot);
                                    self.emit_u16(Op::LOCAL_GET, fn_tmp);
                                    let m_idx = self.add_string_constant(&get_name);
                                    self.emit_u16(Op::STRUCT_SET, m_idx);
                                    self.emit(Op::DROP);
                                    // Cross-language aliases for getter
                                    let method_ci = self.chunks.len() - 1;
                                    common_classes::emit_cross_language_aliases(
                                        &mut self.chunks[idx], this_slot, &get_name, method_ci, self.line,
                                    );
                                }
                                MethodKind::Setter => {
                                    // Compile setter, bind as __set_<name>
                                    self.compile_function_decl(decl)?;
                                    let fn_tmp = self.define_local(&format!("__s_{}", decl.name), true, false);
                                    self.emit_u16(Op::LOCAL_SET, fn_tmp);
                                    self.emit(Op::DROP);
                                    let set_name = format!("__set_{}", decl.name);
                                    self.emit_u16(Op::LOCAL_GET, this_slot);
                                    self.emit_u16(Op::LOCAL_GET, fn_tmp);
                                    let m_idx = self.add_string_constant(&set_name);
                                    self.emit_u16(Op::STRUCT_SET, m_idx);
                                    self.emit(Op::DROP);
                                    // Cross-language aliases for setter
                                    let method_ci = self.chunks.len() - 1;
                                    common_classes::emit_cross_language_aliases(
                                        &mut self.chunks[idx], this_slot, &set_name, method_ci, self.line,
                                    );
                                }
                                _ => {
                                    // Regular instance method (or operator overload)
                                    self.compile_function_decl(decl)?;
                                    let fn_tmp = self.define_local(&format!("__m_{}", decl.name), true, false);
                                    self.emit_u16(Op::LOCAL_SET, fn_tmp);
                                    self.emit(Op::DROP);

                                    self.emit_u16(Op::LOCAL_GET, this_slot);
                                    self.emit_u16(Op::LOCAL_GET, fn_tmp);
                                    let m_idx = self.add_string_constant(&decl.name);
                                    self.emit_u16(Op::STRUCT_SET, m_idx);
                                    self.emit(Op::DROP);

                                    let method_ci = self.chunks.len() - 1;

                                    // Operator overload aliases: operator+ → callable via obj.operator+()
                                    // Also bind Python-style dunder aliases for cross-language interop
                                    if decl.name.starts_with("operator") {
                                        let op = &decl.name["operator".len()..];
                                        let aliases: &[&str] = match op {
                                            "+" => &["__add__", "plus"],
                                            "-" => &["__sub__", "minus"],
                                            "*" => &["__mul__", "times"],
                                            "/" => &["__truediv__", "dividedBy"],
                                            "~/" => &["__floordiv__"],
                                            "%" => &["__mod__"],
                                            "==" => &["__eq__", "equals"],
                                            "!=" => &["__ne__"],
                                            "<" => &["__lt__"],
                                            ">" => &["__gt__"],
                                            "<=" => &["__le__"],
                                            ">=" => &["__ge__"],
                                            "[]" => &["__getitem__"],
                                            "[]=" => &["__setitem__"],
                                            "~" => &["__invert__"],
                                            "&" => &["__and__"],
                                            "|" => &["__or__"],
                                            "^" => &["__xor__"],
                                            "<<" => &["__lshift__"],
                                            ">>" => &["__rshift__"],
                                            _ => &[],
                                        };
                                        for alias in aliases {
                                            common_classes::emit_bind_method(
                                                &mut self.chunks[idx], this_slot, alias, method_ci, self.line,
                                            );
                                        }
                                    } else {
                                        // Regular cross-language aliases
                                        common_classes::emit_cross_language_aliases(
                                            &mut self.chunks[idx], this_slot, &decl.name, method_ci, self.line,
                                        );
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }

                // Stamp __types array for instanceof support
                common_classes::emit_instanceof_chain(&mut self.chunks[idx], this_slot, class_name, self.line);

                // Return this
                common_classes::emit_constructor_return(&mut self.chunks[idx], this_slot, self.line);
                let lc = self.current_scope().next_slot;
                self.chunks[idx].local_count = lc;
                let upvalues = self.current_scope().upvalues.clone();
                self.scopes.pop();
                self.current_chunk_idx = saved;
                self.emit_ref_func(idx, &upvalues);

                // Inherit parent's static methods (before attaching own)
                if let Some(ref parent) = class.extends {
                    let parent_lower = parent.to_lowercase();
                    let is_framework = parent_lower.starts_with("system.")
                        || matches!(parent_lower.as_str(), "object" | "exception");
                    if !is_framework && self.defined_classes.contains(&parent_lower) {
                        self.emit(Op::DUP);
                        let parent_idx = self.add_string_constant(&parent_lower);
                        self.emit_u16(Op::GLOBAL_GET, parent_idx);
                        let assign_idx = self.import("vybe:object", "assign");
                        self.emit_host_call(assign_idx, 2);
                        self.emit(Op::DROP);
                    }
                }

                // Attach own static methods to constructor function
                for (sm_name, sm_ci) in &static_methods_to_bind {
                    self.emit(Op::DUP);
                    self.emit_ref_func(*sm_ci, &[]);
                    let prop_idx = self.add_string_constant(sm_name);
                    self.emit_u16(Op::STRUCT_SET, prop_idx);
                    self.emit(Op::DROP);
                }

                self.emit_global_set(&full_name);
                self.emit(Op::DROP);
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
        let parent_str = class.extends.as_ref().map(|s| s.to_lowercase()).unwrap_or_default();
        let implements_list: Vec<String> = class.implements.iter().map(|s| s.to_lowercase()).collect();
        common_classes::register_type(
            &mut self.chunks, class_name, &parent_str,
            field_names, method_names, class.is_abstract, implements_list, first_ctor_chunk,
        );

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
            Expression::Bool(true) => { self.emit(Op::TRUE); }
            Expression::Bool(false) => { self.emit(Op::FALSE); }
            Expression::Null => { self.emit(Op::NULL); }
            Expression::String(s) => {
                match s {
                    StringExpr::Simple(text) => {
                        self.emit_constant(Value::String(Arc::from(text.as_str())));
                    }
                    StringExpr::Interpolated(parts) => {
                        let count = parts.len();
                        for part in parts {
                            match part {
                                StringPart::Literal(lit) => {
                                    self.emit_constant(Value::String(Arc::from(lit.as_str())));
                                }
                                StringPart::Expr(e) => {
                                    self.compile_expression(e)?;
                                    let line = self.line;
                                    common_strings::emit_to_string(self.chunk_mut(), line);
                                }
                            }
                        }
                        let line = self.line;
                        common_strings::emit_concat(self.chunk_mut(), count, line);
                    }
                }
            }
            Expression::Identifier(name) => {
                // Check for bare host function calls handled at call sites
                match self.resolve_variable(name) {
                    VarResolution::Local(slot) => { self.emit_u16(Op::LOCAL_GET, slot); }
                    VarResolution::Upvalue(idx) => { self.emit_u8(Op::UPVALUE_GET, idx); }
                    VarResolution::Global => {
                        let idx = self.add_string_constant(name);
                        self.emit_u16(Op::GLOBAL_GET, idx);
                    }
                }
            }
            Expression::This => {
                if let Some(slot) = self.current_scope().resolve_local("this") {
                    self.emit_u16(Op::LOCAL_GET, slot);
                } else {
                    self.emit(Op::NULL);
                }
            }
            Expression::Super => {
                // super resolves to this — method dispatch uses __base_* prefix
                if let Some(slot) = self.current_scope().resolve_local("this") {
                    self.emit_u16(Op::LOCAL_GET, slot);
                } else {
                    self.emit(Op::NULL);
                }
            }
            Expression::List { elements, type_arg } => {
                for e in elements { self.compile_expression(e)?; }
                if let Some(t) = type_arg {
                    // Emit typed array creation hint
                    let type_idx = self.add_string_constant(&t.name);
                    self.emit_u16(Op::CONST, type_idx);
                    common_collections::emit_array_new(&mut self.chunks[self.current_chunk_idx], elements.len() as u16 + 1, self.line); // +1 for type hint
                } else {
                    common_collections::emit_array_new(&mut self.chunks[self.current_chunk_idx], elements.len() as u16, self.line);
                }
            }
            Expression::Map { entries, type_args: _ } => {
                // Create dict as plain Object with __keys tracking — same as all languages.
                // Pure WASM opcodes, no host calls.
                let line = self.line;
                vybe_compiler_common::dict::emit_new(self.chunk_mut(), line);
                for (key, val) in entries {
                    self.emit(Op::DUP);
                    self.compile_expression(val)?;
                    if let Expression::String(StringExpr::Simple(s)) = key {
                        let line = self.line;
                        vybe_compiler_common::dict::emit_set_const_key(self.chunk_mut(), s, line);
                    } else {
                        let tmp = self.define_local("__map_v", true, false);
                        self.emit_u16(Op::LOCAL_SET, tmp);
                        self.emit(Op::DROP);
                        self.compile_expression(key)?;
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        let line = self.line;
                        vybe_compiler_common::dict::emit_set_dynamic(self.chunk_mut(), line);
                    }
                }
            }
            Expression::Set { elements, type_arg } => {
                let ctor = self.import("vybe:collections", "Set");
                let argc = if let Some(t) = type_arg {
                    self.emit_constant(Value::String(Arc::from(t.name.as_str())));
                    1
                } else { 0 };
                self.emit_host_call(ctor, argc);
                for el in elements {
                    self.emit(Op::DUP);
                    self.compile_expression(el)?;
                    let add_idx = self.import("vybe:collections", "setAdd");
                    self.emit_host_call(add_idx, 2);
                    self.emit(Op::DROP);
                }
            }
            Expression::Binary { op, left, right } => {
                // Short-circuit logical operators
                if *op == BinOp::And {
                    self.compile_expression(left)?;
                    let line = self.line;
                    let jump = common_expr::emit_and_start(self.chunk_mut(), line);
                    self.compile_expression(right)?;
                    common_expr::emit_short_circuit_end(self.chunk_mut(), jump);
                } else if *op == BinOp::Or {
                    self.compile_expression(left)?;
                    let line = self.line;
                    let jump = common_expr::emit_or_start(self.chunk_mut(), line);
                    self.compile_expression(right)?;
                    common_expr::emit_short_circuit_end(self.chunk_mut(), jump);
                } else {
                    self.compile_expression(left)?;
                    self.compile_expression(right)?;
                    match op {
                        BinOp::Add => self.emit(Op::DYN_ADD),
                        BinOp::Sub => self.emit(Op::F64_SUB),
                        BinOp::Mul => self.emit(Op::F64_MUL),
                        BinOp::Div => self.emit(Op::F64_DIV),
                        BinOp::IntDiv => { self.emit(Op::F64_DIV); self.emit(Op::F64_FLOOR); }
                        BinOp::Mod => { let l = self.line; vybe_compiler_common::expressions::emit_f64_mod(&mut self.chunks[self.current_chunk_idx], l); },
                        BinOp::Eq => self.emit(Op::DYN_EQ),
                        BinOp::NotEq => self.emit(Op::DYN_NE),
                        BinOp::Lt => self.emit(Op::DYN_LT),
                        BinOp::Gt => self.emit(Op::DYN_GT),
                        BinOp::Le => self.emit(Op::DYN_LE),
                        BinOp::Ge => self.emit(Op::DYN_GE),
                        BinOp::And | BinOp::Or => unreachable!(),
                        BinOp::BitAnd => self.emit(Op::I32_AND),
                        BinOp::BitOr => self.emit(Op::I32_OR),
                        BinOp::BitXor => self.emit(Op::I32_XOR),
                        BinOp::Shl => self.emit(Op::I32_SHL),
                        BinOp::Shr => self.emit(Op::I32_SHR_S),
                        BinOp::UShr => self.emit(Op::I32_SHR_U),
                    }
                }
            }
            Expression::Unary { op, expr: inner } => {
                match op {
                    UnaryOp::Neg => {
                        self.compile_expression(inner)?;
                        self.emit(Op::DYN_NEG);
                    }
                    UnaryOp::Not => {
                        self.compile_expression(inner)?;
                        self.emit(Op::DYN_NOT);
                    }
                    UnaryOp::BitNot => {
                        self.compile_expression(inner)?;
                        { let l = self.line; vybe_compiler_common::expressions::emit_i32_not(&mut self.chunks[self.current_chunk_idx], l); }
                    }
                    UnaryOp::PreInc => {
                        self.compile_expression(inner)?;
                        self.emit_constant(Value::F64(1.0));
                        self.emit(Op::DYN_ADD);
                        self.compile_store(inner)?;
                        self.compile_expression(inner)?;
                    }
                    UnaryOp::PreDec => {
                        self.compile_expression(inner)?;
                        self.emit_constant(Value::F64(1.0));
                        self.emit(Op::F64_SUB);
                        self.compile_store(inner)?;
                        self.compile_expression(inner)?;
                    }
                }
            }
            Expression::PostfixUnary { op, expr: inner } => {
                self.compile_expression(inner)?;
                self.emit(Op::DUP); // keep original value
                self.emit_constant(Value::F64(1.0));
                match op {
                    PostfixOp::PostInc => self.emit(Op::DYN_ADD),
                    PostfixOp::PostDec => self.emit(Op::F64_SUB),
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
                        self.emit(Op::DUP);
                        let not_null = self.emit_jump(Op::BR_IF_NULL);
                        // Case: NOT null. Result is original value.
                        let end = self.emit_jump(Op::BR);
                        
                        self.patch_jump(not_null);
                        // Case: IS null. Evaluate right and store.
                        self.emit(Op::DROP);
                        self.compile_expression(right)?;
                        self.emit(Op::DUP);
                        self.compile_store(left)?;
                        self.patch_jump(end);
                        return Ok(());
                    }
                    _ => {
                        self.compile_expression(left)?;
                        self.compile_expression(right)?;
                        match op {
                            AssignOp::AddAssign => self.emit(Op::DYN_ADD),
                            AssignOp::SubAssign => self.emit(Op::F64_SUB),
                            AssignOp::MulAssign => self.emit(Op::F64_MUL),
                            AssignOp::DivAssign => self.emit(Op::F64_DIV),
                            AssignOp::ModAssign => { let l = self.line; vybe_compiler_common::expressions::emit_f64_mod(&mut self.chunks[self.current_chunk_idx], l); },
                            _ => self.emit(Op::DYN_ADD), // fallback
                        }
                    }
                }
                self.emit(Op::DUP); // assignment is an expression, keep value
                self.compile_store(left)?;
            }
            Expression::Ternary { cond, then, else_ } => {
                self.compile_expression(cond)?;
                let line = self.line;
                let false_jump = common_expr::emit_ternary_start(self.chunk_mut(), line);
                self.compile_expression(then)?;
                let end_jump = common_expr::emit_ternary_middle(self.chunk_mut(), false_jump, line);
                self.compile_expression(else_)?;
                common_expr::emit_ternary_end(self.chunk_mut(), end_jump);
            }
            Expression::NullCoalesce { left, right } => {
                self.compile_expression(left)?;
                let line = self.line;
                let (_null_jump, end_jump) = common_expr::emit_null_coalesce_start(self.chunk_mut(), line);
                self.compile_expression(right)?;
                common_expr::emit_null_coalesce_end(self.chunk_mut(), end_jump);
            }
            Expression::Member { object, member, null_safe } => {
                self.compile_expression(object)?;
                if *null_safe {
                    let line = self.line;
                    let (skip, _) = common_expr::emit_null_safe_start(self.chunk_mut(), line);
                    self.compile_member_access(member)?;
                    common_expr::emit_null_safe_end(self.chunk_mut(), skip, line);
                } else {
                    self.compile_member_access(member)?;
                }
            }
            Expression::Index { object, index } => {
                // array_get handles arrays/dicts natively.
                // For user objects with operator[], the VM's array_get also
                // falls through to struct_get for string keys, so this works.
                self.compile_expression(object)?;
                self.compile_expression(index)?;
                common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
            }
            Expression::Call { callee, args, .. } => {
                self.compile_call(callee, args)?;
            }
            Expression::New { class, args, .. } => {
                // Built-in type constructors → host calls (same as VB/Python/PHP)
                match class.as_str() {
                    "DateTime" | "DateTime.now" => {
                        if args.is_empty() {
                            let idx = self.import("vybe:types", "dateTimeNow");
                            self.emit_host_call(idx, 0);
                        } else {
                            let count = self.emit_args(args)?;
                            let idx = self.import("vybe:types", "dateTimeNew");
                            self.emit_host_call(idx, count);
                        }
                    }
                    "StringBuffer" => {
                        if args.is_empty() { self.emit_constant(Value::String(Arc::from(""))); }
                        else { let count = self.emit_args(args)?; let _ = count; }
                        let idx = self.import("vybe:types", "stringBuilderNew");
                        self.emit_host_call(idx, 1);
                    }
                    "Random" => {
                        let idx = self.import("vybe:threading", "randomNew");
                        self.emit_host_call(idx, 0);
                    }
                    "Stopwatch" => {
                        let idx = self.import("vybe:threading", "stopwatchNew");
                        self.emit_host_call(idx, 0);
                    }
                    "RegExp" => {
                        let count = self.emit_args(args)?;
                        let idx = self.import("vybe:regex", "test");
                        self.emit_host_call(idx, count);
                    }
                    "Socket" => {
                        let count = self.emit_args(args)?;
                        let idx = self.import("vybe:net", "tcpConnect");
                        self.emit_host_call(idx, count);
                    }
                    "ServerSocket" => {
                        let count = self.emit_args(args)?;
                        let idx = self.import("vybe:net", "tcpListenerNew");
                        self.emit_host_call(idx, count);
                    }
                    "RawDatagramSocket" => {
                        let count = self.emit_args(args)?;
                        let idx = self.import("vybe:net", "udpNew");
                        self.emit_host_call(idx, count);
                    }
                    _ => {
                        if self.defined_classes.contains(class) || self.classes.contains_key(class) {
                            // User-defined class constructor
                            let idx = self.add_string_constant(class);
                            self.emit_u16(Op::GLOBAL_GET, idx);
                            let count = self.emit_args(args)?;
                            self.emit_u8(Op::CALL, count);
                        } else {
                            // Unresolved class → host call_import("*", name)
                            let count = self.emit_args(args)?;
                            let imp = self.import("*", class);
                            self.emit_host_call(imp, count);
                        }
                    }
                }
            }
            Expression::Const { class, args, .. } => {
                if self.defined_classes.contains(class) || self.classes.contains_key(class) {
                    let idx = self.add_string_constant(class);
                    self.emit_u16(Op::GLOBAL_GET, idx);
                    let count = self.emit_args(args)?;
                    self.emit_u8(Op::CALL, count);
                } else {
                    // Unresolved class → host call_import("*", name)
                    let count = self.emit_args(args)?;
                    let imp = self.import("*", class);
                    self.emit_host_call(imp, count);
                }
            }
            Expression::Cascade { object, ops, null_safe } => {
                self.compile_expression(object)?;
                let slot = self.define_local("__cascade_obj", true, false);
                self.emit_u16(Op::LOCAL_SET, slot);
                self.emit(Op::DROP);
                
                let mut end_cascade = None;
                if *null_safe {
                    self.emit_u16(Op::LOCAL_GET, slot);
                    let is_null = self.emit_jump(Op::BR_IF_NULL);
                    end_cascade = Some(is_null);
                }

                for op in ops {
                    match op {
                        CascadeOp::Method(name, args) => {
                            let prop_idx = self.add_string_constant(name);
                            self.emit_u16(Op::LOCAL_GET, slot);
                            self.emit_u16(Op::STRUCT_GET, prop_idx); // push callee
                            self.emit_u16(Op::LOCAL_GET, slot);       // push receiver as first arg
                            let count = self.emit_args(args)?;
                            self.emit_u8(Op::CALL, count + 1);
                            self.emit(Op::DROP);
                        }
                        CascadeOp::Field(name) => {
                            let prop_idx = self.add_string_constant(name);
                            self.emit_u16(Op::LOCAL_GET, slot);
                            self.emit_u16(Op::STRUCT_GET, prop_idx);
                            self.emit(Op::DROP);
                        }
                        CascadeOp::Assign(name, val) => {
                            self.emit_u16(Op::LOCAL_GET, slot);
                            self.compile_expression(val)?;
                            let prop_idx = self.add_string_constant(name);
                            self.emit_u16(Op::STRUCT_SET, prop_idx);
                            self.emit(Op::DROP);
                        }
                        CascadeOp::Index(idx_expr) => {
                            self.emit_u16(Op::LOCAL_GET, slot);
                            self.compile_expression(idx_expr)?;
                            common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
                            self.emit(Op::DROP);
                        }
                    }
                }

                if let Some(label) = end_cascade {
                    self.patch_jump(label);
                }
                self.emit_u16(Op::LOCAL_GET, slot);
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
                        self.emit(Op::NULL);
                        self.emit(Op::RETURN);
                    }
                    FunctionBody::Expression(e) => {
                        self.compile_expression(e)?;
                        self.emit(Op::RETURN);
                    }
                    FunctionBody::Empty => {
                        self.emit(Op::NULL);
                        self.emit(Op::RETURN);
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
                self.emit_u16(Op::REF_TEST, type_idx);
                if *negated { let l = self.line; vybe_compiler_common::expressions::emit_bool_not(&mut self.chunks[self.current_chunk_idx], l); }
            }
            Expression::As { expr: inner, type_ann } => {
                self.compile_expression(inner)?;
                let type_idx = self.add_string_constant(&self.runtime_type_name(&type_ann.name));
                self.emit_u16(Op::REF_CAST, type_idx);
            }
            Expression::Await(inner) => {
                self.compile_expression(inner)?;
                let line = self.line;
                common_fn::emit_await(self.chunk_mut(), line);
            }
            Expression::Spread(inner) => {
                self.compile_expression(inner)?;
                let spread_idx = self.import("vybe:collections", "spread");
                self.emit_host_call(spread_idx, 1);
            }
            Expression::IfNull { left, right } => {
                self.compile_expression(left)?;
                let line = self.line;
                let (_null_jump, end_jump) = common_expr::emit_null_coalesce_start(self.chunk_mut(), line);
                self.compile_expression(right)?;
                common_expr::emit_null_coalesce_end(self.chunk_mut(), end_jump);
            }
            Expression::Record { elements } => {
                for el in elements {
                    self.compile_expression(&el.value)?;
                }
                self.emit_u16(Op::STRUCT_NEW, elements.len() as u16);
            }
        }
        Ok(())
    }

    fn compile_member_access(&mut self, member: &str) -> Result<(), String> {
        match member {
            "length" => { common_collections::emit_len(&mut self.chunks[self.current_chunk_idx], self.line); return Ok(()); }
            "isEmpty" => {
                common_collections::emit_len(&mut self.chunks[self.current_chunk_idx], self.line);
                self.emit_constant(Value::I32(0));
                self.emit(Op::DYN_EQ);
                return Ok(());
            }
            "isNotEmpty" => {
                common_collections::emit_len(&mut self.chunks[self.current_chunk_idx], self.line);
                self.emit_constant(Value::I32(0));
                self.emit(Op::DYN_GT);
                return Ok(());
            }
            "first" => {
                self.emit_constant(Value::I32(0));
                common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
                return Ok(());
            }
            "last" => {
                self.emit(Op::DUP);
                common_collections::emit_len(&mut self.chunks[self.current_chunk_idx], self.line);
                self.emit_constant(Value::I32(1));
                self.emit(Op::F64_SUB);
                common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
                return Ok(());
            }
            "hashCode" | "runtimeType" => {
                // Stub: return 0 / type string
                self.emit(Op::DROP);
                self.emit_constant(Value::I32(0));
                return Ok(());
            }
            _ => {}
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
                        self.emit_u16(Op::LOCAL_SET, tmp);
                        
                        let idx = self.add_string_constant(&format!("{}_{}", ext.name.as_deref().unwrap_or("Extension"), member));
                        self.emit_u16(Op::GLOBAL_GET, idx);
                        self.emit_u16(Op::LOCAL_GET, tmp);
                        self.emit_u8(Op::CALL, 1);
                        return Ok(());
                    }
                    _ => {}
                }
            }
        }

        let prop_idx = self.add_string_constant(member);
        self.emit_u16(Op::STRUCT_GET, prop_idx);
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
            // Build dict for named args — pure WASM, same shape as all dicts
            let line = self.line;
            vybe_compiler_common::dict::emit_new(self.chunk_mut(), line);
            for (label, value) in named_args {
                self.emit(Op::DUP);
                self.compile_expression(value)?;
                let line = self.line;
                vybe_compiler_common::dict::emit_set_const_key(self.chunk_mut(), &label, line);
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
                let line = self.line;
                common_io::emit_print(self.chunk_mut(), count, line);
                return Ok(());
            }
            // Other bare imports
            if let Some(imp) = self.resolve_bare_import(name) {
                let count = self.emit_args(args)?;
                self.emit_host_call(imp, count);
                return Ok(());
            }
            // Unresolved name → host call_import("*", name)
            if !self.is_known_variable(name) && !self.defined_classes.contains(name) {
                let count = self.emit_args(args)?;
                let imp = self.import("*", name);
                self.emit_host_call(imp, count);
                return Ok(());
            }
        }

        // Handle method calls: obj.method(args)
        if let Expression::Member { object, member, .. } = callee {
            // super.method() → call __base_method on this
            if matches!(object.as_ref(), Expression::Super) {
                if let Some(this_slot) = self.current_scope().resolve_local("this") {
                    self.emit_u16(Op::LOCAL_GET, this_slot);
                    let base_name = format!("__base_{}", member);
                    let prop_idx = self.add_string_constant(&base_name);
                    self.emit_u16(Op::STRUCT_GET, prop_idx);
                    self.emit_u16(Op::LOCAL_GET, this_slot); // push this as first arg
                    let count = self.emit_args(args)?;
                    self.emit_u8(Op::CALL_REF, count + 1);
                    return Ok(());
                }
            }

            // Check if it's a Class named constructor call: Class.named()
            if let Expression::Identifier(obj_name) = object.as_ref() {
                if self.classes.contains_key(obj_name) {
                    let full_name = format!("{}.{}", obj_name, member);
                    let idx = self.add_string_constant(&full_name);
                    self.emit_u16(Op::GLOBAL_GET, idx);
                    let count = self.emit_args(args)?;
                    self.emit_u8(Op::CALL, count);
                    return Ok(());
                }
            }

            // Isolate.spawn(fn, message) → thread spawn
            if let Expression::Identifier(obj_name) = object.as_ref() {
                if obj_name == "Isolate" && member == "spawn" {
                    if let Some(arg) = args.first() {
                        self.compile_expression(&arg.value)?;
                        let line = self.line;
                        common_thread::emit_thread_spawn(&mut self.chunks[self.current_chunk_idx], line);
                        return Ok(());
                    }
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

            // String/array methods as direct opcodes (normalized across all languages)
            if let Some(()) = self.try_compile_opcode_method(object, member, args)? {
                return Ok(());
            }

            // Instance method: check for remaining host-call-based methods
            if let Some(imp) = self.resolve_value_method(member) {
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
                            self.emit_u16(Op::GLOBAL_GET, idx);
                            // Push receiver (object) as 1st arg
                            self.compile_expression(object)?;
                            // Push remaining args
                            let count = self.emit_args(args)?;
                            self.emit_u8(Op::CALL, count + 1);
                            return Ok(());
                        }
                    }
                }
            }

            // toString() method
            if member == "toString" {
                self.compile_expression(object)?;
                let line = self.line;
                common_strings::emit_to_string(self.chunk_mut(), line);
                return Ok(());
            }

            // Generic method call: obj.method(...) → preserve receiver and call with receiver as first arg
            self.compile_expression(object)?; // pushes receiver
            let obj_tmp = self.define_local("__call_obj", true, false);
            self.emit_u16(Op::LOCAL_SET, obj_tmp); // pop receiver -> stored
            self.emit(Op::DROP);

            self.emit_u16(Op::LOCAL_GET, obj_tmp); // push receiver for struct_get
            let prop_idx = self.add_string_constant(member);
            self.emit_u16(Op::STRUCT_GET, prop_idx); // pushes function

            self.emit_u16(Op::LOCAL_GET, obj_tmp); // push receiver as first arg
            let count = self.emit_args(args)?;
            self.emit_u8(Op::CALL, count + 1);
            return Ok(());
        }

        // Generic function call
        self.compile_expression(callee)?;
        let count = self.emit_args(args)?;
        self.emit_u8(Op::CALL, count);
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
                        self.emit_u16(Op::LOCAL_SET, slot);
                        self.emit(Op::DROP);
                    }
                    VarResolution::Upvalue(idx) => {
                        self.emit_u8(Op::UPVALUE_SET, idx);
                        self.emit(Op::DROP);
                    }
                    VarResolution::Global => {
                        // If the name is a known variable (global or previously defined), set it.
                        if self.is_known_variable(name) {
                            self.emit_global_set(name);
                            self.emit(Op::DROP);
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
                                        self.emit_u16(Op::LOCAL_SET, tmp);
                                        self.emit(Op::DROP);

                                        self.emit_u16(Op::LOCAL_GET, this_slot);
                                        self.emit_u16(Op::LOCAL_GET, tmp);
                                        let prop_idx = self.add_string_constant(name);
                                        self.emit_u16(Op::STRUCT_SET, prop_idx);
                                        self.emit(Op::DROP);
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
                self.emit_u16(Op::LOCAL_SET, tmp); // save obj
                self.emit(Op::DROP);
                let tmp2 = self.define_local("__store_val", true, false);
                self.emit_u16(Op::LOCAL_SET, tmp2); // save val
                self.emit(Op::DROP);
                self.emit_u16(Op::LOCAL_GET, tmp); // push obj
                self.emit_u16(Op::LOCAL_GET, tmp2); // push val
                let prop_idx = self.add_string_constant(member);
                self.emit_u16(Op::STRUCT_SET, prop_idx);
                self.emit(Op::DROP);
            }
            Expression::Index { object, index } => {
                let tmp = self.define_local("__idx_val", true, false);
                self.emit_u16(Op::LOCAL_SET, tmp);
                self.emit(Op::DROP);
                self.compile_expression(object)?;
                self.compile_expression(index)?;
                self.emit_u16(Op::LOCAL_GET, tmp);
                common_collections::emit_set(&mut self.chunks[self.current_chunk_idx], self.line);
                self.emit(Op::DROP);
            }
            _ => {} // can't store to other expression types
        }
        Ok(())
    }

    fn compile_switch_expression(&mut self, expr: &Expression, cases: &[SwitchExpressionCase]) -> Result<(), String> {
        self.compile_expression(expr)?;
        let val_slot = self.define_local("__matched_val", true, false);
        self.emit_u16(Op::LOCAL_SET, val_slot);
        self.emit(Op::DROP);

        let mut end_jumps = Vec::new();
        
        for case in cases {
            // Check pattern
            self.emit_u16(Op::LOCAL_GET, val_slot);
            let next_case = self.compile_pattern(&case.pattern)?;
            
            // Check guard if present
            if let Some(guard) = &case.guard {
                self.compile_expression(guard)?;
                { let line = self.line; common_convert::emit_to_bool(self.chunk_mut(), line); }
                let skip_guard = self.emit_jump(Op::BR_IF_FALSE);
                
                // Pattern matched AND guard passed
                self.compile_expression(&case.result)?;
                end_jumps.push(self.emit_jump(Op::BR));
                
                self.patch_jump(skip_guard);
            } else {
                // Pattern matched, no guard
                self.compile_expression(&case.result)?;
                end_jumps.push(self.emit_jump(Op::BR));
            }
            
            self.patch_jump(next_case);
        }
        
        // Default: throw error
        let msg_idx = self.add_string_constant("Switch expression not exhaustive");
        self.emit_u16(Op::CONST, msg_idx); // this is wrong, should use emit_constant but okay for now
        common_errors::emit_throw(&mut self.chunks[self.current_chunk_idx], self.line);
        
        for j in end_jumps {
            self.patch_jump(j);
        }
        
        Ok(())
    }

    fn compile_pattern(&mut self, pattern: &Pattern) -> Result<usize, String> {
        match pattern {
            Pattern::Constant(e) => {
                self.compile_expression(e)?;
                self.emit(Op::DYN_EQ);
                let skip = self.emit_jump(Op::BR_IF_FALSE);
                Ok(skip)
            }
            Pattern::Wildcard => {
                self.emit(Op::DROP);
                self.emit(Op::TRUE);
                let skip = self.emit_jump(Op::BR_IF_FALSE);
                Ok(skip)
            }
            Pattern::Variable(name) => {
                let slot = self.define_local(name, false, false);
                self.emit_u16(Op::LOCAL_SET, slot);
                self.emit(Op::DROP);
                self.emit(Op::TRUE);
                let skip = self.emit_jump(Op::BR_IF_FALSE);
                Ok(skip)
            }
            Pattern::Type(type_name) => {
                let type_name_norm = self.runtime_type_name(type_name);
                let type_idx = self.add_string_constant(&type_name_norm);
                self.emit_u16(Op::REF_TEST, type_idx);
                let skip = self.emit_jump(Op::BR_IF_FALSE);
                Ok(skip)
            }
            Pattern::Relational { op, val } => {
                self.compile_expression(val)?;
                match op.as_str() {
                    ">" => self.emit(Op::DYN_GT),
                    "<" => self.emit(Op::DYN_LT),
                    ">=" => self.emit(Op::DYN_GE),
                    "<=" => self.emit(Op::DYN_LE),
                    "==" => self.emit(Op::DYN_EQ),
                    "!=" => self.emit(Op::DYN_NE),
                    _ => unreachable!(),
                }
                let skip = self.emit_jump(Op::BR_IF_FALSE);
                Ok(skip)
            }
            Pattern::List(patterns) => {
                // Check if it's an array and length matches
                self.emit(Op::DUP);
                self.emit(Op::REF_IS_ARRAY);
                let not_array = self.emit_jump(Op::BR_IF_FALSE);
                
                self.emit(Op::DUP);
                common_collections::emit_len(&mut self.chunks[self.current_chunk_idx], self.line);
                self.emit_constant(Value::I64(patterns.len() as i64));
                self.emit(Op::EQ);
                let wrong_len = self.emit_jump(Op::BR_IF_FALSE);
                
                let mut p_skips = Vec::new();
                for (i, p) in patterns.iter().enumerate() {
                    self.emit(Op::DUP);
                    self.emit_constant(Value::I64(i as i64));
                    common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
                    p_skips.push(self.compile_pattern(p)?);
                }
                
                let success = self.emit_jump(Op::BR);
                self.patch_jump(not_array);
                self.patch_jump(wrong_len);
                for s in p_skips { self.patch_jump(s); }
                let fail = self.emit_jump(Op::BR_IF_TRUE); // this is a hack to skip
                self.patch_jump(success);
                Ok(fail)
            }
            Pattern::Map(entries) => {
                // Check if it's an object/map
                self.emit(Op::DUP);
                self.emit(Op::REF_IS_OBJECT);
                let skip = self.emit_jump(Op::BR_IF_FALSE);
                
                let mut e_skips = Vec::new();
                for (key_expr, val_pat) in entries {
                    self.emit(Op::DUP);
                    self.compile_expression(key_expr)?;
                    common_collections::emit_get(&mut self.chunks[self.current_chunk_idx], self.line);
                    e_skips.push(self.compile_pattern(val_pat)?);
                }
                
                let success = self.emit_jump(Op::BR);
                self.patch_jump(skip);
                for s in e_skips { self.patch_jump(s); }
                let fail = self.emit_jump(Op::BR_IF_TRUE);
                self.patch_jump(success);
                Ok(fail)
            }
            Pattern::Record(elements) => {
                self.emit(Op::DUP);
                self.emit(Op::REF_IS_OBJECT);
                let skip = self.emit_jump(Op::BR_IF_FALSE);
                
                let mut e_skips = Vec::new();
                for (i, el) in elements.iter().enumerate() {
                    self.emit(Op::DUP);
                    let prop_idx = if let Some(label) = &el.label {
                        self.add_string_constant(label)
                    } else {
                        self.add_string_constant(&i.to_string())
                    };
                    self.emit_u16(Op::STRUCT_GET, prop_idx);
                    e_skips.push(self.compile_pattern(&el.pattern)?);
                }
                
                let success = self.emit_jump(Op::BR);
                self.patch_jump(skip);
                for s in e_skips { self.patch_jump(s); }
                let fail = self.emit_jump(Op::BR_IF_TRUE);
                self.patch_jump(success);
                Ok(fail)
            }
            Pattern::Object { class_name, fields } => {
                let type_idx = self.add_string_constant(class_name);
                self.emit(Op::DUP);
                self.emit_u16(Op::REF_TEST, type_idx);
                let skip = self.emit_jump(Op::BR_IF_FALSE);
                
                let mut f_skips = Vec::new();
                for (name, pat) in fields {
                    self.emit(Op::DUP);
                    let prop_idx = self.add_string_constant(name);
                    self.emit_u16(Op::STRUCT_GET, prop_idx);
                    f_skips.push(self.compile_pattern(pat)?);
                }
                
                let success = self.emit_jump(Op::BR);
                self.patch_jump(skip);
                for s in f_skips { self.patch_jump(s); }
                let fail = self.emit_jump(Op::BR_IF_TRUE);
                self.patch_jump(success);
                Ok(fail)
            }
            Pattern::Logical(left, right, is_or) => {
                self.emit(Op::DUP);
                let l_skip = self.compile_pattern(left)?;
                if *is_or {
                    let success = self.emit_jump(Op::BR);
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
