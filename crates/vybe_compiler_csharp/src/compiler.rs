//! C# to bytecode compiler.
//!
//! Follows the same patterns as the VB and JS compilers:
//! - One Chunk per function/constructor/lambda
//! - Chunk 0 is the script chunk (top-level code)
//! - Classes compile to constructor functions (struct_new + call)
//! - Same import table (vybe:gui, wasi:cli, vybe:math, etc.)
//! - Same scope/local tracking via Scope

use std::collections::{HashMap, HashSet};
use std::rc::Rc;

use vybe_bytecode::{Chunk, Op, Value};
use vybe_parser_csharp::ast::*;

// ============================================================
// Scope — same design as VB/JS compilers
// ============================================================

#[derive(Debug, Clone)]
struct Local {
    name: String,
    depth: u32,
    slot: u16,
}

#[derive(Debug, Clone)]
struct UpvalueDesc {
    index: u8,
    is_local: bool,
}

#[derive(Debug)]
struct Scope {
    locals: Vec<Local>,
    upvalues: Vec<UpvalueDesc>,
    depth: u32,
    next_slot: u16,
}

impl Scope {
    fn new() -> Self {
        Scope { locals: Vec::new(), upvalues: Vec::new(), depth: 0, next_slot: 0 }
    }

    fn new_function() -> Self {
        let mut s = Self::new();
        // Slot 0 reserved for the implicit receiver (unused in free functions)
        s.locals.push(Local { name: String::new(), depth: 0, slot: 0 });
        s.next_slot = 1;
        s
    }

    fn define_local(&mut self, name: &str) -> u16 {
        let slot = self.next_slot;
        self.locals.push(Local { name: name.to_lowercase(), depth: self.depth, slot });
        self.next_slot += 1;
        slot
    }

    fn resolve_local(&self, name: &str) -> Option<u16> {
        let lower = name.to_lowercase();
        for local in self.locals.iter().rev() {
            if local.name == lower { return Some(local.slot); }
        }
        None
    }

    fn begin_scope(&mut self) { self.depth += 1; }

    fn end_scope(&mut self) -> Vec<Local> {
        let mut popped = Vec::new();
        while let Some(local) = self.locals.last() {
            if local.depth < self.depth { break; }
            popped.push(self.locals.pop().unwrap());
        }
        self.depth -= 1;
        popped
    }

    fn add_upvalue(&mut self, index: u8, is_local: bool) -> u8 {
        for (i, uv) in self.upvalues.iter().enumerate() {
            if uv.index == index && uv.is_local == is_local { return i as u8; }
        }
        let idx = self.upvalues.len() as u8;
        self.upvalues.push(UpvalueDesc { index, is_local });
        idx
    }
}

// ============================================================
// Loop context for break/continue
// ============================================================

struct LoopContext {
    start: usize,
    break_jumps: Vec<usize>,
    continue_jumps: Vec<usize>,
}

// ============================================================
// Variable resolution
// ============================================================

#[derive(Clone, Copy)]
enum VarResolution {
    Local(u16),
    Global,
}

// ============================================================
// Compiler
// ============================================================

pub struct Compiler {
    chunks: Vec<Chunk>,
    scopes: Vec<Scope>,
    current_chunk_idx: usize,
    line: u32,
    defined_globals: HashSet<String>,
    defined_classes: HashSet<String>,
    defined_interfaces: HashSet<String>,
    function_name_stack: Vec<String>,
    loop_stack: Vec<LoopContext>,
    class_fields: HashSet<String>,
    class_methods: HashSet<String>,
    class_field_map: HashMap<String, HashSet<String>>,
    class_method_map: HashMap<String, HashSet<String>>,
    interface_imports: Vec<String>,
    known_types: HashMap<String, (&'static str, &'static str)>,
}

impl Compiler {
    pub fn new() -> Self {
        Compiler {
            chunks: vec![Chunk::new("<script>")],
            scopes: vec![Scope::new()],
            current_chunk_idx: 0,
            line: 1,
            defined_globals: HashSet::new(),
            defined_classes: HashSet::new(),
            defined_interfaces: HashSet::new(),
            function_name_stack: Vec::new(),
            loop_stack: Vec::new(),
            class_fields: HashSet::new(),
            class_methods: HashSet::new(),
            class_field_map: HashMap::new(),
            class_method_map: HashMap::new(),
            known_types: Self::init_known_types(),
            interface_imports: vec![
                "system".into(),
                "system.console".into(),
                "system.math".into(),
                "system.io".into(),
                "system.io.file".into(),
                "system.io.path".into(),
                "system.io.directory".into(),
                "system.windows.forms".into(),
                "system.collections.generic".into(),
                "system.text".into(),
                "system.drawing".into(),
                "system.net".into(),
                "system.threading".into(),
                "system.diagnostics".into(),
                "system.data".into(),
                "system.security.cryptography".into(),
                "system.xml.linq".into(),
                "system.linq".into(),
            ],
        }
    }

    // ================================================================
    // Public entry point
    // ================================================================

    pub fn compile(mut self, unit: &CompilationUnit) -> Result<Vec<Chunk>, String> {
        // Register using directives as interface imports
        for u in &unit.usings {
            let lower = u.to_lowercase();
            if !self.interface_imports.contains(&lower) {
                self.interface_imports.push(lower);
            }
        }

        // First pass: register class and interface names
        for member in &unit.members {
            match member {
                TypeDecl::Class(class) => {
                    self.defined_classes.insert(class.name.to_lowercase());
                }
                TypeDecl::Interface(iface) => {
                    self.defined_interfaces.insert(iface.name.to_lowercase());
                }
                _ => {}
            }
        }

        // Compile type declarations
        for member in &unit.members {
            self.compile_type_decl(member)?;
        }

        // Compile top-level statements (C# 9+)
        for stmt in &unit.top_level_statements {
            self.compile_statement(stmt)?;
        }

        // If there's a Main method defined, call it
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

    // ================================================================
    // Emit helpers (same as VB compiler)
    // ================================================================

    fn import(&mut self, module: &str, name: &str) -> u16 {
        self.chunks[0].add_import(module, name)
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
        self.defined_globals.insert(name.to_lowercase());
    }

    fn emit_ref_func(&mut self, func_idx: usize, upvalues: &[UpvalueDesc]) {
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit_op_u16(Op::ref_func, func_idx as u16, line);
        self.chunks[self.current_chunk_idx].emit(upvalues.len() as u8, line);
        for uv in upvalues {
            self.chunks[self.current_chunk_idx].emit(if uv.is_local { 1 } else { 0 }, line);
            self.chunks[self.current_chunk_idx].emit(uv.index, line);
        }
    }

    // ================================================================
    // Scope helpers
    // ================================================================

    fn current_scope(&self) -> &Scope { self.scopes.last().unwrap() }
    fn current_scope_mut(&mut self) -> &mut Scope { self.scopes.last_mut().unwrap() }
    fn define_local(&mut self, name: &str) -> u16 { self.current_scope_mut().define_local(name) }

    fn resolve_variable(&self, name: &str) -> VarResolution {
        let lower = name.to_lowercase();
        if let Some(slot) = self.current_scope().resolve_local(&lower) {
            return VarResolution::Local(slot);
        }
        for scope in self.scopes.iter().rev().skip(1) {
            if let Some(slot) = scope.resolve_local(&lower) {
                return VarResolution::Local(slot);
            }
        }
        VarResolution::Global
    }

    fn is_namespace(&self, name: &str) -> bool {
        matches!(name.to_lowercase().as_str(),
            "math" | "console" | "convert" | "string" | "array"
            | "file" | "io" | "directory" | "path"
            | "system" | "application" | "environment"
            | "thread" | "json" | "color" | "datetime"
            | "stringbuilder" | "process" | "timespan"
            | "guid" | "point" | "size" | "font" | "random"
            | "messagebox" | "encoding"
        )
    }

    fn is_namespace_expr(&self, expr: &Expression) -> bool {
        match expr {
            Expression::Identifier(name) => self.is_namespace(name),
            Expression::MemberAccess(inner, _) => self.is_namespace_expr(inner),
            _ => false,
        }
    }

    // ================================================================
    // Known types (same table as VB compiler)
    // ================================================================

    fn init_known_types() -> HashMap<String, (&'static str, &'static str)> {
        let mut m = HashMap::new();
        for (name, module, func) in &[
            ("list", "vybe:types", "listNew"),
            ("dictionary", "vybe:types", "dictNew"),
            ("queue", "vybe:types", "queueNew"),
            ("stack", "vybe:types", "stackNew"),
            ("hashset", "vybe:types", "hashSetNew"),
            ("arraylist", "vybe:types", "listNew"),
            ("hashtable", "vybe:types", "dictNew"),
            ("collection", "vybe:types", "listNew"),
            ("sortedlist", "vybe:types", "dictNew"),
            ("datetime", "vybe:types", "dateTimeNew"),
            ("stringbuilder", "vybe:types", "stringBuilderNew"),
            ("datatable", "vybe:data", "dataTableNew"),
            ("dataset", "vybe:data", "dataSetNew"),
            ("point", "vybe:drawing", "pointNew"),
            ("size", "vybe:drawing", "sizeNew"),
            ("sizef", "vybe:drawing", "sizeNew"),
            ("font", "vybe:drawing", "fontNew"),
            ("random", "vybe:threading", "randomNew"),
            ("stopwatch", "vybe:threading", "stopwatchNew"),
            ("sqlconnection", "vybe:database", "connect"),
            ("tcpclient", "vybe:net", "tcpConnect"),
            ("tcplistener", "vybe:net", "tcpListenerNew"),
            ("udpclient", "vybe:net", "udpNew"),
            ("streamreader", "vybe:net", "streamReaderNew"),
            ("streamwriter", "vybe:net", "streamWriterNew"),
            ("form", "vybe:gui", "newForm"),
        ] {
            m.insert(name.to_string(), (*module, *func));
        }
        m
    }

    // ================================================================
    // Interface resolution (Component Model)
    // ================================================================

    /// Resolve a dotted C# name to a (module, function) host import.
    /// e.g. "Console.WriteLine" → ("wasi:cli", "log")
    /// e.g. "Math.Floor" → ("vybe:math", "floor")
    fn resolve_interface_call(&self, parts: &[&str]) -> Option<(String, String)> {
        let lower_parts: Vec<String> = parts.iter().map(|p| p.to_lowercase()).collect();

        for prefix_len in (1..lower_parts.len()).rev() {
            let prefix = lower_parts[..prefix_len].join(".");
            if self.interface_imports.contains(&prefix) {
                let func = lower_parts[prefix_len..].join(".");
                let module = match prefix.as_str() {
                    "system.console" => "wasi:cli",
                    "system.math" => "vybe:math",
                    "system.io.file" => "wasi:filesystem",
                    "system.io.path" => "wasi:filesystem",
                    "system.io.directory" => "wasi:filesystem",
                    "system.io" => "wasi:filesystem",
                    "system.convert" => "vybe:convert",
                    "system.string" => "vybe:string",
                    "system.array" => "vybe:array",
                    "system.environment" => "wasi:cli",
                    "system.threading.thread" => "wasi:clocks",
                    "system.threading" => "vybe:threading",
                    "system.diagnostics" => "wasi:cli",
                    "system.net" => "wasi:http",
                    "system.net.sockets" => "vybe:net",
                    "system.text.regularexpressions" => "vybe:regex",
                    "system.text" => "vybe:string",
                    "system.collections.generic" => "vybe:types",
                    "system.data" => "vybe:data",
                    "system.security.cryptography" => "vybe:crypto",
                    "system.xml.linq" => "vybe:xml",
                    "system.drawing" => "vybe:drawing",
                    "system.windows.forms" => "vybe:gui",
                    _ => &prefix,
                };
                let mapped_func = map_interface_func(module, &func);
                return Some((module.to_string(), mapped_func));
            }
        }
        None
    }

    // ================================================================
    // Type declarations
    // ================================================================

    fn compile_type_decl(&mut self, decl: &TypeDecl) -> Result<(), String> {
        match decl {
            TypeDecl::Class(class) => self.compile_class(class),
            TypeDecl::Enum(e) => self.compile_enum(e),
            TypeDecl::Struct(s) => {
                // Compile structs like classes
                let class = ClassDecl {
                    name: s.name.clone(),
                    is_partial: false,
                    is_static: false,
                    is_abstract: false,
                    base_type: None,
                    interfaces: vec![],
                    members: s.members.clone(),
                };
                self.compile_class(&class)
            }
            TypeDecl::Interface(iface) => self.compile_interface(iface),
        }
    }

    fn compile_enum(&mut self, e: &EnumDecl) -> Result<(), String> {
        // Compile enum as an object with integer fields
        let mut val = 0i64;
        self.emit_u16(Op::struct_new, 0);
        for (name, explicit_val) in &e.members {
            if let Some(expr) = explicit_val {
                // Evaluate constant expression
                if let Expression::IntLiteral(n) = expr {
                    val = *n;
                }
            }
            self.emit(Op::dup);
            self.emit_constant(Value::I32(val as i32));
            let idx = self.add_string_constant(&name.to_lowercase());
            self.emit_u16(Op::struct_set, idx);
            self.emit(Op::drop);
            val += 1;
        }
        self.emit_global_set(&e.name.to_lowercase());
        self.emit(Op::drop);
        Ok(())
    }

    fn compile_interface(&mut self, iface: &InterfaceDecl) -> Result<(), String> {
        // Compile interface as an empty marker object stored as a global.
        // Duck typing means any class with the right methods "implements" it.
        self.emit_u16(Op::struct_new, 0);
        // Store the interface name so it can be referenced
        let name_idx = self.add_string_constant("__interface_name");
        self.emit(Op::dup);
        self.emit_constant(Value::String(Rc::from(iface.name.as_str())));
        self.emit_u16(Op::struct_set, name_idx);
        self.emit(Op::drop);
        self.emit_global_set(&iface.name.to_lowercase());
        self.emit(Op::drop);
        Ok(())
    }

    // ================================================================
    // Class compilation (same pattern as VB)
    // ================================================================

    fn compile_class(&mut self, class: &ClassDecl) -> Result<(), String> {
        let name = &class.name;

        // Collect constructor, instance methods, static methods
        let mut ctor: Option<&ConstructorDecl> = None;
        let mut instance_methods: Vec<&MethodDecl> = Vec::new();
        let mut static_methods: Vec<&MethodDecl> = Vec::new();
        let mut fields: Vec<(&String, &Option<String>, &Option<Expression>, bool)> = Vec::new();
        let mut properties: Vec<&PropertyDecl> = Vec::new();

        for member in &class.members {
            match member {
                MemberDecl::Constructor(c) => { ctor = Some(c); }
                MemberDecl::Method(m) => {
                    if m.is_static || class.is_static {
                        static_methods.push(m);
                    } else {
                        instance_methods.push(m);
                    }
                }
                MemberDecl::Field { name: fname, type_name, initializer, is_static, .. } => {
                    fields.push((fname, type_name, initializer, *is_static));
                }
                MemberDecl::Property(p) => {
                    properties.push(p);
                }
                MemberDecl::Event { .. } => {}
            }
        }

        // Check if this class contains Main (static entry point)
        let has_main = static_methods.iter().any(|m| m.name.eq_ignore_ascii_case("Main"));

        // If purely static class, compile as a namespace object
        if class.is_static || (ctor.is_none() && instance_methods.is_empty() && !fields.iter().any(|(_, _, _, is_static)| !is_static)) {
            self.compile_static_class(name, &static_methods, &fields, &properties)?;
            if has_main {
                // Register Main as a global
                let main_idx = self.add_string_constant("main");
                let cls_idx = self.add_string_constant(&name.to_lowercase());
                self.emit_u16(Op::global_get, cls_idx);
                let prop_idx = self.add_string_constant("main");
                self.emit_u16(Op::struct_get, prop_idx);
                self.emit_u16(Op::global_set, main_idx);
                self.defined_globals.insert("main".into());
                self.emit(Op::drop);
            }
            return Ok(());
        }

        // --- Compile the constructor chunk ---
        let ctor_params: Vec<&Parameter> = ctor.map(|c| c.params.iter().collect()).unwrap_or_default();
        let ctor_body: Vec<&Statement> = ctor.map(|c| c.body.iter().collect()).unwrap_or_default();
        let base_args: Option<&Vec<Expression>> = ctor.and_then(|c| c.base_args.as_ref());

        let mut chunk = Chunk::new(name.as_str());
        chunk.arity = (1 + ctor_params.len()) as u8; // this + params
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        scope.define_local("this");
        for param in &ctor_params {
            scope.define_local(&param.name);
        }

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        let this_slot = self.current_scope().resolve_local("this").unwrap();

        // Track class fields/methods
        let saved_fields = std::mem::take(&mut self.class_fields);
        let saved_methods = std::mem::take(&mut self.class_methods);
        for (fname, _, _, _) in &fields {
            self.class_fields.insert(fname.to_lowercase());
        }
        // Auto-properties are treated as fields (direct struct_get/struct_set)
        for p in &properties {
            if p.is_auto {
                self.class_fields.insert(p.name.to_lowercase());
            }
        }
        // Inherit parent fields/methods
        if let Some(ref parent) = class.base_type {
            let parent_lower = parent.to_lowercase();
            if let Some(pf) = self.class_field_map.get(&parent_lower) {
                for f in pf { self.class_fields.insert(f.clone()); }
            }
            if let Some(pm) = self.class_method_map.get(&parent_lower) {
                for m in pm { self.class_methods.insert(m.clone()); }
            }
        }
        for m in &instance_methods {
            self.class_methods.insert(m.name.to_lowercase());
        }
        for m in &static_methods {
            self.class_methods.insert(m.name.to_lowercase());
        }

        // Call base constructor if Inherits
        if let Some(ref parent) = class.base_type {
            let parent_lower = parent.to_lowercase();
            let is_framework = parent_lower.starts_with("system.")
                || parent_lower.contains("windows.forms")
                || matches!(parent_lower.as_str(),
                    "form" | "control" | "usercontrol" | "panel" | "component"
                    | "object" | "eventargs" | "exception"
                );
            let is_interface = self.defined_interfaces.contains(&parent_lower);
            if !parent_lower.is_empty() && !is_framework && !is_interface {
                // Store __super
                self.emit_u16(Op::local_get, this_slot);
                let parent_idx = self.add_string_constant(&parent_lower);
                self.emit_u16(Op::global_get, parent_idx);
                let super_idx = self.add_string_constant("__super");
                self.emit_u16(Op::struct_set, super_idx);
                self.emit(Op::drop);

                // Call base constructor with args
                let parent_idx = self.add_string_constant(&parent_lower);
                self.emit_u16(Op::global_get, parent_idx);
                self.emit_u16(Op::local_get, this_slot);
                if let Some(args) = base_args {
                    for arg in args { self.compile_expression(arg)?; }
                    self.emit_u8(Op::call, (args.len() + 1) as u8);
                } else {
                    self.emit_u8(Op::call, 1);
                }
                self.emit(Op::drop);
            }
        }

        // Initialize instance fields
        for (fname, _, initializer, is_static) in &fields {
            if *is_static { continue; }
            self.emit_u16(Op::local_get, this_slot);
            if let Some(ref init) = initializer {
                self.compile_expression(init)?;
            } else {
                self.emit(Op::null);
            }
            let prop_idx = self.add_string_constant(&fname.to_lowercase());
            self.emit_u16(Op::struct_set, prop_idx);
            self.emit(Op::drop);
        }

        // Save base methods for override (skip for interface implementations)
        let base_is_class = class.base_type.as_ref().map(|b| !self.defined_interfaces.contains(&b.to_lowercase())).unwrap_or(false);
        if base_is_class {
            for method in &instance_methods {
                let base_name = format!("__base_{}", method.name.to_lowercase());
                self.emit_u16(Op::local_get, this_slot);
                self.emit_u16(Op::local_get, this_slot);
                let prop_idx = self.add_string_constant(&method.name.to_lowercase());
                self.emit_u16(Op::struct_get, prop_idx);
                let base_idx = self.add_string_constant(&base_name);
                self.emit_u16(Op::struct_set, base_idx);
                self.emit(Op::drop);
            }
        }

        // Initialize auto-properties to null BEFORE constructor body
        for prop in &properties {
            if prop.is_auto {
                let prop_name = prop.name.to_lowercase();
                self.emit_u16(Op::local_get, this_slot);
                self.emit(Op::null);
                let pidx = self.add_string_constant(&prop_name);
                self.emit_u16(Op::struct_set, pidx);
                self.emit(Op::drop);
            }
        }

        // Attach instance methods BEFORE constructor body
        for method in &instance_methods {
            self.emit_u16(Op::local_get, this_slot);
            self.compile_instance_method(method)?;
            let prop_idx = self.add_string_constant(&method.name.to_lowercase());
            self.emit_u16(Op::struct_set, prop_idx);
            self.emit(Op::drop);
        }

        // Compile constructor body
        for stmt in &ctor_body {
            self.compile_statement(stmt)?;
        }

        // Attach property getters/setters (non-auto only)
        for prop in &properties {
            let prop_name = prop.name.to_lowercase();
            if prop.is_auto {
                continue;
            }
            if let Some(ref getter_body) = prop.getter {
                self.emit_u16(Op::local_get, this_slot);
                self.compile_property_getter(&prop_name, getter_body)?;
                let get_name = format!("__get_{}", prop_name);
                let pidx = self.add_string_constant(&get_name);
                self.emit_u16(Op::struct_set, pidx);
                self.emit(Op::drop);
            }
            if let Some((ref value_param, ref setter_body)) = prop.setter {
                self.emit_u16(Op::local_get, this_slot);
                self.compile_property_setter(&prop_name, value_param, setter_body)?;
                let set_name = format!("__set_{}", prop_name);
                let pidx = self.add_string_constant(&set_name);
                self.emit_u16(Op::struct_set, pidx);
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

        // Store field/method sets for inheritance
        self.class_field_map.insert(name.to_lowercase(), self.class_fields.clone());
        self.class_method_map.insert(name.to_lowercase(), self.class_methods.clone());
        self.class_fields = saved_fields;
        self.class_methods = saved_methods;

        self.emit_ref_func(idx, &upvalues);

        // Inherit parent's static members (skip for interface implementations)
        if let Some(ref parent) = class.base_type {
            let parent_lower = parent.to_lowercase();
            let is_fw = parent_lower.starts_with("system.")
                || matches!(parent_lower.as_str(), "form" | "control" | "usercontrol" | "panel" | "component" | "object" | "exception");
            let is_iface = self.defined_interfaces.contains(&parent_lower);
            if !parent_lower.is_empty() && !is_fw && !is_iface {
                self.emit(Op::dup);
                let parent_idx = self.add_string_constant(&parent_lower);
                self.emit_u16(Op::global_get, parent_idx);
                let assign_idx = self.import("vybe:object", "assign");
                self.emit_host_call(assign_idx, 2);
                self.emit(Op::drop);
            }
        }

        // Attach static methods to the constructor function
        for method in &static_methods {
            self.emit(Op::dup);
            self.compile_static_method(method)?;
            let prop_idx = self.add_string_constant(&method.name.to_lowercase());
            self.emit_u16(Op::struct_set, prop_idx);
            self.emit(Op::drop);
        }

        // Store as global
        self.emit_global_set(&name.to_lowercase());
        self.emit(Op::drop);
        self.defined_classes.insert(name.to_lowercase());

        Ok(())
    }

    /// Compile a purely static class (like `static class Program`).
    fn compile_static_class(
        &mut self,
        name: &str,
        static_methods: &[&MethodDecl],
        fields: &[(&String, &Option<String>, &Option<Expression>, bool)],
        _properties: &[&PropertyDecl],
    ) -> Result<(), String> {
        self.emit_u16(Op::struct_new, 0);

        // Static fields
        for (fname, _, initializer, _) in fields {
            self.emit(Op::dup);
            if let Some(ref init) = initializer {
                self.compile_expression(init)?;
            } else {
                self.emit(Op::null);
            }
            let pidx = self.add_string_constant(&fname.to_lowercase());
            self.emit_u16(Op::struct_set, pidx);
            self.emit(Op::drop);
        }

        // Static methods — attach to class object AND register as globals for bare calls
        for method in static_methods {
            self.emit(Op::dup);
            self.compile_static_method(method)?;
            let mname = method.name.to_lowercase();
            // Attach to class object
            let pidx = self.add_string_constant(&mname);
            self.emit_u16(Op::struct_set, pidx);
            self.emit(Op::drop);
            // Also store as global for bare calls from other static methods
            self.emit(Op::dup);
            let pidx2 = self.add_string_constant(&mname);
            self.emit_u16(Op::struct_get, pidx2);
            self.emit_global_set(&mname);
            self.emit(Op::drop);
        }

        self.emit_global_set(&name.to_lowercase());
        self.emit(Op::drop);
        self.defined_classes.insert(name.to_lowercase());
        Ok(())
    }

    /// Compile an instance method as a closure (this as first param).
    fn compile_instance_method(&mut self, method: &MethodDecl) -> Result<(), String> {
        let mut chunk = Chunk::new(method.name.as_str());
        chunk.arity = (method.params.len() + 1) as u8; // this + params
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        scope.define_local("this");
        for param in &method.params {
            scope.define_local(&param.name);
        }

        let has_return = method.return_type.is_some()
            && method.return_type.as_deref() != Some("void");

        if has_return {
            scope.define_local("__return_val");
        }

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        if has_return {
            self.function_name_stack.push(method.name.to_lowercase());
            self.emit(Op::null);
            let rv_slot = self.current_scope().resolve_local("__return_val").unwrap();
            self.emit_u16(Op::local_set, rv_slot);
            self.emit(Op::drop);
        }

        for stmt in &method.body {
            self.compile_statement(stmt)?;
        }

        if has_return {
            self.function_name_stack.pop();
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

    /// Compile a static method (no this parameter).
    fn compile_static_method(&mut self, method: &MethodDecl) -> Result<(), String> {
        let mut chunk = Chunk::new(method.name.as_str());
        chunk.arity = method.params.len() as u8;
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        for param in &method.params {
            scope.define_local(&param.name);
        }

        let has_return = method.return_type.is_some()
            && method.return_type.as_deref() != Some("void");

        if has_return {
            scope.define_local("__return_val");
        }

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        if has_return {
            self.function_name_stack.push(method.name.to_lowercase());
            self.emit(Op::null);
            let rv_slot = self.current_scope().resolve_local("__return_val").unwrap();
            self.emit_u16(Op::local_set, rv_slot);
            self.emit(Op::drop);
        }

        for stmt in &method.body {
            self.compile_statement(stmt)?;
        }

        if has_return {
            self.function_name_stack.pop();
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

    /// Compile property getter body as closure (this as param).
    fn compile_property_getter(&mut self, prop_name: &str, body: &[Statement]) -> Result<(), String> {
        let label = format!("get_{}", prop_name);
        let mut chunk = Chunk::new(label.as_str());
        chunk.arity = 1; // this
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        scope.define_local("this");
        scope.define_local("__return_val");

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        self.emit(Op::null);
        let rv_slot = self.current_scope().resolve_local("__return_val").unwrap();
        self.emit_u16(Op::local_set, rv_slot);
        self.emit(Op::drop);

        self.function_name_stack.push(prop_name.to_string());
        for stmt in body { self.compile_statement(stmt)?; }
        self.function_name_stack.pop();

        let rv_slot = self.current_scope().resolve_local("__return_val").unwrap();
        self.emit_u16(Op::local_get, rv_slot);
        self.emit(Op::r#return);

        let lc = self.current_scope().next_slot;
        self.chunks[idx].local_count = lc;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved;
        self.emit_ref_func(idx, &upvalues);
        Ok(())
    }

    /// Compile property setter body as closure (this + value).
    fn compile_property_setter(&mut self, prop_name: &str, value_param: &str, body: &[Statement]) -> Result<(), String> {
        let label = format!("set_{}", prop_name);
        let mut chunk = Chunk::new(label.as_str());
        chunk.arity = 2; // this + value
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        scope.define_local("this");
        scope.define_local(&value_param.to_lowercase());

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        for stmt in body { self.compile_statement(stmt)?; }
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

    // ================================================================
    // Statements
    // ================================================================

    fn compile_statement(&mut self, stmt: &Statement) -> Result<(), String> {
        match stmt {
            Statement::LocalDecl { name, initializer, .. } => {
                let slot = self.define_local(name);
                if let Some(ref init) = initializer {
                    self.compile_expression(init)?;
                } else {
                    self.emit(Op::null);
                }
                self.emit_u16(Op::local_set, slot);
                self.emit(Op::drop);
            }

            Statement::Assignment { target, value } => {
                self.compile_assignment(target, value)?;
            }

            Statement::CompoundAssignment { target, op, value } => {
                self.compile_compound_assignment(target, op, value)?;
            }

            Statement::If { condition, then_body, else_if, else_body } => {
                self.compile_expression(condition)?;
                self.emit(Op::dyn_to_bool);
                let else_j = self.emit_jump(Op::br_if_false);

                for s in then_body { self.compile_statement(s)?; }

                if else_if.is_empty() && else_body.is_none() {
                    self.patch_jump(else_j);
                } else {
                    // Collect all jumps that need to go to the very end
                    let mut end_jumps: Vec<usize> = Vec::new();

                    // Jump from the then-branch to end
                    end_jumps.push(self.emit_jump(Op::br));
                    self.patch_jump(else_j);

                    for (cond, body) in else_if {
                        self.compile_expression(cond)?;
                        self.emit(Op::dyn_to_bool);
                        let next_j = self.emit_jump(Op::br_if_false);
                        for s in body { self.compile_statement(s)?; }
                        // Jump from this branch to end
                        end_jumps.push(self.emit_jump(Op::br));
                        self.patch_jump(next_j);
                    }

                    if let Some(ref else_stmts) = else_body {
                        for s in else_stmts { self.compile_statement(s)?; }
                    }

                    // Patch all end jumps to here
                    for j in &end_jumps { self.patch_jump(*j); }
                }
            }

            Statement::For { init, condition, update, body } => {
                self.current_scope_mut().begin_scope();

                if let Some(ref init_stmt) = init {
                    self.compile_statement(init_stmt)?;
                }

                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext {
                    start: loop_start,
                    break_jumps: Vec::new(),
                    continue_jumps: Vec::new(),
                });

                let exit_j = if let Some(ref cond) = condition {
                    self.compile_expression(cond)?;
                    self.emit(Op::dyn_to_bool);
                    Some(self.emit_jump(Op::br_if_false))
                } else {
                    None
                };

                for s in body { self.compile_statement(s)?; }

                // Patch continue jumps to here (before update)
                let ctx = self.loop_stack.last().unwrap();
                let continue_jumps: Vec<usize> = ctx.continue_jumps.clone();
                for j in &continue_jumps { self.patch_jump(*j); }

                if let Some(ref upd) = update {
                    self.compile_statement(upd)?;
                }

                self.emit_loop(loop_start);

                if let Some(ej) = exit_j { self.patch_jump(ej); }

                let ctx = self.loop_stack.pop().unwrap();
                for j in &ctx.break_jumps { self.patch_jump(*j); }

                let _popped = self.current_scope_mut().end_scope();
            }

            Statement::ForEach { var_name, iterable, body } => {
                // Desugar: foreach (var x in arr) → { var __arr = arr; var __i = 0; while (__i < __arr.length) { var x = __arr[__i]; body; __i++; } }
                self.current_scope_mut().begin_scope();

                self.compile_expression(iterable)?;
                let arr_slot = self.define_local("__foreach_arr");
                self.emit_u16(Op::local_set, arr_slot);
                self.emit(Op::drop);

                self.emit_constant(Value::F64(0.0));
                let i_slot = self.define_local("__foreach_i");
                self.emit_u16(Op::local_set, i_slot);
                self.emit(Op::drop);

                let var_slot = self.define_local(var_name);

                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext {
                    start: loop_start,
                    break_jumps: Vec::new(),
                    continue_jumps: Vec::new(),
                });

                // i < arr.length
                self.emit_u16(Op::local_get, i_slot);
                self.emit_u16(Op::local_get, arr_slot);
                self.emit(Op::str_length);
                self.emit(Op::dyn_lt);
                let exit_j = self.emit_jump(Op::br_if_false);

                // var x = arr[i]
                self.emit_u16(Op::local_get, arr_slot);
                self.emit_u16(Op::local_get, i_slot);
                self.emit(Op::array_get);
                self.emit_u16(Op::local_set, var_slot);
                self.emit(Op::drop);

                for s in body { self.compile_statement(s)?; }

                let ctx = self.loop_stack.last().unwrap();
                let continue_jumps: Vec<usize> = ctx.continue_jumps.clone();
                for j in &continue_jumps { self.patch_jump(*j); }

                // i++
                self.emit_u16(Op::local_get, i_slot);
                self.emit_constant(Value::F64(1.0));
                self.emit(Op::dyn_add);
                self.emit_u16(Op::local_set, i_slot);
                self.emit(Op::drop);

                self.emit_loop(loop_start);
                self.patch_jump(exit_j);

                let ctx = self.loop_stack.pop().unwrap();
                for j in &ctx.break_jumps { self.patch_jump(*j); }

                let _popped = self.current_scope_mut().end_scope();
            }

            Statement::While { condition, body } => {
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext {
                    start: loop_start,
                    break_jumps: Vec::new(),
                    continue_jumps: Vec::new(),
                });

                self.compile_expression(condition)?;
                self.emit(Op::dyn_to_bool);
                let exit_j = self.emit_jump(Op::br_if_false);

                for s in body { self.compile_statement(s)?; }

                let ctx = self.loop_stack.last().unwrap();
                let continue_jumps: Vec<usize> = ctx.continue_jumps.clone();
                for j in &continue_jumps { self.patch_jump(*j); }

                self.emit_loop(loop_start);
                self.patch_jump(exit_j);

                let ctx = self.loop_stack.pop().unwrap();
                for j in &ctx.break_jumps { self.patch_jump(*j); }
            }

            Statement::DoWhile { body, condition } => {
                let loop_start = self.current_offset();
                self.loop_stack.push(LoopContext {
                    start: loop_start,
                    break_jumps: Vec::new(),
                    continue_jumps: Vec::new(),
                });

                for s in body { self.compile_statement(s)?; }

                let ctx = self.loop_stack.last().unwrap();
                let continue_jumps: Vec<usize> = ctx.continue_jumps.clone();
                for j in &continue_jumps { self.patch_jump(*j); }

                self.compile_expression(condition)?;
                self.emit(Op::dyn_to_bool);
                let exit = self.emit_jump(Op::br_if_false);
                self.emit_loop(loop_start);
                self.patch_jump(exit);

                let ctx = self.loop_stack.pop().unwrap();
                for j in &ctx.break_jumps { self.patch_jump(*j); }
            }

            Statement::Switch { expr, cases } => {
                self.compile_switch(expr, cases)?;
            }

            Statement::Return(expr) => {
                if let Some(ref e) = expr {
                    self.compile_expression(e)?;
                } else {
                    self.emit(Op::null);
                }
                self.emit(Op::r#return);
            }

            Statement::Break => {
                let j = self.emit_jump(Op::br);
                if let Some(ctx) = self.loop_stack.last_mut() {
                    ctx.break_jumps.push(j);
                }
            }

            Statement::Continue => {
                let j = self.emit_jump(Op::br);
                if let Some(ctx) = self.loop_stack.last_mut() {
                    ctx.continue_jumps.push(j);
                }
            }

            Statement::Throw(expr) => {
                self.compile_expression(expr)?;
                self.emit(Op::throw);
            }

            Statement::TryCatchFinally { try_body, catches, finally_body } => {
                self.compile_try_catch(try_body, catches, finally_body)?;
            }

            Statement::Expression(expr) => {
                self.compile_expression(expr)?;
                self.emit(Op::drop);
            }

            Statement::Using { var_name, initializer, body } => {
                self.current_scope_mut().begin_scope();
                let slot = self.define_local(var_name);
                self.compile_expression(initializer)?;
                self.emit_u16(Op::local_set, slot);
                self.emit(Op::drop);
                for s in body { self.compile_statement(s)?; }
                // In a real impl we'd call Dispose() here — simplified
                let _popped = self.current_scope_mut().end_scope();
            }

            Statement::Block(stmts) => {
                self.current_scope_mut().begin_scope();
                for s in stmts { self.compile_statement(s)?; }
                let _popped = self.current_scope_mut().end_scope();
            }

            Statement::Empty => {}
        }
        Ok(())
    }

    fn compile_assignment(&mut self, target: &Expression, value: &Expression) -> Result<(), String> {
        match target {
            Expression::Identifier(name) => {
                let lower = name.to_lowercase();
                // Class field → this.field = value
                if self.class_fields.contains(&lower) {
                    if let VarResolution::Local(this_slot) = self.resolve_variable("this") {
                        self.emit_u16(Op::local_get, this_slot);
                        self.compile_expression(value)?;
                        let idx = self.add_string_constant(&lower);
                        self.emit_u16(Op::struct_set, idx);
                        self.emit(Op::drop);
                        return Ok(());
                    }
                }
                self.compile_expression(value)?;
                match self.resolve_variable(name) {
                    VarResolution::Local(slot) => {
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    }
                    VarResolution::Global => {
                        let idx = self.add_string_constant(&lower);
                        self.emit_u16(Op::global_set, idx);
                        self.emit(Op::drop);
                    }
                }
            }
            Expression::MemberAccess(obj, prop) => {
                self.compile_expression(obj)?;
                self.compile_expression(value)?;
                let idx = self.add_string_constant(&prop.to_lowercase());
                self.emit_u16(Op::struct_set, idx);
                self.emit(Op::drop);
            }
            Expression::Index(obj, index) => {
                self.compile_expression(obj)?;
                self.compile_expression(index)?;
                self.compile_expression(value)?;
                self.emit(Op::array_set);
                self.emit(Op::drop);
            }
            Expression::This => {
                // this.field = value — but target is just `this`, unusual
                self.compile_expression(value)?;
                self.emit(Op::drop);
            }
            _ => {
                return Err(format!("Cannot assign to {:?}", target));
            }
        }
        Ok(())
    }

    fn compile_compound_assignment(&mut self, target: &Expression, op: &CompoundOp, value: &Expression) -> Result<(), String> {
        // Event subscription: btn.Click += handler → onEvent("btn", "Click", handler)
        if *op == CompoundOp::AddAssign {
            if let Expression::MemberAccess(obj, event_name) = target {
                let ev_lower = event_name.to_lowercase();
                if matches!(ev_lower.as_str(),
                    "click" | "textchanged" | "selectedindexchanged" | "checkedchanged"
                    | "load" | "formclosing" | "formclosed" | "resize" | "paint"
                    | "keydown" | "keyup" | "keypress" | "mousedown" | "mouseup"
                    | "mousemove" | "mouseclick" | "doubleclick" | "enter" | "leave"
                    | "validating" | "validated" | "tick" | "valuechanged"
                ) {
                    // Get the control name
                    self.compile_expression(obj)?;
                    let name_idx = self.add_string_constant("name");
                    self.emit_u16(Op::struct_get, name_idx);
                    // Event name
                    self.emit_constant(Value::String(Rc::from(event_name.as_str())));
                    // Handler
                    self.compile_expression(value)?;
                    let idx = self.import("vybe:gui", "onEvent");
                    self.emit_host_call(idx, 3);
                    self.emit(Op::drop);
                    return Ok(());
                }
            }
        }

        // Desugar: target op= value → target = target op value
        match target {
            Expression::Identifier(name) => {
                self.compile_expression(target)?;
                self.compile_expression(value)?;
                self.emit_compound_op(op);
                match self.resolve_variable(name) {
                    VarResolution::Local(slot) => {
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                    }
                    VarResolution::Global => {
                        let idx = self.add_string_constant(&name.to_lowercase());
                        self.emit_u16(Op::global_set, idx);
                        self.emit(Op::drop);
                    }
                }
            }
            Expression::MemberAccess(obj, prop) => {
                self.compile_expression(obj)?;
                self.emit(Op::dup);
                let pidx = self.add_string_constant(&prop.to_lowercase());
                self.emit_u16(Op::struct_get, pidx);
                self.compile_expression(value)?;
                self.emit_compound_op(op);
                let pidx2 = self.add_string_constant(&prop.to_lowercase());
                self.emit_u16(Op::struct_set, pidx2);
                self.emit(Op::drop);
            }
            Expression::Index(obj, index) => {
                self.compile_expression(obj)?;
                self.compile_expression(index)?;
                // Duplicate obj+index for the get
                // Simplified: use dyn ops
                self.compile_expression(target)?;
                self.compile_expression(value)?;
                self.emit_compound_op(op);
                self.emit(Op::array_set);
                self.emit(Op::drop);
            }
            _ => {
                return Err(format!("Cannot compound-assign to {:?}", target));
            }
        }
        Ok(())
    }

    fn emit_compound_op(&mut self, op: &CompoundOp) {
        match op {
            CompoundOp::AddAssign => self.emit(Op::dyn_add),
            CompoundOp::SubAssign => self.emit(Op::f64_sub),
            CompoundOp::MulAssign => self.emit(Op::f64_mul),
            CompoundOp::DivAssign => self.emit(Op::f64_div),
            CompoundOp::ModAssign => self.emit(Op::f64_mod),
            CompoundOp::AndAssign => self.emit(Op::i32_and),
            CompoundOp::OrAssign => self.emit(Op::i32_or),
            CompoundOp::XorAssign => self.emit(Op::i32_xor),
            CompoundOp::ShlAssign => self.emit(Op::i32_shl),
            CompoundOp::ShrAssign => self.emit(Op::i32_shr_s),
        }
    }

    fn compile_switch(&mut self, expr: &Expression, cases: &[SwitchCase]) -> Result<(), String> {
        self.compile_expression(expr)?;
        let switch_slot = self.define_local("__switch_val");
        self.emit_u16(Op::local_set, switch_slot);
        self.emit(Op::drop);

        // Push loop context so break statements work inside switch
        self.loop_stack.push(LoopContext {
            start: 0,
            break_jumps: Vec::new(),
            continue_jumps: Vec::new(),
        });

        let mut end_jumps: Vec<usize> = Vec::new();

        for case in cases {
            let mut is_default = false;
            let mut case_jumps: Vec<usize> = Vec::new();

            for label in &case.labels {
                match label {
                    Some(label_expr) => {
                        self.emit_u16(Op::local_get, switch_slot);
                        self.compile_expression(label_expr)?;
                        self.emit(Op::eq);
                        case_jumps.push(self.emit_jump(Op::br_if_true));
                    }
                    None => {
                        is_default = true;
                    }
                }
            }

            if !is_default && !case_jumps.is_empty() {
                let skip_j = self.emit_jump(Op::br);
                for j in &case_jumps { self.patch_jump(*j); }
                for s in &case.body { self.compile_statement(s)?; }
                end_jumps.push(self.emit_jump(Op::br));
                self.patch_jump(skip_j);
            } else {
                // Default case — always matches
                for j in &case_jumps { self.patch_jump(*j); }
                for s in &case.body { self.compile_statement(s)?; }
                end_jumps.push(self.emit_jump(Op::br));
            }
        }

        let ctx = self.loop_stack.pop().unwrap();
        for j in &ctx.break_jumps { self.patch_jump(*j); }
        for j in &end_jumps { self.patch_jump(*j); }
        Ok(())
    }

    fn compile_try_catch(
        &mut self,
        try_body: &[Statement],
        catches: &[CatchClause],
        finally_body: &Option<Vec<Statement>>,
    ) -> Result<(), String> {
        // Emit try_start with placeholder offsets
        let try_start_offset = self.current_offset();
        let catch_offset = if !catches.is_empty() { 0xFFFF_u16 } else { 0 };
        let finally_offset = if finally_body.is_some() { 0xFFFF_u16 } else { 0 };
        self.emit_u16(Op::try_start, catch_offset);
        // Emit finally offset
        let line = self.line;
        self.chunks[self.current_chunk_idx].emit((finally_offset >> 8) as u8, line);
        self.chunks[self.current_chunk_idx].emit((finally_offset & 0xff) as u8, line);

        // Try body
        for s in try_body { self.compile_statement(s)?; }
        self.emit(Op::try_end);
        let end_j = self.emit_jump(Op::br);

        // Patch catch offset (relative to IP after try_start's 4 operand bytes)
        if !catches.is_empty() {
            let catch_pos = self.current_offset();
            let ip_after_try_start = try_start_offset + 5; // 1 op byte + 4 operand bytes
            let relative_offset = catch_pos as i16 - ip_after_try_start as i16;
            let c = &mut self.chunks[self.current_chunk_idx];
            c.code[try_start_offset + 1] = (relative_offset >> 8) as u8;
            c.code[try_start_offset + 2] = (relative_offset & 0xff) as u8;

            for catch_clause in catches {
                self.current_scope_mut().begin_scope();
                if let Some(ref var_name) = catch_clause.var_name {
                    let slot = self.define_local(var_name);
                    // Exception object is on stack from the VM
                    self.emit_u16(Op::local_set, slot);
                    self.emit(Op::drop);
                } else {
                    self.emit(Op::drop); // discard exception
                }
                for s in &catch_clause.body { self.compile_statement(s)?; }
                let _popped = self.current_scope_mut().end_scope();
            }
        }

        let end_catch_j = self.emit_jump(Op::br);

        // Finally
        if let Some(ref fb) = finally_body {
            let finally_pos = self.current_offset();
            // Patch finally offset
            let try_start_code_pos = try_start_offset + 4; // after op bytes + catch offset
            let c = &mut self.chunks[self.current_chunk_idx];
            c.code[try_start_code_pos] = (finally_pos >> 8) as u8;
            c.code[try_start_code_pos + 1] = (finally_pos & 0xff) as u8;

            for s in fb { self.compile_statement(s)?; }
        }

        self.patch_jump(end_j);
        self.patch_jump(end_catch_j);
        Ok(())
    }

    // ================================================================
    // Expressions
    // ================================================================

    fn compile_expression(&mut self, expr: &Expression) -> Result<(), String> {
        match expr {
            // -- Literals --
            Expression::IntLiteral(n) => {
                if *n == 0 {
                    self.emit(Op::i32_const_0);
                } else if *n == 1 {
                    self.emit(Op::i32_const_1);
                } else if *n >= i32::MIN as i64 && *n <= i32::MAX as i64 {
                    self.emit_constant(Value::I32(*n as i32));
                } else {
                    self.emit_constant(Value::I64(*n));
                }
            }
            Expression::DoubleLiteral(n) => {
                if *n == 0.0 {
                    self.emit(Op::f64_const_0);
                } else {
                    self.emit_constant(Value::F64(*n));
                }
            }
            Expression::StringLiteral(s) => {
                self.emit_constant(Value::String(Rc::from(s.as_str())));
            }
            Expression::CharLiteral(c) => {
                self.emit_constant(Value::String(Rc::from(c.to_string().as_str())));
            }
            Expression::BoolLiteral(b) => {
                self.emit(if *b { Op::r#true } else { Op::r#false });
            }
            Expression::NullLiteral => {
                self.emit(Op::null);
            }
            Expression::InterpolatedString(parts) => {
                self.compile_interpolated_string(parts)?;
            }

            // -- Names --
            Expression::Identifier(name) => {
                self.compile_identifier(name)?;
            }
            Expression::This => {
                match self.resolve_variable("this") {
                    VarResolution::Local(slot) => self.emit_u16(Op::local_get, slot),
                    VarResolution::Global => self.emit(Op::null),
                }
            }
            Expression::Base => {
                // base → this.__super
                match self.resolve_variable("this") {
                    VarResolution::Local(slot) => {
                        self.emit_u16(Op::local_get, slot);
                        let idx = self.add_string_constant("__super");
                        self.emit_u16(Op::struct_get, idx);
                    }
                    VarResolution::Global => self.emit(Op::null),
                }
            }

            // -- Binary --
            Expression::Binary(op, left, right) => {
                self.compile_binary(op, left, right)?;
            }

            // -- Unary --
            Expression::Unary(op, operand) => {
                self.compile_expression(operand)?;
                match op {
                    UnaryOp::Neg => self.emit(Op::dyn_neg),
                    UnaryOp::Not => self.emit(Op::dyn_not),
                    UnaryOp::BitNot => self.emit(Op::i32_not),
                }
            }

            Expression::PostIncrement(operand) => {
                self.compile_increment_decrement(operand, true, true)?;
            }
            Expression::PostDecrement(operand) => {
                self.compile_increment_decrement(operand, false, true)?;
            }
            Expression::PreIncrement(operand) => {
                self.compile_increment_decrement(operand, true, false)?;
            }
            Expression::PreDecrement(operand) => {
                self.compile_increment_decrement(operand, false, false)?;
            }

            // -- Member access --
            Expression::MemberAccess(obj, member) => {
                self.compile_member_access(obj, member)?;
            }
            Expression::NullConditionalAccess(obj, member) => {
                self.compile_expression(obj)?;
                self.emit(Op::dup);
                let null_j = self.emit_jump(Op::br_if_null);
                let idx = self.add_string_constant(&member.to_lowercase());
                self.emit_u16(Op::struct_get, idx);
                let end_j = self.emit_jump(Op::br);
                self.patch_jump(null_j);
                // Already null on stack
                self.patch_jump(end_j);
            }

            // -- Indexer --
            Expression::Index(obj, index) => {
                self.compile_expression(obj)?;
                self.compile_expression(index)?;
                self.emit(Op::array_get);
            }

            // -- Invocation --
            Expression::Call(callee, args) => {
                self.compile_call(callee, args)?;
            }

            // -- Object creation --
            Expression::New(class_name, args) => {
                self.compile_new(class_name, args)?;
            }
            Expression::NewArray(_type_name, size) => {
                self.compile_expression(size)?;
                let idx = self.import("vybe:array", "newArray");
                self.emit_host_call(idx, 1);
            }
            Expression::ArrayInit(elements) => {
                for e in elements { self.compile_expression(e)?; }
                self.emit_u16(Op::array_new, elements.len() as u16);
            }
            Expression::ObjectInit(base_expr, inits) => {
                self.compile_expression(base_expr)?;
                for (prop, val) in inits {
                    self.emit(Op::dup);
                    self.compile_expression(val)?;
                    let idx = self.add_string_constant(&prop.to_lowercase());
                    self.emit_u16(Op::struct_set, idx);
                    self.emit(Op::drop);
                }
            }

            // -- Cast / type check --
            Expression::Cast(_type_name, inner) => {
                // Simplified: just compile inner (runtime coercion)
                self.compile_expression(inner)?;
            }
            Expression::Is(inner, type_name) => {
                self.compile_expression(inner)?;
                let idx = self.add_string_constant(&type_name.to_lowercase());
                self.emit_u16(Op::ref_test, idx);
            }
            Expression::As(inner, _type_name) => {
                // Simplified: compile inner, null if wrong type (would need ref_cast)
                self.compile_expression(inner)?;
            }
            Expression::TypeOf(type_name) => {
                self.emit_constant(Value::String(Rc::from(type_name.as_str())));
            }

            // -- Ternary --
            Expression::Conditional(cond, then_expr, else_expr) => {
                self.compile_expression(cond)?;
                self.emit(Op::dyn_to_bool);
                let else_j = self.emit_jump(Op::br_if_false);
                self.compile_expression(then_expr)?;
                let end_j = self.emit_jump(Op::br);
                self.patch_jump(else_j);
                self.compile_expression(else_expr)?;
                self.patch_jump(end_j);
            }

            // -- Null coalescing --
            Expression::NullCoalescing(left, right) => {
                self.compile_expression(left)?;
                self.emit(Op::dup);
                let end_j = self.emit_jump(Op::br_if_null);
                // Left is not null — keep it, skip right
                let done_j = self.emit_jump(Op::br);
                self.patch_jump(end_j);
                self.emit(Op::drop); // drop the null
                self.compile_expression(right)?;
                self.patch_jump(done_j);
            }

            // -- Lambda --
            Expression::Lambda(params, body) => {
                self.compile_lambda_expr(params, body)?;
            }
            Expression::LambdaBlock(params, body) => {
                self.compile_lambda_block(params, body)?;
            }

            // -- Await --
            Expression::Await(inner) => {
                self.compile_expression(inner)?;
                self.emit(Op::r#await);
            }

            // -- NameOf --
            Expression::NameOf(name) => {
                self.emit_constant(Value::String(Rc::from(name.as_str())));
            }

            // -- Default --
            Expression::Default(type_name) => {
                match type_name.as_deref() {
                    Some("int") | Some("Int32") | Some("long") | Some("Int64") => {
                        self.emit(Op::i32_const_0);
                    }
                    Some("double") | Some("Double") | Some("float") | Some("Single") => {
                        self.emit(Op::f64_const_0);
                    }
                    Some("bool") | Some("Boolean") => {
                        self.emit(Op::r#false);
                    }
                    _ => {
                        self.emit(Op::null);
                    }
                }
            }
        }
        Ok(())
    }

    fn compile_identifier(&mut self, name: &str) -> Result<(), String> {
        let lower = name.to_lowercase();

        // Check if it's a class field (this.field) — inside class methods
        if self.class_fields.contains(&lower) {
            match self.resolve_variable("this") {
                VarResolution::Local(slot) => {
                    self.emit_u16(Op::local_get, slot);
                    let idx = self.add_string_constant(&lower);
                    self.emit_u16(Op::struct_get, idx);
                    return Ok(());
                }
                _ => {}
            }
        }

        // Check if it's a class method (this.method)
        if self.class_methods.contains(&lower) {
            match self.resolve_variable("this") {
                VarResolution::Local(slot) => {
                    self.emit_u16(Op::local_get, slot);
                    let idx = self.add_string_constant(&lower);
                    self.emit_u16(Op::struct_get, idx);
                    return Ok(());
                }
                _ => {}
            }
        }

        match self.resolve_variable(name) {
            VarResolution::Local(slot) => {
                self.emit_u16(Op::local_get, slot);
            }
            VarResolution::Global => {
                let idx = self.add_string_constant(&lower);
                self.emit_u16(Op::global_get, idx);
            }
        }
        Ok(())
    }

    fn compile_binary(&mut self, op: &BinaryOp, left: &Expression, right: &Expression) -> Result<(), String> {
        // Short-circuit for And/Or
        match op {
            BinaryOp::And => {
                self.compile_expression(left)?;
                self.emit(Op::dyn_to_bool);
                let short_j = self.emit_jump(Op::br_if_false);
                self.compile_expression(right)?;
                self.emit(Op::dyn_to_bool);
                let end_j = self.emit_jump(Op::br);
                self.patch_jump(short_j);
                self.emit(Op::r#false);
                self.patch_jump(end_j);
                return Ok(());
            }
            BinaryOp::Or => {
                self.compile_expression(left)?;
                self.emit(Op::dyn_to_bool);
                let short_j = self.emit_jump(Op::br_if_true);
                self.compile_expression(right)?;
                self.emit(Op::dyn_to_bool);
                let end_j = self.emit_jump(Op::br);
                self.patch_jump(short_j);
                self.emit(Op::r#true);
                self.patch_jump(end_j);
                return Ok(());
            }
            _ => {}
        }

        self.compile_expression(left)?;
        self.compile_expression(right)?;

        match op {
            BinaryOp::Add => self.emit(Op::dyn_add),
            BinaryOp::Sub => self.emit(Op::f64_sub),
            BinaryOp::Mul => self.emit(Op::f64_mul),
            BinaryOp::Div => self.emit(Op::f64_div),
            BinaryOp::Mod => self.emit(Op::f64_mod),
            BinaryOp::Eq => self.emit(Op::dyn_eq),
            BinaryOp::Neq => self.emit(Op::dyn_ne),
            BinaryOp::Lt => self.emit(Op::dyn_lt),
            BinaryOp::Gt => self.emit(Op::dyn_gt),
            BinaryOp::Le => self.emit(Op::dyn_le),
            BinaryOp::Ge => self.emit(Op::dyn_ge),
            BinaryOp::BitAnd => self.emit(Op::i32_and),
            BinaryOp::BitOr => self.emit(Op::i32_or),
            BinaryOp::BitXor => self.emit(Op::i32_xor),
            BinaryOp::Shl => self.emit(Op::i32_shl),
            BinaryOp::Shr => self.emit(Op::i32_shr_s),
            BinaryOp::NullCoalescing => {
                // Already handled by NullCoalescing expression, but just in case
                // left ?? right — left is already on stack, right on top
                // This shouldn't normally be reached
            }
            BinaryOp::And | BinaryOp::Or => unreachable!(),
        }
        Ok(())
    }

    fn compile_increment_decrement(&mut self, operand: &Expression, is_inc: bool, is_post: bool) -> Result<(), String> {
        match operand {
            Expression::Identifier(name) => {
                match self.resolve_variable(name) {
                    VarResolution::Local(slot) => {
                        if is_post {
                            self.emit_u16(Op::local_get, slot); // old value
                        }
                        self.emit_u16(Op::local_get, slot);
                        self.emit(Op::i32_const_1);
                        if is_inc { self.emit(Op::dyn_add); } else { self.emit(Op::f64_sub); }
                        self.emit_u16(Op::local_set, slot);
                        self.emit(Op::drop);
                        if !is_post {
                            self.emit_u16(Op::local_get, slot); // new value
                        }
                    }
                    VarResolution::Global => {
                        let idx = self.add_string_constant(&name.to_lowercase());
                        if is_post {
                            self.emit_u16(Op::global_get, idx);
                        }
                        let idx2 = self.add_string_constant(&name.to_lowercase());
                        self.emit_u16(Op::global_get, idx2);
                        self.emit(Op::i32_const_1);
                        if is_inc { self.emit(Op::dyn_add); } else { self.emit(Op::f64_sub); }
                        let idx3 = self.add_string_constant(&name.to_lowercase());
                        self.emit_u16(Op::global_set, idx3);
                        self.emit(Op::drop);
                        if !is_post {
                            let idx4 = self.add_string_constant(&name.to_lowercase());
                            self.emit_u16(Op::global_get, idx4);
                        }
                    }
                }
            }
            Expression::MemberAccess(obj, prop) => {
                self.compile_expression(obj)?;
                self.emit(Op::dup);
                let pidx = self.add_string_constant(&prop.to_lowercase());
                self.emit_u16(Op::struct_get, pidx);
                if is_post { self.emit(Op::dup); } // keep old value under
                self.emit(Op::i32_const_1);
                if is_inc { self.emit(Op::dyn_add); } else { self.emit(Op::f64_sub); }
                let pidx2 = self.add_string_constant(&prop.to_lowercase());
                self.emit_u16(Op::struct_set, pidx2);
                self.emit(Op::drop);
                if !is_post {
                    // Need to get new value — re-read
                    self.compile_expression(obj)?;
                    let pidx3 = self.add_string_constant(&prop.to_lowercase());
                    self.emit_u16(Op::struct_get, pidx3);
                }
            }
            _ => {
                return Err(format!("Cannot increment/decrement {:?}", operand));
            }
        }
        Ok(())
    }

    fn compile_member_access(&mut self, obj: &Expression, member: &str) -> Result<(), String> {
        let member_lower = member.to_lowercase();

        // Check for namespace-style access: Console.WriteLine, Math.Floor, etc.
        if let Some(parts) = self.flatten_member_access(obj, member) {
            // Check Math intrinsics first
            if let Some(math_op) = self.try_math_intrinsic(&parts) {
                // This is a reference to a math function — it will be called
                // The call site handles it; here we just note it's a namespace
                // Actually, member access alone isn't callable — it gets used in Call
                // Just push as a global reference
                let full = parts.join(".");
                self.emit_constant(Value::String(Rc::from(full.as_str())));
                let _ = math_op;
                return Ok(());
            }

            // Try interface resolution
            let parts_refs: Vec<&str> = parts.iter().map(|s| s.as_str()).collect();
            if let Some((module, func)) = self.resolve_interface_call(&parts_refs) {
                // It's a known namespace function reference
                let full = format!("{}:{}", module, func);
                self.emit_constant(Value::String(Rc::from(full.as_str())));
                return Ok(());
            }
        }

        // String/array built-in properties
        match member_lower.as_str() {
            "length" => {
                self.compile_expression(obj)?;
                self.emit(Op::str_length);
                return Ok(());
            }
            _ => {}
        }

        // Regular member access
        self.compile_expression(obj)?;

        // Check for property getter (__get_prop)
        let get_name = format!("__get_{}", member_lower);
        // We emit struct_get for the property; the VM handles __get_ delegation
        let idx = self.add_string_constant(&member_lower);
        self.emit_u16(Op::struct_get, idx);
        let _ = get_name;
        Ok(())
    }

    /// Flatten a member access chain into parts: Console.WriteLine → ["Console", "WriteLine"]
    fn flatten_member_access(&self, obj: &Expression, member: &str) -> Option<Vec<String>> {
        let mut parts = Vec::new();
        let mut current = obj;
        loop {
            match current {
                Expression::Identifier(name) => {
                    parts.push(name.clone());
                    break;
                }
                Expression::MemberAccess(inner, m) => {
                    parts.push(m.clone());
                    current = inner;
                }
                _ => return None,
            }
        }
        parts.reverse();
        parts.push(member.to_string());
        Some(parts)
    }

    /// Check if a dotted access is a Math intrinsic, returning the opcode.
    fn try_math_intrinsic(&self, parts: &[String]) -> Option<Op> {
        if parts.len() == 2 && parts[0].eq_ignore_ascii_case("Math") {
            return match parts[1].to_lowercase().as_str() {
                "floor" => Some(Op::f64_floor),
                "ceil" | "ceiling" => Some(Op::f64_ceil),
                "abs" => Some(Op::f64_abs),
                "sqrt" => Some(Op::f64_sqrt),
                "min" => Some(Op::f64_min),
                "max" => Some(Op::f64_max),
                "truncate" | "trunc" => Some(Op::f64_trunc),
                "round" => Some(Op::f64_nearest),
                _ => None,
            };
        }
        None
    }

    fn compile_call(&mut self, callee: &Expression, args: &[Expression]) -> Result<(), String> {
        // Console.* — uses call_import so tests can override wasi:cli log
        if let Expression::MemberAccess(obj, method) = callee {
            if let Expression::Identifier(ref obj_name) = **obj {
                let obj_lower = obj_name.to_lowercase();
                let meth_lower = method.to_lowercase();
                if obj_lower == "console" {
                    let host = match meth_lower.as_str() {
                        "writeline" | "write" => Some(("wasi:cli", "log")),
                        "readline" => Some(("wasi:cli", "readLine")),
                        "error" => Some(("wasi:cli", "error")),
                        _ => None,
                    };
                    if let Some((module, func)) = host {
                        for arg in args { self.compile_expression(arg)?; }
                        let idx = self.import(module, func);
                        self.emit_host_call(idx, args.len() as u8);
                        return Ok(());
                    }
                }
                // Other namespaces — resolve via global namespace objects
                if self.is_namespace(&obj_lower) && obj_lower != "console" {
                    let ns_idx = self.add_string_constant(&obj_lower);
                    self.emit_u16(Op::global_get, ns_idx);
                    let meth_idx = self.add_string_constant(&meth_lower);
                    self.emit_u16(Op::struct_get, meth_idx);
                    for arg in args { self.compile_expression(arg)?; }
                    self.emit_u8(Op::call, args.len() as u8);
                    return Ok(());
                }
            }
        }

        // Namespace-style calls: Math.Floor, etc.
        if let Expression::MemberAccess(obj, method) = callee {
            if let Some(parts) = self.flatten_member_access(obj, method) {
                // Math intrinsics (single-arg ops become direct opcodes)
                if let Some(math_op) = self.try_math_intrinsic(&parts) {
                    match math_op {
                        Op::f64_min | Op::f64_max => {
                            // These take 2 args
                            if args.len() >= 2 {
                                self.compile_expression(&args[0])?;
                                self.compile_expression(&args[1])?;
                                self.emit(math_op);
                                return Ok(());
                            }
                        }
                        _ => {
                            // Single arg
                            if let Some(arg) = args.first() {
                                self.compile_expression(arg)?;
                                self.emit(math_op);
                                return Ok(());
                            }
                        }
                    }
                }

                // Math functions that map to host calls
                if parts.len() == 2 && parts[0].eq_ignore_ascii_case("Math") {
                    let func_lower = parts[1].to_lowercase();
                    match func_lower.as_str() {
                        "pow" | "log" | "log10" | "sin" | "cos" | "tan"
                        | "asin" | "acos" | "atan" | "atan2" | "exp"
                        | "sign" | "clamp" => {
                            for arg in args { self.compile_expression(arg)?; }
                            let idx = self.import("vybe:math", &func_lower);
                            self.emit_host_call(idx, args.len() as u8);
                            return Ok(());
                        }
                        _ => {}
                    }
                }

                // Try interface resolution for full dotted names
                let parts_refs: Vec<&str> = parts.iter().map(|s| s.as_str()).collect();
                if let Some((module, func)) = self.resolve_interface_call(&parts_refs) {
                    for arg in args { self.compile_expression(arg)?; }
                    let idx = self.import(&module, &func);
                    self.emit_host_call(idx, args.len() as u8);
                    return Ok(());
                }
            }

            // Instance method call: obj.Method(args) → get method, call with obj as first arg
            return self.compile_method_call_expr(obj, method, args);
        }

        // Simple function call: foo(args)
        if let Expression::Identifier(name) = callee {
            let lower = name.to_lowercase();

            // Inside a class: bare method call → this.method(this, args)
            if self.class_methods.contains(&lower) {
                if let VarResolution::Local(this_slot) = self.resolve_variable("this") {
                    self.emit_u16(Op::local_get, this_slot);
                    let idx = self.add_string_constant(&lower);
                    self.emit_u16(Op::struct_get, idx);
                    self.emit_u16(Op::local_get, this_slot);
                    for arg in args { self.compile_expression(arg)?; }
                    self.emit_u8(Op::call, (args.len() + 1) as u8);
                    return Ok(());
                }
            }

            // Check if it's a known global function
            match self.resolve_variable(name) {
                VarResolution::Local(slot) => {
                    self.emit_u16(Op::local_get, slot);
                    for arg in args { self.compile_expression(arg)?; }
                    self.emit_u8(Op::call, args.len() as u8);
                    return Ok(());
                }
                VarResolution::Global => {
                    let idx = self.add_string_constant(&lower);
                    self.emit_u16(Op::global_get, idx);
                    for arg in args { self.compile_expression(arg)?; }
                    self.emit_u8(Op::call, args.len() as u8);
                    return Ok(());
                }
            }
        }

        // Generic callee
        self.compile_expression(callee)?;
        for arg in args { self.compile_expression(arg)?; }
        self.emit_u8(Op::call, args.len() as u8);
        Ok(())
    }

    fn compile_method_call_expr(&mut self, obj: &Expression, method: &str, args: &[Expression]) -> Result<(), String> {
        let method_lower = method.to_lowercase();

        // Controls.Add(ctrl) → controlsAdd host call
        if method_lower == "add" {
            if let Expression::MemberAccess(parent, member) = obj {
                if member.to_lowercase() == "controls" {
                    // parent.Controls.Add(ctrl) → controlsAdd(parent, ctrl)
                    self.compile_expression(parent)?;
                    for arg in args { self.compile_expression(arg)?; }
                    let idx = self.import("vybe:gui", "controlsAdd");
                    self.emit_host_call(idx, (args.len() + 1) as u8);
                    return Ok(());
                }
            }
        }

        // No-op layout methods
        if matches!(method_lower.as_str(),
            "suspendlayout" | "resumelayout" | "performlayout"
            | "begininit" | "endinit" | "dispose"
        ) {
            self.emit(Op::null);
            return Ok(());
        }

        // String built-in methods → direct opcodes
        match method_lower.as_str() {
            "tostring" => {
                self.compile_expression(obj)?;
                let idx = self.import("vybe:convert", "toString");
                self.emit_host_call(idx, 1);
                return Ok(());
            }
            "toupper" => {
                self.compile_expression(obj)?;
                self.emit(Op::str_to_upper);
                return Ok(());
            }
            "tolower" => {
                self.compile_expression(obj)?;
                self.emit(Op::str_to_lower);
                return Ok(());
            }
            "trim" => {
                self.compile_expression(obj)?;
                self.emit(Op::str_trim);
                return Ok(());
            }
            "trimstart" => {
                self.compile_expression(obj)?;
                self.emit(Op::str_trim_start);
                return Ok(());
            }
            "trimend" => {
                self.compile_expression(obj)?;
                self.emit(Op::str_trim_end);
                return Ok(());
            }
            "startswith" => {
                self.compile_expression(obj)?;
                if let Some(arg) = args.first() { self.compile_expression(arg)?; }
                self.emit(Op::str_starts_with);
                return Ok(());
            }
            "endswith" => {
                self.compile_expression(obj)?;
                if let Some(arg) = args.first() { self.compile_expression(arg)?; }
                self.emit(Op::str_ends_with);
                return Ok(());
            }
            "contains" => {
                self.compile_expression(obj)?;
                if let Some(arg) = args.first() { self.compile_expression(arg)?; }
                self.emit(Op::str_contains);
                return Ok(());
            }
            "replace" => {
                self.compile_expression(obj)?;
                for arg in args { self.compile_expression(arg)?; }
                self.emit(Op::str_replace);
                return Ok(());
            }
            "split" => {
                self.compile_expression(obj)?;
                for arg in args { self.compile_expression(arg)?; }
                self.emit(Op::str_split);
                return Ok(());
            }
            "substring" => {
                self.compile_expression(obj)?;
                for arg in args { self.compile_expression(arg)?; }
                self.emit(Op::str_substring);
                return Ok(());
            }
            "indexof" => {
                self.compile_expression(obj)?;
                if let Some(arg) = args.first() { self.compile_expression(arg)?; }
                self.emit(Op::str_index_of);
                return Ok(());
            }
            "lastindexof" => {
                self.compile_expression(obj)?;
                if let Some(arg) = args.first() { self.compile_expression(arg)?; }
                self.emit(Op::str_last_index_of);
                return Ok(());
            }
            "padleft" => {
                self.compile_expression(obj)?;
                for arg in args { self.compile_expression(arg)?; }
                self.emit(Op::str_pad_start);
                return Ok(());
            }
            "padright" => {
                self.compile_expression(obj)?;
                for arg in args { self.compile_expression(arg)?; }
                self.emit(Op::str_pad_end);
                return Ok(());
            }
            _ => {}
        }

        // Convert methods
        match method_lower.as_str() {
            "toint32" | "parseint" if self.is_namespace_expr(obj) => {
                for arg in args { self.compile_expression(arg)?; }
                let idx = self.import("vybe:convert", "cint");
                self.emit_host_call(idx, args.len() as u8);
                return Ok(());
            }
            "todouble" | "parsefloat" if self.is_namespace_expr(obj) => {
                for arg in args { self.compile_expression(arg)?; }
                let idx = self.import("vybe:convert", "cdbl");
                self.emit_host_call(idx, args.len() as u8);
                return Ok(());
            }
            _ => {}
        }

        // Try runtime callMethod dispatch (handles List, Map, Set, Dict, Queue, Stack)
        self.compile_expression(obj)?;
        self.emit_constant(Value::String(Rc::from(method)));
        for arg in args { self.compile_expression(arg)?; }
        let cm_idx = self.import("vybe:runtime", "callMethod");
        self.emit_host_call(cm_idx, (args.len() + 2) as u8);
        // If callMethod returned non-Undefined, it handled it — we're done
        // (Null is a valid return value from handled methods; Undefined means "not handled")
        self.emit(Op::dup);
        self.emit(Op::undefined);
        self.emit(Op::eq);
        let not_handled = self.emit_jump(Op::br_if_true); // undefined = not handled
        let done = self.emit_jump(Op::br); // handled — skip fallback
        self.patch_jump(not_handled);
        self.emit(Op::drop); // drop undefined

        // Fallback: general instance method call via struct_get
        self.compile_expression(obj)?;
        let method_idx = self.add_string_constant(&method_lower);
        self.emit(Op::dup);
        self.emit_u16(Op::struct_get, method_idx);
        // Stack: [obj, method_fn] — need [method_fn, obj, args...]
        // Swap: not a direct opcode, so we use local
        let tmp_slot = self.define_local("__tmp_method");
        self.emit_u16(Op::local_set, tmp_slot);
        self.emit(Op::drop);
        // Now stack: [obj]. Get method back
        let tmp2_slot = self.define_local("__tmp_obj");
        self.emit_u16(Op::local_set, tmp2_slot);
        self.emit(Op::drop);
        // Push method, then obj (this), then args
        self.emit_u16(Op::local_get, tmp_slot);
        self.emit_u16(Op::local_get, tmp2_slot);
        for arg in args { self.compile_expression(arg)?; }
        self.emit_u8(Op::call, (args.len() + 1) as u8); // +1 for this
        self.patch_jump(done);
        Ok(())
    }

    fn compile_new(&mut self, class_name: &str, args: &[Expression]) -> Result<(), String> {
        let lower = class_name.to_lowercase();

        // Strip generic type params: List<int> → list
        let bare = lower
            .find('<').map(|p| lower[..p].to_string()).unwrap_or_else(|| lower.clone());
        // Strip namespace prefixes
        let bare = bare
            .strip_prefix("system.data.sqlclient.").or_else(|| bare.strip_prefix("system.data.oledb."))
            .or_else(|| bare.strip_prefix("system.net.sockets."))
            .or_else(|| bare.strip_prefix("system.io."))
            .or_else(|| bare.strip_prefix("system.collections.generic."))
            .or_else(|| bare.strip_prefix("system.collections."))
            .or_else(|| bare.strip_prefix("system.text."))
            .or_else(|| bare.strip_prefix("system.windows.forms."))
            .or_else(|| bare.strip_prefix("system.drawing."))
            .unwrap_or(&bare)
            .to_string();

        // Exception types → simple object with message
        if matches!(bare.as_str(), "exception" | "argumentexception" | "invalidoperationexception"
            | "notimplementedexception" | "notsupportedexception" | "nullreferenceexception"
            | "indexoutofrangeexception" | "argumentnullexception") {
            self.emit_u16(Op::struct_new, 0);
            self.emit(Op::dup);
            if let Some(msg_arg) = args.first() {
                self.compile_expression(msg_arg)?;
            } else {
                self.emit_constant(Value::String(Rc::from("")));
            }
            let msg_idx = self.add_string_constant("message");
            self.emit_u16(Op::struct_set, msg_idx);
            self.emit(Op::drop);
            return Ok(());
        }

        // Unified new: look up constructor from globals.
        // - User-defined classes: global is a Function → struct_new + call(args+1)
        // - Host types (Button, Point, List): global is a HostFunction → call(args)
        // - Both resolve via the same global_get.
        let idx = self.add_string_constant(&bare);
        self.emit_u16(Op::global_get, idx);

        if self.defined_classes.contains(&bare) {
            // User class: create empty object, call constructor with it
            self.emit_u16(Op::struct_new, 0);
            // Set __type for instanceof
            self.emit(Op::dup);
            self.emit_constant(Value::String(Rc::from(bare.as_str())));
            let type_idx = self.add_string_constant("__type");
            self.emit_u16(Op::struct_set, type_idx);
            self.emit(Op::drop);
            for arg in args { self.compile_expression(arg)?; }
            self.emit_u8(Op::call, (args.len() + 1) as u8);
        } else {
            // Host constructor (Button, Point, List, etc): just call it
            for arg in args { self.compile_expression(arg)?; }
            self.emit_u8(Op::call, args.len() as u8);
        }
        Ok(())
    }

    fn compile_interpolated_string(&mut self, parts: &[StringPart]) -> Result<(), String> {
        if parts.is_empty() {
            self.emit_constant(Value::String(Rc::from("")));
            return Ok(());
        }
        if parts.len() == 1 {
            match &parts[0] {
                StringPart::Text(s) => {
                    self.emit_constant(Value::String(Rc::from(s.as_str())));
                }
                StringPart::Expr(e) => {
                    self.compile_expression(e)?;
                    let idx = self.import("vybe:convert", "toString");
                    self.emit_host_call(idx, 1);
                }
            }
            return Ok(());
        }

        let count = parts.len();
        for part in parts {
            match part {
                StringPart::Text(s) => {
                    self.emit_constant(Value::String(Rc::from(s.as_str())));
                }
                StringPart::Expr(e) => {
                    self.compile_expression(e)?;
                    let idx = self.import("vybe:convert", "toString");
                    self.emit_host_call(idx, 1);
                }
            }
        }
        self.emit_u8(Op::str_concat_n, count as u8);
        Ok(())
    }

    fn compile_lambda_expr(&mut self, params: &[String], body: &Expression) -> Result<(), String> {
        let mut chunk = Chunk::new("<lambda>");
        chunk.arity = params.len() as u8;
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        for param in params {
            scope.define_local(&param.to_lowercase());
        }

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        self.compile_expression(body)?;
        self.emit(Op::r#return);

        let lc = self.current_scope().next_slot;
        self.chunks[idx].local_count = lc;
        let upvalues = self.current_scope().upvalues.clone();
        self.scopes.pop();
        self.current_chunk_idx = saved;
        self.emit_ref_func(idx, &upvalues);
        Ok(())
    }

    fn compile_lambda_block(&mut self, params: &[String], body: &[Statement]) -> Result<(), String> {
        let mut chunk = Chunk::new("<lambda>");
        chunk.arity = params.len() as u8;
        let idx = self.chunks.len();
        self.chunks.push(chunk);

        let mut scope = Scope::new_function();
        for param in params {
            scope.define_local(&param.to_lowercase());
        }

        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);

        for stmt in body { self.compile_statement(stmt)?; }
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
}

// ================================================================
// Helper functions
// ================================================================

/// Map C#/CLR method names to actual host function names.
fn map_interface_func(module: &str, func: &str) -> String {
    match (module, func) {
        // Console
        ("wasi:cli", "writeline") => "log".into(),
        ("wasi:cli", "write") => "log".into(),
        ("wasi:cli", "readline") => "readLine".into(),
        ("wasi:cli", "error") => "error".into(),
        // Math
        ("vybe:math", f) => f.to_string(),
        // Filesystem
        ("wasi:filesystem", "readalltext") => "readFile".into(),
        ("wasi:filesystem", "writealltext") => "writeFile".into(),
        ("wasi:filesystem", "appendalltext") => "appendFile".into(),
        ("wasi:filesystem", "exists") => "exists".into(),
        ("wasi:filesystem", "delete") => "remove".into(),
        ("wasi:filesystem", "copy") => "copy".into(),
        ("wasi:filesystem", "move") => "rename".into(),
        ("wasi:filesystem", "combine") => "pathCombine".into(),
        ("wasi:filesystem", "getfilename") => "pathGetFileName".into(),
        ("wasi:filesystem", "getextension") => "pathGetExtension".into(),
        ("wasi:filesystem", "getdirectoryname") => "pathGetDirectory".into(),
        ("wasi:filesystem", "getfilenamewithoutextension") => "pathGetFileNameWithoutExt".into(),
        ("wasi:filesystem", "changeextension") => "pathChangeExtension".into(),
        ("wasi:filesystem", "getfullpath") => "pathGetFullPath".into(),
        ("wasi:filesystem", "gettemppath") => "pathGetTempPath".into(),
        ("wasi:filesystem", "createdirectory") => "mkdir".into(),
        ("wasi:filesystem", "getfiles") => "listDir".into(),
        ("wasi:filesystem", "getcurrentdirectory") => "cwd".into(),
        // Convert
        ("vybe:convert", "toint32") => "cint".into(),
        ("vybe:convert", "todouble") => "cdbl".into(),
        ("vybe:convert", "tostring") => "toString".into(),
        ("vybe:convert", "toboolean") => "cbool".into(),
        ("vybe:convert", "todatetime") => "toString".into(),
        // Environment
        ("wasi:cli", "getenvironmentvariable") => "getEnv".into(),
        ("wasi:cli", "machinename") => "machineName".into(),
        ("wasi:cli", "currentdirectory") => "cwd".into(),
        // Threading
        ("wasi:clocks", "sleep") => "sleep".into(),
        // GUI
        ("vybe:gui", f) => {
            let cap = capitalize_control_name(f);
            if !cap.is_empty() && cap != f {
                format!("new_{}", cap)
            } else {
                f.to_string()
            }
        }
        // Default
        (_, f) => f.to_string(),
    }
}

/// Map lowercase control name to proper cased name.
fn capitalize_control_name(name: &str) -> String {
    match name {
        "button" => "Button", "label" => "Label", "textbox" => "TextBox",
        "checkbox" => "CheckBox", "radiobutton" => "RadioButton",
        "combobox" => "ComboBox", "listbox" => "ListBox",
        "panel" => "Panel", "groupbox" => "GroupBox",
        "tabcontrol" => "TabControl", "tabpage" => "TabPage",
        "datagridview" => "DataGridView", "progressbar" => "ProgressBar",
        "trackbar" => "TrackBar", "numericupdown" => "NumericUpDown",
        "datetimepicker" => "DateTimePicker", "richtextbox" => "RichTextBox",
        "picturebox" => "PictureBox", "menustrip" => "MenuStrip",
        "toolstrip" => "ToolStrip", "statusstrip" => "StatusStrip",
        "splitcontainer" => "SplitContainer",
        "flowlayoutpanel" => "FlowLayoutPanel",
        "tablelayoutpanel" => "TableLayoutPanel",
        "linklabel" => "LinkLabel", "maskedtextbox" => "MaskedTextBox",
        "listview" => "ListView", "webbrowser" => "WebBrowser",
        "monthcalendar" => "MonthCalendar",
        "contextmenustrip" => "ContextMenuStrip",
        "timer" => "Timer", "bindingsource" => "BindingSource",
        "tooltip" => "ToolTip", "imagelist" => "ImageList",
        "openfiledialog" => "OpenFileDialog",
        "savefiledialog" => "SaveFileDialog",
        "folderbrowserdialog" => "FolderBrowserDialog",
        "colordialog" => "ColorDialog", "fontdialog" => "FontDialog",
        _ => return name.to_string(),
    }.to_string()
}
