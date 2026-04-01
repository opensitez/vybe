use std::rc::Rc;
use std::collections::HashSet;

use vybe_bytecode::{Chunk, Value, Op};
use vybe_bytecode::chunk::TypeEntry;
use vybe_compiler_common::functions as common_fn;
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
    pub(crate) loop_stack: Vec<LoopContext>,
    pub(crate) class_fields: HashSet<String>,
    pub(crate) class_methods: HashSet<String>,
    /// Stores each class's field/method names so derived classes can inherit them
    pub(crate) class_field_map: std::collections::HashMap<String, HashSet<String>>,
    pub(crate) class_method_map: std::collections::HashMap<String, HashSet<String>>,
    /// Component Model: imported interface prefixes from `Imports` statements.
    pub(crate) interface_imports: Vec<String>,
    /// Known built-in types: name → (constructor_module, constructor_fn)
    pub(crate) known_types: std::collections::HashMap<String, (&'static str, &'static str)>,
    /// Declared function signatures: name → vec of is_byref per param
    /// Used to box/unbox ByRef args at call sites.
    pub(crate) func_signatures: std::collections::HashMap<String, Vec<bool>>,
    /// Names known to hold arrays (from Dim arr(N) declarations)
    pub(crate) known_arrays: HashSet<String>,
    /// WASM GC type table: compile-time type definitions for classes.
    /// Loaded into VM's TypeRegistry before execution.
    pub(crate) type_entries: Vec<TypeEntry>,
    /// Class name → index into type_entries (for set_type_id at construction sites).
    pub(crate) class_type_ids: std::collections::HashMap<String, usize>,
}

pub(crate) struct LoopContext {
    pub _start: usize,
    pub break_jumps: Vec<usize>,
    pub continue_jumps: Vec<usize>,
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
            loop_stack: Vec::new(),
            class_fields: HashSet::new(),
            class_methods: HashSet::new(),
            class_field_map: std::collections::HashMap::new(),
            class_method_map: std::collections::HashMap::new(),
            known_types: Self::init_known_types(),
            func_signatures: std::collections::HashMap::new(),
            known_arrays: HashSet::new(),
            type_entries: Vec::new(),
            class_type_ids: std::collections::HashMap::new(),
            interface_imports: vec![
                // Default imports — always available
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
                "microsoft.visualbasic".into(),
            ],
        }
    }

    pub fn compile(mut self, program: &Program) -> Result<Vec<Chunk>, String> {
        // Merge partial classes before compilation
        let declarations = Self::merge_partial_classes(&program.declarations);
        for decl in &declarations {
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
        // Attach WASM GC type table to script chunk
        self.chunks[0].types = self.type_entries;
        vybe_compiler_common::bundle::finalize_with_stdlib(&mut self.chunks);
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
    /// Print N args on the stack via wasi:cli/log (import routed to chunk 0).
    pub(crate) fn emit_print(&mut self, arg_count: u8) {
        let idx = self.import("wasi:cli", "log");
        self.emit_host_call(idx, arg_count);
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
            | "timespan" | "guid" | "point" | "size" | "font" | "random"
            | "path" | "messagebox" | "encoding"
        )
    }

    pub(crate) fn is_namespace_expr(&self, expr: &Expression) -> bool {
        match expr {
            Expression::Variable(name) => self.is_namespace(&name.as_str().to_lowercase()),
            Expression::MemberAccess(inner, _) => self.is_namespace_expr(inner),
            _ => false,
        }
    }

    /// Merge partial class declarations into single classes.
    fn merge_partial_classes(declarations: &[Declaration]) -> Vec<Declaration> {
        use std::collections::HashMap as Map;
        let mut class_map: Map<String, ClassDecl> = Map::new();
        let mut result: Vec<Declaration> = Vec::new();
        let mut class_order: Vec<String> = Vec::new();

        for decl in declarations {
            if let Declaration::Class(class) = decl {
                let key = class.name.as_str().to_lowercase();
                if let Some(existing) = class_map.get_mut(&key) {
                    // Merge: add fields, methods, properties from this partial
                    existing.fields.extend(class.fields.clone());
                    existing.methods.extend(class.methods.clone());
                    existing.properties.extend(class.properties.clone());
                    if existing.inherits.is_none() && class.inherits.is_some() {
                        existing.inherits = class.inherits.clone();
                    }
                    existing.implements.extend(class.implements.clone());
                } else {
                    class_order.push(key.clone());
                    class_map.insert(key, class.clone());
                }
            } else {
                result.push(decl.clone());
            }
        }

        // Insert merged classes in original order
        let mut final_result: Vec<Declaration> = Vec::new();
        let mut class_inserted: std::collections::HashSet<String> = std::collections::HashSet::new();
        for decl in declarations {
            if let Declaration::Class(class) = decl {
                let key = class.name.as_str().to_lowercase();
                if !class_inserted.contains(&key) {
                    if let Some(merged) = class_map.remove(&key) {
                        final_result.push(Declaration::Class(merged));
                        class_inserted.insert(key);
                    }
                }
                // Skip duplicate partials
            } else {
                final_result.push(decl.clone());
            }
        }
        final_result
    }

    /// Component Model: resolve a dotted name to a (module, function) host import.
    /// Tries to match the name against registered interface prefixes.
    ///
    /// "System.Windows.Forms.Button" with import "system.windows.forms"
    ///   → Some(("system.windows.forms", "new_Button"))
    ///
    /// "Console.WriteLine" with import "system.console"
    ///   → Some(("wasi:cli", "log")) via builtin table
    ///
    /// Returns None if no interface matches.
    pub(crate) fn resolve_interface_call(&self, parts: &[&str]) -> Option<(String, String)> {
        // Build progressively longer prefixes and check against imports
        // e.g. for ["System", "Windows", "Forms", "Button"]:
        //   try "system" → no match for "windows.forms.button"
        //   try "system.windows" → no match
        //   try "system.windows.forms" → match! func = "button"
        let lower_parts: Vec<String> = parts.iter().map(|p| p.to_lowercase()).collect();

        for prefix_len in (1..lower_parts.len()).rev() {
            let prefix = lower_parts[..prefix_len].join(".");
            if self.interface_imports.contains(&prefix) {
                let func = lower_parts[prefix_len..].join(".");
                // Map to actual host module
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
                    "microsoft.visualbasic" => "vybe:string",
                    _ => &prefix,
                };
                // Map VB method names to actual host function names
                let mapped_func = map_interface_func(module, &func);
                return Some((module.to_string(), mapped_func));
            }
        }
        None
    }

    fn init_known_types() -> std::collections::HashMap<String, (&'static str, &'static str)> {
        let mut m = std::collections::HashMap::new();
        // All built-in types with their constructor host functions
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
            // WinForms controls (handled separately via capitalize)
            ("form", "vybe:gui", "newForm"),
        ] {
            m.insert(name.to_string(), (*module, *func));
        }
        m
    }

} // end impl Compiler (part 1)

/// Map VB/CLR method names to actual host function names.
/// e.g. "writeline" → "log", "getdirectoryname" → "pathGetDirectory"
fn map_interface_func(module: &str, func: &str) -> String {
    match (module, func) {
        // Console
        ("wasi:cli", "writeline") => "log".into(),
        ("wasi:cli", "write") => "log".into(),
        ("wasi:cli", "readline") => "readLine".into(),
        ("wasi:cli", "error") => "error".into(),
        // Math
        ("vybe:math", f) => f.to_string(), // math names match
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
        ("vybe:convert", "todatetime") => "toString".into(), // simplified
        // Environment
        ("wasi:cli", "getenvironmentvariable") => "getEnv".into(),
        ("wasi:cli", "machinename") => "machineName".into(),
        ("wasi:cli", "currentdirectory") => "cwd".into(),
        ("wasi:cli", "print") => "log".into(),
        // Threading
        ("wasi:clocks", "sleep") => "sleep".into(),
        // GUI
        ("vybe:gui", f) => {
            // WinForms control constructors: button → new_Button
            let cap = capitalize_control_name_for_interface(f);
            if !cap.is_empty() {
                format!("new_{}", cap)
            } else {
                f.to_string()
            }
        }
        // Default: use as-is
        (_, f) => f.to_string(),
    }
}

fn capitalize_control_name_for_interface(name: &str) -> String {
    match name {
        "button" | "label" | "textbox" | "checkbox" | "radiobutton"
        | "combobox" | "listbox" | "panel" | "groupbox" | "tabcontrol"
        | "tabpage" | "datagridview" | "progressbar" | "trackbar"
        | "numericupdown" | "datetimepicker" | "richtextbox" | "picturebox"
        | "menustrip" | "toolstrip" | "statusstrip" | "splitcontainer"
        | "flowlayoutpanel" | "tablelayoutpanel" | "linklabel" | "maskedtextbox"
        | "listview" | "webbrowser" | "monthcalendar" | "contextmenustrip"
        | "timer" | "bindingsource" | "tooltip" | "imagelist" => {
            // Use the capitalize function from expressions.rs
            super::expressions::capitalize_control_name(name)
        }
        _ => String::new(),
    }
}

impl Compiler {
    /// Emit set_type_id for the TOS object using __tid_<name> global.
    /// If the global doesn't exist at runtime, the I32(0) default is harmless (Object type).
    /// Emit set_type_id for the TOS object using __tid_<name> global.
    /// Stack: [obj] → [obj] (type_id stamped in-place).
    pub(crate) fn emit_set_type_id(&mut self, type_name: &str) {
        let tid_name = format!("__tid_{}", type_name.to_lowercase());
        let tid_idx = self.add_string_constant(&tid_name);
        // set_type_id pops type_id, peeks obj — leaves obj on stack
        self.emit_u16(Op::global_get, tid_idx);
        self.emit(Op::set_type_id);
    }

    // ---- Declarations ----

    pub(crate) fn compile_declaration(&mut self, decl: &Declaration) -> Result<(), String> {
        match decl {
            Declaration::Sub(sub) => {
                // Record signature for ByRef call-site boxing
                let sig: Vec<bool> = sub.parameters.iter()
                    .map(|p| p.pass_type == ParameterPassType::ByRef)
                    .collect();
                let name = sub.name.as_str().to_lowercase();
                self.func_signatures.insert(name.clone(), sig);
                self.compile_sub(sub)?;
                self.emit_global_set(&name);
                self.emit(Op::drop);
            }
            Declaration::Function(func) => {
                let sig: Vec<bool> = func.parameters.iter()
                    .map(|p| p.pass_type == ParameterPassType::ByRef)
                    .collect();
                let name = func.name.as_str().to_lowercase();
                self.func_signatures.insert(name.clone(), sig);
                self.compile_function(func)?;
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
                    if let Some(ref bounds) = var.array_bounds {
                        // Dim arr(N) — VB arrays are 0..N inclusive, so size = N+1
                        if let Some(bound_expr) = bounds.first() {
                            self.compile_expression(bound_expr)?;
                            self.emit_constant(Value::F64(1.0));
                            self.emit(Op::dyn_add);
                            self.emit(Op::array_new_default);
                        } else {
                            self.emit_u16(Op::array_new, 0);
                        }
                    } else if let Some(ref init) = var.initializer {
                        self.compile_expression(init)?;
                    } else {
                        self.emit(Op::null);
                    }
                    let name = var.name.as_str().to_lowercase();
                    if var.array_bounds.is_some() {
                        self.known_arrays.insert(name.clone());
                    }
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
            Declaration::Enum(e) => {
                // Compile enum as an object with named constants
                self.emit_u16(Op::struct_new, 0);
                for (i, member) in e.members.iter().enumerate() {
                    self.emit(Op::dup);
                    let val = member.value.as_ref()
                        .map(|v| v.clone())
                        .unwrap_or(Expression::IntegerLiteral(i as i32));
                    self.compile_expression(&val)?;
                    let prop_idx = self.add_string_constant(&member.name.as_str().to_lowercase());
                    self.emit_u16(Op::struct_set, prop_idx);
                    self.emit(Op::drop);
                }
                let name = e.name.as_str().to_lowercase();
                self.emit_global_set(&name);
                self.emit(Op::drop);
            }
            Declaration::Structure(s) => {
                // Structure is like a lightweight class — compile same way
                // Reuse class compilation by treating fields as class fields
                let class = ClassDecl {
                    visibility: s.visibility.clone(),
                    name: s.name.clone(),
                    is_partial: false,
                    inherits: None,
                    implements: vec![],
                    properties: s.properties.clone(),
                    methods: s.methods.clone(),
                    fields: s.fields.clone(),
                    is_must_inherit: false,
                    is_not_inheritable: false,
                    nested_classes: vec![],
                    nested_enums: vec![],
                };
                self.compile_class(&class)?;
                let name = s.name.as_str().to_lowercase();
                self.defined_classes.insert(name.clone());
                self.emit_global_set(&name);
                self.emit(Op::drop);
            }
            Declaration::Imports(imp) => {
                // Register as interface prefix for Component Model resolution
                let path = imp.path.to_lowercase();
                if !self.interface_imports.contains(&path) {
                    self.interface_imports.push(path);
                }
            }
            Declaration::Namespace(ns) => {
                // Flatten namespace contents — compile all nested declarations
                for decl in &ns.declarations {
                    self.compile_declaration(decl)?;
                }
            }
            Declaration::Interface(iface) => {
                // Register interface as a type entry in the type table.
                let method_entries: Vec<(String, usize)> = iface.methods.iter().map(|m| {
                    let name = match m {
                        vybe_parser_basic::ast::InterfaceMember::Sub { name, .. } => name.as_str().to_lowercase(),
                        vybe_parser_basic::ast::InterfaceMember::Function { name, .. } => name.as_str().to_lowercase(),
                        vybe_parser_basic::ast::InterfaceMember::Property { name, .. } => name.as_str().to_lowercase(),
                        vybe_parser_basic::ast::InterfaceMember::Event { name, .. } => name.as_str().to_lowercase(),
                    };
                    (name, 0usize) // chunk index 0 = placeholder for interface methods
                }).collect();
                self.type_entries.push(TypeEntry {
                    name: iface.name.as_str().to_lowercase(),
                    parent: String::new(),
                    fields: Vec::new(),
                    methods: method_entries,
                    is_interface: true,
                    implements: Vec::new(),
                    constructor_chunk: None,
                });
            }
            Declaration::Delegate(_) | Declaration::Event(_) => {
                // Type declarations — no bytecode needed
            }
        }
        Ok(())
    }

    // ---- Sub / Function / Class compilation ----

    pub(crate) fn compile_store_ident(&mut self, target: &Identifier) -> Result<(), String> {
        let name = target.as_str().to_lowercase();
        // VB convention: assigning to the function name sets the return value
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
                if self.current_scope().is_byref(&name) {
                    // ByRef: write to box[0] — stack has [value]
                    // Need: [box, 0, value] for array_set
                    let tmp = self.define_local("__byref_tmp");
                    self.emit_u16(Op::local_set, tmp);
                    self.emit(Op::drop);
                    self.emit_u16(Op::local_get, slot); // box
                    self.emit(Op::i32_const_0);           // index 0
                    self.emit_u16(Op::local_get, tmp);    // value
                    self.emit(Op::array_set);
                    self.emit(Op::drop);
                } else {
                    self.emit_u16(Op::local_set, slot);
                    self.emit(Op::drop);
                }
            }
            VarResolution::Global => {
                // Inside a class: unresolved name that's a class field → Me.field = value
                if self.class_fields.contains(&name) {
                    if let Some(me_slot) = self.current_scope().resolve_local("me") {
                        self.emit_u16(Op::local_get, me_slot);
                        // Stack: [value, me] — but struct_set needs [obj, val]
                        // value is already on stack from caller, me is on top
                        // Need to swap: not directly available. Use a workaround:
                        // Actually, the caller already pushed value before calling us.
                        // Stack: [..., value]. We need to emit: [me, value] struct_set
                        // But value is already on top. Let's use a temp local.
                        let tmp = self.define_local("__field_tmp");
                        // Save value to temp
                        // Wait — the value is on the stack BEFORE this function.
                        // The pattern: compile_expression(value), then compile_store_ident(target)
                        // So stack has [..., value]. We need:
                        // local_set tmp (saves value), drop, local_get me, local_get tmp, struct_set
                        self.emit_u16(Op::local_set, tmp); // save me to tmp (wrong — me is on top)
                        // This is getting messy. Let me restructure.
                        // Actually we pushed me AFTER value. Stack: [..., value, me]
                        // We need struct_set which pops [obj, val]: expects obj below val
                        // Stack: [value, me] — me is on top, value below. struct_set sees obj=value, val=me. Wrong order.
                        // Need to swap. No swap opcode. Let me pop both and re-push.
                        self.emit(Op::drop); // drop me, stack: [value]
                        // Actually this whole approach is wrong. Let me handle it differently.
                        // Start over: value is on stack top. We need Me under it.
                        // Pop value to temp, push Me, push value back
                        // But we already defined tmp and pushed me...
                        // Let me just redo the whole thing cleanly:
                        return self.compile_store_field(&name);
                    }
                }
                self.emit_global_set(&name);
                self.emit(Op::drop);
            }
        }
        Ok(())
    }

    /// Store to Me.field_name. Value is on top of stack.
    fn compile_store_field(&mut self, field_name: &str) -> Result<(), String> {
        // Stack: [..., value]
        // Need: Me on stack, then value, then struct_set
        // Use a temp to reorder
        let tmp = self.define_local(&format!("__st_{}", field_name));
        self.emit_u16(Op::local_set, tmp); // save value
        self.emit(Op::drop);
        // Now push Me, then value
        if let Some(me_slot) = self.current_scope().resolve_local("me") {
            self.emit_u16(Op::local_get, me_slot);
        }
        self.emit_u16(Op::local_get, tmp); // push value back
        let prop_idx = self.add_string_constant(field_name);
        self.emit_u16(Op::struct_set, prop_idx);
        self.emit(Op::drop);
        Ok(())
    }

    fn compile_sub(&mut self, sub: &SubDecl) -> Result<(), String> {
        let name = sub.name.as_str();
        let chunk = common_fn::create_function_chunk(name, sub.parameters.len() as u8);
        let idx = self.chunks.len();
        self.chunks.push(chunk);
        let mut scope = Scope::new_function();
        for param in &sub.parameters {
            if param.pass_type == ParameterPassType::ByRef {
                scope.define_byref_local(&param.name.as_str().to_lowercase());
            } else {
                scope.define_local(&param.name.as_str().to_lowercase());
            }
        }
        let saved = self.current_chunk_idx;
        self.current_chunk_idx = idx;
        self.scopes.push(scope);
        for stmt in &sub.body { self.compile_statement(stmt)?; }
        common_fn::emit_function_epilogue(&mut self.chunks[idx], self.line);
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
        let chunk = common_fn::create_function_chunk(fname, params.len() as u8);
        let idx = self.chunks.len();
        self.chunks.push(chunk);
        let mut scope = Scope::new_function();
        for param in params {
            if param.pass_type == ParameterPassType::ByRef {
                scope.define_byref_local(&param.name.as_str().to_lowercase());
            } else {
                scope.define_local(&param.name.as_str().to_lowercase());
            }
        }
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
