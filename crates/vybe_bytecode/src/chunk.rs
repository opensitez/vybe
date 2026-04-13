use crate::opcode::Op;
use crate::value::Value;

/// A host function import declaration — (module, name).
/// Like WASM: (import "vybe:math" "floor" (func ...))
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Import {
    pub module: String,
    pub name: String,
}

/// A compile-time type definition — WASM GC type section entry.
/// Describes a class/struct with named fields and vtable methods.
/// Loaded into TypeRegistry before execution.
#[derive(Debug, Clone)]
pub struct TypeEntry {
    pub name: String,
    /// Parent type name (for inheritance). Empty = inherits from Object.
    pub parent: String,
    /// Field names in order. Field i is at `fields[i]` in the object's indexed storage.
    pub fields: Vec<String>,
    /// Vtable: method_name → chunk_index. Methods are shared across all instances.
    pub methods: Vec<(String, usize)>,
    /// Whether this is an interface definition (not a concrete class).
    pub is_interface: bool,
    /// Interface names this type implements.
    pub implements: Vec<String>,
    /// Constructor chunk index (if any). Resolved during load_type_table.
    pub constructor_chunk: Option<usize>,
}

/// A constant initialization expression (Extended Const Expressions proposal).
/// Evaluated at module instantiation time, before code execution.
#[derive(Debug, Clone)]
pub enum ConstExpr {
    /// A literal value.
    Value(Value),
    /// Reference another global: global_name.
    GlobalGet(String),
    /// Add two const exprs: left + right.
    Add(Box<ConstExpr>, Box<ConstExpr>),
    /// Multiply two const exprs: left * right.
    Mul(Box<ConstExpr>, Box<ConstExpr>),
    /// Create a function reference from a chunk index.
    /// Evaluated at load time — produces a callable Function object.
    RefFunc(usize),
}

/// A global variable initializer — evaluated at link/load time.
#[derive(Debug, Clone)]
pub struct GlobalInit {
    /// Global name (stored in VM.globals).
    pub name: String,
    /// Initialization expression (evaluated before code runs).
    pub init: ConstExpr,
}

/// A continuation tag — defines the type contract for typed continuations.
#[derive(Debug, Clone)]
pub struct ContinuationTag {
    /// Tag name (for debugging and cross-language matching).
    pub name: String,
    /// Expected yield value type name (empty = any).
    pub yield_type: String,
    /// Expected resume value type name (empty = any).
    pub resume_type: String,
}

/// A compiled chunk of bytecode — one per function/script.
#[derive(Debug, Clone)]
pub struct Chunk {
    pub code: Vec<u8>,
    pub constants: Vec<Value>,
    pub lines: Vec<u32>,
    pub name: String,
    pub arity: u8,
    pub local_count: u16,
    /// Import table — only on the script chunk (chunk 0).
    /// Each entry is a (module, name) pair.
    /// CallHost operand indexes into this table.
    pub imports: Vec<Import>,
    /// Type table — WASM GC type section. Only on the script chunk (chunk 0).
    /// Each entry defines a class type with fields and vtable methods.
    /// Loaded into VM's TypeRegistry before execution.
    pub types: Vec<TypeEntry>,
    /// Exception tag table — maps tag index to type name for typed exception handling.
    /// Tag 0 = catch-all (matches any exception).
    /// Tag N = matches exceptions whose type is `exception_tags[N]` or a subtype.
    pub exception_tags: Vec<String>,
    /// Type imports — types from other components this chunk needs.
    /// Each entry is (interface_name, type_name).
    pub type_imports: Vec<(String, String)>,
    /// Type exports — types this component makes available to others.
    /// Each entry is (interface_name, type_name, type_id).
    pub type_exports: Vec<(String, String, usize)>,
    /// Global initializers — const expressions evaluated at load time.
    pub global_inits: Vec<GlobalInit>,
    /// Continuation tags — typed contracts for suspend/resume.
    pub continuation_tags: Vec<ContinuationTag>,
}

impl Chunk {
    pub fn new(name: impl Into<String>) -> Self {
        Chunk {
            code: Vec::new(),
            constants: Vec::new(),
            lines: Vec::new(),
            name: name.into(),
            arity: 0,
            local_count: 0,
            imports: Vec::new(),
            types: Vec::new(),
            exception_tags: Vec::new(),
            type_imports: Vec::new(),
            type_exports: Vec::new(),
            global_inits: Vec::new(),
            continuation_tags: Vec::new(),
        }
    }

    /// Add an import and return its index (used by CallHost operand).
    pub fn add_import(&mut self, module: impl Into<String>, name: impl Into<String>) -> u16 {
        let import = Import { module: module.into(), name: name.into() };
        // Deduplicate — return existing index if already imported
        for (i, existing) in self.imports.iter().enumerate() {
            if *existing == import {
                return i as u16;
            }
        }
        self.imports.push(import);
        (self.imports.len() - 1) as u16
    }

    /// Add an exception tag and return its index.
    /// Tag 0 = catch-all. Tag N = typed catch for exceptions matching `type_name`.
    pub fn add_exception_tag(&mut self, type_name: impl Into<String>) -> u8 {
        let name = type_name.into();
        // Deduplicate
        for (i, existing) in self.exception_tags.iter().enumerate() {
            if *existing == name { return i as u8; }
        }
        self.exception_tags.push(name);
        (self.exception_tags.len() - 1) as u8
    }

    /// Add a global initializer (evaluated at load time).
    pub fn add_global_init(&mut self, name: impl Into<String>, init: ConstExpr) {
        self.global_inits.push(GlobalInit { name: name.into(), init });
    }

    /// Add a continuation tag and return its index.
    pub fn add_continuation_tag(&mut self, name: impl Into<String>, yield_type: impl Into<String>, resume_type: impl Into<String>) -> u16 {
        let tag = ContinuationTag {
            name: name.into(),
            yield_type: yield_type.into(),
            resume_type: resume_type.into(),
        };
        self.continuation_tags.push(tag);
        (self.continuation_tags.len() - 1) as u16
    }

    pub fn emit(&mut self, byte: u8, line: u32) {
        self.code.push(byte);
        self.lines.push(line);
    }

    pub fn emit_op(&mut self, op: Op, line: u32) {
        let (b1, b2) = op.encode();
        self.emit(b1, line);
        if let Some(b) = b2 { self.emit(b, line); }
    }

    pub fn emit_op_u16(&mut self, op: Op, operand: u16, line: u32) {
        self.emit_op(op, line);
        self.emit((operand >> 8) as u8, line);
        self.emit((operand & 0xff) as u8, line);
    }

    pub fn emit_op_u8(&mut self, op: Op, operand: u8, line: u32) {
        self.emit_op(op, line);
        self.emit(operand, line);
    }

    pub fn add_constant(&mut self, value: Value) -> u16 {
        self.constants.push(value);
        (self.constants.len() - 1) as u16
    }

    pub fn emit_jump(&mut self, op: Op, line: u32) -> usize {
        self.emit_op(op, line);
        self.emit(0xff, line);
        self.emit(0xff, line);
        self.code.len() - 2
    }

    pub fn patch_jump(&mut self, offset: usize) {
        let jump = self.code.len() as i32 - (offset as i32 + 2);
        self.code[offset] = (jump >> 8) as u8;
        self.code[offset + 1] = (jump & 0xff) as u8;
    }

    pub fn current_offset(&self) -> usize {
        self.code.len()
    }

    /// Get the source line number for a bytecode offset.
    /// Returns None if no line info is available for that offset.
    pub fn get_line(&self, offset: usize) -> Option<u32> {
        self.lines.get(offset).copied().filter(|&l| l > 0)
    }

    pub fn emit_loop(&mut self, target_offset: usize, line: u32) {
        self.emit_op(Op::br, line);
        let jump = target_offset as i32 - (self.code.len() as i32 + 2);
        self.emit((jump >> 8) as u8, line);
        self.emit((jump & 0xff) as u8, line);
    }

    pub fn read_u16(&self, offset: usize) -> u16 {
        ((self.code[offset] as u16) << 8) | (self.code[offset + 1] as u16)
    }

    pub fn read_i16(&self, offset: usize) -> i16 {
        self.read_u16(offset) as i16
    }
}
