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
