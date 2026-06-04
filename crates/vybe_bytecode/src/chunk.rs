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
    /// Number of results this function returns. Default 1 (single
    /// externref) matches the pre-multi-value ABI. A chunk that wants
    /// to take advantage of the multi-value proposal sets this >1; the
    /// WASM emitter then declares the function type with that many
    /// externref results.
    pub result_arity: u8,
    /// JSPI: this function is `async` in its source language. The
    /// compiler sets this flag when compiling an `async function` /
    /// `async def` / `Async Function`. The WASM emitter writes a
    /// `vybe.jspi` custom section listing every async chunk so a JS
    /// host can wrap the corresponding export with
    /// `WebAssembly.promising(fn)` and therefore have it return a real
    /// JS Promise that resolves when the Vybe fiber completes. Non-Vybe
    /// WASM engines ignore unknown custom sections, so setting this on
    /// a sync function is at worst no-op on those engines. On Vybe VM
    /// it's informational only — the runtime already knows via
    /// `PROMISE_SUSPEND` opcodes in the body.
    pub is_async: bool,
    /// When true, `call_value` on a function bound to this chunk
    /// returns a fresh generator continuation instead of executing the
    /// body. Compilers set this for `def` with `yield`, `function*`,
    /// C# `yield return`, Ruby `Enumerator::new`, Dart `sync*`. Lowers
    /// to WASM stack-switching `cont.new` at emit time.
    pub is_generator: bool,
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
            result_arity: 1,
            is_async: false,
            is_generator: false,
        }
    }

    /// Add an import and return its index (used by CallHost operand).
    pub fn add_import(&mut self, module: impl Into<String>, name: impl Into<String>) -> u16 {
        let import = Import {
            module: module.into(),
            name: name.into(),
        };
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
            if *existing == name {
                return i as u8;
            }
        }
        self.exception_tags.push(name);
        (self.exception_tags.len() - 1) as u8
    }

    /// Add a global initializer (evaluated at load time).
    pub fn add_global_init(&mut self, name: impl Into<String>, init: ConstExpr) {
        self.global_inits.push(GlobalInit {
            name: name.into(),
            init,
        });
    }

    /// Add a continuation tag and return its index.
    pub fn add_continuation_tag(
        &mut self,
        name: impl Into<String>,
        yield_type: impl Into<String>,
        resume_type: impl Into<String>,
    ) -> u16 {
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
        let (prefix, sub) = op.encode();
        self.emit(prefix, line);
        self.emit(sub, line);
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

    pub fn emit_leb_u32(&mut self, mut value: u32, line: u32) {
        loop {
            let mut byte = (value & 0x7f) as u8;
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            self.emit(byte, line);
            if value == 0 {
                break;
            }
        }
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
        self.emit_op(Op::BR, line);
        let jump = target_offset as i32 - (self.code.len() as i32 + 2);
        self.emit((jump >> 8) as u8, line);
        self.emit((jump & 0xff) as u8, line);
    }

    // ── Structured control flow (WASM-compliant) ───────────────────────
    //
    // BLOCK, LOOP, IF all carry a single blocktype byte (WASM §5.4.1):
    //   0x40 = void (no result), 0x6F = externref, 0x7F = i32, etc.
    //
    // The VM pre-scans each function body to build a block table mapping
    // every BLOCK/LOOP/IF/ELSE opcode position to its ELSE/END target ip.
    // No size headers or patching needed — nesting is the structure.

    /// Emit a void BLOCK (no value produced). The VM uses the pre-scanned
    /// block table to find the matching END; `patch_block` is a no-op.
    pub fn emit_block(&mut self, line: u32) -> usize {
        self.emit_block_typed(line, 0)
    }

    /// Emit BLOCK with explicit result count.
    /// 0 = void, 1 = single externref, 2+ = multi-value (WASM encoder registers type).
    pub fn emit_block_typed(&mut self, line: u32, result_count: u8) -> usize {
        self.emit_op(Op::BLOCK, line);
        self.emit(result_count, line); // raw count; WASM encoder translates to blocktype
        self.code.len() - 1 // dummy patch pos — patch_block is a no-op
    }

    /// Emit a void LOOP. Returns (dummy_patch, loop_body_start).
    /// `loop_body_start` is the ip right after the result_count byte —
    /// `br 0` inside the loop restarts there.
    pub fn emit_loop_s(&mut self, line: u32) -> (usize, usize) {
        self.emit_loop_typed(line, 0)
    }

    /// Emit LOOP with explicit result count.
    pub fn emit_loop_typed(&mut self, line: u32, result_count: u8) -> (usize, usize) {
        self.emit_op(Op::LOOP, line);
        self.emit(result_count, line);
        let dummy_patch = self.code.len() - 1;
        let loop_body_start = self.code.len();
        (dummy_patch, loop_body_start)
    }

    /// Close a BLOCK, LOOP, IF, or IF/ELSE.
    pub fn emit_end(&mut self, line: u32) {
        self.emit_op(Op::END, line);
    }

    /// No-op — block table replaces size-header patching.
    #[inline(always)]
    pub fn patch_block(&mut self, _patch_pos: usize) {}

    /// No-op — block table replaces size-header patching.
    #[inline(always)]
    pub fn patch_loop(&mut self, _patch_pos: usize) {}

    /// Emit IF that produces no value (void).
    /// Caller must have an i32 on the stack (non-zero = enter then-body).
    /// Use `emitter::ops::emit_dyn_to_bool` first to coerce a Value → i32.
    pub fn emit_if(&mut self, line: u32) -> usize {
        self.emit_op(Op::IF, line);
        self.emit(0u8, line); // result_count=0 → void
        self.code.len() - 1
    }

    /// Emit IF that leaves one externref on the stack (then/else both push one Value).
    pub fn emit_if_value(&mut self, line: u32) -> usize {
        self.emit_op(Op::IF, line);
        self.emit(1u8, line); // result_count=1 → externref
        self.code.len() - 1
    }

    /// Emit ELSE. Must be matched with emit_if + emit_end.
    pub fn emit_else(&mut self, line: u32) {
        self.emit_op(Op::ELSE, line);
    }

    /// Emit WASM `br` with a structured label depth.
    pub fn emit_br(&mut self, depth: u32, line: u32) {
        self.emit_op(Op::BR, line);
        self.emit_leb_u32(depth, line);
    }

    /// Emit WASM `br_if` with a structured label depth. Expects i32 on stack.
    pub fn emit_br_if(&mut self, depth: u32, line: u32) {
        self.emit_op(Op::BR_IF, line);
        self.emit_leb_u32(depth, line);
    }

    /// Emit WASM `br_table` with structured label depths.
    pub fn emit_br_table(&mut self, depths: &[u32], default_depth: u32, line: u32) {
        self.emit_op(Op::BR_TABLE, line);
        self.emit_leb_u32(depths.len() as u32, line);
        for &depth in depths {
            self.emit_leb_u32(depth, line);
        }
        self.emit_leb_u32(default_depth, line);
    }

    pub fn read_u16(&self, offset: usize) -> u16 {
        ((self.code[offset] as u16) << 8) | (self.code[offset + 1] as u16)
    }

    pub fn read_i16(&self, offset: usize) -> i16 {
        self.read_u16(offset) as i16
    }
}
